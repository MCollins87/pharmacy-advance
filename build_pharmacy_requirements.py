#!/usr/bin/env python3
"""Build a pharmacy advance-preparation workbook from ARIA reports.

Operational source hierarchy:
1. sch_by_inst_pt_excel_vprov.xls supplies attendance date/time, NHS number,
   scheduled event, and visit provider.
2. pharm_reqmt.xls supplies drug, dose, prescription physician, administration
   date, and the dispensed/verified indicators.
3. ptmeds_sch_time.pdf is used only as a fallback for scheduled patients who
   have no included preparation item in pharm_reqmt.xls.

The script requires no command-line arguments. Edit BASE_DIR below if needed.
Inputs are archived only after a non-empty workbook has been saved.

This is a planning/reconciliation aid. It does not calculate doses and does
not replace verification against the current authorised ARIA prescription.
"""
from __future__ import annotations

import csv
import logging
import re
import shutil
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Operational configuration
# ---------------------------------------------------------------------------

BASE_DIR = Path(r"C:\IDR\PharmacyAdvance")

INPUT_DIR = BASE_DIR / "Input"
OUTPUT_DIR = BASE_DIR / "Output"
ARCHIVE_DIR = BASE_DIR / "Archive"
CONFIG_DIR = BASE_DIR / "Config"
LOG_DIR = BASE_DIR / "Logs"

SCHEDULE_FILENAME = "sch_by_inst_pt_excel_vprov.xls"
PHARMACY_FILENAME = "pharm_reqmt.xls"
PDF_FILENAME = "ptmeds_sch_time.pdf"
AGENT_RULES_FILENAME = "agent_preparation_rules.csv"
EVENT_RULES_FILENAME = "event_drug_rules.csv"

PHARMACY_REPORT_TITLE = (
    "Pharmacy Requirements - by Agent, Rx Type, Administration Date and Patient"
)
PDF_REPORT_TITLE = "Patient Medications - Patients Scheduled to Visit - by Time"

PDF_PATIENT_RE = re.compile(
    r"Start Time:\s*(?P<date>[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4})\s+"
    r"(?P<time>\d{1,2}:\d{2})\s+Event Provider:.*?\n"
    r"(?P<patient>.+?)\s+NHS Number:\s*(?P<nhs>\d{10})\b",
    re.S,
)
PDF_SECTION_RE = re.compile(
    r"^(Chemo|Hormone|Immunotherapy|Supportive|Pharmacy:)\s*$", re.I
)
PDF_FOOTER_RE = re.compile(r"^Report Name:", re.I)
PDF_ADMIN_RE = re.compile(r"^\s*Administration Instructions:", re.I)
PDF_DATE_AT_END_RE = re.compile(
    r"\s+[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4}"
    r"(?:\s+[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4})?\s*$"
)

OUTPUT_HEADERS = [
    "Administration date",
    "Appointment time",
    "Patient ID",
    "Patient",
    "Prescription physician",
    "Visit provider",
    "Scheduled event(s)",
    "Drug",
    "Dose",
    "Route",
    "Dispensed",
    "Verified",
    "Source",
    "Source reference",
    "Status / reason",
]


@dataclass(frozen=True)
class ScheduleRow:
    institution: str
    nhs_number: str
    patient: str
    event_dt: datetime
    start_dt: datetime
    end_dt: datetime
    event: str
    visit_provider: str
    comments: str
    source_row: int


@dataclass(frozen=True)
class PharmacyRequirement:
    administration_date: date
    agent: str
    dose: str
    patient: str
    physician: str
    dispensed: str
    verified: str
    source_row: int


@dataclass(frozen=True)
class PdfMedication:
    visit_date: date
    visit_time: datetime
    nhs_number: str
    patient: str
    section: str
    agent: str
    course_description: str
    dose: str
    route: str
    page: int


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------

def clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalise(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean(value).casefold())


def normalise_nhs_number(value) -> str:
    text = clean(value)
    if text.casefold() == "reserved":
        return "Reserved"
    digits = re.sub(r"\D", "", text)
    return digits or text


def patient_match_key(value) -> str:
    """Use normalised surname and first forename for cross-report matching."""
    text = clean(value).upper()
    if not text:
        return ""
    if "," in text:
        surname, forenames = text.split(",", 1)
    else:
        parts = text.split()
        surname = parts[0]
        forenames = " ".join(parts[1:])
    first_names = re.findall(r"[A-Z0-9]+", forenames)
    first_name = first_names[0] if first_names else ""
    return f"{normalise(surname)}|{normalise(first_name)}"


def parse_excel_datetime(value, datemode: int) -> datetime | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return xlrd.xldate_as_datetime(value, datemode)
    text = clean(value)
    for date_format in ("%b %d, %Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, date_format)
        except ValueError:
            pass
    return None


def parse_text_date(value) -> date | None:
    text = clean(value)
    for date_format in ("%b %d, %Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            pass
    return None


def format_dose(amount, unit) -> str:
    if amount in (None, ""):
        return ""
    if isinstance(amount, float) and amount.is_integer():
        amount_text = str(int(amount))
    else:
        amount_text = clean(amount)
    return clean(f"{amount_text} {unit}")


def extract_pdf_dose(full_text: str) -> str:
    text = re.sub(r"^Prescribed:\s*", "", clean(full_text), flags=re.I)
    match = re.match(
        r"(?P<dose>\d[\d,]*(?:\.\d+)?"
        r"(?:\s*-\s*\d[\d,]*(?:\.\d+)?)?\s*"
        r"(?:mg|mcg|g|mmol|mL|IU|Units|dose\(s\)|tablet|capsule|injection)"
        r"(?:\s*\([^)]*\))?)\b",
        text,
        re.I,
    )
    return clean(match.group("dose")) if match else ""


def extract_route(full_text: str) -> str:
    for route in (
        "Subcutaneous",
        "Intravenous",
        "Intramuscular",
        "Oral",
        "Oromucosal",
        "Topical",
        "Ocular",
        "Inhalation",
        "Neb",
        "Unknown",
        "Not Assigned",
    ):
        if re.search(rf"\b{re.escape(route)}\b", full_text, re.I):
            return route
    return ""


# ---------------------------------------------------------------------------
# Logging, validation, and archiving
# ---------------------------------------------------------------------------

def configure_logging() -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_DIR / f"pharmacy_advance_{timestamp}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )
    return log_path


def ensure_folders() -> None:
    for folder in (INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR, CONFIG_DIR, LOG_DIR):
        folder.mkdir(parents=True, exist_ok=True)


def validate_required_files() -> dict[str, Path]:
    files = {
        "schedule": INPUT_DIR / SCHEDULE_FILENAME,
        "pharmacy": INPUT_DIR / PHARMACY_FILENAME,
        "pdf": INPUT_DIR / PDF_FILENAME,
        "agent_rules": CONFIG_DIR / AGENT_RULES_FILENAME,
        "event_rules": CONFIG_DIR / EVENT_RULES_FILENAME,
    }
    missing = [path for path in files.values() if not path.exists()]
    if missing:
        formatted = "\n".join(f"  - {path}" for path in missing)
        raise FileNotFoundError(f"Required files are missing:\n{formatted}")
    return files


def get_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{timestamp}{path.suffix}")


def archive_input_files(input_paths: list[Path], administration_date: date) -> Path:
    run_date = datetime.now().date()
    archive_folder = (
        ARCHIVE_DIR
        / run_date.isoformat()
        / f"Administration_{administration_date.isoformat()}"
    )
    archive_folder.mkdir(parents=True, exist_ok=True)
    for source_path in input_paths:
        if not source_path.exists():
            raise FileNotFoundError(f"Cannot archive missing file: {source_path}")
        destination = get_available_path(archive_folder / source_path.name)
        logging.info("Archiving %s to %s", source_path, destination)
        shutil.move(str(source_path), str(destination))
    return archive_folder


# ---------------------------------------------------------------------------
# Rules
# ---------------------------------------------------------------------------

def load_csv_rules(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as source:
        return [
            {key: clean(value) for key, value in row.items()}
            for row in csv.DictReader(source)
        ]


def get_agent_decision(agent: str, rules: list[dict[str, str]]) -> tuple[str, str]:
    for rule in rules:
        pattern = rule.get("agent_pattern", "")
        if pattern and re.search(pattern, agent, re.I):
            return (
                rule.get("decision", "REVIEW").upper(),
                rule.get("notes", ""),
            )
    return "REVIEW", "Agent is absent from the preparation rules"


def get_pdf_event_decision(
    events: list[str], medication: PdfMedication, rules: list[dict[str, str]]
) -> tuple[str, str]:
    matches = []
    for rule in rules:
        event_pattern = rule.get("event_pattern", "")
        agent_pattern = rule.get("agent_pattern", "")
        if (
            event_pattern
            and agent_pattern
            and any(re.search(event_pattern, event, re.I) for event in events)
            and re.search(agent_pattern, medication.agent, re.I)
        ):
            matches.append(rule)
    if matches:
        decisions = {
            rule.get("decision", "REVIEW").upper() for rule in matches
        }
        if len(decisions) == 1:
            return decisions.pop(), matches[0].get(
                "notes", "Matched event-to-drug rule"
            )
    return "REVIEW", "PDF fallback requires event-to-drug review"


# ---------------------------------------------------------------------------
# Schedule XLS parser
# ---------------------------------------------------------------------------

def parse_schedule(path: Path) -> list[ScheduleRow]:
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    output = []
    for row_index in range(sheet.nrows):
        row = sheet.row_values(row_index) + [""] * 13
        event_datetime = parse_excel_datetime(row[5], workbook.datemode)
        start_datetime = parse_excel_datetime(row[7], workbook.datemode)
        if not (
            clean(row[0])
            and normalise_nhs_number(row[1])
            and clean(row[4])
            and event_datetime
            and start_datetime
            and clean(row[9])
        ):
            continue
        output.append(
            ScheduleRow(
                institution=clean(row[0]),
                nhs_number=normalise_nhs_number(row[1]),
                patient=clean(row[4]),
                event_dt=event_datetime,
                start_dt=start_datetime,
                end_dt=parse_excel_datetime(row[8], workbook.datemode)
                or start_datetime,
                event=clean(row[9]).lstrip("*"),
                visit_provider=clean(row[11]),
                comments=clean(row[12]),
                source_row=row_index + 1,
            )
        )
    if not output:
        raise ValueError(f"No schedule records were found in {path.name}")
    return output


# ---------------------------------------------------------------------------
# Pharmacy Requirements XLS parser
# ---------------------------------------------------------------------------

def parse_pharmacy_requirements(path: Path) -> list[PharmacyRequirement]:
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    sample = " ".join(
        clean(sheet.cell_value(row_index, column_index))
        for row_index in range(min(10, sheet.nrows))
        for column_index in range(sheet.ncols)
    )
    if PHARMACY_REPORT_TITLE.casefold() not in sample.casefold():
        raise ValueError(
            f"The supplied XLS is not the expected Pharmacy Requirements report: {path.name}"
        )
    output = []
    for row_index in range(sheet.nrows):
        row = sheet.row_values(row_index) + [""] * 11
        agent = clean(row[0])
        amount = row[1]
        unit = clean(row[2])
        patient = clean(row[4])
        physician = clean(row[6])
        administration_date = parse_text_date(row[7])
        dispensed = clean(row[9])
        verified = clean(row[10])
        if not (agent and patient and administration_date):
            continue
        if agent.casefold().startswith("total "):
            continue
        if "oncology day unit" in agent.casefold():
            continue
        output.append(
            PharmacyRequirement(
                administration_date=administration_date,
                agent=agent,
                dose=format_dose(amount, unit),
                patient=patient,
                physician=physician,
                dispensed=dispensed,
                verified=verified,
                source_row=row_index + 1,
            )
        )
    if not output:
        raise ValueError(f"No medication records were found in {path.name}")

    # Collapse exact duplicate report lines while preserving different doses.
    seen = set()
    distinct = []
    for medication in output:
        duplicate_key = (
            medication.administration_date,
            patient_match_key(medication.patient),
            normalise(medication.agent),
            normalise(medication.dose),
            normalise(medication.physician),
        )
        if duplicate_key in seen:
            continue
        seen.add(duplicate_key)
        distinct.append(medication)
    return distinct


# ---------------------------------------------------------------------------
# Patient Medications PDF parser
# ---------------------------------------------------------------------------

def likely_pdf_agent_start(line: str) -> tuple[str, str] | None:
    if not line.strip() or PDF_ADMIN_RE.match(line) or PDF_FOOTER_RE.match(line):
        return None
    match = re.match(
        r"^\s*(?P<agent>[A-Za-z0-9][A-Za-z0-9 /&().,'+\-]{1,70}?)"
        r"\s{2,}(?P<course>.+?)\s*$",
        line,
    )
    if not match:
        return None
    agent = clean(match.group("agent"))
    course = PDF_DATE_AT_END_RE.sub("", clean(match.group("course")))
    if agent.casefold() in {"active", "agent", "nhs number", "allergies"}:
        return None
    if not re.search(
        r"\b(mg|mcg|g|mmol|mL|IU|Units|dose\(s\)|tablet|capsule|"
        r"injection|infusion|prescribed:)\b",
        course,
        re.I,
    ):
        return None
    return agent, course


def parse_pdf(path: Path, included_sections: set[str]) -> list[PdfMedication]:
    reader = PdfReader(str(path))
    if not reader.pages:
        raise ValueError("The PDF contains no pages")
    first_page_text = (
        reader.pages[0].extract_text(extraction_mode="layout") or ""
    )
    if PDF_REPORT_TITLE.casefold() not in first_page_text.casefold():
        raise ValueError(
            f"The PDF is not the expected Patient Medications report: {path.name}"
        )

    output = []
    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text(extraction_mode="layout") or ""
        patient_match = PDF_PATIENT_RE.search(text)
        if not patient_match:
            continue
        visit_date = datetime.strptime(
            patient_match.group("date"), "%b %d, %Y"
        ).date()
        visit_time_value = datetime.strptime(
            patient_match.group("time"), "%H:%M"
        ).time()
        patient = clean(patient_match.group("patient"))
        patient_nhs = normalise_nhs_number(patient_match.group("nhs"))
        current_section = ""
        current_medication = None

        def emit_medication() -> None:
            nonlocal current_medication
            if (
                current_medication
                and normalise(current_medication["section"]) in included_sections
            ):
                description = clean(
                    " ".join(current_medication["course_parts"])
                )
                output.append(
                    PdfMedication(
                        visit_date=visit_date,
                        visit_time=datetime.combine(visit_date, visit_time_value),
                        nhs_number=patient_nhs,
                        patient=patient,
                        section=current_medication["section"],
                        agent=current_medication["agent"],
                        course_description=description,
                        dose=extract_pdf_dose(description),
                        route=extract_route(description),
                        page=page_number,
                    )
                )
            current_medication = None

        for raw_line in text.splitlines():
            line = raw_line.rstrip()
            stripped = clean(line)
            section_match = PDF_SECTION_RE.match(stripped)
            if section_match:
                emit_medication()
                current_section = section_match.group(1)
                continue
            if PDF_FOOTER_RE.match(stripped):
                emit_medication()
                break
            if PDF_ADMIN_RE.match(stripped):
                emit_medication()
                continue
            medication_start = likely_pdf_agent_start(line)
            if medication_start:
                emit_medication()
                current_medication = {
                    "section": current_section,
                    "agent": medication_start[0],
                    "course_parts": [medication_start[1]],
                }
            elif current_medication and stripped:
                continuation = PDF_DATE_AT_END_RE.sub("", stripped)
                current_medication["course_parts"].append(continuation)
        emit_medication()

    seen = set()
    distinct = []
    for medication in output:
        duplicate_key = (
            medication.visit_date,
            medication.nhs_number,
            normalise(medication.section),
            normalise(medication.agent),
            normalise(medication.course_description),
        )
        if duplicate_key in seen:
            continue
        seen.add(duplicate_key)
        distinct.append(medication)
    return distinct


# ---------------------------------------------------------------------------
# Matching and workbook formatting
# ---------------------------------------------------------------------------

def get_schedule_context(appointments: list[ScheduleRow]) -> tuple:
    ordered = sorted(appointments, key=lambda item: item.start_dt)
    appointment_time = min(
        (item.start_dt.time() for item in ordered), default=None
    )
    events = " | ".join(
        dict.fromkeys(item.event for item in ordered if item.event)
    )
    visit_providers = " | ".join(
        dict.fromkeys(
            item.visit_provider for item in ordered if item.visit_provider
        )
    )
    comments = " | ".join(
        dict.fromkeys(item.comments for item in ordered if item.comments)
    )
    return appointment_time, events, visit_providers, comments


def add_table(worksheet, table_name: str) -> None:
    if worksheet.max_row < 2:
        return
    reference = f"A1:{worksheet.cell(worksheet.max_row, worksheet.max_column).coordinate}"
    table = Table(displayName=table_name, ref=reference)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True
    )
    worksheet.add_table(table)


def style_worksheet(worksheet) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    worksheet.row_dimensions[1].height = 32
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    widths = {
        "A": 18,
        "B": 14,
        "C": 15,
        "D": 31,
        "E": 35,
        "F": 35,
        "G": 42,
        "H": 38,
        "I": 22,
        "J": 18,
        "K": 12,
        "L": 12,
        "M": 24,
        "N": 18,
        "O": 54,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width


def write_output_sheet(
    workbook: Workbook, title: str, rows: list[list], table_name: str
):
    worksheet = workbook.create_sheet(title)
    worksheet.append(OUTPUT_HEADERS)
    for row in rows:
        worksheet.append(row)
    style_worksheet(worksheet)
    for cell in worksheet["A"][1:]:
        cell.number_format = "dd/mm/yyyy"
    for cell in worksheet["B"][1:]:
        cell.number_format = "hh:mm"
    add_table(worksheet, table_name)
    return worksheet


# ---------------------------------------------------------------------------
# Main build
# ---------------------------------------------------------------------------

def build_workbook(
    schedule_path: Path,
    pharmacy_path: Path,
    pdf_path: Path,
    agent_rules_path: Path,
    event_rules_path: Path,
    output_path: Path,
) -> dict:
    schedule = parse_schedule(schedule_path)
    pharmacy_requirements = parse_pharmacy_requirements(pharmacy_path)
    pdf_medications = parse_pdf(pdf_path, included_sections={"chemo"})
    agent_rules = load_csv_rules(agent_rules_path)
    event_rules = load_csv_rules(event_rules_path)

    schedule_dates = {row.event_dt.date() for row in schedule}
    pharmacy_dates = {
        row.administration_date for row in pharmacy_requirements
    }
    if len(schedule_dates) != 1:
        raise ValueError("The timed schedule must cover exactly one treatment date")
    if len(pharmacy_dates) != 1:
        raise ValueError(
            "The Pharmacy Requirements report must cover exactly one administration date"
        )
    schedule_date = next(iter(schedule_dates))
    pharmacy_date = next(iter(pharmacy_dates))
    if schedule_date != pharmacy_date:
        raise ValueError(
            f"The report dates do not match. Schedule: {schedule_date}; "
            f"Pharmacy Requirements: {pharmacy_date}"
        )
    administration_date = schedule_date

    schedule_by_name = defaultdict(list)
    schedule_by_nhs = defaultdict(list)
    for row in schedule:
        schedule_by_name[
            (row.event_dt.date(), patient_match_key(row.patient))
        ].append(row)
        schedule_by_nhs[(row.event_dt.date(), row.nhs_number)].append(row)

    pharmacy_order = []
    pharmacy_review = []
    not_scheduled = []
    included_schedule_keys = set()

    for medication in pharmacy_requirements:
        decision, rule_note = get_agent_decision(
            medication.agent, agent_rules
        )
        appointments = schedule_by_name.get(
            (
                medication.administration_date,
                patient_match_key(medication.patient),
            ),
            [],
        )
        appointment_time, events, visit_providers, comments = (
            get_schedule_context(appointments)
        )
        patient_id = appointments[0].nhs_number if appointments else ""
        status_reasons = []
        if not appointments:
            status_reasons.append(
                "No matching timed schedule record by administration date, surname and first forename"
            )
        if decision == "REVIEW":
            status_reasons.append(rule_note)
        if not medication.dose:
            status_reasons.append("Dose is missing")
        if medication.dispensed:
            status_reasons.append(
                f"Dispensed indicator: {medication.dispensed}"
            )
        if medication.verified:
            status_reasons.append(
                f"Verified indicator: {medication.verified}"
            )
        status = "; ".join(status_reasons) if status_reasons else "Ready"
        row = [
            medication.administration_date,
            appointment_time,
            patient_id,
            medication.patient,
            medication.physician,
            visit_providers,
            events,
            medication.agent,
            medication.dose,
            "",
            medication.dispensed,
            medication.verified,
            "Pharmacy Requirements XLS",
            f"Row {medication.source_row}",
            status,
        ]
        if appointments and decision == "INCLUDE" and medication.dose:
            pharmacy_order.append(row)
            for appointment in appointments:
                included_schedule_keys.add(
                    (appointment.event_dt.date(), appointment.nhs_number)
                )
        elif not appointments and decision == "INCLUDE":
            not_scheduled.append(row)
        elif appointments and decision == "REVIEW":
            pharmacy_review.append(row)
        # EXCLUDE items are intentionally omitted.

    pdf_by_key = defaultdict(list)
    for medication in pdf_medications:
        pdf_by_key[(medication.visit_date, medication.nhs_number)].append(
            medication
        )

    Not_Approved = []
    no_candidate = []
    for schedule_key, appointments in schedule_by_nhs.items():
        if schedule_key in included_schedule_keys:
            continue
        if schedule_key[1].casefold() == "reserved":
            continue
        appointment_time, events_text, visit_providers, comments = (
            get_schedule_context(appointments)
        )
        events = [
            appointment.event for appointment in appointments if appointment.event
        ]
        pdf_candidates = pdf_by_key.get(schedule_key, [])
        if not pdf_candidates:
            no_candidate.append(
                [
                    schedule_key[0],
                    appointment_time,
                    schedule_key[1],
                    appointments[0].patient,
                    "",
                    visit_providers,
                    events_text,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "No medication source",
                    "",
                    "Scheduled patient has no included Pharmacy Requirements line and no PDF Chemo candidate",
                ]
            )
            continue
        for medication in pdf_candidates:
            event_decision, event_note = get_pdf_event_decision(
                events, medication, event_rules
            )
            if event_decision == "INCLUDE":
                status = (
                    "Pharmacy Requirements has no included preparation line; "
                    f"{event_note}"
                )
            else:
                status = (
                    "Pharmacy Requirements has no included preparation line; "
                    f"manual review required; {event_note}"
                )
            if not medication.dose:
                status = (
                    "Dose could not be parsed from PDF; review the source page"
                )
            Not_Approved.append(
                [
                    medication.visit_date,
                    appointment_time,
                    medication.nhs_number,
                    medication.patient,
                    "",
                    visit_providers,
                    events_text,
                    medication.agent,
                    medication.dose,
                    medication.route,
                    "",
                    "",
                    "Patient Medications PDF",
                    f"Page {medication.page}",
                    status,
                ]
            )

    def sort_key(row):
        return (
            row[0],
            row[1] or datetime.max.time(),
            clean(row[3]),
            clean(row[7]),
        )

    for rows in (
        pharmacy_order,
        pharmacy_review,
        Not_Approved,
        not_scheduled,
        no_candidate,
    ):
        rows.sort(key=sort_key)

    workbook = Workbook()
    workbook.remove(workbook.active)
    write_output_sheet(
        workbook, "Pharmacy Order", pharmacy_order, "PharmacyOrder"
    )
    write_output_sheet(
        workbook, "Not Approved", Not_Approved, "NotApproved"
    )
    write_output_sheet(
        workbook,
        "Pharmacy Report Review",
        pharmacy_review,
        "PharmacyReportReview",
    )
    write_output_sheet(
        workbook, "Not Scheduled", not_scheduled, "NotScheduled"
    )
    write_output_sheet(
        workbook,
        "No Medication Candidate",
        no_candidate,
        "NoMedicationCandidate",
    )

    summary = workbook.create_sheet("Run Summary")
    summary.append(["Item", "Value"])
    summary_rows = [
        ("Generated", datetime.now()),
        ("Administration date", administration_date),
        ("Schedule source", schedule_path.name),
        ("Pharmacy source", pharmacy_path.name),
        ("PDF source", pdf_path.name),
        ("Agent rules", agent_rules_path.name),
        ("Event rules", event_rules_path.name),
        ("Schedule rows", len(schedule)),
        ("Distinct Pharmacy Requirements rows", len(pharmacy_requirements)),
        ("PDF Chemo medication rows", len(pdf_medications)),
        ("Pharmacy order lines", len(pharmacy_order)),
        ("Not Approved lines", len(Not_Approved)),
        ("Pharmacy report-review lines", len(pharmacy_review)),
        ("Included medication lines not scheduled", len(not_scheduled)),
        ("Scheduled patients with no candidate", len(no_candidate)),
        (
            "Important",
            "Planning extract only. Verify drug and dose against the current authorised ARIA prescription before preparation, release or administration.",
        ),
    ]
    for row in summary_rows:
        summary.append(row)
    summary.freeze_panes = "A2"
    summary.sheet_view.showGridLines = False
    summary.column_dimensions["A"].width = 48
    summary.column_dimensions["B"].width = 110
    for cell in summary[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for row in summary.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    summary["B2"].number_format = "dd/mm/yyyy hh:mm"
    summary["B3"].number_format = "dd/mm/yyyy"
    add_table(summary, "RunSummary")

    workbook.save(output_path)
    return {
        "administration_date": administration_date,
        "order": len(pharmacy_order),
        "not_approved": len(Not_Approved),
        "pharmacy_review": len(pharmacy_review),
        "not_scheduled": len(not_scheduled),
        "no_candidate": len(no_candidate),
    }


# ---------------------------------------------------------------------------
# No-parameter operational execution
# ---------------------------------------------------------------------------

def main() -> int:
    log_path = configure_logging()
    logging.info("Starting pharmacy advance-preparation process")
    try:
        ensure_folders()
        files = validate_required_files()
        logging.info("Schedule input: %s", files["schedule"])
        logging.info("Pharmacy input: %s", files["pharmacy"])
        logging.info("PDF input: %s", files["pdf"])

        schedule_rows = parse_schedule(files["schedule"])
        treatment_dates = {row.event_dt.date() for row in schedule_rows}
        if len(treatment_dates) != 1:
            raise ValueError(
                "The schedule must cover exactly one administration date"
            )
        administration_date = next(iter(treatment_dates))
        requested_output = OUTPUT_DIR / (
            f"Pharmacy_Advance_Preparation_{administration_date.isoformat()}.xlsx"
        )
        output_path = get_available_path(requested_output)
        if output_path != requested_output:
            logging.warning(
                "The standard output already exists. Creating: %s", output_path
            )

        statistics = build_workbook(
            schedule_path=files["schedule"],
            pharmacy_path=files["pharmacy"],
            pdf_path=files["pdf"],
            agent_rules_path=files["agent_rules"],
            event_rules_path=files["event_rules"],
            output_path=output_path,
        )
        if not output_path.exists():
            raise RuntimeError(
                "The build completed but the output workbook was not found"
            )
        if output_path.stat().st_size == 0:
            raise RuntimeError("The generated workbook is empty")

        logging.info("Output created: %s", output_path)
        logging.info("Pharmacy order lines: %s", statistics["order"])
        logging.info("Not Approved lines: %s", statistics["not_approved"])
        logging.info(
            "Pharmacy report-review lines: %s",
            statistics["pharmacy_review"],
        )
        logging.info("Not-scheduled lines: %s", statistics["not_scheduled"])
        logging.info("No-candidate patients: %s", statistics["no_candidate"])

        archive_folder = archive_input_files(
            input_paths=[
                files["schedule"],
                files["pharmacy"],
                files["pdf"],
            ],
            administration_date=administration_date,
        )
        logging.info("Inputs archived to: %s", archive_folder)

        print()
        print("Pharmacy preparation process completed successfully.")
        print(f"Output:  {output_path}")
        print(f"Archive: {archive_folder}")
        print(f"Log:     {log_path}")
        print()
        return 0

    except Exception:
        logging.exception("Pharmacy advance-preparation process failed")
        print()
        print("The pharmacy preparation process failed.")
        print("The input files have not been archived.")
        print(f"Review the log file: {log_path}")
        print()
        return 1


if __name__ == "__main__":
    sys.exit(main())
