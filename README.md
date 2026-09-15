# Pharmacy Advance Preparation

This utility builds an Excel planning workbook for pharmacy advance preparation from ARIA schedule, pharmacy requirements, and patient medication reports.

It is a planning and reconciliation aid. It does not calculate doses and does not replace verification against the current authorised ARIA prescription before preparation, release, or administration.

## Purpose
Following the failure of the Pharmacy Manufacturing Isolator used to produce SACT for PHU, Oncology havebeen tasked with providing required medications at least four days in advance.

This process generates a pharmacy advance-preparation spreadsheet for treatments scheduled on a specifiedadministration date.
The workbook combines multiple ARIA reports to identify:
- scheduled patients;
- treatment medications;
- prescribed doses;
- prescribing clinicians;
- appointments requiring review.

The objective is to provide pharmacy with advance notice of medicines that may require preparation while also highlighting records that require manual checking before inclusion.

## What It Does

The process:

1. Reads the timed schedule as the authoritative appointment source.
2. Reads Pharmacy Requirements as the primary medication source.
3. Applies the configured agent preparation rules.
4. Uses the Patient Medications PDF only as a fallback for scheduled patients without an included Pharmacy Requirements line.
5. Creates a dated workbook with preparation lines and exception views.
6. Archives the three report files only after a non-empty workbook has been saved.

The schedule and Pharmacy Requirements reports must each contain exactly one date, and those dates must match.

## Requirements

- Windows machine with Python 3.10 or later recommended.
- Python packages:

```text
xlrd
openpyxl
pypdf
```

Install them with:

```powershell
python -m pip install xlrd openpyxl pypdf
```

The script uses the fixed base directory `C:\IDR\PharmacyAdvance`. If the folder is moved, update `BASE_DIR` near the top of `build_pharmacy_requirements.py` before running it.

## Directory Layout

```text
C:\IDR\PharmacyAdvance\
|-- build_pharmacy_requirements.py
|-- Input\
|   |-- sch_by_inst_pt_excel_vprov.xls
|   |-- pharm_reqmt.xls
|   `-- ptmeds_sch_time.pdf
|-- Output\
|-- Archive\
|-- Config\
|   |-- agent_preparation_rules.csv
|   `-- event_drug_rules.csv
`-- Logs\
```

The script creates missing `Input`, `Output`, `Archive`, `Config`, and `Logs` folders, but the two configuration CSV files must already exist.

## Input Reports

Export these reports from ARIA and save them in `Input` using the exact filenames below.

| File | ARIA report | Main use |
| --- | --- | --- |
| `sch_by_inst_pt_excel_vprov.xls` | Schedule - Institution by Time - Exportable to Excel | Date, appointment time, NHS number, patient, event, and visit provider |
| `pharm_reqmt.xls` | Pharmacy Requirements - by Agent, Rx Type, Administration Date and Patient | Drug, dose, prescribing physician, administration date, dispensed, and verified indicators |
| `ptmeds_sch_time.pdf` | Patient Medications - Patients Scheduled to Visit - by Time | Fallback chemotherapy candidates for scheduled patients missing an included Pharmacy Requirements line |

The script validates the report titles before processing. It ignores report header rows, totals, and the `Oncology Day Unit` line where applicable. Exact duplicate Pharmacy Requirements lines are collapsed while different doses are retained.

## Configuration

### `agent_preparation_rules.csv`

Controls automatic decisions for medication agents. Rules are evaluated from top to bottom; the first matching regular expression wins.

Required columns:

```csv
agent_pattern,decision,notes
```

Supported decisions are `INCLUDE`, `REVIEW`, and `EXCLUDE`.

- `INCLUDE`: eligible for `Pharmacy Order` when a scheduled patient and dose are present.
- `REVIEW`: written to `Pharmacy Report Review` when the patient is scheduled.
- `EXCLUDE`: omitted from the output workbook.

The final catch-all rule normally provides the default decision for agents not otherwise listed. An agent absent from the rules defaults to `REVIEW`.

### `event_drug_rules.csv`

Controls fallback decisions for candidates parsed from the PDF. A candidate is included in `Not Approved` with an automatic-inclusion note only when its scheduled event and agent match a rule with a single consistent decision. Otherwise it is marked for manual review.

Required columns:

```csv
event_pattern,agent_pattern,decision,notes
```

Regular expressions are case-insensitive in the script. Keep these rules under clinical/pharmacy ownership and review them when ARIA event names or regimens change.

## Run The Process

1. Run the three ARIA reports for the required administration date.
2. Remove or archive any prior files from `Input`.
3. Save the new reports with the exact filenames in the input table.
4. Confirm both CSV rule files are present in `Config`.
5. From `C:\IDR\PharmacyAdvance`, run:

```powershell
python build_pharmacy_requirements.py
```

The script takes no command-line arguments. On success it prints the output workbook, archive folder, and log path. On failure it returns exit code `1` and leaves the report files in `Input`.

## Output Workbook

The workbook is written to:

```text
Output\Pharmacy_Advance_Preparation_YYYY-MM-DD.xlsx
```

If that filename already exists, a timestamp is appended instead of overwriting it, for example `Pharmacy_Advance_Preparation_2026-09-17_20260914_180001.xlsx`.

Every data sheet uses the same columns:

`Administration date`, `Appointment time`, `Patient ID`, `Patient`, `Prescription physician`, `Visit provider`, `Scheduled event(s)`, `Drug`, `Dose`, `Route`, `Dispensed`, `Verified`, `Source`, `Source reference`, and `Status / reason`.

### Workbook Tabs

Review the tabs in this order:

1. **Run Summary**: generation details, source files, row counts, and exception totals.
2. **Pharmacy Order**: Pharmacy Requirements lines that match a scheduled patient, are `INCLUDE`, and have a dose. These are the primary advance-preparation planning lines.
3. **Not Approved**: PDF fallback medication candidates for scheduled patients without an included Pharmacy Requirements line. Confirm the event-to-drug match and source PDF page manually.
4. **No Medication Candidate**: scheduled patients with neither an included Pharmacy Requirements line nor a PDF chemotherapy candidate. Treat these as high-priority exceptions.
5. **Not Scheduled**: included Pharmacy Requirements lines with no matching timed schedule record. Investigate cancellations, deferrals, moved appointments, or stale prescriptions.
6. **Pharmacy Report Review**: scheduled Pharmacy Requirements lines whose agents are configured as `REVIEW`.

The `Status / reason` column records missing doses, dispensed or verified indicators, unmatched schedules, rule notes, and fallback-review reasons.

## Archiving And Logs

After a successful build, the three report files are moved to:

```text
Archive\YYYY-MM-DD\Administration_YYYY-MM-DD\
```

The first date is the run date; the second is the administration date. Existing archive files are not overwritten. Each run creates a log at:

```text
Logs\pharmacy_advance_YYYYMMDD_HHMMSS.log
```

If validation or workbook creation fails, the source reports are not archived. Check the log for the exception and correct the input or configuration before rerunning.

## Validation Checklist

Before pharmacy preparation:

- Confirm the administration date and source filenames on `Run Summary`.
- Compare `Pharmacy Order` with the current authorised ARIA prescription.
- Verify patient identity, drug, dose, route, appointment time, and prescribing physician.
- Review every line in `Not Approved`, `No Medication Candidate`, `Not Scheduled`, and `Pharmacy Report Review`.
- Confirm dispensed and verified indicators where populated.
- Record discrepancies and update the configuration rules only after the appropriate clinical/pharmacy review.

For reconciliation, compare at minimum:

```text
Administration date + Patient ID + Drug + Dose
```

## Ownership

Clinical/pharmacy owners should approve the medication and event rules and the final preparation list. The technical owner should maintain the script, dependencies, and run logs.


## Known Issues

1. Report `ptmeds_sch_time` will only provide one A4 pahe per patient. For patients where the list of medications exceed the provided length, admissions are ommitted. These ptients will therfore appear in the tab `No Medication Candidate`. 
    - **Proposed action**: The tab `No Medication Candidate` should be reviewed for events indicating a treatment and mannually checked in Aria Medical Oncology. 