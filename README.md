# Pharmacy Advance Preparation

This utility builds an Excel planning and reconciliation workbook from ARIA schedule, pharmacy requirements, and patient medication reports. It does not calculate doses and must not replace verification against the current authorised ARIA prescription before preparation, release, or administration.

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

## Daily Run
To be run daily for the schedule fiur working days in advance:

| Day report run   | Shedule parameters |
| ---------------- | ------------------ |
| Monday week n    | Friday week n      |
| Tuesday week n   | Monday week n+1    |
| Wednesday week n | Tuesday week n+1   |
| Thursday week n  | Wednesday week n+1 |
| Friday week n    | Thursday week n_1  |


## Requirements

- Windows with Python 3.10 or later recommended.
- Python packages:

```powershell
python -m pip install xlrd openpyxl pypdf
```

The script uses the fixed base directory `C:\IDR\PharmacyAdvance`. Change `BASE_DIR` near the top of `build_pharmacy_requirements.py` if the working directory is different.

## Input Reports

Place these files in `C:\IDR\PharmacyAdvance\Input` using the exact filenames:

| File | Purpose |
| --- | --- |
| `sch_inst_by_pt_sum_aal.xls` | Authoritative daily schedule, including appointment time, NHS number, patient, event, and provider |
| `pharm_reqmt.xls` | Preferred medication source, including agent, dose, physician, dispensed, and verified fields |
| `ptmeds_sch_time.pdf` | Fallback medication source for scheduled patients without matching Pharmacy Requirements lines |

The report titles are validated before parsing. The schedule and Pharmacy Requirements report must each contain exactly one administration date, and those dates must match. The PDF date must also match when patient records are present.

## Directory Layout

```text
C:\IDR\PharmacyAdvance\
|-- Input\
|   |-- sch_inst_by_pt_sum_adel.xls
|   |-- pharm_reqmt.xls
|   `-- ptmeds_sch_time.pdf
|-- Output\
|-- Archive\
`-- Logs\
```

The script creates missing folders. It takes no command-line arguments.

## Run

1. Export the three reports from ARIA for one administration date.
2. Save them in `Input` with the exact filenames above.
3. From `C:\IDR\PharmacyAdvance`, run:

```powershell
python build_pharmacy_requirements.py
```

On success, the script prints the workbook, archive folder, and log paths. On failure it returns exit code `1` and leaves the input reports in `Input`.

## Source Precedence

1. The schedule is the authoritative list of appointments.
2. Matching Pharmacy Requirements lines are included as `Approved`.
3. For scheduled patients without matching Pharmacy Requirements lines, parsed chemotherapy lines from the Patient Medications PDF are included as `Planned`.
4. Patients or medication records that cannot be reconciled are placed in a review tab.

Patient matching uses the NHS number for schedule/PDF reconciliation and a normalised surname plus first-forename key for Pharmacy Requirements reconciliation. Exact duplicate medication lines are removed.

## Output Workbook

The workbook is written to:

```text
Output\Pharmacy_Advance_Preparation_YYYY-MM-DD.xlsx
```

If that name already exists, a timestamp is appended rather than overwriting it.

The medication and review tabs contain the following fields:

Administration date, Appointment time, Patient ID, Patient,
Prescription physician, Visit provider, Scheduled event(s),
Drug, Dose, Route, Dispensed, Verified, Source,
Source reference, and Status / reason.

To provide a simplified operational view for Pharmacy users,
the following reconciliation and troubleshooting columns are
hidden by default:

- Prescription physician
- Visit provider
- Dispensed
- Verified
- Source
- Source reference
- Status / reason

These columns remain present within the workbook and may be
unhidden by authorised users for investigation, data-quality
checks, reconciliation, or troubleshooting purposes.

### Workbook Tabs

- **Pharmacy List**: approved Pharmacy Requirements lines plus planned PDF fallback lines for scheduled patients. Reconciliation fields are hidden by default.
- **Patient Review**: scheduled patients not found in either medication source.
- **Drug Review**: approved or planned medication records whose patient is not found in the schedule.
- **Summary Sheet**: generation details, source filenames, row counts, and the planning-only warning.

Review `Pharmacy List` against the current authorised ARIA prescription. Review every row in `Patient Review` and `Drug Review` before preparation.

## Archiving and Logs

After a successful non-empty workbook is saved, the three input reports are moved to:

```text
Archive\YYYY-MM-DD\Administration_YYYY-MM-DD\
```

The first date is the run date and the second is the administration date. Existing archive files are not overwritten. Each run creates a log at:

```text
Logs\pharmacy_advance_YYYYMMDD_HHMMSS.log
```

If validation or workbook creation fails, the input reports are not archived. Correct the source reports and rerun after reviewing the log.

## Safety Checklist

Before pharmacy preparation, manually confirm:

- administration date, patient identity, appointment time, and event;
- drug, dose, route, prescribing physician, dispensed, and verified fields;
- every planned PDF fallback line and every row in both review tabs;
- all records against the current authorised ARIA prescription.