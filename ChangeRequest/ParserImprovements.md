Project: Pharmacy Advance

Current branch:
feature/pdf-parser-improvements

Status:

- Drug exclusion lookup complete and merged.
- Issue #8 created: Improve PDF medication parsing.
- Issue #9 created: Replace Patient Medications PDF with XLS.

Current findings:

- parse_patient_medication_xls() prototype implemented.
- XLS parser returns 113 medication records.
- PDF parser returns 110 medication records.
- PHESGO loading and maintenance doses identified correctly.
- Pembrolizumab identified correctly.
- Azacitidine identified correctly.
- Structured XLS data appears more reliable than PDF parsing.

Current implementation:

- XlsMedication dataclass created.
- parse_patient_medication_xls() extracts patient + agent.
- Branch pushed to GitHub.

Next objective:

Enhance parse_patient_medication_xls() to populate:
- visit_date
- visit_time
- course_description
- dose
- route

Then build xls_meds_by_key and begin replacing PDF medication usage inside build_workbook().