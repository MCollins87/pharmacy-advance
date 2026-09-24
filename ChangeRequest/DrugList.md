# Change Request - Pre-med and Main Treatment

**Requestor:** Rob Williams
**Date:** 2026-09-21

## Request Details

    In the drug column it pulls drugs which are not the main treatment but just pre meds but doesn’t do this for every regimen. Is it possible to lose these from the report?
    One way we could ‘tidy’ it would be filtering using the route column but this isn’t consistently populated
    Happy to discuss

## Current Logic

1. Uses Pharmacy Requirements XLS as the preferred source.
2. Uses Patient Medications PDF as the fallback.
3. Outputs whatever drug records are present.
4. Makes no distinction between:
    - Treatment drugs
    - Premeds
    - Supportive medications
    - Hydration

``` mermaid
flowchart TD

    A[Start] --> B[Load Schedule XLS]

    B --> C[Extract Scheduled Patients]

    C --> D[Load Pharmacy Requirements XLS]

    D --> E[Load Patient Medications PDF]

    E --> F[Validate Administration Dates Match]

    F --> G[Build Patient Lookup Tables]

    G --> H{Patient in Schedule?}

    H -->|Yes| I{Found in Pharmacy Requirements?}

    I -->|Yes| J["Add to Pharmacy List<br/>Status = Approved"]

    I -->|No| K{Found in Patient Medications PDF?}

    K -->|Yes| L["Add to Pharmacy List<br/>Status = Planned"]

    K -->|No| M[Add to Patient Review]

    H -->|No| N[Drug Review]

    J --> O[Generate Workbook]
    L --> O
    M --> O
    N --> O

    O --> P[Archive Source Files]

    P --> Q[Complete]
```

### Drug Handling:
``` mermaid
flowchart TD

    A[Drug Record Found] --> B[Output Drug]

    B --> C[Workbook]

    C --> D[Pharmacy List]
```

## Proppsed Changes

Include a config file to identify drugs that need manufactured. 

Drug Handling becomes: 
``` mermaid
flowchart TD

    A[Drug Record Found]

    A --> B[Lookup ManufactureRequired.csv]

    B --> C{Drug in Lookup?}

    C -->|Yes| D[Manufacture Required]

    C -->|No| E[Off The Shelf]

    D --> F[Highlight in Workbook]

    E --> G[Normal Display]

    F --> H[Pharmacy List]
    G --> H
```
This will put all drugs in output, but if we want pre-meds removed, then:

``` mermaid
flowchart TD

    A[Drug Record]

    A --> B[Supportive Drug Lookup]

    B --> C{Supportive?}

    C -->|Yes| D[Hide from Drug Column]

    C -->|No| E[Display Drug]

    D --> F[Optional Audit Tab]

    E --> G[Pharmacy List]
```

## Meeting 2026-09-24 11:00

Start with
    Given the current isolator outage, would it be more useful if the report identified drugs requiring aseptic manufacture versus off-the-shelf drugs?

### Executive summary

The agreed direction is to introduce a maintained exclusion lookup containing drugs that Pharmacy does not need to see. The report will then focus on drugs that may require advance manufacture, while the exclusion list is progressively refined from Pharmacy feedback.

### Pharmacy Advance discussion summary
#### Purpose

The report is currently being used to help the manufacturing unit anticipate the number of chemotherapy items that may need to be prepared. It is intended as a planning indicator, because doses may still change before treatment.

#### Current problem
For approved prescriptions, drug information comes from the Pharmacy Requirements report.
When a scheduled patient is not found there, the script falls back to the Patient Medications PDF.
The PDF fallback can include premedications, PRN medications and take-home medication alongside the relevant chemotherapy drugs.
ARIA’s medication sections do not consistently distinguish the relevant items. Some premedications and PRN items can appear in the same section as chemotherapy.

This means filtering by section or route alone would not be sufficiently reliable.

#### Agreed approach

Use an exclusion lookup rather than trying to maintain a definitive list of drugs requiring manufacture.

The proposed logic is:
``` mermaid
flowchart TD
    A[Drug extracted from ARIA report] --> B{Drug in exclusion lookup?}
    B -->|Yes| C[Exclude from Pharmacy List]
    B -->|No| D[Retain in Pharmacy List]
    D --> E[Pharmacy reviews retained item]
    E --> F{Not required in future?}
    F -->|Yes| G[Add drug to exclusion lookup]
    F -->|No| H[Continue to display]
```

This was preferred because manufacture requirements are not completely binary. Some drugs may normally be purchased as stock but still require particular doses to be manufactured. An inclusion-only list could therefore omit something important.

#### Lookup maintenance

The proposed iterative process is:

1. Provide Pharmacy with the current list of drug names.
2. Pharmacy identifies drugs that are never relevant to advance manufacture.
3. Add those drugs to the exclusion lookup.
4. Run the next report.
5. Pharmacy reports any remaining unwanted drugs.
6. Update and progressively refine the lookup.

The configuration file should remain separate from the packaged application so Pharmacy requirements can be updated without rebuilding the program.

#### Output requirements
Retain the simplified workbook layout and currently hidden technical columns.
Remove superfluous medication rows using the exclusion lookup.
Continue showing drugs not yet classified, rather than silently excluding them.
Preserve the report as a planning tool rather than treating it as a definitive manufacturing instruction.
#### Actions
| Owner | Action|
| --- | --- |
|Pharmacy | Review the supplied drug list and highlight drugs that do not need to appear. |
| Mark |	Clean and deduplicate the returned list.|
| Mark	|Create an external drug-exclusion lookup file.|
| Mark	|Update the Python script to filter medication rows against that lookup.|
| Mark	|Keep the lookup separate from any packaged executable.|
| Pharmacy and Mark	|Refine the lookup following subsequent report runs.|
| Mark	|Later investigate packaging the tool so other authorised users can run it.|

The first five actions were directly discussed. Packaging was identified as a later enhancement because the tool currently runs only on your computer, while the source ARIA reports still require manual generation.

#### Recommended lookup structure
```
drug_name,exclude,reason,reviewed_by,review_date
Dexamethasone,Yes,Premedication,,
Ondansetron,Yes,Supportive medication,,
```

**Recommendation:** default to retaining any drug that is missing from the lookup. This is the safer failure mode because an unclassified drug remains visible for Pharmacy review rather than being silently removed.

## Decision Log

Decision 1
```
Drug filtering will use an exclusion lookup.

Rationale:
Pharmacy cannot reliably define all manufacture-required drugs.
Safer to exclude known unwanted drugs and leave unknown drugs visible.
```

Decision 2
```
Lookup maintained by Pharmacy.

Format:
CSV or Excel file.

Behaviour:
Missing drugs remain visible.
```

Decision 3
```
Configuration stored separately from executable.

Allows Pharmacy updates without code changes.
```
