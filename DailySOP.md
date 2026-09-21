# Daily Standard Operating Procedures

## 1. Run Reports

From ARIA MO, run the following reports 

| Report Name | Parameter | Export format |
| --- | --- | --- |
| `sch_inst_by_pt_sum_aal` | CHOC & Today + 4 Working Days | `.xls` |
| `pharm_reqmt` | CHOC & Today + 4 Working Days | `.xls` |
| `ptmeds_sch_time` | CHOC & Today + 4 Working Days | `.pdf` |

### Export Settings

- Format: `Microsoft Excel 97-2000 (XLS)
- Accept default settings

### Save Location

`RHU-D090232\IDR\PharmacyAdvance\Input\`

## 2. Run Python Script

To be run on RHU-D090232

``` poweshell
python C:\git-repos\pharmacy-advance\build_pharmacy_requirements.py
```

## 3. Output

- Saved to `RHU-090232\IDR\PharmacyAdvance\Output`
- Send to robert.williams53@nhs.net