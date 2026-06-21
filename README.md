# Microsoft Excel Formula and Calculation Troubleshooter

Created by **Dewald Pretorius**.

The repository includes the original formula, calculation-mode, circular-reference, link, and workbook diagnostics plus a guarded `Repair.ps1` helper.

Supported actions are `Diagnose`, `ResetOfficeCache`, and `RepairOffice`.

```powershell
.\Repair.ps1 -Action Diagnose
.\Repair.ps1 -Action ResetOfficeCache -WhatIf
.\Repair.ps1 -Action RepairOffice -Confirm
```

Close Excel before cache repair. Existing cache data is preserved as a timestamped backup, and current calculation options are recorded before changes. Source-reviewed for PowerShell 5.1; not runtime-tested against every workbook or Office build.
