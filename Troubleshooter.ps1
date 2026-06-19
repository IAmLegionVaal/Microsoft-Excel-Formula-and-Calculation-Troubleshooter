#requires -Version 5.1
<# Created by Dewald Pretorius #>
[CmdletBinding()]param([string]$OutputPath)
$ErrorActionPreference='SilentlyContinue';if(-not $OutputPath){$OutputPath=Join-Path ([Environment]::GetFolderPath('Desktop')) 'Excel_Formula_Reports'};New-Item $OutputPath -ItemType Directory -Force|Out-Null
$stamp=Get-Date -Format yyyyMMdd_HHmmss;$report=Join-Path $OutputPath "Excel_Formula_$stamp.txt"
$excel=Get-Process EXCEL; $events=Get-WinEvent -FilterHashtable @{LogName='Application';StartTime=(Get-Date).AddDays(-5)}|Where-Object Message -match 'EXCEL|calculation|formula'|Select-Object -First 30 TimeCreated,Id,ProviderName,Message
$calc=Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
@('MICROSOFT EXCEL FORMULA AND CALCULATION TROUBLESHOOTER','Created by Dewald Pretorius',"Generated: $(Get-Date)",'',"Excel running: $([bool]$excel)","Calculation registry state:",($calc|Format-List|Out-String -Width 220),'Recent events:',($events|Format-List|Out-String -Width 220),'Guidance: verify calculation mode, circular references, broken links, volatile formulas, named ranges, and workbook-specific corruption.')|Set-Content $report -Encoding UTF8
Write-Host "Report created: $report" -ForegroundColor Green