$ErrorActionPreference = "Stop"

$root = Split-Path $PSScriptRoot -Parent
$outFile = Join-Path $root "data\files\TeklifPollHost.xlsm"
$vbaFile = Join-Path $root "data\modules-source\PollScheduler.bas"
$vbaCode = Get-Content $vbaFile -Raw

$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $module = $wb.VBProject.VBComponents.Add(1)
    $module.Name = "PollScheduler"
    $module.CodeModule.AddFromString($vbaCode.Trim())

    if (Test-Path $outFile) { Remove-Item $outFile -Force }
    $wb.SaveAs($outFile, 52)
    Write-Output "Olusturuldu: $outFile"
}
finally {
    if ($wb -ne $null) { $wb.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($excel -ne $null) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
