param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [string]$OwnerPidFile
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$excel = $null
$workbook = $null

try {
    $resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $existingExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    if ($OwnerPidFile) {
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class SheetPilotWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
        [uint32]$ownedExcelId = 0
        [void][SheetPilotWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$ownedExcelId)
        if ($ownedExcelId -gt 0 -and $ownedExcelId -notin $existingExcelIds) {
            [IO.File]::WriteAllText($OwnerPidFile, [string]$ownedExcelId)
        }
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($resolvedPath)
    $excel.Calculation = -4105
    $excel.CalculateBeforeSave = $true
    $excel.CalculateFullRebuild()
    $workbook.Save()
    $workbook.Close($false)
    $workbook = $null
    $excel.Quit()
    $excel = $null
    @{ success = $true; message = "Microsoft Excel completed full recalculation and saved cached values." } |
        ConvertTo-Json -Compress
    exit 0
}
catch {
    @{ success = $false; message = $_.Exception.Message } | ConvertTo-Json -Compress
    exit 1
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
