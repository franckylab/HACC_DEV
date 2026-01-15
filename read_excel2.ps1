# Kill any existing Excel processes
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds 2

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $file = 'D:\RAPPORT MENSUEL HACC 2025\REPORTING ADMINISTRATIF-FINANCIER-VENTES AOUT 2025.xlsx'
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "FILE: AOUT 2025" -ForegroundColor Cyan
    Write-Host "========================================"
    
    if (Test-Path $file) {
        $workbook = $excel.Workbooks.Open($file)
        
        foreach ($sheet in $workbook.Worksheets) {
            Write-Host "`n--- Sheet: $($sheet.Name) ---"
            $usedRange = $sheet.UsedRange
            $rows = $usedRange.Rows.Count
            $cols = $usedRange.Columns.Count
            Write-Host "Dimensions: $rows rows x $cols columns"
            
            # Read first 50 rows
            for ($i = 1; $i -le [Math]::Min(50, $rows); $i++) {
                $rowValues = @()
                for ($j = 1; $j -le [Math]::Min(25, $cols); $j++) {
                    $val = $sheet.Cells.Item($i, $j).Value2
                    if ($val -ne $null) {
                        $rowValues += $val.ToString()
                    } else {
                        $rowValues += ""
                    }
                }
                $joined = $rowValues -join ' '
                if ($joined.Trim()) {
                    Write-Host $joined
                }
            }
        }
        
        $workbook.Close($false)
    } else {
        Write-Host "File not found!"
    }
    
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $($_.ScriptStackTrace) -ForegroundColor Yellow
    
    # Try to cleanup
    try {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    } catch {}
}
