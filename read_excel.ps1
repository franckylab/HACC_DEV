try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $files = @(
        'D:\RAPPORT MENSUEL HACC 2025\REPORTING ADMINISTRATIF-FINANCIER-VENTES AOUT 2025.xlsx',
        'D:\RAPPORT MENSUEL HACC 2025\REPORTING ADMINISTRATIF-FINANCIER-VENTES SEPTEMBRE 2025.xlsx',
        'D:\RAPPORT MENSUEL HACC 2025\REPORTING ADMINISTRATIF-FINANCIER-VENTES OCTOBRE 2025.xlsx',
        'D:\RAPPORT MENSUEL HACC 2025\REPORTING ADMINISTRATIF-FINANCIER-VENTES NOVEMBRE 2025.xlsx',
        'D:\RAPPORT MENSUEL HACC 2025\REPORTING ADMINISTRATIF-FINANCIER-VENTES DECEMBRE 2025.xlsx'
    )
    
    foreach ($file in $files) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "FILE: $file" -ForegroundColor Cyan
        Write-Host "========================================"
        
        if (Test-Path $file) {
            $workbook = $excel.Workbooks.Open($file)
            foreach ($sheet in $workbook.Worksheets) {
                Write-Host "`n--- Sheet: $($sheet.Name) ---"
                $usedRange = $sheet.UsedRange
                $rows = $usedRange.Rows.Count
                $cols = $usedRange.Columns.Count
                Write-Host "Dimensions: $rows rows x $cols columns"
                
                # Read first 30 rows
                for ($i = 1; $i -le [Math]::Min(30, $rows); $i++) {
                    $rowValues = @()
                    for ($j = 1; $j -le [Math]::Min(20, $cols); $j++) {
                        $val = $sheet.Cells.Item($i, $j).Value2
                        if ($val -ne $null) {
                            $rowValues += $val.ToString()
                        } else {
                            $rowValues += ""
                        }
                    }
                    $joined = $rowValues -join ''
                    if ($joined.Trim()) {
                        Write-Host ($rowValues -join ' | ')
                    }
                }
            }
            $workbook.Close($false)
        } else {
            Write-Host "File not found!"
        }
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $($_.ScriptStackTrace) -ForegroundColor Yellow
}
