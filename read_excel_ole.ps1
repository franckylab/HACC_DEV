# Essayer d'utiliser OLE DB pour lire Excel directement sans Excel application
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='D:\RAPPORT MENSUEL HACC 2025\REPORTING ADMINISTRATIF-FINANCIER-VENTES AOUT 2025.xlsx';Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1'"

try {
    $connection = New-Object System.Data.OleDb.OleDbConnection
    $connection.ConnectionString = $connectionString
    $connection.Open()
    
    $schema = $connection.GetSchema("Tables")
    Write-Host "Available sheets:"
    foreach ($row in $schema) {
        Write-Host "  - $($row['TABLE_NAME'])"
    }
    
    foreach ($row in $schema) {
        $sheetName = $row['TABLE_NAME']
        $query = "SELECT * FROM [$sheetName]"
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $query, $connection
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        
        Write-Host "`n========================================"
        Write-Host "Sheet: $sheetName"
        Write-Host "========================================"
        
        $table = $dataset.Tables[0]
        for ($i = 0; $i -lt [Math]::Min(50, $table.Rows.Count); $i++) {
            $row = $table.Rows[$i]
            $values = @()
            foreach ($col in $table.Columns) {
                $val = $row[$col]
                if ($val -ne $null -and $val -ne [System.DBNull]::Value) {
                    $values += $val.ToString()
                } else {
                    $values += ""
                }
            }
            $joined = $values -join ' | '
            if ($joined.Trim()) {
                Write-Host $joined
            }
        }
    }
    
    $connection.Close()
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "ACE OLEDB driver not available. Trying alternative approach..."
}
