Import-Module PCXLab.Core
Import-Module PCXLab.Excel

$config = Read-Config -Path ".\config.json"

Write-Log "Scanning folder..."

$files = Get-ICICIFiles -Path $config.InputFolder

foreach ($file in $files) {

    try {
        Write-Log "Processing $($file.Name)"

        $result = Convert-ICICIFormat -File $file

        Export-Excel `
            -Path (Join-Path $config.OutputFolder ($file.BaseName + "_Transformed.xlsx")) `
            -InputObject $result `
            -AutoSize -BoldTopRow

        Write-Log "Completed $($file.Name)" "SUCCESS"
    }
    catch {
        Write-Log $_.Exception.Message "ERROR"
    }
}