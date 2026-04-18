function Convert-ICICIFormat {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File
    )

    # Step 1: Convert XLS → XLSX if needed
    $workingFile = Convert-XlsToXlsx -File $File

    # Step 2: Get MOP
    $mop = Get-MOPFromFileName -FileName $workingFile.Name

    # Step 3: Read raw Excel
    $raw = Import-Excel $workingFile.FullName -NoHeader

    # Step 4: Detect header row
    $headerIndex = Get-ICICIHeader -RawData $raw

    if ($null -eq $headerIndex) {
        throw "Header not found in $($File.Name)"
    }

    # Step 5: Extract header row
    $headerRow = $raw[$headerIndex].PSObject.Properties.Value

    # Step 6: Build column map (ROBUST)
    $colMap = @{}

    for ($i = 0; $i -lt $headerRow.Count; $i++) {
        $val = $headerRow[$i]

        if ($val -match "Transaction Date") { $colMap.Date = $i }
        elseif ($val -match "Details")      { $colMap.Details = $i }
        elseif ($val -match "Amount")       { $colMap.Amount = $i }
        elseif ($val -match "Reference")    { $colMap.Ref = $i }
    }

    # Step 7: Validate mapping
    if ($colMap.Count -lt 4) {
        throw "Column mapping incomplete. Found: $($colMap.Keys -join ', ')"
    }

    # Step 8: Extract data rows
    $data = $raw[($headerIndex + 2)..($raw.Count - 1)]

    $prevDate = $null

    foreach ($row in $data) {

        $values = $row.PSObject.Properties.Value

        $date    = $values[$colMap.Date]
        $details = $values[$colMap.Details]
        $amount  = $values[$colMap.Amount]
        $ref     = $values[$colMap.Ref]

        # Handle merged/blank date rows
        if (-not $date) {
            $date = $prevDate
        } else {
            $prevDate = $date
        }

        # Skip completely empty rows
        if (-not $date -and -not $details -and -not $amount) {
            continue
        }

        # Parse amount
        $amtDr = 0
        $amtCr = 0

        if ($amount -match "([\d\.]+)\s*Dr") {
            $amtDr = [decimal]$matches[1]
        }
        elseif ($amount -match "([\d\.]+)\s*Cr") {
            $amtCr = [decimal]$matches[1]
        }

        # Output object (your schema)
        [PSCustomObject]@{
            Date           = $date
            Narration      = $details
            Item           = ""
            Category       = ""
            Place          = ""
            Freq           = ""
            For            = ""
            MOP            = $mop
            "Amt (Dr)"     = $amtDr
            "Amt (Cr)"     = $amtCr
            "Value Dt"     = if ($amtDr -gt 0) { "Dr." } elseif ($amtCr -gt 0) { "Cr." };
            "Chq./Ref.No." = $ref
        }
    }
}