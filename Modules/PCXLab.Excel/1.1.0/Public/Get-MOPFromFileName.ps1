function Get-MOPFromFileName {
    param([string]$FileName)

    # Bank
    if ($FileName -match "ICICI") { $bank = "ICICI" }
    elseif ($FileName -match "HDFC") { $bank = "HDFC" }
    else { $bank = "UNK" }

    # Type
    if ($FileName -match "CC") { $type = "CC" }
    elseif ($FileName -match "SB") { $type = "SB" }
    else { $type = "OT" }

    # Initials (e.g., _HP_)
    if ($FileName -match "_([A-Z]{2})_") {
        $initials = $matches[1]
    }
    else {
        $initials = "NA"
    }

    return "$bank-$type-$initials"
}