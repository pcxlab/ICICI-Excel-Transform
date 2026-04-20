function Get-MOPFromFileName {

    param([string]$FileName)

    $name = $FileName.ToUpper()

    # Bank
    switch -Regex ($name) {
        "ICICI" { $bank = "ICICI"; break }
        "HDFC"  { $bank = "HDFC"; break }
        default { $bank = "UNK" }
    }

    # Type
    switch -Regex ($name) {
        "_CC_" { $type = "CC"; break }
        "_SB_" { $type = "SB"; break }
        default { $type = "OT" }
    }

    # Initials
    if ($name -match "^[A-Z]+_[A-Z]+_([A-Z]{2})_") {
        $initials = $matches[1]
    }
    else {
        $initials = "NA"
    }

    # 🔥 CHANGED HERE → underscore instead of dash
    return "${bank}_${type}_${initials}"
}