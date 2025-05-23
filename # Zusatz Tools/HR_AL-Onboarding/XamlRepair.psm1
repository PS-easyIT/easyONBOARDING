function Repair-XamlMargins {
    param([string]$XamlContent)
    return $XamlContent -replace '(Margin=")(\d+),(\d+),(\d+)(")', '$1$2,$3,$4,0$5'
}
