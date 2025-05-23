param(
    [Parameter(Mandatory=$true)]
    [string]$htmlFile,
    [Parameter(Mandatory=$true)]
    [string]$pdfFile,
    [Parameter(Mandatory=$true)]
    [string]$wkhtmltopdfPath
)

# Lokalen Pfad in URL umwandeln
$htmlFileURL = "file:///" + ($htmlFile -replace '\\', '/')
$arguments = @("--enable-local-file-access", $htmlFileURL, $pdfFile)
Write-Host "Starte wkhtmltopdf mit folgenden Argumenten: $arguments"
Start-Process -FilePath $wkhtmltopdfPath -ArgumentList $arguments -Wait -NoNewWindow
if (Test-Path $pdfFile) {
    Write-Host "PDF erfolgreich erstellt: $pdfFile"
} else {
    Write-Error "PDF konnte nicht erstellt werden."
}
