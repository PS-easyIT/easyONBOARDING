param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$htmlFile,
    
    [Parameter(Mandatory=$true)]
    [string]$pdfFile,
    
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$wkhtmltopdfPath
)

try {
    # Überprüfen, ob die PDF-Datei bereits existiert
    if (Test-Path $pdfFile) {
        Write-Warning "Die PDF-Datei existiert bereits und wird überschrieben: $pdfFile"
    }

    # Lokalen Pfad in URL umwandeln und URL-kodieren
    $htmlFileURL = "file:///" + ($htmlFile -replace '\\', '/') -replace ' ', '%20'
    
    # Argumente für wkhtmltopdf vorbereiten
    $arguments = @("--enable-local-file-access", $htmlFileURL, $pdfFile)
    
    Write-Host "Starte wkhtmltopdf mit folgenden Argumenten: $($arguments -join ' ')"
    
    # Prozess starten und Ergebnis erfassen
    $process = Start-Process -FilePath $wkhtmltopdfPath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
    
    # Prüfen des Prozess-Exit-Codes
    if ($process.ExitCode -ne 0) {
        Write-Error "wkhtmltopdf wurde mit Exit-Code $($process.ExitCode) beendet."
        return
    }
    
    # Überprüfen, ob die PDF-Datei erstellt wurde
    if (Test-Path $pdfFile) {
        Write-Host "PDF erfolgreich erstellt: $pdfFile"
    } else {
        Write-Error "PDF konnte nicht erstellt werden, obwohl wkhtmltopdf erfolgreich beendet wurde."
    }
}
catch {
    Write-Error "Ein Fehler ist aufgetreten: $_"
}