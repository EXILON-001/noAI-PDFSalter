Add-Type -AssemblyName System.Windows.Forms

# Hauptformular erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = "Anti AI PDF Salting Tool"
$form.Width = 600
$form.Height = 400
$form.StartPosition = "CenterScreen"

# Labels und Textfelder für Quell- und Ziel-PDF
$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Text = "Quell-PDF:"
$labelSource.Location = New-Object System.Drawing.Point(10, 20)
$labelSource.AutoSize = $true

$textBoxSource = New-Object System.Windows.Forms.TextBox
$textBoxSource.Location = New-Object System.Drawing.Point(100, 18)
$textBoxSource.Width = 380

$buttonBrowseSource = New-Object System.Windows.Forms.Button
$buttonBrowseSource.Text = "Durchsuchen..."
$buttonBrowseSource.Location = New-Object System.Drawing.Point(490, 15)
$buttonBrowseSource.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "PDF-Dateien (*.pdf)|*.pdf"
    $dialog.Title = "Quell-PDF auswählen"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxSource.Text = $dialog.FileName
    }
})

$labelDestination = New-Object System.Windows.Forms.Label
$labelDestination.Text = "Ziel-PDF:"
$labelDestination.Location = New-Object System.Drawing.Point(10, 60)
$labelDestination.AutoSize = $true

$textBoxDestination = New-Object System.Windows.Forms.TextBox
$textBoxDestination.Location = New-Object System.Drawing.Point(100, 58)
$textBoxDestination.Width = 380

$buttonBrowseDestination = New-Object System.Windows.Forms.Button
$buttonBrowseDestination.Text = "Durchsuchen..."
$buttonBrowseDestination.Location = New-Object System.Drawing.Point(490, 55)
$buttonBrowseDestination.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "PDF-Dateien (*.pdf)|*.pdf"
    $dialog.Title = "Ziel-PDF auswählen"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxDestination.Text = $dialog.FileName
    }
})

# Qualitätsauswahl hinzufügen
$labelQuality = New-Object System.Windows.Forms.Label
$labelQuality.Text = "Qualität:"
$labelQuality.Location = New-Object System.Drawing.Point(10, 100)
$labelQuality.AutoSize = $true

$comboBoxQuality = New-Object System.Windows.Forms.ComboBox
$comboBoxQuality.Location = New-Object System.Drawing.Point(100, 98)
$comboBoxQuality.Width = 150
$comboBoxQuality.DropDownStyle = 'DropDownList'
$comboBoxQuality.Items.AddRange(@("Niedrig (kleinere Datei)", "Mittel", "Hoch (größere Datei)"))
$comboBoxQuality.SelectedIndex = 1  # Standardmäßig "Mittel"

# Hinweistext zur Qualität
$labelQualityNote = New-Object System.Windows.Forms.Label
$labelQualityNote.Text = "Hinweis: Höhere Qualität kann zu größeren Dateien führen."
$labelQualityNote.Location = New-Object System.Drawing.Point(260, 100)
$labelQualityNote.Width = 320
$labelQualityNote.AutoSize = $false

# Start-Button erstellen
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Start"
$buttonStart.Location = New-Object System.Drawing.Point(250, 140)
$buttonStart.Enabled = $true

# Schließen-Button erstellen (anfangs unsichtbar)
$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Schließen"
$buttonClose.Location = New-Object System.Drawing.Point(250, 140)
$buttonClose.Visible = $false
$buttonClose.Add_Click({
    $form.Close()
})

# Fortschrittsbalken erstellen
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 190)
$progressBar.Width = 560
$progressBar.Minimum = 0
$progressBar.Maximum = 100

# Status-Label erstellen
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "Status: Warten auf Eingabe"
$labelStatus.Location = New-Object System.Drawing.Point(10, 230)
$labelStatus.Width = 560

# Steuerelemente zum Formular hinzufügen
$form.Controls.Add($labelSource)
$form.Controls.Add($textBoxSource)
$form.Controls.Add($buttonBrowseSource)
$form.Controls.Add($labelDestination)
$form.Controls.Add($textBoxDestination)
$form.Controls.Add($buttonBrowseDestination)
$form.Controls.Add($labelQuality)
$form.Controls.Add($comboBoxQuality)
$form.Controls.Add($labelQualityNote)
$form.Controls.Add($buttonStart)
$form.Controls.Add($buttonClose)
$form.Controls.Add($progressBar)
$form.Controls.Add($labelStatus)

# Funktion zum Aktualisieren des Status
function Update-Status {
    param (
        [string]$message
    )
    $labelStatus.Text = "Status: $message"
    $labelStatus.Refresh()
}

# Funktion zum Aktualisieren des Fortschrittsbalkens
function Update-Progress {
    param (
        [int]$value,
        [int]$max
    )
    $progressBar.Maximum = $max
    $progressBar.Value = $value
    $progressBar.Refresh()
}

# Funktion zur Qualitätsauswahl
function Get-QualitySettings {
    switch ($comboBoxQuality.SelectedIndex) {
        0 { return @{ dpi = 72; quality = 30; } }   # Niedrig
        1 { return @{ dpi = 150; quality = 50; } }  # Mittel
        2 { return @{ dpi = 300; quality = 80; } }  # Hoch
        default { return @{ dpi = 150; quality = 50; } }
    }
}

# Funktion zur PDF-Verarbeitung
$buttonStart.Add_Click({
    # Start-Button deaktivieren
    $buttonStart.Enabled = $false

    # Eingabefelder deaktivieren
    $textBoxSource.Enabled = $false
    $buttonBrowseSource.Enabled = $false
    $textBoxDestination.Enabled = $false
    $buttonBrowseDestination.Enabled = $false
    $comboBoxQuality.Enabled = $false

    # Quell- und Zielpfade abrufen
    $pdfPath = $textBoxSource.Text
    $outputPdfPath = $textBoxDestination.Text

    if (-not (Test-Path $pdfPath)) {
        Update-Status "Die Quell-PDF existiert nicht."
        $buttonStart.Enabled = $true
        $textBoxSource.Enabled = $true
        $buttonBrowseSource.Enabled = $true
        $textBoxDestination.Enabled = $true
        $buttonBrowseDestination.Enabled = $true
        $comboBoxQuality.Enabled = $true
        return
    }

    if ([string]::IsNullOrEmpty($outputPdfPath)) {
        Update-Status "Bitte wählen Sie eine Ziel-PDF aus."
        $buttonStart.Enabled = $true
        $textBoxSource.Enabled = $true
        $buttonBrowseSource.Enabled = $true
        $textBoxDestination.Enabled = $true
        $buttonBrowseDestination.Enabled = $true
        $comboBoxQuality.Enabled = $true
        return
    }

    # Qualitätsparameter abrufen
    $qualitySettings = Get-QualitySettings
    $dpi = $qualitySettings.dpi
    $quality = $qualitySettings.quality

    try {
        # Fortschrittsvariablen initialisieren
        $progressCurrent = 0
        $progressTotal = 2  # Vorläufige Gesamtzahl der Schritte (wird später angepasst)
        Update-Progress -value $progressCurrent -max $progressTotal

        Update-Status "Erstelle temporäres Verzeichnis..."
        $tempDir = New-Item -ItemType Directory -Path ([System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName()))
        $progressCurrent++
        Update-Progress -value $progressCurrent -max $progressTotal

        Update-Status "Konvertiere PDF in Bilder..."
        magick -density $dpi "$pdfPath" "$($tempDir.FullName)\page_%03d.jpg"
        $progressCurrent++
        Update-Progress -value $progressCurrent -max $progressTotal

        $images = Get-ChildItem -Path $tempDir -Filter *.jpg | Sort-Object Name
        $totalImages = $images.Count
        if ($totalImages -eq 0) {
            Update-Status "Keine Bilder gefunden. Verarbeitung abgebrochen."
            $buttonStart.Enabled = $true
            $textBoxSource.Enabled = $true
            $buttonBrowseSource.Enabled = $true
            $textBoxDestination.Enabled = $true
            $buttonBrowseDestination.Enabled = $true
            $comboBoxQuality.Enabled = $true
            return
        }

        # Gesamtzahl der Schritte aktualisieren
        $progressTotal += $totalImages + 1  # +1 für das Erstellen der finalen PDF
        Update-Progress -value $progressCurrent -max $progressTotal

        $currentImage = 0

        foreach ($image in $images) {
            $currentImage++
            $progressCurrent++
            $outputImagePath = Join-Path $tempDir $image.Name.Replace(".jpg", "_processed.jpg")

            Update-Status "Verarbeite Bild $currentImage von $totalImages..."
            magick "$($image.FullName)" `
                -resize "$($dpi * 8.27)x$($dpi * 11.69)" `
                -attenuate 0.01 +noise Gaussian `
                -quality $quality `
                -colorspace sRGB `
                -strip `
                "$outputImagePath"

            Update-Progress -value $progressCurrent -max $progressTotal
        }

        Update-Status "Erstelle finale PDF..."
        magick "$tempDir\*_processed.jpg" `
            -units PixelsPerInch `
            -density $dpi `
            -compress JPEG `
            -quality $quality `
            -page A4 `
            "$outputPdfPath"

        $progressCurrent++
        Update-Progress -value $progressCurrent -max $progressTotal

        Update-Status "Verarbeitung abgeschlossen."
        $buttonClose.Visible = $true
    }
    catch {
        Update-Status "Fehler: $_"
    }
    finally {
        if ($tempDir -and (Test-Path $tempDir)) {
            Remove-Item -Recurse -Force $tempDir
        }
        $buttonStart.Enabled = $true
        $buttonStart.Visible = $false
        # Eingabefelder wieder aktivieren
        $textBoxSource.Enabled = $true
        $buttonBrowseSource.Enabled = $true
        $textBoxDestination.Enabled = $true
        $buttonBrowseDestination.Enabled = $true
        $comboBoxQuality.Enabled = $true
    }
})

# Formular anzeigen
[void]$form.ShowDialog()
