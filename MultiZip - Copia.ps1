Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

### ===== FORM PRINCIPALE ===== ###
$form = New-Object System.Windows.Forms.Form
$form.Text = "Estrattore ZIP Avanzato"
$form.Size = New-Object System.Drawing.Size(650,500)
$form.StartPosition = "CenterScreen"

# Font generale
$font = New-Object System.Drawing.Font("Segoe UI",10)
$form.Font = $font

### ------ LABELS E INPUT ------ ###
# Sorgente
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Cartella sorgente:"
$lblSource.Location = "10,20"
$lblSource.AutoSize = $true
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = "150,18"
$txtSource.Size = "350,25"
$form.Controls.Add($txtSource)

$btnSource = New-Object System.Windows.Forms.Button
$btnSource.Text = "Sfoglia"
$btnSource.Location = "510,17"
$btnSource.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if($dialog.ShowDialog() -eq "OK"){ $txtSource.Text = $dialog.SelectedPath }
})
$form.Controls.Add($btnSource)

# Destinazione
$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Text = "Cartella destinazione:"
$lblDest.Location = "10,60"
$lblDest.AutoSize = $true
$form.Controls.Add($lblDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = "150,58"
$txtDest.Size = "350,25"
$form.Controls.Add($txtDest)

$btnDest = New-Object System.Windows.Forms.Button
$btnDest.Text = "Sfoglia"
$btnDest.Location = "510,57"
$btnDest.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if($dialog.ShowDialog() -eq "OK"){ $txtDest.Text = $dialog.SelectedPath }
})
$form.Controls.Add($btnDest)

# Apri destinazione
$btnApriDest = New-Object System.Windows.Forms.Button
$btnApriDest.Text = "Apri destinazione"
$btnApriDest.Location = "510,90"
$btnApriDest.Add_Click({
    if(Test-Path -LiteralPath $txtDest.Text){
        Start-Process $txtDest.Text
    }
})
$form.Controls.Add($btnApriDest)

### ------ OPZIONI ------ ###
$chkDelete = New-Object System.Windows.Forms.CheckBox
$chkDelete.Text = "Elimina ZIP dopo estrazione"
$chkDelete.Location = "10,100"
$form.Controls.Add($chkDelete)

$chkOverwrite = New-Object System.Windows.Forms.CheckBox
$chkOverwrite.Text = "Sovrascrivi file esistenti"
$chkOverwrite.Location = "10,130"
$form.Controls.Add($chkOverwrite)

$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Cerca ZIP anche nelle sottocartelle"
$chkRecursive.Location = "10,160"
$form.Controls.Add($chkRecursive)

$chkKeepStructure = New-Object System.Windows.Forms.CheckBox
$chkKeepStructure.Text = "Mantieni la struttura delle sottocartelle"
$chkKeepStructure.Location = "10,190"
$form.Controls.Add($chkKeepStructure)

### ------ LOG ------ ###
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log operazioni:"
$lblLog.Location = "10,230"
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = "10,260"
$txtLog.Size = "620,160"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

### ------ PROGRESS BAR ------ ###
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = "10,430"
$progress.Size = "620,20"
$form.Controls.Add($progress)

### ------ BOTTONE ESEGUI ------ ###
$btnExtract = New-Object System.Windows.Forms.Button
$btnExtract.Text = "ESTRAI ZIP"
$btnExtract.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
$btnExtract.Location = "250,455"
$btnExtract.Size = "150,35"

# Helper per normalizzare i percorsi
function Get-NormalPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    try {
        $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
        return $resolved.Path
    } catch {
        # Se non risolvibile (p.es. non esiste ancora), prova a costruire percorso assoluto
        try { return [System.IO.Path]::GetFullPath($Path) } catch { return $Path }
    }
}

$btnExtract.Add_Click({

    $source = Get-NormalPath $txtSource.Text
    $dest   = Get-NormalPath $txtDest.Text

    $txtLog.AppendText("=== Avvio estrazione ===`r`n")

    if([string]::IsNullOrWhiteSpace($source) -or -not (Test-Path -LiteralPath $source)){
        [System.Windows.Forms.MessageBox]::Show("La cartella sorgente non esiste o è vuota.","Errore")
        return
    }

    if([string]::IsNullOrWhiteSpace($dest)){
        [System.Windows.Forms.MessageBox]::Show("La cartella destinazione non è valida.","Errore")
        return
    }

    # Crea la destinazione principale (se non esiste)
    [void][System.IO.Directory]::CreateDirectory($dest)

    # Cerca ZIP
    if($chkRecursive.Checked){
        $zipFiles = Get-ChildItem -LiteralPath $source -Recurse -Filter *.zip -File
    } else {
        $zipFiles = Get-ChildItem -LiteralPath $source -Filter *.zip -File
    }

    if(-not $zipFiles -or $zipFiles.Count -eq 0){
        [System.Windows.Forms.MessageBox]::Show("Nessun file ZIP trovato.","Info")
        return
    }

    # Imposta progress bar a 0-100
    $progress.Minimum = 0
    $progress.Maximum = 100
    $progress.Value = 0

    foreach ($zip in $zipFiles) {

        $txtLog.AppendText("Apro ZIP: $($zip.FullName)`r`n")
        [System.Windows.Forms.Application]::DoEvents()

        # Determina la destinazione
        if($chkKeepStructure.Checked){
            $relativePath = $zip.DirectoryName.Substring($source.Length).TrimStart('\')
            $extractPath  = Join-Path $dest $relativePath
        } else {
            $extractPath  = $dest
        }

        # Normalizza/crea la cartella di estrazione
        [void][System.IO.Directory]::CreateDirectory($extractPath)

        # Risolvi il percorso per la Shell
        $extractPathResolved = Get-NormalPath $extractPath

        try {
            $shell = New-Object -ComObject Shell.Application
            $zipFolder  = $shell.NameSpace($zip.FullName)
            $destFolder = $shell.NameSpace($extractPathResolved)

            if ($zipFolder -eq $null){
                $txtLog.AppendText("ERRORE: Il file ZIP è corrotto o non accessibile.`r`n")
                continue
            }
            if ($destFolder -eq $null){
                $txtLog.AppendText("ERRORE: Impossibile aprire cartella di destinazione: $extractPathResolved`r`n")
                continue
            }

            $items = $zipFolder.Items()
            $totalEntries = $items.Count
            if ($totalEntries -eq 0) {
                $txtLog.AppendText("ZIP vuoto: $($zip.Name)`r`n")
                continue
            }

            $counter = 0

            foreach ($item in $items) {
                $counter++
                $percent = [math]::Floor(($counter / $totalEntries) * 100)
                $progress.Value = [Math]::Min($progress.Maximum, $percent)
                [System.Windows.Forms.Application]::DoEvents()

                $txtLog.AppendText("Estrazione: $($item.Name)`r`n")
                [System.Windows.Forms.Application]::DoEvents()

                # Opzioni CopyHere:
                # 0x10 = FOF_NOCONFIRMATION (no prompt)
                # 0x4  = FOF_SILENT (no UI)
                # 0x400= FOF_NOERRORUI (no error UI)
                $options = 0x10 + 0x4 + 0x400
                $destFolder.CopyHere($item, $options)

                Start-Sleep -Milliseconds 50
            }

            if ($chkDelete.Checked){
                Remove-Item -LiteralPath $zip.FullName -Force
                $txtLog.AppendText("ZIP eliminato.`r`n")
            }

        } catch {
            $txtLog.AppendText("ERRORE durante l'estrazione: $($_.Exception.Message)`r`n")
        }
    }

    # Assicura progress bar a 100% a fine processo
    $progress.Value = $progress.Maximum
    [System.Windows.Forms.Application]::DoEvents()

    [System.Windows.Forms.MessageBox]::Show("Estrazione completata!","Fatto")
    $txtLog.AppendText("=== Completato ===`r`n")
})

$form.Controls.Add($btnExtract)

### ------ MOSTRA FORM ------ ###
$form.ShowDialog()
