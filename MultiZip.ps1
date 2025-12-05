Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

### ===== FORM PRINCIPALE ===== ###
$form = New-Object System.Windows.Forms.Form
$form.Text = "Estrattore ZIP Avanzato"
$form.Size = New-Object System.Drawing.Size(750,580)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)
$form.AutoScroll = $true

### ------ LABELS E INPUT ------ ###
# Sorgente
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Cartella sorgente:"
$lblSource.Location = "10,20"
$lblSource.AutoSize = $true
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = "150,18"
$txtSource.Size = "450,25"
$form.Controls.Add($txtSource)

$btnSource = New-Object System.Windows.Forms.Button
$btnSource.Text = "Sfoglia"
$btnSource.Location = "610,17"
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
$txtDest.Size = "450,25"
$form.Controls.Add($txtDest)

$btnDest = New-Object System.Windows.Forms.Button
$btnDest.Text = "Sfoglia"
$btnDest.Location = "610,57"
$btnDest.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if($dialog.ShowDialog() -eq "OK"){ $txtDest.Text = $dialog.SelectedPath }
})
$form.Controls.Add($btnDest)

# Apri destinazione
$btnApriDest = New-Object System.Windows.Forms.Button
$btnApriDest.Text = "Apri destinazione"
$btnApriDest.Location = "610,90"
$btnApriDest.Size = "100,25"
$btnApriDest.Add_Click({
    if(Test-Path $txtDest.Text){
        Start-Process $txtDest.Text
    }
})
$form.Controls.Add($btnApriDest)

### ------ OPZIONI ------ ###
$chkDelete = New-Object System.Windows.Forms.CheckBox
$chkDelete.Text = "Elimina ZIP dopo estrazione"
$chkDelete.Location = "10,100"
$chkDelete.AutoSize = $false
$chkDelete.Size = New-Object System.Drawing.Size(300,25)
$form.Controls.Add($chkDelete)

$chkOverwrite = New-Object System.Windows.Forms.CheckBox
$chkOverwrite.Text = "Sovrascrivi file esistenti"
$chkOverwrite.Location = "10,130"
$chkOverwrite.AutoSize = $false
$chkOverwrite.Size = New-Object System.Drawing.Size(300,25)
$form.Controls.Add($chkOverwrite)

$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Cerca ZIP anche nelle sottocartelle"
$chkRecursive.Location = "10,160"
$chkRecursive.AutoSize = $false
$chkRecursive.Size = New-Object System.Drawing.Size(350,25)
$form.Controls.Add($chkRecursive)

$chkKeepStructure = New-Object System.Windows.Forms.CheckBox
$chkKeepStructure.Text = "Mantieni la struttura delle sottocartelle"
$chkKeepStructure.Location = "10,190"
$chkKeepStructure.AutoSize = $false
$chkKeepStructure.Size = New-Object System.Drawing.Size(380,25)
$form.Controls.Add($chkKeepStructure)

### ------ LOG ------ ###
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log operazioni:"
$lblLog.Location = "10,230"
$lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = "10,260"
$txtLog.Size = "680,160"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

### ------ PROGRESS BAR ------ ###
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = "10,430"
$progress.Size = "680,20"
$form.Controls.Add($progress)

### ------ BOTTONE ESEGUI ------ ###
$btnExtract = New-Object System.Windows.Forms.Button
$btnExtract.Text = "ESTRAI ZIP"
$btnExtract.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
$btnExtract.Location = "280,455"
$btnExtract.Size = "150,35"

$btnExtract.Add_Click({

    $source = $txtSource.Text
    $dest = $txtDest.Text

    $txtLog.AppendText("=== Avvio estrazione ===`r`n")

    if(-not (Test-Path $source)){
        [System.Windows.Forms.MessageBox]::Show("La cartella sorgente non esiste!","Errore")
        return
    }

    if(-not (Test-Path $dest)){
        [System.Windows.Forms.MessageBox]::Show("La cartella destinazione non esiste!","Errore")
        return
    }

    # Cerca ZIP
    if($chkRecursive.Checked){
        $zipFiles = Get-ChildItem $source -Recurse -Filter *.zip -File
    } else {
        $zipFiles = Get-ChildItem $source -Filter *.zip -File
    }

    if($zipFiles.Count -eq 0){
        [System.Windows.Forms.MessageBox]::Show("Nessun file ZIP trovato.","Info")
        return
    }

    # Shell COM
    $shell = New-Object -ComObject Shell.Application

    # Calcola numero totale di file in tutti gli ZIP
    $totalItems = 0
    foreach ($zip in $zipFiles) {
        $zipFolder = $shell.NameSpace($zip.FullName)
        if ($zipFolder -ne $null) {
            $totalItems += $zipFolder.Items().Count
        }
    }

    if ($totalItems -eq 0){
        [System.Windows.Forms.MessageBox]::Show("Nessun file da estrarre.","Info")
        return
    }

    # Imposta progress bar globale
    $progress.Minimum = 0
    $progress.Maximum = $totalItems
    $progress.Value = 0

    # Determina flag per CopyHere
    # $copyFlags = 0x10   # non mostra UI
	$copyFlags = 0x0
    if ($chkOverwrite.Checked) {
        # $copyFlags = $copyFlags -bor 0x4  # aggiunge sovrascrittura automatica
		$copyFlags = 0x10   # non mostra UI
		$copyFlags = $copyFlags -bor 0x4  # aggiunge sovrascrittura automatica
    }

    foreach ($zip in $zipFiles) {

        $txtLog.AppendText("Apro ZIP: $($zip.FullName)`r`n")
        [System.Windows.Forms.Application]::DoEvents()

        if($chkKeepStructure.Checked){
            $relativePath = $zip.DirectoryName.Substring($source.Length).TrimStart("\")
            $extractPath = Join-Path $dest $relativePath
        } else {
            $extractPath = $dest
        }

        if (!(Test-Path $extractPath)){
            New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        }

        try {
            $zipFolder  = $shell.NameSpace($zip.FullName)
            $destFolder = $shell.NameSpace($extractPath)

            if ($zipFolder -eq $null){
                $txtLog.AppendText("ERRORE: Il file ZIP è corrotto o non accessibile.`r`n")
                continue
            }

            foreach ($item in $zipFolder.Items()) {

                # Estrazione con flag
                $destFolder.CopyHere($item, $copyFlags)

                # Aggiorna progress bar globale
                $progress.Value++
                [System.Windows.Forms.Application]::DoEvents()

                # Log
                $txtLog.AppendText("Estrazione: $($item.Name)`r`n")
                [System.Windows.Forms.Application]::DoEvents()

                Start-Sleep -Milliseconds 50
            }

            if ($chkDelete.Checked){
                Remove-Item $zip.FullName -Force
                $txtLog.AppendText("ZIP eliminato.`r`n")
            }

        } catch {
            $txtLog.AppendText("ERRORE durante l'estrazione: $($_.Exception.Message)`r`n")
        }
    }

    $progress.Value = $progress.Maximum
    [System.Windows.Forms.Application]::DoEvents()

    [System.Windows.Forms.MessageBox]::Show("Estrazione completata!","Fatto")
    $txtLog.AppendText("=== Completato ===`r`n")
})



$form.Controls.Add($btnExtract)

### ------ MOSTRA FORM ------ ###
$form.ShowDialog()
