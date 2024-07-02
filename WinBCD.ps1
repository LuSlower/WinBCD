# Check administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell "-File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Console {
    param (
        [Switch]$Show,
        [Switch]$Hide,
        [string]$Text
    )

    if (-not ("Console.Window" -as [type])) {
        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
    }

    if ($Show) {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        $null = [Console.Window]::ShowWindow($consolePtr, 5)
    }

    if ($Hide) {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        $null = [Console.Window]::ShowWindow($consolePtr, 0)
    }

    if ($Text) {
        $syncHash.TextBoxOutPut.AppendText("$Text`r`n")
    }
}

# Funcion para mostrar un dialogo desplegable
function Show-SelectionDialog {
    param (
        [string]$title,
        [string]$prompt,
        [array]$options
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(410, 290)
    $form.StartPosition = 'CenterScreen'
    $form.MaximizeBox = $false
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $prompt
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Size = New-Object System.Drawing.Size(360, 150)
    $listBox.Location = New-Object System.Drawing.Point(10, 50)
    $listBox.Items.AddRange($options)
    $form.Controls.Add($listBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(310, 220)
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(220, 220)
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    $form.Add_Shown({$form.Activate()})
    [void]$form.ShowDialog()

    if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK -and $listBox.SelectedItem -ne $null) {
        return $listBox.SelectedItem
    }

    return $null
}

# Funcion para refrescar los nombres de las particiones
function Refresh-Partitions {
    $comboBox.Items.Clear()
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 -or $_.DriveType -eq 5  -and $_.DeviceID -ne "$env:SystemDrive" }
    foreach ($drive in $drives) {
        $label = $drive.VolumeName
        if ([string]::IsNullOrWhiteSpace($label)) {
            $label = "Drive " + $drive.DeviceID
        }
        $comboBox.Items.Add("$($drive.DeviceID) - $label")
    }
    if ($comboBox.Items.Count -gt 0) {
        $comboBox.SelectedIndex = 0
    }
}


# Funcion para crear la particion a partir del nombre la iso
function Create-Partition {
    param (
        [string]$filePath
    )

    # Verificar si existe la particion
    $volumeName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $checkpartition = Get-WMIObject Win32_Volume | Where-Object { $_.Label -eq $volumeName}

    if ($checkpartition) {
        [System.Windows.Forms.MessageBox]::Show("The partition already exists", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # calcular el valor en GB, redondear para estar más seguro, verificar en que disco se encuentra
    $fileSizeGB = [math]::Ceiling((Get-Item $filePath).Length / 1GB)
    $systemDrive = $env:SystemDrive.TrimEnd(':')
    $disk = Get-Partition -DriveLetter $systemDrive | Get-Disk

    if ($disk -ne $null) {
        # reducir el tamaño de la particion principal,
        $systemPartition = Get-Partition -DriveLetter $systemDrive
        Resize-Partition -InputObject $systemPartition -Size ($systemPartition.Size - ($fileSizeGB * 3GB))

        # Crear la nueva particion
        $partition = New-Partition -DiskNumber $disk.Number -Size ($partitionSizeGB * 3GB) -AssignDriveLetter

        if ($partition -ne $null) {
            $driveLetter = $partition.DriveLetter

            # Determinar el sistema de archivos, obtener el nombre de la iso
            $fileSystem = if ($fileSizeGB -lt 4) { "FAT32" } else { "NTFS" }

            # formatear
            Format-Volume -DriveLetter $driveLetter -FileSystem $fileSystem -NewFileSystemLabel $volumeName

            [System.Windows.Forms.MessageBox]::Show("Partition created and formatted successfully", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            # refrescar
            Refresh-Partitions

        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to create partition", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Failed to find the system disk", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Crear form, ocultar consola
Console -Hide
[System.Windows.Forms.Application]::EnableVisualStyles();
$form = New-Object System.Windows.Forms.Form
$form.Text = "WinBCD"
$form.Size = New-Object System.Drawing.Size(330, 135)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.AllowDrop = $true
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# ComboBox para enumerar particiones
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 10)
$comboBox.Size = New-Object System.Drawing.Size(250, 20)
$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList


# TextBox para mostrar el nombre de la imagen
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 40)
$textBox.Size = New-Object System.Drawing.Size(250, 20)
$textBox.Enabled = $false

# OpenFileDialog para seleccionar la imagen 
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "ISO Files (*.iso)|*.iso|WIM Files (*.wim)|*.wim|ESD Files (*.esd)|*.esd"

# Boton para mostrar el OpenFileDialog
$buttonSelectFile = New-Object System.Windows.Forms.Button
$buttonSelectFile.Location = New-Object System.Drawing.Point(270, 40)
$buttonSelectFile.Size = New-Object System.Drawing.Size(25, 20)
$buttonSelectFile.Text = "..."
$buttonSelectFile.Add_Click({
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text = $openFileDialog.FileName
        $buttonCreatePartition.Enabled = $true
    }
})

# Boton para instalar la image en la particion seleccionada
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Location = New-Object System.Drawing.Point(10, 70)
$buttonInstall.Size = New-Object System.Drawing.Size(75, 23)
$buttonInstall.Text = "Install"
$buttonInstall.Add_Click({
    
    # Obtener la ruta de la imagen
    $filePath = $textBox.Text
    if ([string]::IsNullOrEmpty($filePath)) {
        [System.Windows.Forms.MessageBox]::Show("An error occurred while trying to mount the ISO.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Step 1: Confirmar montaje de la imagen
    $driveLetter = ($comboBox.SelectedItem.ToString() -split ' ')[0]
    $warningMessage = "Mounting the image ($filePath) will erase all files on drive $driveLetter. Do you want to continue?"
    $confirmation = [System.Windows.Forms.MessageBox]::Show($warningMessage, "Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Verificar si la imagen es .iso
        if ([System.IO.Path]::GetExtension($filePath) -eq ".iso") {
            try {
                Mount-DiskImage -ImagePath "$filePath" | Out-Null 
            } catch {
                Write-Host "An error occurred while trying to mount the ISO.`nError: $($_.Exception.Message)`n" -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show("An error occurred while trying to mount the ISO.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $mountedDrive = (Get-DiskImage -ImagePath "$filePath" | Get-Volume).DriveLetter

            $esdPath = "$mountedDrive`:\sources\install.esd"
            $wimPath = "$mountedDrive`:\sources\install.wim"

            if (Test-Path $esdPath) {
                $imgPath = $esdPath
            } elseif (Test-Path $wimPath) {
                $imgPath = $wimPath
            } else {
                [System.Windows.Forms.MessageBox]::Show("The selected ISO does not appear to be a Windows ISO. The ISO must contain either install.esd or install.wim in the 'sources' folder.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Dismount-DiskImage -ImagePath "$filePath"
                return
            }
        } else {
            $imgPath = $filePath
        }

        # Step 2: Seleccionar la edicion
        $windowsEditions = Get-WindowsImage -ImagePath $imgPath
        $editionsList = $windowsEditions | ForEach-Object { $_.ImageName }
        $selectedEdition = Show-SelectionDialog -title "Select a Windows edition" -prompt "Available editions:" -options $editionsList

        if (-not $selectedEdition -or $selectedEdition -eq "Cancel") {
            Write-Host "Edition selection canceled." -ForegroundColor Yellow
            Dismount-DiskImage -ImagePath "$filePath"
            return
        }

        $selectedIndex = $windowsEditions | Where-Object { $_.ImageName -eq $selectedEdition } | Select-Object -ExpandProperty ImageIndex
        
        # Step 3: Aplicar la imagen
        # Formatear, extraer imagen en la particion
        Write-Host "Applying image (installing/extracting Windows)..." -ForegroundColor Yellow
        Expand-WindowsImage -ImagePath "$imgPath" -ApplyPath "$driveLetter\" -Index $selectedIndex -LogLevel 2

        # Step 7: Aplicar autounattend por si existe
        if (Test-Path "$driveLetter\autounattend.xml") {
            Write-Host "Applying unattend answer files..." -ForegroundColor Yellow
            Use-WindowsUnattend -Path "$driveLetter\" -UnattendPath "$driveLetter\autounattend.xml"
            if (Test-Path "$mountedDrive\sources\$OEM$\$$\Setup\Scripts") {
                Copy-Item "$mountedDrive\sources\$OEM$\$$\Setup\Scripts\*" -Destination "$driveLetter\Windows\Setup\Scripts" -Recurse
            }
            Copy-Item "$mountedDrive\autounattend.xml*" -Destination "$driveLetter\Windows\Setup\Scripts" -Recurse
        }

        # Step 8: añadir la instalacion de windows al bootmgr
        Write-Host "Adding new Windows installation to boot loader..." -ForegroundColor Yellow
        bcdboot "$driveLetter\Windows" | Out-Null

        Write-Host "Installation completed!" -ForegroundColor Green
        $confirm = [System.Windows.Forms.MessageBox]::Show("Installation completed!`ndo you wish to reboot now?", "Info", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            Restart-Computer -Force
        }
        Dismount-DiskImage -ImagePath "$filePath"
    } else {
        Write-Host "Mounting canceled." -ForegroundColor Yellow
        return
    }
})


$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Location = New-Object System.Drawing.Point(100, 70)
$buttonRefresh.Size = New-Object System.Drawing.Size(75, 23)
$buttonRefresh.Text = "Refresh"
$buttonRefresh.Add_Click({
    Refresh-Partitions
})

$buttonCreatePartition = New-Object System.Windows.Forms.Button
$buttonCreatePartition.Location = New-Object System.Drawing.Point(190, 70)
$buttonCreatePartition.Size = New-Object System.Drawing.Size(120, 23)
$buttonCreatePartition.Text = "Create Partition"
$buttonCreatePartition.Enabled = $false
$buttonCreatePartition.Add_Click({
    Create-Partition -filePath $textBox.Text
})

# añadir Controles
$form.Controls.Add($comboBox)
$form.Controls.Add($textBox)
$form.Controls.Add($buttonSelectFile)
$form.Controls.Add($buttonInstall)
$form.Controls.Add($buttonRefresh)
$form.Controls.Add($buttonCreatePartition)

# Drag & Drop
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
    }
})

$form.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -eq 1) {
        $filePath = $files[0]
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        if ($extension -eq ".iso" -or $extension -eq ".wim" -or $extension -eq ".esd") {
            $textBox.Text = $filePath
            $buttonCreatePartition.Enabled = $true
        } else {
            [System.Windows.Forms.MessageBox]::Show("Unsupported file format.`nPlease drag and drop an ISO, WIM, or ESD file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})


Refresh-Partitions

[void]$form.ShowDialog()