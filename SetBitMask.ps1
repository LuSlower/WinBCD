# Check administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell "-File `"$PSCommandPath`"" -Verb RunAs
    exit
}

function Console
{
    param ([Switch]$Show,[Switch]$Hide)
    if (-not ("Console.Window" -as [type])) { 

        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
    }

    if ($Show)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()

        $null = [Console.Window]::ShowWindow($consolePtr, 5)
    }

    if ($Hide)
    {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        $null = [Console.Window]::ShowWindow($consolePtr, 0)
    }
}

Add-Type -AssemblyName System.Windows.Forms

function Convert-HexToDec ($hex) {
    try {
        if ([string]::IsNullOrEmpty($hex)) {
            throw "Input string is empty or null."
        }

        # Convertir el hex a un número decimal
        $dec = [Convert]::ToInt64($hex, 16)

        # Devolver el valor como un string para mantener los ceros a la izquierda
        return $dec
    } catch {
        return
    }
}

function Convert-DecToHex ($dec) {
    try {
        if ([string]::IsNullOrEmpty($dec)) {
            throw "Input string is empty or null."
        }

        # Convertir el decimal a hexadecimal
        $hex = "{0:X}" -f [int64]$dec
        return $hex
    } catch {
        return
    }
}

function Convert-HexToBin ($hex) {
    try {
        if ([string]::IsNullOrEmpty($hex)) {
            throw "Input string is empty or null."
        }

        # Convertir el hex a binario
        $bin = [Convert]::ToString([Convert]::ToInt64($hex, 16), 2)

        if ($global:lzeros) {
            # Calcular la longitud esperada del binario (4 bits por cada dígito hexadecimal)
            $desiredLength = $hex.Length * 4

            # Asegurar que el binario tenga la longitud esperada
            $bin = $bin.PadLeft($desiredLength, '0')
        }

        return $bin
    } catch {
        return
    }
}


function Convert-BinToDec ($bin) {
    try {
        if ([string]::IsNullOrEmpty($bin)) {
            throw "Input string is empty or null."
        }

        # Convertir el binario a decimal
        return [Convert]::ToInt64($bin, 2)
    } catch {
        return
    }
}

# función para actualizar valores
function Update-Values {
    param (
        [string]$source,
        [string]$value
    )

    if ($source -eq 'Hex') {
        $dec = Convert-HexToDec $value
        if ($dec -ne "") {
            $textBoxDec.Text = $dec
            $textBoxBinary.Text = Convert-HexToBin $value
            Update-BitPairs -value $textBoxBinary.Text
        }
    } elseif ($source -eq 'Dec') {
        $hex = Convert-DecToHex $value
        if ($hex -ne "") {
            $textBoxHex.Text = $hex
            $textBoxBinary.Text = Convert-HexToBin $hex
            Update-BitPairs -value $textBoxBinary.Text
        }
    } elseif ($source -eq 'Binary') {
        $dec = Convert-BinToDec $value
        if ($dec -ne "") {
            $textBoxDec.Text = $dec
            $hex = Convert-DecToHex $dec
            $textBoxHex.Text = $hex
            Update-BitPairs -value $value
        }
    }
}

function Update-BitPairs {
    param (
        [string]$value
    )

    # Guardar el índice del TextBox actualmente enfocado
    $focusedTextBox = $global:textBoxPairs | Where-Object { $_.Focused }
    $focusedIndex = $global:textBoxPairs.IndexOf($focusedTextBox)

    # Dividir el valor binario en pares de bits y actualizar los TextBoxes correspondientes
    for ($i = 0; $i -lt $global:textBoxPairs.Count; $i++) {
        if ($i * 2 -lt $value.Length) {
            # Verificar si hay suficientes caracteres para extraer una subcadena de longitud 2
            if ($i * 2 + 2 -le $value.Length) {
                $global:textBoxPairs[$i].Text = $value.Substring($i * 2, 2)
            } else {
                # Si no hay suficientes caracteres, establecer el TextBox en blanco o manejar según lo necesario
                $global:textBoxPairs[$i].Text = ""
            }
        } else {
            # Si el índice es mayor o igual que la longitud del valor binario, establecer en blanco o manejar según lo necesario
            $global:textBoxPairs[$i].Text = ""
        }
    }

    # Restaurar el foco al TextBox que estaba enfocado anteriormente
    if ($focusedIndex -ne -1) {
        $global:textBoxPairs[$focusedIndex].Focus()
    }
}

# función para crear y mostrar pares de bits y sumas
function Create-BitGroups {
    param (
        [int]$numGroups,
        [System.Windows.Forms.Form]$form
    )

    $left = 25
    $top = 30
    $bitWidth = 50
    $textBoxWidth = 20 
    $textBoxHeight = 20

    for ($i = 0; $i -lt $numGroups; $i++) {
        # Calcular posiciones para cada par de bits
        $labelIndex1 = $i * 2
        $labelIndex2 = $labelIndex1 + 1

        # crear pares de bits
        $labelPair = New-Object System.Windows.Forms.Label
        $labelPair.Text = "${labelIndex1}:${labelIndex2}"
        $labelPair.AutoSize = $true
        $labelPair.Location = New-Object System.Drawing.Point($left, $top)
        $form.Controls.Add($labelPair)

        # Crear TextBox debajo de la etiqueta
        $textBoxPair = New-Object System.Windows.Forms.TextBox
        $textBoxPair.Size = New-Object System.Drawing.Size($textBoxWidth, $textBoxHeight)
        $textBoxPair.Location = New-Object System.Drawing.Point(($left + 1), ($top + 20))  # Espacio entre la etiqueta y el TextBox
        $textBoxPair.MaxLength = 2
        $textBoxPair.Name = "TextBoxPair_$i"

        # Agregar el evento TextChanged
        $textBoxPair.Add_TextChanged({
            # Actualizar $binaryValue con los valores actuales de todos los TextBox
            $binaryValue = ""
            foreach ($pair in $global:textBoxPairs) {
                $binaryValue += $pair.Text
            }

            if ($this.Focused) {
                # Actualizar los TextBox de Binario, Decimal y Hexadecimal
                $textBoxBinary.Text = $binaryValue
                $textBoxDec.Text = Convert-BinToDec -bin $binaryValue
                $textBoxHex.Text = Convert-DecToHex -dec $textBoxDec.Text
            }
            
            # Mover el foco al siguiente TextBox si el actual tiene un par de caracteres y está enfocado
            if ($this.Text.Length -eq 2 -and $this.Focused) {
                $actIndex = $global:textBoxPairs.IndexOf($this)
                if ($actIndex -lt ($global:textBoxPairs.Count - 1)) {
                    $nextIndex = $actIndex + 1
                    $global:textBoxPairs[$nextIndex].Focus()
                }
            }
        })

        $textBoxPair.Add_KeyDown({
            param($sender, $e)
            # Verificar si se presionó backspace
            if ($e.KeyCode -eq "Back") {
                # Manejar la tecla de retroceso para borrar y mover el foco al TextBox anterior si es el primer carácter
                if ($sender.Text.Length -eq 0) {
                    $prevIndex = [math]::Max($global:textBoxPairs.IndexOf($sender) - 1, 0)
                    $global:textBoxPairs[$prevIndex].Focus()
                }

                # Actualizar $binaryValue con los valores actuales de todos los TextBox
                $binaryValue = ""
                foreach ($pair in $global:textBoxPairs) {
                    $binaryValue += $pair.Text
                }

                # Actualizar los TextBox de Binario, Decimal y Hexadecimal
                $textBoxBinary.Text = $binaryValue
                $textBoxDec.Text = Convert-BinToDec -bin $binaryValue
                $textBoxHex.Text = Convert-DecToHex -dec $textBoxDec.Text
            }
        })

        $form.Controls.Add($textBoxPair)
        $global:textBoxPairs.Add($textBoxPair)

        # Ajustar posición izquierda para el próximo par de bits
        $left += $bitWidth

        # Saltar a la siguiente línea después de 8 pares de bits
        if (($i + 1) % 8 -eq 0) {
            $left = 25
            $top += 70
        }
    }
}

# lista global BitPairs
$global:textBoxPairs = New-Object System.Collections.Generic.List[System.Windows.Forms.TextBox]

# crear form
Console -Hide
[System.Windows.Forms.Application]::EnableVisualStyles();
$form = New-Object System.Windows.Forms.Form
$form.Text = "SetBitmask"
$form.Size = New-Object System.Drawing.Size(435, 485)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# crear los pares de bits
Create-BitGroups -numGroups 32 -form $form

# leer ceros a la izquierda
$global:lzeros = $false
$chkbits = New-Object System.Windows.Forms.CheckBox
$chkbits.Text = "Leading Zeros Hex"
$chkbits.AutoSize = $true
$chkbits.Location = New-Object System.Drawing.Point(20, 300)
$chkbits.Add_CheckedChanged({
    if ($chkbits.Checked) {
        $global:lzeros = $true
    } else {
        $global:lzeros = $false
    }

    Update-Values -source 'Hex' -value $textBoxHex.Text
})
$form.Controls.Add($chkbits)

# valor hex
$labelHex = New-Object System.Windows.Forms.Label
$labelHex.Text = "Hexadecimal"
$labelHex.Size = New-Object System.Drawing.Size(70, 15)
$labelHex.Location = New-Object System.Drawing.Point(10, 340)
$form.Controls.Add($labelHex)

$textBoxHex = New-Object System.Windows.Forms.TextBox
$textBoxHex.Size = New-Object System.Drawing.Size(190, 10)
$textBoxHex.Location = New-Object System.Drawing.Point(10, 360)
$textBoxHex.Add_KeyPress({
    param($sender, $e)
    if ((($e.KeyChar -lt '0' -or $e.KeyChar -gt '9') -and
         ($e.KeyChar -lt 'A' -or $e.KeyChar -gt 'F') -and
         ($e.KeyChar -lt 'a' -or $e.KeyChar -gt 'f')) -and
         $e.KeyChar -ne [char][System.Windows.Forms.Keys]::Back -and
         $e.KeyChar -ne [System.Windows.Forms.Keys]::Right -and
         $e.KeyChar -ne [System.Windows.Forms.Keys]::Left) {
        $e.Handled = $true
    }
})
$textBoxHex.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        # Ctrl+C
        $textBoxHex.Copy()
    }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        # Ctrl+V
        $textBoxHex.Paste()
        Update-Values -source 'Hex' -value $textBoxHex.Text
        $e.Handled = $true
    }
})
$textBoxHex.Add_TextChanged({
    if ($textBoxHex.Focused) {
        Update-Values -source 'Hex' -value $textBoxHex.Text
    }
})
$form.Controls.Add($textBoxHex)

# valor dec
$labelDec = New-Object System.Windows.Forms.Label
$labelDec.Text = "Decimal"
$labelDec.Size = New-Object System.Drawing.Size(70, 15)
$labelDec.Location = New-Object System.Drawing.Point(215, 340)
$form.Controls.Add($labelDec)

$textBoxDec = New-Object System.Windows.Forms.TextBox
$textBoxDec.Size = New-Object System.Drawing.Size(190, 10)
$textBoxDec.Location = New-Object System.Drawing.Point(215, 360)
$textBoxDec.Add_KeyPress({
    param($sender, $e)
    if ($e.KeyChar -lt '0' -or $e.KeyChar -gt '9' -and
    $e.KeyChar -ne [char][System.Windows.Forms.Keys]::Back -and
    $e.KeyChar -ne [System.Windows.Forms.Keys]::Right -and
    $e.KeyChar -ne [System.Windows.Forms.Keys]::Left) {
        $e.Handled = $true
    }
})
$textBoxDec.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        # Ctrl+C
        $textBoxDec.Copy()
    }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        # Ctrl+V
        $textBoxDec.Paste()
        Update-Values -source 'Dec' -value $textBoxDec.Text
        $e.Handled = $true
    }
})
$textBoxDec.Add_TextChanged({
    if ($textBoxDec.Focused) {
        Update-Values -source 'Dec' -value $textBoxDec.Text
    }
})
$form.Controls.Add($textBoxDec)

# valor binario
$labelBitmask = New-Object System.Windows.Forms.Label
$labelBitmask.Text = "Binary"
$labelBitmask.Size = New-Object System.Drawing.Size(70, 15)
$labelBitmask.Location = New-Object System.Drawing.Point(10, 400)
$form.Controls.Add($labelBitmask)

$textBoxBinary = New-Object System.Windows.Forms.TextBox
$textBoxBinary.Size = New-Object System.Drawing.Size(395, 10)
$textBoxBinary.Location = New-Object System.Drawing.Point(10, 420)
$textBoxBinary.MaxLength = 64 #64bits
$textBoxBinary.Add_KeyDown({
    param($sender, $e)
    if (-not ($e.KeyCode -eq [System.Windows.Forms.Keys]::D0 -or
    $e.KeyCode -eq [System.Windows.Forms.Keys]::D1 -or
    $e.KeyCode -eq [System.Windows.Forms.Keys]::NumPad0 -or
    $e.KeyCode -eq [System.Windows.Forms.Keys]::NumPad1 -or
    $e.KeyCode -eq [System.Windows.Forms.Keys]::Back -or 
    $e.KeyCode -eq [System.Windows.Forms.Keys]::Right -or
    $e.KeyCode -eq [System.Windows.Forms.Keys]::Left)) {
        $e.SuppressKeyPress = $true
    }
})
$textBoxBinary.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        # Ctrl+C
        $textBoxBinary.Copy()
    }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        # Ctrl+V
        $textBoxBinary.Paste()
        Update-Values -source 'Binary' -value $textBoxBinary.Text
        $e.Handled = $true
    }
})
$textBoxBinary.Add_TextChanged({
    if ($textBoxBinary.Focused) {
        Update-Values -source 'Binary' -value $textBoxBinary.Text
    }
})
$form.Controls.Add($textBoxBinary)

# iniciar form
$form.ShowDialog()
