#Set-ExecutionPolicy RemoteSigned -force bei erlaubnis Problemen
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

# New array of "printer objects", rather than $strPrinter
$Printers = @(
    @{
        Name = "Update Powershell"
        SetCommand = {Update-Module -Name SpeculationControl -Force}
    }, 
    @{
        Name = "Admin Local"
        SetCommand = {
            $Password = ConvertTo-SecureString "Ad_Local_28Di" -AsPlainText -Force
            New-LocalUser "admin_local" -Password $Password -FullName "Administrator Local" -Description "Localer Administrator"
            Add-LocalGroupMember -Group "Administratoren" -Member "admin_local"
         }
    },
    @{ 
        Name = "Seveco_Installer"
        SetCommand = {Sevecoinst}
    }, 
    @{
        Name = "Seveco_Supportfiles_2013"
        SetCommand = {Start-Process \\srvospelthaas01\SEVECO_Client\SupportFiles\vcredist_x86_2013.exe}
    },
    @{
        Name = "Seveco_Supportfiles_2010"
        SetCommand = {Start-Process \\srvospelthaas01\SEVECO_Client\SupportFiles\vcredist_x86_2010.exe}
    },
    @{
        Name = "Seveco_Supportfiles_2008"
        SetCommand = {Start-Process \\srvospelthaas01\SEVECO_Client\SupportFiles\vcredist_x86_2008.exe}
    },
    @{
        Name = "AHK_Standart_WZ_Neu"
        SetCommand = {NeueAHK}
    } #Achten auf kein Komma
)

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Powershell Skriptrunner"
$objForm.Size = New-Object System.Drawing.Size(1200,800) 
$objForm.StartPosition = "CenterScreen"
$objForm.

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objListBox.SelectedItem;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

#Ok Button
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "Run"
$OKButton.Add_Click({
    # Grab the printer array index
    $index = $objListBox.SelectedIndex

    # Execute the appropriate command
    & $Printers[$index].SetCommand 

})
$objForm.Controls.Add($OKButton)

#Cancel Button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please select a Skript:"
$objForm.Controls.Add($objLabel) 

#List box showing printer options
$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(10,40) 
$objListBox.Size = New-Object System.Drawing.Size(360,20) 
$objListBox.Height = 80

foreach($Printer in $Printers){
    [void] $objListBox.Items.Add($Printer.Name)
}

function NeueAHK {

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(1200,800)
$Form.text                       = "AutoHotkey Vorlage WZ"
$Form.TopMost                    = $false

$Run                             = New-Object system.Windows.Forms.Button
$Run.text                        = "Run"
$Run.width                       = 200
$Run.height                      = 80
$Run.location                    = New-Object System.Drawing.Point(350,500)
$Run.Font                        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Cancel                          = New-Object system.Windows.Forms.Button
$Cancel.text                     = "Abbrechen"
$Cancel.width                    = 200
$Cancel.height                   = 80
$Cancel.location                 = New-Object System.Drawing.Point(650,500)
$Cancel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$WZName                          = New-Object system.Windows.Forms.TextBox
$WZName.multiline                = $false
$WZName.width                    = 300
$WZName.height                   = 20
$WZName.location                 = New-Object System.Drawing.Point(500,235)
$WZName.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',20)

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "WZ-Name:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(50,220)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',20)

$WZID                            = New-Object system.Windows.Forms.TextBox
$WZID.multiline                  = $false
$WZID.width                      = 300
$WZID.height                     = 20
$WZID.location                   = New-Object System.Drawing.Point(500,155)
$WZID.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif',20)

$Wdsfo                           = New-Object system.Windows.Forms.Label
$Wdsfo.text                      = "WZ-ID:"
$Wdsfo.AutoSize                  = $true
$Wdsfo.width                     = 25
$Wdsfo.height                    = 10
$Wdsfo.location                  = New-Object System.Drawing.Point(50,140)
$Wdsfo.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',20)

$Masch                           = New-Object system.Windows.Forms.ListBox
$Masch.text                      = "listBox"
$Masch.width                     = 300
$Masch.height                    = 70
$Masch.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',20)
$Masch.location                  = New-Object System.Drawing.Point(500,50)

$Maschine                        = New-Object system.Windows.Forms.Label
$Maschine.text                   = "Maschine:"
$Maschine.AutoSize               = $true
$Maschine.width                  = 25
$Maschine.height                 = 10
$Maschine.location               = New-Object System.Drawing.Point(50,40)
$Maschine.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',20)

$ErrorProvider1                  = New-Object system.Windows.Forms.ErrorProvider


$Form.controls.AddRange(@($Run,$WZName,$WZID,$Wdsfo,$Label1,$Cancel,$Masch,$Maschine))

$Run.Add_Click({ strart })
$Cancel.Add_Click({ Cancel })


function Cancel {$Form.Close()}
function strart {
    $neueName = $WZName.text
    $Codewort = "WZ-Nummer"
    $vorlagendatei = "\\srvospelthaas01\Allgemein\AutoHotKey\Vorlage\Vorlage.ahk"
    $neuedateiordner = "\\srvospelthaas01\Allgemein\AutoHotKey\Mill E 700U"
    $neuedatei = $neuedateiordner+ "\vorlage.ahk"
    Copy-Item -Path $vorlagendatei -Destination $neuedateiordner
    Get-Item -Path $neuedatei
    (Get-Content -path $neuedatei -raw) -replace 'WZ-Nummer',$WZID.text | Set-Content -Path $neuedatei
    Rename-Item -Path $neuedatei -NewName "$neueName.ahK" 
    }

$Maschiins = @(
   @{
        Name = "Mill P"
        SetCommand = {$neuedateiordner = "C:\Powershell Test\Neue Datei"}
    },
   @{
        Name = "Mill E"
        SetCommand = {$neuedateiordner = "C:\Powershell Test\Neue Datei"}
   }
)

foreach($Maschiin in $Maschiins){
    [void] $Masch.Items.Add($Maschiin.Name)
}


[void]$Form.ShowDialog()
} #V1

function Sevecoinst {
New-Item -Path C:\ -Name sevecodaten -ItemType Directory 
Copy-Item -Path "\\srvospelthaas01\SEVECO_Client\XceedSco.dll" -Destination "C:\sevecodaten" -Force
Copy-Item -Path "\\srvospelthaas01\SEVECO_Client\BarcodeC.dll" -Destination "C:\sevecodaten" -Force
Copy-Item -Path "\\srvospelthaas01\SEVECO_Client\BarcodeC.ocx" -Destination "C:\sevecodaten" -Force
Copy-Item -Path "\\srvospelthaas01\SEVECO_Client\XceedCry.dll" -Destination "C:\sevecodaten" -Force
Start-Process regsvr32 -ArgumentList "C:\sevecodaten\BarcodeC.ocx" 
Start-Process regsvr32 -ArgumentList "C:\sevecodaten\XceedCry.dll" 
Start-Process regsvr32 -ArgumentList "C:\sevecodaten\XceedSco.dll"

$linkedfile = "\\srvospelthaas01\SEVECO_Client\SEC.exe"
$ShortcutFile = "$env:USERPROFILE\Desktop\Seveco ERP.lnk"
 
if (!(Test-Path $ShortcutFile)) {
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.Description = "Outlook-Config"
    $shortcut.WorkingDirectory = "c:\temp"
    $Shortcut.TargetPath = $linkedfile
    $Shortcut.Save()
    }
} #V0

$objForm.Controls.Add($objListBox) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

Return