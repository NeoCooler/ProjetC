$ShortcutObject = New-Object -comObject WScript.Shell
$Shortcut = $ShortcutObject.CreateShortcut($Link)
$Shortcut.TargetPath ="C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -windowstyle hidden -noprofile -File $Scripts"
$Shortcut.Save()

$myDesktopPath = [Environment]::GetFolderPath("Desktop")
$myDocFolderPath=[Environment]::GetFolderPath("MyDocuments")
$myProgramDataPath=$env:ProgramData
$myCurrentFolderPath=Split-Path -parent $MyInvocation.MyCommand.Definition

$myScriptsFolderPath="$myProgramDataPath\Scripts\"
$myFacturesFolderPath="$myDocFolderPath\Factures\"
$myResourcesFolderPath="$myProgramDataPath\Scripts\Resources\"

$Source_Icone="$myCurrentFolderPath\Resources\Icone.ico"
$Source_Scripts="$myCurrentFolderPath\Resources\NewFacture.ps1"
$Source_Template="$myCurrentFolderPath\Resources\Template.xltx"

$Icone="$myResourcesFolderPath\Icone.ico"
$Link="$myDesktopPath\Nouvelle_Facture.lnk"
$Script="$myScriptsFolderPath\NewFacture.ps1"
$Template="$myResourcesFolderPath\Template.xltx"

New-Item -ItemType Directory -Force -Path "$myScriptsFolderPath"
New-Item -ItemType Directory -Force -Path "$myFacturesFolderPath"
New-Item -ItemType Directory -Force -Path "$myResourcesFolderPath"

Copy-Item $Source_Icone -Recurse $Icone
Copy-Item $Source_Scripts -Recurse $Scripts
Copy-Item $Source_Template -Recurse $Template

$ShortcutObject=New-Object -comObject WScript.Shell
$Shortcut=$ShortcutObject.CreateShortcut($Link)
$Shortcut.TargetPath="powershell.exe"
$Shortcut.Arguments="-windowstyle hidden -noprofile -File $Script"
$Shortcut.IconLocation=$Icone
$Shortcut.Save()

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") 
[reflection.assembly]::loadwithpartialname("System.Drawing")
$path = Get-Process -id $pid | Select-Object -ExpandProperty Path            		
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)    		
$notify = new-object system.windows.forms.notifyicon
$notify.icon = $icon
$notify.visible = $true
$Title = "Nouvelle Facture"
$message = "Vous pouvez vous servir dés a présent du system automatique de création de facture. Le lien ce trouve sur le bureau. Enjoy ;-)"
$notify.showballoontip(10,$title,$Message, [system.windows.forms.tooltipicon]::info)