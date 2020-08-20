$NumOperateur="01"
$myDocFolderPath=[Environment]::GetFolderPath("MyDocuments")
$myFactFolderPath="$myDocFolderPath\Factures\"
$myProgramDataPath=$env:ProgramData
$myResourcesFolderPath="$myProgramDataPath\Scripts\Resources"
$ext=".xltx"
$Template="$myResourcesFolderPath\Template$ext"
$Date=Get-Date -UFormat "%d %B %Y"
$FactureFolderPath="$myFactFolderPath\$Date"
New-Item -ItemType Directory -Force -Path $FactureFolderPath
$factures_count = [System.IO.Directory]::GetFiles("$FactureFolderPath", "*"+"$ext").Count
$NF=$factures_count + 1
$NumFacture="{0:d4}" -f $NF
$NewFactName="$NumOperateur $NumFacture$ext"

$NewFacturePatch="$FactureFolderPath\$NewFactName"

Copy-Item $Template -Recurse $NewFacturePatch

$excel=new-object -comobject Excel.Application

$excel.visible=$true

$excel.DisplayAlerts=$False

$workbook=$excel.Workbooks.open($NewFacturePatch)

$diskSpacewksht= $workbook.Worksheets.Item(1)

$diskSpacewksht.Cells.Item(5,5)="[" + $NumFacture + "]"

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") 
[reflection.assembly]::loadwithpartialname("System.Drawing")
$path = Get-Process -id $pid | Select-Object -ExpandProperty Path            		
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)    		
$notify = new-object system.windows.forms.notifyicon
$notify.icon = $icon
$notify.visible = $true
$Title = "Nouvelle Facture"
$message = "Nouvelle Facture créer dans :           $Date\$Nomdufichier                       Vous n'avez plus qu'a la completé ;-)"
$notify.showballoontip(10,$title,$Message, [system.windows.forms.tooltipicon]::info)

