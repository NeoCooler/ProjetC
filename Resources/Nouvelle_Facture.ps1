$NumOperateur="01"

$myDocFolder=[Environment]::GetFolderPath("MyDocuments")

$Dossier_Facture=($myDocFolder + "\ProjetC\Factures\")

$ext = ".xltx"

$Template=$myDocFolder + "\ProjetC\Resources\Template$ext"

$Date=Get-Date -UFormat "%d %B %Y"

New-Item -ItemType Directory -Force -Path $Dossier_Facture\$Date

$file_count = [System.IO.Directory]::GetFiles("$Dossier_Facture\$Date", "*"+"$ext").Count

$NF=$file_count + 1

$NumeroFacture="{0:d4}" -f $NF

$Nomdufichier=($NumOperateur + "_" + $NumeroFacture + $ext)

$mynewdocument=$Dossier_Facture+"\"+$Date+"\"+$Nomdufichier

Copy-Item $Template -Recurse $mynewdocument

$excel=new-object -comobject Excel.Application

$excel.visible=$true

$excel.DisplayAlerts=$False

$workbook=$excel.Workbooks.open($mynewdocument)

$diskSpacewksht= $workbook.Worksheets.Item(1)

$diskSpacewksht.Cells.Item(5,5)="[" + $NumeroFacture + "]"

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

exit