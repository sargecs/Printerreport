# Powershell script check OFFLINE status of printers  
#
#################################################################################
# Created = Quick and dirty , does what it have to do
#
################################################################################# 
#Variables
$printserver = $env:computername
$printerdetails = Get-Printer | select -ExpandProperty name
$date = Get-Date -Format "dd/MM/yyyy"
$Path = "E:\Report\$printserver.csv"
#
#Script
foreach ( $printer in $printerdetails){

                $address = Get-PrinterPort | where { $_.Name -eq $printer } | select -expandproperty printerhostaddress
                $portname= Get-PrinterPort | where { $_.Name -eq $printer } |select -ExpandProperty name
                $Details = Get-Printer | where { $_.Name -eq $printer } | select DriverName , Location , Comment

   foreach ( $entry in $address){

                    $ok = Test-Connection -IPAddress $entry -Count 4 -quiet

if ($ok -eq $false){

	   [PSCustomObject]@{
       Printserver = $printserver
       Printername = (($printer) -join ",")
       Portname = $portname
       IP = $entry
       Status = "Offline"
       Reportdate = $Date
       Driver = $Details.drivername
       Location = $Details.location
       Comment = $Details.comment
       } | Export-Csv $Path -notype -Append }
       }
}