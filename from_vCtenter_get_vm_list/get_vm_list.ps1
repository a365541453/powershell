Import-Module VMware.DeployAutomation
Import-Module VMware.ImageBuilder
Import-Module VMware.VimAutomation.Core
$ConfirmPreference = 'None'

$vm_ip = '192.167.0.4'
$vm_username = 'administrator@vsphere.local'
$vm_password = 'Win2008.cn'

Connect-VIServer -Server $vm_ip -User $vm_username -Password $vm_password

$vm_list = Get-VMGuest * | Select-Object -ExpandProperty VmName,IPAddress

$excel = New-Object -ComObject Excel.Application
$workbooks = $excel.workbooks.add()
$sheet = $workbooks.activesheet
$cell = $sheet.cells
#$excel.Visible = True  让excel程序视图可见 

$row = 2
foreach($vm in $vm_list)
{
$hostname = $vm | Select-Object -ExpandProperty VmName
$vm_ip_type = ($vm | Select-Object -ExpandProperty IPAddress).gettype().name

if( $vm_ip_type -eq 'String' )
{ $vm_ip = ($vm | Select-Object -ExpandProperty IPAddress) }
elseif( $vm_ip_type -eq 'Object[]' )
{ $vm_ip = ($vm | Select-Object -ExpandProperty IPAddress)[0] }
else
{ $vm_ip = ' ' }
$hostip = $vm | Select-Object -ExpandProperty IPAddress


$cell.item($row,1) = $hostname
$cell.item($row,2) = $hostip
$row = $row + 1
}

Disconnect-VIServer -Server 192.167.0.4

$workbooks.saveas("c:\vm_list.xlsx")
$workbooks.close()
$excel.quit()
$excel = $null
[GC]::Collect()