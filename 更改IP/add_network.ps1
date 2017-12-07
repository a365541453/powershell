param($vm_name)
$ConfirmPreference = 'None'

Add-PSSnapin VMware.DeployAutomation
Add-PSSnapin VMware.ImageBuilder
Add-PSSnapin VMware.VimAutomation.Core

connect-viserver -server 192.167.0.4 -user administrator@vsphere.local -password Win2008.cn
$snapshot = "$vm_name-snapshot"
$vm = get-vm -name $vm_name 
stop-vm -vm $vm
new-snapshot -vm $vm -name $snapshot
$networkname = (get-networkadapter -vm $vm).networkname
new-networkadapter -vm $vm -networkname $networkname.ToString() -StartConnected -WakeOnLan
start-vm -vm $vm
