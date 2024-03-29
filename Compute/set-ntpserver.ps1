cls
###################################################################
# Update ESXi NTP addresses
# Args usage: script.ps1 vcenter.fqdn "desired ntp server"
###################################################################

# Variables
$vcenter = $args[0]
$ntpserver = $args[1]

# Connect to vCenter and gather data
If (-not $global:DefaultVIServer) {Connect-VIServer -Server $vcenter -Credential (Get-Credential)}
$vmhosts = Get-VMHost

# Update NTP server info
$i = 1
foreach ($server in $vmhosts)
	{
	# Everyone loves progress bars, so here is a progress bar
	Write-Progress -Activity "Configuring NTP Settings" -Status $server -PercentComplete (($i / $vmhosts.Count) * 100)
	
	# Determine existing ntp config
	$ntpold = $server | Get-VMHostNtpServer
	
	# Check to see if an NTP entry exists; if so, delete the value(s)
	If ($ntpold) {$server | Remove-VMHostNtpServer -NtpServer $ntpold -Confirm:$false}
	
	# Add desired NTP value to the host
	Add-VmHostNtpServer -VMHost $server -NtpServer $ntpserver | Out-Null
	
	# Output to console (optional)
	Write-Host "Host $server is now using NTP server" (Get-VMHostNtpServer -VMHost $server)
	
	$i++
	}