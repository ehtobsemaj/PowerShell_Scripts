cls
<###################################################
vSphere Update Utility for UCS Environments
###################################################>

### Variables and Setup

# UCS connection

$ucsdomain = $args[0]
if (-not $ucscreds) {$ucscreds = Get-Credential "UCS Credentials"}
Disconnect-Ucs
Connect-Ucs -Name $ucsdomain -Credential $ucscreds

# Ensure that UCS is healty for upgrade

$ucsstatus = Get-UcsStatus
if ($ucsstatus.HaReadiness -ne "ready" -or $ucsstatus.HaReady -ne "yes")
	{
	throw "UCS System is not in an HA state. Halting."	
	}

# vCenter connection

$vcenter = $args[1]
if (-not $vcentercreds) {$vcentercreds = Get-Credential "vCenter Credentials"}
Connect-VIServer -Server $vcenter -Credential $vcentercreds


### Put ESXi host into maintenance mode



### Shut down the ESXi host



### Validate host to blade relationship



### Set UCS maintenance policy to UserAck

foreach ($maintpolicy in Get-UcsMaintenancePolicy)
	{
	Set-UcsMaintenancePolicy -MaintenancePolicy $maintpolicy -UptimeDisr "user-ack" -Confirm:$false -Force
	}

### Send firmware package to UCS

<# Not ready yet
Send-UcsFirmware -LiteralPath C:\work\Images\ucs-k9-bundle-b-series.1.4.2b.B.bin | Watch-Ucs -Property TransferState -SuccessValue downloaded -PollSec 30 -TimeoutSec 600 
#>

### Begin firmware update of the blade

function global:testing {
# get the blade
Get-UcsBlade -Dn "sys/chassis-1/blade-1" | Set-UcsFirmware -Version 2.2 -Type cimc
$item = Get-UcsFirmwareUpdatable
}

$item = Get-UcsServiceProfile | where {$_.ConfigState -eq "applying"}
Get-UcsBlade -Dn $item.PnDn | Acknowledge-UcsBlade
$item = Get-UcsFault

### Validate blade is updated

### Power on the blade

### Validate ESXi host is online