cls
<#
RVTools Queries
Args usage = PSfile "VDS name to configure" "Mbps limit value"
#>

# Connect to vCenter
If (-not $global:DefaultVIServer) {Connect-VIServer -Server $vcenter}

# Check args
	if (-not $args[1])
		{
		throw "Incorrect format. Provide both the VDS name and Mbps for the replication limit. Example: `"script.ps1 `"VDS1`" 500`" would set VR bandwidth on VDS1 to 500 Mbps"
		exit		
		}

# Variables
$dvsName = $args[0]
$rpName = "vSphere Replication"
# Set to -1 to go back to unlimited
$newLimit = $args[1]

# Get the VDS details
$dvs = Get-VDSwitch -Name $dvsName
 
# Set the VR network pool to the value provided in args
# The section below was written by Luc Dekens (@LucD22), all credit to him
$rp = $dvs.ExtensionData.NetworkResourcePool | Where {$_.Name -match $rpName}
if($rp){
    $spec = New-Object VMware.Vim.DVSNetworkResourcePoolConfigSpec
    $spec.AllocationInfo = $rp.AllocationInfo
    $spec.AllocationInfo.Limit = [long]$newLimit
    $spec.ConfigVersion = $rp.ConfigVersion
    $spec.Key = $rp.Key
    $dvs.ExtensionData.UpdateNetworkResourcePool(@($spec))
}

# Verify
$rp = $dvs.ExtensionData.NetworkResourcePool | Where {$_.Name -match $rpName}
Write-Host "A limit of" $rp.AllocationInfo.Limit "Mbps has been set on $vds"