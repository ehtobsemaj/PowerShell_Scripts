cls

<##########################################
Toggles the light on a UCS domain :)
args = script "cluster name" "on or off"
##########################################>

# UCS Cluster details
[int]$i = 0
$ucscluster = $args[0]
$toggle = $args[1]
If (-not $ucscluster) {throw "Append the UCS cluster IP as the argument to this script."}

# UCS credentials and connection
if (-not $creds) {$creds = Get-Credential}
Disconnect-Ucs
Connect-Ucs -Credential $creds -Name $ucscluster

# Turn off the LEDs
function global:Off {

foreach ($led in Get-UcsLocatorLed)
	{
	Set-UcsLocatorLed -LocatorLed $led -AdminState "off" -Confirm:$false -ErrorAction SilentlyContinue -Force
	}
	
	}

# Turn on the LEDs
function global:On {

foreach ($led in Get-UcsLocatorLed)
	{
	Set-UcsLocatorLed -LocatorLed $led -AdminState "on" -Confirm:$false -ErrorAction SilentlyContinue -Force
	}
	
	}

If ($toggle -match "on") {on}
	else {off}