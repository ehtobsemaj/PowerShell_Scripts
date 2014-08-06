cls


# UCS Cluster details
$args = $ucscluster
If (-not $ucscluster) {throw "Append the UCS cluster IP as the argument to this script."}

# UCS credentials and connection
if (-not $creds) {$creds = Get-Credential}
Disconnect-Ucs
Connect-Ucs -Credential $creds -Name $ucscluster

# Turn off the LEDs
function Off {

foreach ($led in Get-UcsLocatorLed)
	{
	Set-UcsLocatorLed -LocatorLed $led -AdminState off -Confirm:$false
	}
	
	}

# Turn on the LEDs
function On {

foreach ($led in Get-UcsLocatorLed)
	{
	Set-UcsLocatorLed -LocatorLed $led -AdminState on -Confirm:$false
	}
	
	}

