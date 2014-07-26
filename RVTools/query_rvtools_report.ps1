cls
###################################################################
# RVTools Queries
# Args usage = PSfile "path to RVTools csv exports"
###################################################################

# Check for args
	if (-not $args)
		{
		throw "Incorrect format. Supply the path to spreadsheet and path to resulting filename as arguments. Example: query_rvtools_report.ps1 `"path to RVTools csv exports`""
		exit		
		}
	
	$csvpath = $args[0]

# Import tabs from RVTools, unless it's already imported (debug)
	if (-not $tabvhost) {$tabvhost = Import-Csv -Path ($csvpath + "`\RVTools_tabvhost.csv")}
	if (-not $tabvpartition) {$tabvpartition = Import-Csv -Path ($csvpath + "`\RVTools_tabvPartition.csv")}
	if (-not $tabvdatastore) {$tabvdatastore = Import-Csv -Path ($csvpath + "`\RVTools_tabvDatastore.csv")}
	if (-not $tabvdisk) {$tabvdisk = Import-Csv -Path ($csvpath + "`\RVTools_tabvDisk.csv")}

###################################################################
# Host Queries

	# Group the clusters together by the cluster name
	$clusters = $tabvhost | Group {$_.Cluster}

	# Reporting setup
	$EndReport = @()

	# Break down into individual clusters
	foreach ($cluster in $clusters)
		{
		
		# Null out all of the counting vars, some just for when debugging code
		# Probably better if I stuck this in a function (blarg)
		$Report = {} | select DataCenter,ClusterName,HostQty,HostVendors,HostModels,Hypervisor,VMs,PCPU,LCPU,CPUUtil,RAM,RAMperHost,RAMUtil,Comments
		$VMs = $null
		$PCPU = $null
		$LCPU = $null
		$CPUUtil = $null
		$RAM = $null
		$RAMUtil = $null
		
		# Arrange the clusters into an array var because it makes things easier
		[array]$entries = $cluster.Group

		# Look at each host in the cluster and run various queries
		foreach ($entry in $entries)
			{			
			$PCPU += $entry.'# CPU' * $entry.'Cores per CPU'
			$LCPU += $entry.'# CPU' * $entry.'# Cores'
			$VMs += $entry.'# VMs'
			$CPUUtil += $entry.'CPU usage %'
			$RAM += $entry.'# Memory' / 1024
			$RAMUtil += $entry.'Memory usage %'
			}
			
		# Add values to the report
		$Report.DataCenter = $entry.Datacenter
		$Report.ClusterName = $cluster.Name 
		$Report.HostQty = $cluster.Count
		$Report.HostVendors = ($entries | ForEach-Object {$_.Vendor} | Select-Object -Unique) -join ','
		$Report.HostModels = ($entries | ForEach-Object {$_.Model} | Select-Object -Unique) -join ','
		$Report.Hypervisor = (($entries | ForEach-Object {$_.'ESX Version'} | Select-Object -Unique) -join ',') -replace 'VMware ',''
			If ($Report.HostVendors -match "," -or $Report.HostModels -match ",") {$Report.Comments += "Cluster contains mixed hardware. "}
			If ($Report.Hypervisor -match ",") {$Report.Comments += "Different versions of ESXi are present. "}
		$Report.VMs = $VMs
		$Report.PCPU = $PCPU
		$Report.LCPU = $LCPU
		$Report.CPUUtil = "{0:N2}" -f ($CPUUtil / $cluster.Count)
		$Report.RAM = "{0:N0}" -f $RAM
		[int]$Report.RAMperHost = "{0:N0}" -f ($RAM / $cluster.Count)
		$Report.RAMUtil = "{0:N2}" -f ($RAMUtil / $cluster.Count)
		If ($Report.RAMperHost -le 100) {$Report.Comments += "Hosts have a low quantity of RAM. "}
		$EndReport += $Report
	
		}

###################################################################
# Storage Queries

	# Ensure we have unique datastores based on the NAA
	[array]$datastores = $tabvdatastore | Sort-Object -Property "Address" -Unique
	
	# Find relationships between datastores and VMs
	foreach ($datastore in $datastores)
		{
		
		# Build a report
		$SReport = {} | select DataCenter,ClusterName,Capacity,Provisioned
		$disk = $null
		
		# Comb through the VMs
		foreach ($disk in $tabvdisk)
			{
			
			# Find VMs that live on the datastore and record the values
			If (($disk.Path).split('`[`]')[1] -match $datastore.Name)
				{
				
				}
			
			$SReport.DataCenter = $disk.Datacenter
			$SReport.ClusterName = $disk.Cluster			
			
			
			}
		
		}

###################################################################
# Final Report
	
	# Export results to CSV
	$EndReport = $EndReport | Sort-Object DataCenter,ClusterName
	$EndReport | Export-Csv -Path ($csvpath + "`\FinalReport.csv") -NoTypeInformation

