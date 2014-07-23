<#

###################################################################
# RVTools Queries
###################################################################

# Excel import function

function Import-Excel
{
  param (
    [string]$FileName,
    [string]$WorksheetName,
    [bool]$DisplayProgress = $true
  )

  if ($FileName -eq "") {
    throw "Please provide path to the Excel file"
    Exit
  }

  if (-not (Test-Path $FileName)) {
    throw "Path '$FileName' does not exist."
    exit
  }

  $FileName = Resolve-Path $FileName
  $excel = New-Object -com "Excel.Application"
  $excel.Visible = $false
  $workbook = $excel.workbooks.open($FileName)

  if (-not $WorksheetName) {
    Write-Warning "Defaulting to the first worksheet in workbook."
    $sheet = $workbook.ActiveSheet
  } else {
    $sheet = $workbook.Sheets.Item($WorksheetName)
  }
  
  if (-not $sheet)
  {
    throw "Unable to open worksheet $WorksheetName"
    exit
  }
  
  $sheetName = $sheet.Name
  $columns = $sheet.UsedRange.Columns.Count
  $lines = $sheet.UsedRange.Rows.Count
  
  Write-Warning "Worksheet $sheetName contains $columns columns and $lines lines of data"
  
  $fields = @()
  
  for ($column = 1; $column -le $columns; $column ++) {
    $fieldName = $sheet.Cells.Item.Invoke(1, $column).Value2
    if ($fieldName -eq $null) {
      $fieldName = "Column" + $column.ToString()
    }
    $fields += $fieldName
  }
  
  $line = 2
  
  
  for ($line = 2; $line -le $lines; $line ++) {
    $values = New-Object object[] $columns
    for ($column = 1; $column -le $columns; $column++) {
      $values[$column - 1] = $sheet.Cells.Item.Invoke($line, $column).Value2
    }  
  
    $row = New-Object psobject
    $fields | foreach-object -begin {$i = 0} -process {
      $row | Add-Member -MemberType noteproperty -Name $fields[$i] -Value $values[$i]; $i++
    }
    $row
    $percents = [math]::round((($line/$lines) * 100), 0)
    if ($DisplayProgress) {
      Write-Progress -Activity:"Importing from Excel file $FileName" -Status:"Imported $line of total $lines lines ($percents%)" -PercentComplete:$percents
    }
  }
  $workbook.Close()
  $excel.Quit()
}

# Location of the spreadsheet
#	$xlspath = $args[0]
	$xlspath = "C:\Dropbox\Projects - Ahead\UW Health\RVTools-AOB-1SPark.xls"  # Uncomment for dev work
	
	$wkstname = "tabvHost"

# Import tabs from RVTools
	$tabvhost = import-excel -FileName $xlspath -WorksheetName $wkstname
#>
###################################################################
# HOST Queries

cls

# Group the clusters together by the cluster name
	$clusters = $tabvhost | Group {$_.Cluster}
	
	# Reporting setup
	$EndReport = @()

	# Break down into individual clusters
	foreach ($cluster in $clusters)
		{
		
		# Null out all of the counting vars
		$Report = {} | select ClusterName,PCPU,LCPU
		$PCPU = $null
		$LCPU = $null
		
		# Arrange the clusters into an array var
		$Report.ClusterName = $cluster.Name 
		[array]$entries = $cluster.Group
		
		# Look at each host in the cluster and run various queries
		foreach ($entry in $entries)
			{			
			$PCPU += $entry.'# CPU' * $entry.'Cores per CPU'
			$LCPU += $entry.'# CPU' * $entry.'# Cores'
			}
			
		# Add values to the report
		$Report.PCPU = $PCPU
		$Report.LCPU = $LCPU
		$EndReport += $Report	
	
		}