#=======================================================================================
# Name: Get-VMSnapshotWithOwner
# Created on: 2017-02-01
# Version 1.0
# Last Updated: 
# Last Updated by: John Shelton | c: 260-410-1200 | e: john.shelton@wegmans.com
#
# Purpose: 
#
# Notes: 
# 
# Change Log:
# 
#
#=======================================================================================
#
# Clear & Define Variables
#
$SnapshotRpt = @()
$CurrentDate = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
$VCenterServers = ("RDC-VMVC-01"),("BDC-VMVC-01"),("TST-VMVC-01")
$ReportName = "c:\Temp\VMSnapshotStatus_$CurrentDate.csv"
$ReportHTML = "c:\Temp\VMSnapshotStatus_$CurrentDate.html"
$Subject = "VMWare Snapshot Report as of $CurrentDate"
#
# EMail Settings
#
#$to = "techwintel-oncall@wegmans.com"
$to = "john.shelton@wegmans.com"
$from = "windows.support@wegmans.com"
$smtpserver = "smtp.wegmans.com"
#
# Email Format Settings
#
$EmailFormat = "<Style>"
$EmailFormat = $EmailFormat + "BODY{background-color:White;}"
$EmailFormat = $EmailFormat + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$EmailFormat = $EmailFormat + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:blue}"
$EmailFormat = $EmailFormat + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:lightgray}"
$EmailFormat = $EmailFormat + "</style>"
#
# Loop Through All VCenter Servers
#
FOREACH ($VCenterServer in $VCenterServers)
	{
	#
	# Connect to VMWare Server
	#
	Connect-VIServer $VCenterServer -Force
	#
	# Get All VMs with Snapshots
	#
	$Snapshots = Get-VM | Get-Snapshot
	#
	# Loop Through All Snapshots
	#
	FOREACH ($Snapshot in $Snapshots)
		{
		#
		# Review the VM Events for each server with a snapshot to determine the owner
		#
		$TaskMgr = Get-View TaskManager
		$Filter = New-Object VMware.Vim.TaskFilterSpec
		$Filter.Time = New-Object VMware.Vim.TaskFilterSpecByTime
		$Filter.Time.beginTime = ((($Snapshot.Created).AddSeconds(-5)).ToUniversalTime())
		$Filter.Time.timeType = "startedTime"
		$Filter.Time.EndTime = ((($Snapshot.Created).AddSeconds(5)).ToUniversalTime())
		$Filter.State = "success"
		$Filter.Entity = New-Object VMware.Vim.TaskFilterSpecByEntity
		$Filter.Entity.recursion = "self"
		$Filter.Entity.entity = (Get-Vm -Name $Snapshot.VM.Name).Extensiondata.MoRef
		$TaskCollector = Get-View ($TaskMgr.CreateCollectorForTasks($Filter))
		$TaskCollector.RewindCollector | Out-Null
		$Tasks = $TaskCollector.ReadNextTasks(100)
		FOREACH ($Task in $Tasks)
			{
			$GuestName = $Snapshot.VM
			$Task = $Task | where {$_.DescriptionId -eq "VirtualMachine.createSnapshot" -and $_.State -eq "success" -and $_.EntityName -eq $GuestName -and $_.Result -eq $Snapshot.ID}
			IF ($Task -ne $null)
				{
				$SnapUser = $Task.Reason.UserName
				}
			ELSE
				{
				$SnapUser = "No Task Found"
				}
			}
		$TaskCollector.DestroyCollector()
		#
		# Create the data log
		#
		$SnapshotTemp = New-Object PSObject
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "VCenterServer" -Value $VCenterServer
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "VMName" -Value $Snapshot.VM
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "SnapshotCreatedBy" -Value $SnapUser
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "SnapshotCreatedOn" -Value $Snapshot.Created
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "SnapshotName" -Value $Snapshot.Name
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "SnapshotDescription" -Value $Snapshot.Description
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "SnapshotIsCurrent" -Value $Snapshot.IsCurrent
		$SnapshotTemp | Add-Member -MemberType NoteProperty -Name "SnapshotSize(GB)" -Value $Snapshot.SizeGB.ToString(".##")
		$SnapshotRpt += $SnapshotTemp
	   }
	#
	# Create CSV File
	#
	$SnapshotRpt | Sort-Object VCenterServer, SnapshotCreatedOn | Export-CSV $ReportName -nti
	Disconnect-VIServer -Confirm:$false
	}
#
# Prepare the body of the email
#
# $Body = Get-Content $ReportName | ConvertFrom-Csv | ConvertTo-Html -head $EmailFormat -body "<H2> VCenter Server Snapshot Report </H2>" | Out-String
$Body = $SnapshotRpt | Sort-Object VCenterServer, SnapshotCreatedOn | ConvertTo-Html -head $EmailFormat -body "<H2> VCenter Server Snapshot Report </H2>" | Out-String
$BodyFile = $SnapshotRpt | Sort-Object VCenterServer, SnapshotCreatedOn | ConvertTo-Html -head $EmailFormat -body "<H2> VCenter Server Snapshot Report </H2>" | Set-Content $ReportHTML
#
# Send the email
#
Send-MailMessage -From $from -To $To -SmtpServer $SmtpServer -Subject $Subject -BodyAsHtml $Body -Attachments $ReportName
