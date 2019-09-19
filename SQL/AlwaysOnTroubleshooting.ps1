###########################################################################

$Servername = 'EUPATCAtLsn.apps.cloud.advent'
$AlertDateTime = "Saturday, June 29, 2019 10:35:10 AM"

$StartDate = (Get-Date $AlertDateTime).AddMinutes(-10)
$EndDate = (Get-Date $AlertDateTime).AddMinutes(+10)

Write-Host "`nServer Name : " $servername
Write-Host "`nAlert Time : " (Get-Date $AlertDateTime)
Write-Host "Start Date : " $StartDate
Write-Host "End Date   : " $EndDate


#######################################################################################
#      Reboot Time
#######################################################################################

Write-host " "
$LastRebootTime = (Get-WmiObject Win32_OperatingSystem -ComputerName $servername | Select @{LABEL = 'LastBootUpTime';Expression={$_.ConverttoDateTime($_.lastbootuptime)}}).lastbootuptime

$DaysSinceReboot = (New-TimeSpan -Start $LastRebootTime -End (Get-Date).DateTime).Days
Write-host "Server Reboot Time : " $LastRebootTime
Write-Host "Server was rebooted $DaysSinceReboot days earlier."

#######################################################################################
#  Always On Configuration
#######################################################################################

$Query = @"
WITH AGStatus AS(
SELECT
name as AGname,
replica_server_name,
CASE WHEN  (primary_replica  = replica_server_name) THEN  1
ELSE  '' END AS IsPrimaryServer,
secondary_role_allow_connections_desc AS ReadableSecondary,
[availability_mode]  AS [Synchronous],
failover_mode_desc
FROM master.sys.availability_groups Groups
INNER JOIN master.sys.availability_replicas Replicas ON Groups.group_id = Replicas.group_id
INNER JOIN master.sys.dm_hadr_availability_group_states States ON Groups.group_id = States.group_id
)
 
Select
[AGname],
[Replica_server_name],
[IsPrimaryServer],
[Synchronous],
[ReadableSecondary],
[Failover_mode_desc]
FROM AGStatus
--WHERE
--IsPrimaryServer = 1
--AND Synchronous = 1
ORDER BY
AGname ASC,
IsPrimaryServer DESC;
"@

Invoke-Sqlcmd -ServerInstance $servername -Query $Query | FT -AutoSize 

#######################################################################################
# SQL Error Log
#######################################################################################

Write-Host "`nSQL error Log : "
$Query = @"
If object_id('tempdb..#TmpErrorLog') IS NOT NULL drop table #TmpErrorLog

CREATE TABLE [dbo].[#TmpErrorLog]
([LogDate] DATETIME NULL,
 [ProcessInfo] VARCHAR(20) NULL,
 [Text] VARCHAR(MAX) NULL )

DECLARE @SQLString NVARCHAR(4000);
Declare @s varchar(20);
Declare @e varchar(20);

set @s = (select CONVERT(datetime, '$StartDate'))
set @e = (select CONVERT(datetime, '$StartDate'))
set @s = Dateadd(hour, -1, @s)
set @e = Dateadd(hour, 1, @e)

set @sqlstring = N'INSERT INTO #TmpErrorLog ([LogDate], [ProcessInfo], [Text]) exec xp_readerrorlog 0,1,NULL,NULL,'''+@s+''','''+@e+''',N''asc'''
execute sp_executesql @sqlstring

set @sqlstring = N'INSERT INTO #TmpErrorLog ([LogDate], [ProcessInfo], [Text]) exec xp_readerrorlog 1,1,NULL,NULL,'''+@s+''','''+@e+''',N''asc'''

execute sp_executesql @sqlstring
select * from #TmpErrorLog
"@

Invoke-Sqlcmd -ServerInstance $servername -Query $Query | FT -AutoSize 

#Get-winEvent -ComputerName ListnerName -filterHashTable @{logname ='Microsoft-Windows-FailoverClustering/Operational'; id=1641}| ft -AutoSize -Wrap

#Write-host "Ping Results"
#test-connection $servername STOMHAXDB6-2.EMEA.ONLINE.ADVENT

#Write-Host "Trace Route Results:"
#tracert -d $servername



#Get-EventLog -ComputerName $servername -After (Get-Date 02/19/2018) -EntryType Error -LogName System
#Write-host "Checking System Logs..."
#Get-EventLog -ComputerName $servername -After $StartDate -Before $EndDate -EntryType Error -LogName System | Select TimeGenerated,Source,Message,Username | FT -AutoSize -ErrorAction SilentlyContinue
#Write-host "Checking Application Logs..."
#Get-EventLog -ComputerName $servername -After $StartDate -Before $EndDate -EntryType Error -LogName Application | Select TimeGenerated,Source,Message,Username | FT -AutoSize -ErrorAction SilentlyContinue
#Write-host "Checking System Time Change Logs..."
#Get-EventLog -ComputerName $servername -After $StartDate -Before $EndDate -LogName System -Source "Microsoft-Windows-Kernel-General" | Select TimeGenerated,Source,Username, Message | FT -AutoSize -ErrorAction SilentlyContinue
#Write-host "Checking Logs Completed..."