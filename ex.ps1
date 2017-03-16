. "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\RunScript\run-funtions.ps1"


$ScriptstartTime = get-date
$logger.Info("********* Start script run {0}",$MyInvocation.MyCommand.Name)

#Stop if any still running
if(AnyInProgress){
	$logger.warn("Agent running in progress... retry in 60s")
	sleep 60
	if(AnyInProgress){
		$logger.error("Agent running in progress stop {0}",$MyInvocation.MyCommand.Name)
		exit
	}
}

#Timeout alert
$timer = New-Object System.Timers.Timer
$timer.Interval = 1200000  #20min
#$timer.Interval = 1800000  #30min
$timer.AutoReset = $false

$Timeoutaction = { 
	$logger.error("Script timeout alert!")
}  
Register-ObjectEvent -InputObject $timer -SourceIdentifier MIMTimeoutElapsed -EventName Elapsed -Action $Timeoutaction
$timer.Enabled = $true

$ImportStartTime = get-date
#Import start stage
$logger.Info("*** Import stage")

start-agent 'WaHSAn' 'Full import' -AsJob
sleep 5
start-agent 'Heroma.KP.PERSON' 'Full import' -AsJob
sleep 5
start-agent 'AD' 'Delta import'
start-agent 'AnvandarID' 'Delta import'
start-agent 'AnvandarID-Special' 'Full import'
start-agent 'ADAM' 'Delta import'
#start-agent 'WaHSAn' 'Delta import' -AsJob
start-agent 'Nice' 'Delta import'
start-agent 'ARX' 'Delta import'
start-agent 'SharePointSite' 'Full import'
start-agent 'AID-Special' 'Full import'
start-agent 'RCARD-M5' 'Delta import'
start-agent 'Heroma.BH.USERID' 'Full import' -AsJob
sleep 5
start-agent 'Heroma.KP.ANSTAELLNING' 'Full import' -AsJob

#Wait on jobb, log & remove
GetWaitLogRemove-jobs
#Import end stage

#Save all change row (hologram and deltas) 
SaveChangeCS $ImportStartTime
#

$ExportStartTime = get-date
#Sync start stage
$logger.Info("*** Sync stage")
start-agent 'AnvandarID' 'Delta sync' -StageCount
start-agent 'AnvandarID-Special' 'Delta sync' -StageCount
# Dalfolke after AnvandarID! So new project has crate uniqueIdentifier i MV
$ImportStartTime = get-date
start-agent 'Dalfolke' 'Delta import'
#Save all change row (hologram and deltas) 
SaveChangeCS $ImportStartTime

start-agent 'Dalfolke' 'Delta sync' -StageCount
#
start-agent 'ADAM' 'Delta sync' -StageCount
start-agent 'AD' 'Delta sync' -StageCount
start-agent 'Heroma.KP.ANSTAELLNING' 'Delta sync' -StageCount
start-agent 'AID-Special' 'Delta sync'
start-agent 'Heroma.BH.USERID' 'Delta sync' -StageCount
start-agent 'Heroma.KP.PERSON' 'Delta sync' -StageCount
start-agent 'Nice' 'Delta sync' -StageCount
start-agent 'ARX' 'Delta sync' -StageCount
start-agent 'WaHSAn' 'Delta sync' -StageCount
start-agent 'RCARD-M5' 'Delta sync' -StageCount
start-agent 'SharePointSite' 'Full sync'
#Sync end stage
SaveChangeCS $ExportStartTime -export

#Export start stage
$logger.Info("*** Export stage")
start-agent 'AnvandarID' 'Export' -ExportCount -AsJob
start-agent 'AnvandarID-Special' 'Export' -ExportCount -AsJob
start-agent 'Synergi.Anvandare' 'Export' -ExportCount -AsJob
GetWaitLogRemove-jobs
start-agent 'WaHSAn' 'Export' -ExportCount -AsJob
$UserOrgExportCount = (export-count 'UserOrgExport')
start-agent 'UserOrgExport' 'Export' -ExportCount -AsJob
start-agent 'Nice' 'Export' -ExportCount -AsJob
GetWaitLogRemove-jobs

start-agent 'Medusa' 'Export' -ExportCount -AsJob
start-agent 'SharePointSite' 'Export' -ExportCount -AsJob
start-agent 'Heroma.KP.PERSON' 'Export' -ExportCount -AsJob
GetWaitLogRemove-jobs
start-agent 'Heroma.BH.USERID' 'Export' -ExportCount -AsJob
start-agent 'AD' 'Export' -ExportCount -AsJob
start-agent 'ARX' 'Export' -ExportCount -AsJob
start-agent 'RCARD-M5' 'Export' -ExportCount -AsJob
GetWaitLogRemove-jobs

#Export end stage

$logger.Info("********* End script run {0} execTime:{1}s",$MyInvocation.MyCommand.Name,((get-date)-$ScriptstartTime).TotalSeconds )
Unregister-Event MIMTimeoutElapsed