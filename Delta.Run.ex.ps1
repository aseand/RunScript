
#Load funtions
. .\MIM.syncservice.run.ma.ps1
. .\MIM.portal.funtions.OP.ps1

function TriggerTemporalEventsJob{
	$startTime = [DateTime]::Now
	$log.Info("Start FIM_TemporalEventsJob")
	#"" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString > "C:\Script\Run\sqlpassword"
    #$sqlpassword = (New-Object System.Net.NetworkCredential "",((gc "C:\Script\Run\sqlpassword") | ConvertTo-SecureString)).Password
	#$resualt = start-SQLJob -ServerName "ServerName.se" -JobName "FIM_TemporalEventsJob" -StepName "step1" -Login "MIM-TaskRun" -LoginPassword $sqlpassword
	
	#$resualt = start-SQLJob -ServerName "ServerName.se" -JobName "FIM_TemporalEventsJob" -StepName "step1"
	$resualt = start-SQLJobSQL -ServerName "ServerName.se" -JobName "FIM_TemporalEventsJob" -Wait
	
	if($resualt -eq "last_run_outcome 1"){
		$log.Info("FIM_TemporalEventsJob run: $resualt")
	}else{
		$log.Error("FIM_TemporalEventsJob run: $resualt")
	}
	$log.Info("End FIM_TemporalEventsJob RunTime: {0}s",([DateTime]::Now - $startTime).TotalSeconds)
}

$ScriptstartTime = [DateTime]::Now

#Set up log
$Name = $MyInvocation.MyCommand.Name
$log = [NLog.LogManager]::GetLogger("MIM.syncservice.run.ma.$Name")
$log.Info("Start run $Name")

if((Get-AnyMAInProgress)){ $log.Error("Agent running in progress: '$((Get-AnyMAInProgress -List)[0])'"); return }

#Set up timer
$timer = New-Object System.Timers.Timer
$timer.Interval = 1200000  #20min
#$timer.Interval = 1800000  #30min
#$timer.Interval = 2400000  #40min
$timer.AutoReset = $false

#Set timeoutacction
$Timeoutaction = Register-ObjectEvent -InputObject $timer -SourceIdentifier $Name -EventName Elapsed -Action { 
	$log.Error("Run time exceeded time frame")
	#exit
	#[Environment]::Exit(0)
} 
$timer.Enabled = $true

#Refresch-synchronizationRule
$result = Refresch-synchronizationRule
foreach($error in $result.Previews.error){
	$log.Error($error)
}

#Run MAs
Start-MA -maName "AD-MA" -profile "Delta Import" -DontWait
Start-MA -maName "AD2-MA" -profile "Delta Import"
#while((Get-AnyMAInProgress)){sleep 5}
do{ sleep 5; $log.Info("AnyMAInProgress: $(Get-AnyMAInProgress -List)") }While((Get-AnyMAInProgress))

Start-MA -maName "AD-MA" -profile "Delta Sync" -RunOnChange
Start-MA -maName "AD2-MA" -profile "Delta Sync" -RunOnChange 
Start-MA -maName "AD-MA" -profile "Export" -RunOnChange

Start-MA -maName "MIMPortal-MA" -profile "Export" -RunOnChange
do{ sleep 5; $log.Info("PostProcessingCount: $(get-PostProcessingCount)") }While((get-PostProcessingCount) -gt 0)
Start-MA -maName "MIMPortal-MA" -profile "Delta Import" 
Start-MA -maName "MIMPortal-MA" -profile "Delta Sync" -RunOnChange -CountChangeOn @(@{ Attribute = "extSyncStatusAD"; Threshold = 500; Ratio = 10 })


$PurgeExecHistory = Run-PurgeExecHistory 14
$log.Info($PurgeExecHistory)

#Unreg timeout event
Unregister-Event $Name

$log.Info("End RunTime: {0}s",([DateTime]::Now - $ScriptstartTime).TotalSeconds)