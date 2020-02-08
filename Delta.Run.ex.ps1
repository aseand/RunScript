
#Load funtions
. .\MIM.syncservice.run.ma.ps1
. .\MIM.portal.funtions.OP.ps1

$ScriptstartTime = [DateTime]::Now

#Set up log
$Name = $MyInvocation.MyCommand.Name
$logg = [NLog.LogManager]::GetLogger("MIM.syncservice.run.ma.$Name")
$logg.Info("Start run $Name")

if((Get-AnyMAInProgress)){ $log.Error("Agent running"); return }

#Set up timer
$timer = New-Object System.Timers.Timer
$timer.Interval = 1200000  #20min
#$timer.Interval = 1800000  #30min
#$timer.Interval = 2400000  #40min
$timer.AutoReset = $false

#Set timeoutacction
$Timeoutaction = Register-ObjectEvent -InputObject $timer -SourceIdentifier $Name -EventName Elapsed -Action { 
	$logg.Error("Error Message")
	#exit
	#[Environment]::Exit(0)
} 
$timer.Enabled = $true


#Run MAs
Start-MA -maName "AD-MA" -profile "Delta Import" 
Start-MA -maName "AD-MA" -profile "Delta Sync" -RunOnChange
Start-MA -maName "AD-MA" -profile "Export" -RunOnChange

Start-MA -maName "MIMPortal-MA" -profile "Export" -RunOnChange
While($PostProcessingCount -gt 0){ $PostProcessingCount = get-PostProcessingCount }
Start-MA -maName "MIMPortal-MA" -profile "Delta Import" 
Start-MA -maName "MIMPortal-MA" -profile "Delta Sync" -RunOnChange


#Unreg timeout
Unregister-Event $Name

$logg.Info("End RunTime: {0}s",([DateTime]::Now - $ScriptstartTime).TotalSeconds)