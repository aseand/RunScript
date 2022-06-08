
. (join-path ($PSScriptRoot) MIM.syncservice.funtions.OP.ps1)

#Variables
$global:SqlConnectionString = "Data Source=localhost;Initial Catalog=FIMSynchronizationService;Integrated Security=SSPI;"

#For name and profile test without running
#$global:NameTest = $true

#Nlog
#Download Nlog
if(-NOT (Test-Path (join-path ($PSScriptRoot) NLog.dll)))
{
	$FileName = "4.6.8"
	$FullPath = (join-path ($PSScriptRoot) $FileName)
	Invoke-WebRequest "https://www.nuget.org/api/v2/package/NLog/$FileName" -OutFile $FullPath
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	$zip = [System.IO.Compression.ZipFile]::OpenRead($FullPath)
	$zip.Entries|?{$_.FullName.StartsWith("lib/net45/NLog.dll")}|%{[System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, (join-path ($PSScriptRoot) $_.Name), $true)}
	$zip.Dispose()
	rm $FullPath
}

Add-Type -Path (join-path ($PSScriptRoot) NLog.dll)

#Nlog 
if((Test-Path (join-path ($PSScriptRoot) "NLog.config.xml")) ){
	([NLog.LogManager]::Configuration) = new-object NLog.Config.XmlLoggingConfiguration((join-path ($PSScriptRoot) "NLog.config.xml"))
	
}else{
	$Configuration = New-Object NLog.Config.LoggingConfiguration
	#$Configuration.AddRule("Trace", "Fatal", (new-object NLog.Targets.ConsoleTarget),"*")
	$Configuration.AddRule("Info", "Fatal", (new-object NLog.Targets.ConsoleTarget),"*")
	[NLog.LogManager]::Configuration = $Configuration
}

function profileArgumentCompleter{
    param ( $commaName,
            $parameterName,
            $wordToComplete,
            $commandAst,
            $fakeBoundParameters )
			
	if ($fakeBoundParameters.ContainsKey('maName')) {
		[Guid]$maGuid = Get-maguid -maName $fakeBoundParameters.maName
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		[array]$profiles = ([xml]$MMSWebService.GetMaData("{$maGuid}".ToUpper(),1048592,539,27)).'ma-data'.'ma-run-data'.'run-configuration'
		$profiles.name.ToLower() | % { "'$_'" }
	}
}

function Start-MA{
	<#
	  .SYNOPSIS
	  Run MA agent in MIM whit run profile can wait until done and cant run on import/export change only
	  .DESCRIPTION
	  Run MA agent in MIM whit run profile can wait until done and cant run on import/export change only
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full import"
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full import" -DontWait
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full import" -TimeOutMin 10
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full sync" -RunOnChange
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER profile
	  Profile to run
	  .PARAMETER TimeOutMin
	  Timeout in min befor timeout action
	  .PARAMETER RunOnChange
	  Run only if import/export count of selectet profile have data to work on
	  .PARAMETER DontWait
	  Start MA and dont wait untill done
	  .PARAMETER MMSWebService
	  MMSWebService object
		#NLog
		Add-Type -Path ("path\NLog.dll")
		([NLog.LogManager]::Configuration) = new-object NLog.Config.XmlLoggingConfiguration("path\NLogNLog.config.xml")
		$global:logger = [NLog.LogManager]::GetLogger("MIM.syncservice.funtions.OP")
	#>
  [CmdletBinding()]
	param
	(
		[parameter(Mandatory=$true)]
		[ArgumentCompleter({ Get-MAList })]
		[String]$maName,
		[Parameter(Mandatory = $true)]
		[ArgumentCompleter({ profileArgumentCompleter @args })]
		[string]$profile,
		[int]$TimeOutMin = 0,
		[int]$DeleteThreshold = -1,
		[int]$DeleteRatio = 10,
		$CountChangeOn, #@(@{ Attribute = "AttributeName"; Threshold = 153; Ratio = 10 })
		[switch]$RunOnChange,
		[switch]$DontWait,
		[switch]$SaveCSChangeData
	)
	process{
		
		$StartTime = [DateTime]::Now
		$logger = [NLog.LogManager]::GetLogger("MIM.syncservice.run.ma.$maName")
		
		if($logger.IsDebugEnabled){ 
			foreach ($Name in ((Get-Command -Name ($PSCmdlet.MyInvocation.InvocationName)).Parameters).Keys) {
				try{$logger.Debug("$Name :" + (Get-Variable $Name -ValueOnly -ErrorAction SilentlyContinue))}Catch{}
			}
		}
		
		#
		if((Get-Service FIMSynchronizationService).Status -ne "Running"){ 
			$logger.Warn("FIMSynchronizationService not running, try Start-Service")
			Start-Service FIMSynchronizationService
			sleep 10
		}
	
		#Get MA guid
		[Guid]$maGuid = Get-maguid -maName $maName
		if(-NOT $maGuid){
			$logger.Fatal("Start-MA Missing MA '$maName'")
			Throw "Start-MA Missing MA '$maName'"
			return
		}
		#

		$RunStatus = Get-MAStatistics -maGuid $maGuid -RunStatus
		$logger.Debug("MA '$maName' '$maGuid' '$profile' runStatus '$RunStatus' ")
		if($RunStatus -ne "MA_EXEC_STATE_IDLE")
		{
			$logger.Fatal("MA is running '$maName' '$maGuid'")
			Throw "MA is running'$maName - $maGuid'"
			return
		}
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		[array]$profiles = ([xml]$MMSWebService.GetMaData("{$maGuid}".ToUpper(),1048592,539,27)).'ma-data'.'ma-run-data'.'run-configuration'
		#$profiles = ([xml]$MMSWebService.GetRunData([uint32]::MaxValue,"{$maGuid}".ToUpper(),"")).'ma-run-data'.'run-configuration'
		$profileIndex = $profiles.name.ToLower().IndexOf($Profile.ToLower())

		if($profileIndex -eq -1 -OR $profiles[$profileIndex].InnerXml.Length -lt 1){
			$logger.Fatal("No profile '$Profile' ")
			$logger.Debug($profiles[$profileIndex].InnerXml)
			Throw "No profile '$Profile' for '$maName'"
			return
		}
		
		#Get Step type
		$steptype = $profiles[$profileIndex].configuration.step.'step-type'.type
		$logger.Debug("steptype $steptype")

		$StatCount = -1
		$MAStatistics = Get-MAStatistics -maGuid $maGuid -ChangeCounters
		if($logger.IsDebugEnabled){
			foreach($name in $MAStatistics.Keys){
				$logger.Debug("{0} {1}",$name,$MAStatistics[$name])
			}
		}
		switch -wildcard ($steptype){
			"*export" {
				$StatCount = $MAStatistics.exportcount
				if($DeleteThreshold -gt 0 -AND $MAStatistics.exportdelete -gt 0 -AND $MAStatistics.exportdelete -gt $DeleteThreshold){
					$logger.Fatal("'$Profile' threshold $DeleteThreshold / " + $MAStatistics.exportdelete )
					Throw "'$maName' '$Profile' threshold $DeleteThreshold / " + $MAStatistics.exportdelete 
					return
				}
				#if($DeleteRatio -gt 0 -AND $MAStatistics.exportdelete -gt 0 -AND ($MAStatistics.exportdelete/$MAStatistics.all * 100) -gt $DeleteRatio){
				if($DeleteRatio -gt 0 -AND $MAStatistics.exportdelete -gt 0 -AND ($MAStatistics.exportdelete/$MAStatistics.totalconnector * 100) -gt $DeleteRatio){
					#$Value = $MAStatistics.exportdelete/$MAStatistics.all * 100
					$Value = $MAStatistics.exportdelete/$MAStatistics.totalconnector * 100
					$logger.Fatal("'$Profile' threshold ratio $DeleteRatio% Current value: $Value%")
					Throw "'$maName' '$Profile' threshold ratio $DeleteRatio% Current value: $Value%"
					return
				}
				
				if($CountChangeOn -AND $MAStatistics.exportupdate -gt 0){
					$CSChangeList = $null

					foreach($CountChange in $CountChangeOn){
						
						$ChangeCount = 0
						$ChangeCountTot = 0
						
						if($CountChange.Attribute -eq "*"){
							$ChangeCount = $MAStatistics.exportupdate
							#$ChangeCountTot = $MAStatistics.importcount
						}else{
							if(!$CSChangeList){
								#$CSChangeList = Get-CSSearch -maGuid $maGuid -SearchMethod PendingExport -Modify -AllAttribute
								
								$AttributeList = $CountChangeOn | % {$_.Attribute}
								$CSChangeList = Get-CSSearch -maGuid $maGuid -SearchMethod PendingExport -Modify -AttributeArray $AttributeList
							}
							
							$CSChangeList.'unapplied-export'.delta.attr|?{ $_.Name -eq $CountChange.Attribute} | %{ $ChangeCount++ }
							#$ChangeCountTot = $CSChangeList.Count
						}

						if($CountChange.Threshold -gt 0 -AND $ChangeCount -gt $CountChange.Threshold){
							$logger.Fatal("'$Profile' change threshold $($CountChange.Attribute): $($CountChange.Threshold) / $ChangeCount")
							Throw "'$maName' '$Profile' change threshold $($CountChange.Attribute): $($CountChange.Threshold) / $ChangeCount" 
							return
						}
						
						if($CountChange.Ratio -gt 0 -AND $ChangeCount -gt 0 -AND ($ChangeCount/$MAStatistics.totalconnector * 100) -gt $CountChange.Ratio){
							$logger.Fatal("'$Profile' change threshold ratio $($CountChange.Attribute): $($CountChange.Ratio) / $($MAStatistics.totalconnector)")
							Throw "'$maName' '$Profile' change threshold ratio $($CountChange.Attribute): $($CountChange.Ratio) / $($MAStatistics.totalconnector)" 
							return
						}
					}
				}
			}
			"*import*" {
			}
			"apply-rules" {
				$DontWait = $false
				$StatCount = $MAStatistics.importcount
				if($DeleteThreshold -gt 0 -AND $MAStatistics.importdelete -gt 0 -AND $MAStatistics.importdelete -gt $DeleteThreshold){
					$logger.Fatal("'$Profile' threshold $DeleteThreshold / " + $MAStatistics.importdelete )
					Throw "'$maName' '$Profile' threshold $DeleteThreshold / " + $MAStatistics.importdelete 
					return
				}
				#if($DeleteRatio -gt 0 -AND $MAStatistics.importdelete -gt 0 -AND ($MAStatistics.importdelete/$MAStatistics.all * 100) -gt $DeleteRatio){
				if($DeleteRatio -gt 0 -AND $MAStatistics.importdelete -gt 0 -AND ($MAStatistics.importdelete/$MAStatistics.totalconnector * 100) -gt $DeleteRatio){
					#$Value = $MAStatistics.importdelete/$MAStatistics.all * 100
					$Value = $MAStatistics.importdelete/$MAStatistics.totalconnector * 100
					$logger.Fatal("'$Profile' threshold ratio $DeleteRatio% Current value: $Value%")
					Throw "'$maName' '$Profile' threshold ratio $DeleteRatio% Current value: $Value%"
					return
				}
				
				if($CountChangeOn -AND $MAStatistics.importupdate -gt 0){
					$CSChangeList = $null

					foreach($CountChange in $CountChangeOn){
						
						$ChangeCount = 0
						$ChangeCountTot = 0
						
						if($CountChange.Attribute -eq "*"){
							$ChangeCount = $MAStatistics.importupdate
							#$ChangeCountTot = $MAStatistics.importcount
						}else{
							if(!$CSChangeList){
								#$CSChangeList = Get-CSSearch -maGuid $maGuid -SearchMethod PendingImport -Modify -AllAttribute
								
								$AttributeList = $CountChangeOn | % {$_.Attribute}
								$CSChangeList = Get-CSSearch -maGuid $maGuid -SearchMethod PendingImport -Modify -AttributeArray $AttributeList
							}
							
							$CSChangeList.'pending-import'.delta.attr|?{ $_.Name -eq $CountChange.Attribute} | %{ $ChangeCount++ }
							#$ChangeCountTot = $CSChangeList.Count
						}

						if($CountChange.Threshold -gt 0 -AND $ChangeCount -gt $CountChange.Threshold){
							$logger.Fatal("'$Profile' change threshold $($CountChange.Attribute): $($CountChange.Threshold) / $ChangeCount")
							Throw "'$maName' '$Profile' change threshold $($CountChange.Attribute): $($CountChange.Threshold) / $ChangeCount" 
							return
						}
						
						if($CountChange.Ratio -gt 0 -AND $ChangeCount -gt 0 -AND ($ChangeCount/$MAStatistics.totalconnector * 100) -gt $CountChange.Ratio){
							$logger.Fatal("'$Profile' change threshold ratio $($CountChange.Attribute): $($CountChange.Ratio) / $($MAStatistics.totalconnector)")
							Throw "'$maName' '$Profile' change threshold ratio $($CountChange.Attribute): $($CountChange.Ratio) / $($MAStatistics.totalconnector)" 
							return
						}
					}
				}
			}
		}
		
		if($RunOnChange -AND $StatCount -eq 0){
			return
		}
		
		#Set timeout
		if($TimeOutMin -gt 0){
			$timer = New-Object System.Timers.Timer
			$timer.Interval = ($TimeOutMin * 60000)
			$timer.AutoReset = $false
			
			$logger.Debug("Start timer $TimeOutMin Min")
			
			$Timeoutaction = Register-ObjectEvent -InputObject $timer -SourceIdentifier $maGuid -MessageData $maName -EventName Elapsed -Action { 
				Stop-MA -maGuid $Event.SourceIdentifier
				#Write-Output -ForegroundColor Red -BackgroundColor Black  ("Timeout MA " + $Event.MessageData)
				$logger.Fatal("'$Profile' Timeout MA " + $Event.MessageData)
				Throw "'$maName' '$Profile' Timeout MA " + $Event.MessageData
				return
			} 
			$timer.Enabled = $true
		}
		
		$logger.Info("Start ma '$maName' '$Profile' DontWait:$DontWait SaveCSChangeData:$SaveCSChangeData")
		if($NameTest){ $DontWait = $false }

		if(-NOT $DontWait){
			Run-Agent -maName $maName -maGuid $maGuid -Profile $Profile -ProfileXml ("<run-configuration>{0}</run-configuration>" -f $profiles[$Profileindex].InnerXml) -SaveCSChangeData:$SaveCSChangeData -logger $logger
			$logger.Info("'$Profile' for " + ([DateTime]::now - $StartTime).TotalSeconds + "s")
		}else{
			
			#$scriptPath = (join-path ($PSScriptRoot) MIM.syncservice.run.ma.ps1)
			#$initScript = [scriptblock]::Create("Import-Module -Name '$scriptPath'")
			#$job = Start-Job -ScriptBlock (${Function:Run-Agent}) -InitializationScript $initScript -ArgumentList $maName,$maGuid,$Profile,("<run-configuration>{0}</run-configuration>" -f $profiles[$Profileindex].InnerXml),$SaveCSChangeData
		
			if(-NOT $runspacesHandels){
				$Global:runspacesHandels = New-Object System.Collections.ArrayList
				$Global:RunspacePool = [RunspaceFactory]::CreateRunspacePool(1,20)
				$RunspacePool.Open()
			}
			
			$scriptPath = (join-path ($PSScriptRoot) MIM.syncservice.funtions.OP.ps1)
			$logger.Debug($scriptPath)
			#$PowerShell = [PowerShell]::Create().AddScript(". $scriptPath")
			#$PowerShell.Invoke()
			#$PowerShell.Commands.Commands.RemoveAt(0)
			
			$PowerShell = [PowerShell]::Create().AddScript(${Function:Run-Agent}).AddArgument($maName).AddArgument($maGuid).AddArgument($Profile).AddArgument($("<run-configuration>{0}</run-configuration>" -f $profiles[$Profileindex].InnerXml)).AddArgument($SaveCSChangeData).AddArgument($scriptPath).AddArgument($logger)
			$PowerShell.RunspacePool = $runspacepool
			$temp = New-Object -TypeName PSObject -Property @{
				PowerShell = $PowerShell 
				Runspace = $PowerShell.BeginInvoke()
			}
			[void]$runspacesHandels.Add($temp)
		}
		
		if($TimeOutMin -gt 0){
			$logger.Debug("Unregister-Event")
			Unregister-Event $maGuid 
		}
	}
}

function Run-Agent{
	<#
	  .SYNOPSIS
	  Run MA agent in MIM whit run profile can wait until done and cant run on import/export change only
	  .DESCRIPTION
	  Run MA agent in MIM whit run profile can wait until done and cant run on import/export change only
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full import"
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full import" -DontWait
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full import" -TimeOutMin 10
	  .EXAMPLE
	  Start-MA -maName "AD" -profile "Full sync" -RunOnChange
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER profile
	  Profile to run
	  .PARAMETER ProfileXml
	  Profile to run XML
	  .PARAMETER MMSWebService
	  MMSWebService object
		#NLog
		Add-Type -Path ("path\NLog.dll")
		([NLog.LogManager]::Configuration) = new-object NLog.Config.XmlLoggingConfiguration("path\NLogNLog.config.xml")
		$global:logger = [NLog.LogManager]::GetLogger("MIM.syncservice.funtions.OP")
	#>
	[CmdletBinding()]
	param
	(
		$maName,
		$maGuid,
		$Profile,
		$ProfileXml,
		$SaveCSChangeData,
		$scriptPath,
		$logger
	)
	process{
	#$logger.Info("Run-Agent $scriptPath")
	if($scriptPath){ Import-Module $scriptPath }
		if($logger.IsDebugEnabled){ 
			foreach ($Name in ((Get-Command -Name ($PSCmdlet.MyInvocation.InvocationName)).Parameters).Keys) {
				try{$logger.Debug("Run-Agent: $Name :" + (Get-Variable $Name -ValueOnly -ErrorAction SilentlyContinue))}Catch{}
			}
		}
		
		#Get MA guid
		if(-NOT $maGuid){
			$returnvalue = Get-MAguid -maName $maName
			if(-NOT $returnvalue){
				$logger.Fatal("Run-Agent Missing MA '$maName'")
				Throw "Run-Agent Missing MA '$maName'"
				return
			}
			$maGuid = $returnvalue
		}
		
		#Get MA name
		if(-NOT $maName){
			$maName = Get-MAname -maGuid $maGuid
			if(-NOT $maName){
				$logger.Fatal("Run-Agent Missing MA '$maGuid'")
				Throw "Run-Agent Missing MA '$maGuid'"
				return
			}
		}
		
		$RunStatus = Get-MAStatistics -maGuid $maGuid -RunStatus
		$logger.Debug("MA '$maName' '$maGuid' $profile runStatus $RunStatus ")
		if($RunStatus -ne "MA_EXEC_STATE_IDLE")
		{
			$logger.Fatal("MA is running '$maName' '$maGuid'")
			Throw "MA is running'$maName - $maGuid'"
			return
		}
		
		
		if(-NOT $ProfileXml){
		
			$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
			$profiles = (,([xml]$MMSWebService.GetMaData("{$maGuid}".ToUpper(),1048592,539,27)).'ma-data'.'ma-run-data'.'run-configuration')
			#$profiles = ([xml]$MMSWebService.GetRunData([uint32]::MaxValue,"{$maGuid}".ToUpper(),"")).'ma-run-data'.'run-configuration'
			$profileIndex = $profiles.name.ToLower().IndexOf($Profile.ToLower())

			if($profileIndex -eq -1){
				$logger.Fatal("No profile $Profile")
				Throw "No profile $Profile for $maName"
				return
			}
			
			$ProfileXml = "<run-configuration>{0}</run-configuration>" -f $profiles[$Profileindex].InnerXml
		}
		
		#Run MA
		$logger.Debug("Run agent '$maName' '$maGuid' '$ProfileXml' SaveCSChangeData:$SaveCSChangeData")
		$logger.Info("Run '$Profile' SaveCSChangeData:$SaveCSChangeData")
		
		if($NameTest){
			return
		}
		
		$Retry = $false
		$RetryCount = 0
		do{
			$MaStartTime = [DateTime]::Now
			
			$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
			$Resualt = $MMSWebService.RunMA("{$maGuid}".ToUpper(),$ProfileXml,$false)
			
			#$logger.Debug("Resualt:$Resualt")
			
			if($Resualt){
				#write-error $Resualt $MAName $Profile
				#$logger.Error("Runing '$Profile'")
				$logger.Fatal("Runing '$Profile' Resualt: $Resualt")
				Throw $Resualt
				return
			}
			
			$runhistory = Get-MAStatistics -maGuid $maGuid
			$runnumber = $runhistory.lastRunXml.'run-history'.'run-details'.'run-number'
			
			$logger.Trace("Wait on '$Profile' runnumber: $runnumber")
			$Done = $false
			while(-NOT $Done){
				sleep 1
				$RunStatus = Get-MAStatistics -maGuid $maGuid -RunStatus
				
				#Error
				if(-NOT $RunStatus){
					$logger.Warn("No run status, retry in 60s...")
					sleep 60
					
					if((Get-Service FIMSynchronizationService).Status -ne "Running"){ 
						$logger.Error("FIMSynchronizationService not running, try Start-Service")
						Start-Service FIMSynchronizationService
						sleep 10
					}

					$RunStatus = Get-MAStatistics -maGuid $maGuid -RunStatus
					if(-NOT $RunStatus){
						$logger.Error("No run status. FIMSynchronizationService status: $((Get-Service FIMSynchronizationService).Status)")
						break
					}
				}
				
				#$logger.Trace("$RunStatus")
				if($RunStatus -eq "MA_EXEC_STATE_IDLE")
				{
					$Done = $true
				}
			}
			

					
			$runhistory = Get-MAStatistics -maGuid $maGuid
			
			if($logger.IsTraceEnabled){
				$logger.Trace((Write-XmlToScreen ($runhistory.lastRunXml)))
			}
				
			#Errors to log
			$result = $runhistory.lastRunXml.'run-history'.'run-details'.result
			$newrunnumber = $runhistory.lastRunXml.'run-history'.'run-details'.'run-number'
			
			$MaStopTime = [DateTime]::Now
			$logger.Info("'$Profile' run for $(($MaStopTime-$MaStartTime).TotalSeconds)s run-number: $newrunnumber")
			
			if($result -ne "success"){
				switch -wildcard ($result){
					"stopped-*" { 
						if($result -ne "stopped-object-limit" -OR $result -ne "stopped-server-down" ){
							$logger.Info("Retry run in 60s ...")
							sleep 60
							$Retry = $true
						}
					}
					
					"no-start-ma" {
						$logger.Fatal("'$Profile' $result run-number: $newrunnumber")
					}
					
					default {
						$logger.Error("'$Profile' $result run-number: $newrunnumber")
					}
				}
			}
			$RetryCount++
		}while($Retry -AND $RetryCount -lt 2)
		
		if($Retry -AND $result -ne "success"){
			$logger.Fatal("'$Profile' $result run-number: $newrunnumber")
		}
		
		$runnumber = $runhistory.lastRunXml.'run-history'.'run-details'.'run-number'
		#$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		#[xml]$ExecutionHistory = $MMSWebService.GetExecutionHistory(("<execution-history-req ma=`"{0}`"><run-number>{1}</run-number><errors-summary>true</errors-summary></execution-history-req>" -f "{$maGuid}".ToUpper(),$runnumber))
		[xml]$ExecutionHistory = Get-ExecutionHistory -maGuid $maGuid -runNumber $runnumber
				
		if($logger.IsTraceEnabled){
			$logger.Trace((Write-XmlToScreen ($ExecutionHistory.'run-history'.InnerXml)))
		}
		
		$securityid = $ExecutionHistory.'run-history'.'run-details'.'security-id'
		$logger.Info("'$Profile' run by '$securityid'")
		
		foreach($Step in $ExecutionHistory.'run-history'.'run-details'.'step-details'){
			#Errors info
			if($Step.'step-result' -ne "success"){
				$stepnumber = $Step.'step-number'
				$steptype = $Step.'step-description'.'step-type'.'type'
				$logger.Error("Step $stepnumber $steptype on $maName")
				
				if($logger.IsInfoEnabled){
					foreach($Node in $Step.ChildNodes){
						if($Node.Name -like '*error*'){
							foreach($ErrorNode in $Node.ChildNodes){
								if($logger.IsDebugEnabled){$logger.Debug((Write-XmlToScreen ($ErrorNode.OuterXml)))}
								foreach($InnerErrorNode in $ErrorNode.ChildNodes){
									if($ErrorNode.Attributes["dn"]){$DN = $ErrorNode.Attributes["dn"].Value} else { $DN = "" }
									$logger.Info($ErrorNode.Name + ": " + $DN +" " + $InnerErrorNode.Name + " " + $InnerErrorNode.InnerText)
								}
							}
						}
					}
				}
			}
			
			#Statistics info
			if($logger.IsInfoEnabled){
				foreach($Node in $Step.ChildNodes){
					if($Node.Name -like '*-counter' -OR $Node.Name -like '*-counters'){
						#$Node.Name
						foreach($CountNode in $Node.ChildNodes){		
							$Value = [int]::Parse($CountNode.InnerText)
							if($Value -gt 0){
								$logger.Info(("{0} {1} {2}" -f $Node.Name,$CountNode.Name,$Value))
							}
						}
					}
				}
			}
		}		
				
		if($SaveCSChangeData){
			#$runhistory = Get-MAStatistics -maGuid $maGuid
			#$runnumber = $runhistory.lastRunXml.'run-history'.'run-details'.'run-number'
			#$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
			#[xml]$ExecutionHistory = $MMSWebService.GetExecutionHistory(("<execution-history-req ma=`"{0}`"><run-number>{1}</run-number><errors-summary>true</errors-summary></execution-history-req>" -f "{$maGuid}".ToUpper(),$runnumber))
			#[xml]$ExecutionHistory = Get-ExecutionHistory -maGuid $maGuid -runNumber $runnumber -MMSWebService $MMSWebService
			$stepid = $ExecutionHistory.'run-history'.'run-details'.'step-details'.'step-id'
			
			$exportType = 402784264
			switch -wildcard ($steptype){
				"*export" {
					$Nodes = $ExecutionHistory.'run-history'.'run-details'.'step-details'.'export-counters'.ChildNodes
					$exportType = 402784257
				}
				"*import*" {
					$Nodes = $ExecutionHistory.'run-history'.'run-details'.'step-details'.'staging-counters'.ChildNodes
				}
				"apply-rules" {
					$Nodes = $ExecutionHistory.'run-history'.'run-details'.'step-details'.'inbound-flow-counters'.ChildNodes
				}
			}
			
			$logger.Debug("maName: '$maName' maGuid: '$maGuid' run-number: $runnumber step-id: $stepid Node count: " + $Nodes.Count)

			foreach($Node in $Nodes){
				$logger.Debug("Node - Name: {0} detail: {1} Count: {2}",$Node.Name,$Node.detail,$Node.InnerText)
				if($Node.detail -AND $Node.InnerText -ne "0"){

					#$tokenGuid = $MMSWebService.ExecuteStepObjectDetailsSearch("<step-object-details-filter step-id='$stepid'><statistics type='" + $Node.Name + "' /></step-object-details-filter>")

					#[xml]$StepObjectResults = $MMSWebService.GetStepObjectResults($tokenGuid,$Node.InnerText)
					$StepObjectResults = Get-StepObjects -stepId $stepid -statisticsType $Node.Name -PageSize $Node.InnerText
					#if($logger.IsTraceEnabled){
						#$logger.Trace($StepObjectResults.InnerXml)
					#}
					#each CS
					#foreach($CSGuid in $StepObjectResults.'step-object-details'.'cs-object'.id){
						#[xml]$CSData = $MMSWebService.GetCSObjects($CSGuid,1,$exportType,17,0,0)
						#$CSData.Save("c:\CSData\$CSGuid.xml")
					#}
					
					
					#All cs in one
					#if($StepObjectResults.'step-object-details'.'cs-object'.Count -gt 0){
					if($StepObjectResults.Count -gt 0){
						#$logger.Debug("cs-object count: {0} {1}",$StepObjectResults.'step-object-details'.'cs-object'.Count, [string]::Join(",",$StepObjectResults.'step-object-details'.'cs-object'.id))
						$logger.Debug("cs-object count: {0} {1}",$StepObjectResults.Count, [string]::Join(",",$StepObjectResults.id))
						
						#[xml]$CSData = $MMSWebService.GetCSObjects($StepObjectResults.id,$StepObjectResults.Count,$exportType,17,0,0)
						[xml]$CSData = Get-CSXml -csGuids $StepObjectResults.id -CSElementBitMask $exportType -CSEntryBitMask 17
						#$CSData.Save("c:\CSData\"+$Node.Name+".xml")
						if($logger.IsTraceEnabled){
							$logger.Trace($CSData.InnerXml)
						}
						
						$DataTable = New-Object System.Data.DataTable("cs-delta")
						[void]$DataTable.Columns.Add("cs-dn",[String])
						[void]$DataTable.Columns.Add("id",[Guid])
						[void]$DataTable.Columns.Add("object-type",[String])
						[void]$DataTable.Columns.Add("ma-id",[Guid])
						[void]$DataTable.Columns.Add("last-import-delta-time",[DateTime])
						[void]$DataTable.Columns.Add("cs-object",[String])
						<#
						CREATE TABLE [dbo].[change_history](
							[cs-dn] [nvarchar](438) NOT NULL,
							[id] [uniqueidentifier] NOT NULL,
							[object-type] [nvarchar](255) NULL,
							[ma-id] [uniqueidentifier] NULL,
							[last-import-delta-time] [datetime] NULL,
							[cs-object] [xml] NULL,
							[rowid] [int] IDENTITY(1,1) NOT NULL,
						 CONSTRAINT [PK_cust_cs-delta] PRIMARY KEY CLUSTERED 
						(
							[rowid] ASC
						)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
						) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
						GO
						
						CREATE NONCLUSTERED INDEX [IX_change_history] ON [dbo].[change_history]
						(
							[cs-dn] ASC
						)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
						GO
						
						CREATE NONCLUSTERED INDEX [IX_change_history_1] ON [dbo].[change_history]
						(
							[id] ASC
						)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
						GO
						
						CREATE NONCLUSTERED INDEX [IX_change_history_2] ON [dbo].[change_history]
						(
							[last-import-delta-time] ASC
						)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
						GO
						
						CREATE NONCLUSTERED INDEX [IX_change_history_3] ON [dbo].[change_history]
						(
							[ma-id] ASC
						)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
						GO
						
						#>
						
						foreach($cs in $CSData.'cs-objects'.ChildNodes){
							$row = $DataTable.NewRow()
							$row["cs-dn"] = $cs.'cs-dn'
							$row["id"] = $cs.'id'
							$row["object-type"] = $cs.'object-type'
							$row["ma-id"] = $cs.'ma-id'
							$row["last-import-delta-time"] = $cs.'last-import-delta-time'
							$row["cs-object"] = $cs.OuterXml
							
							$DataTable.Rows.Add($row)
						}
						
						$Connection = New-Object System.Data.SqlClient.SqlConnection $global:SqlConnectionString
						$Connection.Open()
						
						$SqlBulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy $Connection
						$SqlBulkCopy.BulkCopyTimeout = $Connection.ConnectionTimeout
						$SqlBulkCopy.DestinationTableName = "change_history"

						foreach ($colum in $DataTable.Columns){
							[void]$SqlBulkCopy.ColumnMappings.Add($colum.ColumnName, $colum.ColumnName)
						}
						
						$SqlBulkCopy.WriteToServer($DataTable)
						$SqlBulkCopy.Close()
						
						$Connection.Dispose()
						$DataTable.Dispose()
						
					}else{
						$logger.Debug("cs-object count: {0}",$StepObjectResults.Count)
					}
				}
			}
		}	
		$logger.Debug("Run '$Profile' done")
	}
}

function Stop-MA{
	<#
	  .SYNOPSIS
	  Stop MA if runing
	  .DESCRIPTION
	  Stop MA if runing
	  .EXAMPLE
	  stop-MA -maName AD
	  .PARAMETER maName
	  Name of MA
	#>
	param
	(
		[string]$maName,
		[Guid]$maGuid
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName
			if(-NOT $maGuid){ Throw "Stop-MA Missing MA '$maName'" }
		}
		if(-NOT $maName){
			$maName = Get-MAname -maGuid $maGuid
			if(-NOT $maName){ Throw "Stop-MA Missing MA '$maGuid'" }
		}
	
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		$Resualt = $MMSWebService.StopMA("{$maGuid}".ToUpper())
		if(-not [string]::IsNullOrEmpty($Resualt)){
			Write-Output "$Resualt $maName"
			Throw "$Resualt $maName"
		}
	}
}
