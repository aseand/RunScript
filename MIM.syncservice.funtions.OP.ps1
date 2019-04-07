
add-type -Path "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\PropertySheetBase.dll"

function Write-XmlToScreen{
	<#
	  .SYNOPSIS
	  write xml object to string indented
	  .DESCRIPTION
	  write xml object to string indented
	  .EXAMPLE
	  Get-MAguid -xml $xmlobject
	  .PARAMETER xml
	  xml object
	#>
  [CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[xml]$xml
	)
    $StringWriter = New-Object System.IO.StringWriter
    $XmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter
    $XmlWriter.Formatting = "indented"
    $xml.WriteTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()
    Write-Output $StringWriter.ToString()
}

function Get-MAList{
	<#
	  .SYNOPSIS
	  Get list of MA names
	  .DESCRIPTION
	  Get list of MA names
	  .EXAMPLE
	  Get-MAList
	#>
  [CmdletBinding()]
	param(
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		$GuidList = $null
		$NameList = $null
		$MMSWebService.GetMAGuidList([ref]$GuidList,[ref]$NameList)
		
		$NameList
	}
}

function Get-MAguid{
	<#
	  .SYNOPSIS
	  Get Guid for MA by name
	  .DESCRIPTION
	  Get Guid for MA by name
	  .EXAMPLE
	  Get-MAguid -maName AD
	  .PARAMETER maName
	  Name of MA
	#>
  [CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[String]$maName,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		#[xml]$MAList = $MMSWebService.GetMAList()
		#$Node = $MAList.SelectSingleNode("/ma_list/ma[@name='$maName']")
		#[Guid]$Node.guid
		$GuidList = $null
		$NameList = $null
		[void]$MMSWebService.GetMAGuidList([ref]$GuidList,[ref]$NameList)
		$list = $NameList.Tolower()
		$index = $list.IndexOf($maName.Tolower())
		[Guid]$GuidList[$index]
	}
}

function Get-MAname{
	<#
	  .SYNOPSIS
	  Get MA for MA by name
	  .DESCRIPTION
	  Get Guid for MA by name
	  .EXAMPLE
	  Get-MAName -maGuid <guid string>
	  .PARAMETER maGuid
	  Guid of MA
	#>
  [CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[Guid]$maGuid,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		#[xml]$MAList = $MMSWebService.GetMAList()
		#$GuidString = "{$maGuid}".ToUpper()
		#$Node = $MAList.SelectSingleNode("/ma_list/ma[@guid='$GuidString']")
		#[string]$Node.name
		$GuidList = $null
		$NameList = $null
		[void]$MMSWebService.GetMAGuidList([ref]$GuidList,[ref]$NameList)
		$list = $GuidList.ToUpper()
		$index = $list.IndexOf("{$maGuid}".ToUpper())
		$NameList[$index]
	}
}

function Get-MAStatistics{
	<#
	  .SYNOPSIS
	  Get MAState and MAStatistics from MA, return xml object and MV object count
	  .DESCRIPTION
	  Get MAState and MAStatistics from MA, can return only ChangeCounters or deletecounts(stagecount, importcount, exportcount) or run status of MA
	  .EXAMPLE
	  Get-MAStatistics -maName AD
	  .EXAMPLE
	  Get-MAStatistics -maName AD -ChangeCounters
		.EXAMPLE
	  Get-MAStatistics -maName AD -RunStatus 
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER ChangeCounters
	  Return only value of stagecount, importcount, exportcount
	  .PARAMETER RunStatus
	  Return run state of MA
	#>
  [CmdletBinding()]
	param(
		[String]$maName,
		[Guid]$maGuid,
		[switch]$ChangeCounters,
		[switch]$DeleteCounters,
		[switch]$RunStatus,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName -MMSWebService $MMSWebService
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}

		[Xml]$GetMAState = $MMSWebService.GetMAState("<guids><guid>"+"{$maGuid}".ToUpper()+"</guid></guids>")
		if($RunStatus){
			return $GetMAState.'mas-state'.ma.state
		}
		
		$lastRunXml = $null
		$mvObjectCount = $null
		[Xml]$MAStatistics = $MMSWebService.GetMAStatistics("{$maGuid}".ToUpper(),[ref]$lastRunXml,[ref]$mvObjectCount)
		if($ChangeCounters){
			$stagecount = -1
			$stagingcounters = $GetMAState.'mas-state'.ma.'run-history'.'run-details'.'step-details'.'staging-counters'
			if($GetMAState.'mas-state'.ma.'run-history'.'run-details'.'step-details'.Count -eq 1){
			$stagecount = [int]([int]$stagingcounters.'stage-add'.InnerText+
								[int]$stagingcounters.'stage-update'.InnerText+
								[int]$stagingcounters.'stage-rename'.InnerText+
								[int]$stagingcounters.'stage-delete'.InnerText+
								[int]$stagingcounters.'stage-delete-add'.InnerText)
			}
			return @{
				stagecount = $stagecount
				importcount = [int]([int]$MAStatistics.'total-summary'.'import-add'+
								 [int]$MAStatistics.'total-summary'.'import-update'+
								 [int]$MAStatistics.'total-summary'.'import-delete')
				exportcount = [int]([int]$MAStatistics.'total-summary'.'export-add'+
								 [int]$MAStatistics.'total-summary'.'export-update'+
								 [int]$MAStatistics.'total-summary'.'export-delete') 
			}
		}
		
		if($DeleteCounters){
			$stagecount = -1
			$stagingcounters = $GetMAState.'mas-state'.ma.'run-history'.'run-details'.'step-details'.'staging-counters'
			if($GetMAState.'mas-state'.ma.'run-history'.'run-details'.'step-details'.Count -eq 1){
			$stagecount = [int]([int]$stagingcounters.'stage-delete'.InnerText+
								[int]$stagingcounters.'stage-delete-add'.InnerText)
			}
			return @{
				stagecountdelete = $stagecount
				importdelete = [int]([int]$MAStatistics.'total-summary'.'import-delete')
				exportdelete = [int]([int]$MAStatistics.'total-summary'.'export-delete') 
			}
		}
		
		return @{
			GetMAState = [Xml]$GetMAState
			MAStatistics = [Xml]$MAStatistics
			lastRunXml = [Xml]$lastRunXml
			mvObjectCount = [int]$mvObjectCount
		}
	}
}

function Get-AnyMAInProgress{
	<#
	  .SYNOPSIS
	  return true if any MA is running
	  .DESCRIPTION
	  return true if any MA is running
	  .EXAMPLE
	  AnyMAInProgress
	  .PARAMETER MMSWebService
	  MMSWebService object
	#>
  [CmdletBinding()]
	param(
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		$MAList = ([xml]$MMSWebService.GetMAList()).ma_list.ma

		for($i=0;$i -lt $MAList.Length;$i++){
			$stat = Get-MAStatistics -maGuid ($MAList[$i].guid) -RunStatus -MMSWebService $MMSWebService
			if($stat -eq "MA_EXEC_STATE_RUNNING")
			{
				$MAList[$i].name
				return $true
			}
		}
		return $false
	}
}

function Get-MIMParameters{
		<#
	  .SYNOPSIS
	  Get FIM/MIM SynchronizationService MSSQL server and instans and tabelname and Portal adress
	  .DESCRIPTION
	  Get FIM/MIM SynchronizationService MSSQL server and instans and tabelname and Portal adress
	  .EXAMPLE
	  Get-SynchronizationServiceSQLParameters
	#>
	process{
		$InstallMIMVersions = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  
		? {$_.DisplayName -like "*Identity Manager*Service*"}|
		%{@{DisplayName = $_.DisplayName;Version = $_.DisplayVersion}}
		
		$PortalParameters = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Forefront Identity Manager\2010\Portal" -ErrorAction SilentlyContinue

		#$ServiceParameters = Get-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\FIMService
		$SynchronizationServiceParameters = Get-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\FIMSynchronizationService\Parameters

		@{ 
			InstallVersions = $InstallMIMVersions
			PortalUrl = $PortalParameters.BaseSiteCollectionURL+"/identitymanagement/";
			SQLServerInstans = ("localhost",$SynchronizationServiceParameters.Server)[$SynchronizationServiceParameters.Server.Length -gt 0]+("",("\"+$SynchronizationServiceParameters.SQLInstance))[$SynchronizationServiceParameters.SQLInstance.Length -gt 0];
			DBName = $SynchronizationServiceParameters.DBName
		}
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
	#>
  [CmdletBinding()]
	param
	(
		[string]$maName,
		[guid]$maGuid,
		[Parameter(Mandatory = $true)]
		[string]$profile,
		[int]$TimeOutMin = 0,
		[switch]$RunOnChange,
		[switch]$DontWait,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-MAguid -maName $maName -MMSWebService $MMSWebService
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		
		if(-NOT $maName){
			$maName = Get-MAname -maGuid $maGuid -MMSWebService $MMSWebService
			if(-NOT $maName){ Throw "Missing MA '$maGuid'" }
		}
		
		$RunStatus = Get-MAStatistics -maGuid $maGuid -RunStatus -MMSWebService $MMSWebService
		if($RunStatus -ne "MA_EXEC_STATE_IDLE")
		{
			Throw "MA is running'$maName - $maGuid'" 
		}

		$profiles = ([xml]$MMSWebService.GetMaData("{$maGuid}".ToUpper(),1048592,539,27)).'ma-data'.'ma-run-data'.'run-configuration'
		#$profiles = ([xml]$MMSWebService.GetRunData([uint32]::MaxValue,"{$maGuid}".ToUpper(),"")).'ma-run-data'.'run-configuration'
		$profileIndex = $profiles.name.ToLower().IndexOf($Profile.ToLower())

		if($profileIndex -eq -1){
			write-error "No profile $Profile for $ma"
			return
		}
		
		#Get Step type
		$steptype = "sync"
		if($profiles[$profileIndex].configuration.step.'step-type'.type -eq "export"){
			$steptype = "export"
		}elseif(-NOT $profiles[$profileIndex].configuration.step.'step-type'.type.Contains("import")){
			$steptype = "import"
		}

		if($RunOnChange){
			$StatCount = -1
			$MAStatistics = Get-MAStatistics -maGuid $maGuid -ChangeCounters -MMSWebService $MMSWebService
			switch($steptype){
				"export" {
					$StatCount = $MAStatistics.exportcount
				}
				"import" {
					$StatCount = $MAStatistics.importcount
				}
			}
			if($StatCount -eq 0){
				return
			}
		}
		
		#Set timeout
		if($TimeOutMin -gt 0){
			$timer = New-Object System.Timers.Timer
			$timer.Interval = ($TimeOutMin * 60000)
			$timer.AutoReset = $false
			
			$Timeoutaction = Register-ObjectEvent -InputObject $timer -SourceIdentifier $maGuid -MessageData $maName -EventName Elapsed -Action { 
				Stop-MA -maGuid $Event.SourceIdentifier
				Write-Output -ForegroundColor Red -BackgroundColor Black  ("Timeout MA " + $Event.MessageData)
			} 
			$timer.Enabled = $true
		}
		
		#Run MA
		$StartDateTime = [DateTime]::Now
		$Resualt = $MMSWebService.RunMA("{$maGuid}".ToUpper(),("<run-configuration>{0}</run-configuration>" -f $profiles[$Profileindex].InnerXml),$false)
		$StopDateTime = [DateTime]::Now
		
		if([string]::IsNullOrEmpty($Resualt)){
			while(-not $DontWait){
				sleep 1
				$RunStatus = Get-MAStatistics -maGuid $maGuid -RunStatus -MMSWebService $MMSWebService
				if($RunStatus -eq "MA_EXEC_STATE_IDLE")
				{
					$DontWait = $true
				}
			}
		}else{
			
			write-error $Resualt $MAName $Profile
			#$MoreData = Get-MAStatistics -maGuid $maGuid
			#write-error $MoreData.lastRunXml.'run-history'.'run-details'
		}
		
		if($TimeOutMin -gt 0){
			Unregister-Event $maGuid 
		}
		
		if($SaveCSChangeData){
			#$runhistory = Get-MAStatistics -maGuid $maGuid
			#$runnumber = $runhistory.lastRunXml.'run-history'.'run-details'.'run-number'
			#[xml]$ExecutionHistory = $MMSWebService.GetExecutionHistory(("<execution-history-req ma=`"{0}`"><run-number>{1}</run-number><errors-summary>true</errors-summary></execution-history-req>" -f "{$maGuid}".ToUpper(),$runnumber))
			#$stepid = $ExecutionHistory.'run-history'.'run-details'.'step-details'.'step-id'
			switch($steptype){
				"sync" {
					#$ExecutionHistory.'run-history'.'run-details'.'step-details'.'outbound-flow-counters'
					#$tokenGuid = $MMSWebService.ExecuteStepObjectDetailsSearch("<step-object-details-filter step-id='{62A3AB8A-7A32-4766-B514-48E6B6AA391F}'><statistics type='connector-flow' ma-id='{5E2BCD35-D191-4AD5-BA25-795BF5FABFF4}'/></step-object-details-filter>")
				}
				"import" {
					#$ExecutionHistory.'run-history'.'run-details'.'step-details'.'staging-counters'
					#$tokenGuid = $MMSWebService.ExecuteStepObjectDetailsSearch("<step-object-details-filter step-id='$stepid'><statistics type='stage-update' /></step-object-details-filter>")
					#"<step-object-details-filter step-id='{AB1EB08B-A24C-4B6D-9144-4C528489B8F5}'><statistics type='stage-update' /></step-object-details-filter>"
				}
				"export" {
					#$ExecutionHistory.'run-histo)ry'.'run-details'.'step-details'.'export-counters'
					#$tokenGuid = $MMSWebService.ExecuteStepObjectDetailsSearch("<step-object-details-filter step-id='{019DBB69-5F50-4B6C-91B2-240D4946629A}'><statistics type='export-update' /></step-object-details-filter>")
				}
			}
			#[xml]$StepObjectResults = $MMSWebService.GetStepObjectResults($tokenGuid,1000)
			#$StepObjectResults.'step-object-details'.'cs-object'
			#foreach($CSGuid in $StepObjectResults.'step-object-details'.'cs-object'.id){
				#[xml]$CSData = $MMSWebService.GetCSObjects($row["object_id"].ToString(),1,$exportType,17,0,0)
			#}
		}
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
		[Guid]$maGuid,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName -MMSWebService $MMSWebService
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		if(-NOT $maName){
			$maName = Get-MAname -maGuid $maGuid -MMSWebService $MMSWebService
			if(-NOT $maName){ Throw "Missing MA '$maGuid'" }
		}
	
		$Resualt = $MMSWebService.StopMA("{$maGuid}".ToUpper())
		if(-not [string]::IsNullOrEmpty($Resualt)){
			Write-Output "$Resualt $maName"
			Throw "$Resualt $maName"
		}
	}
}

function Run-PurgeExecHistory{
	<#
	  .SYNOPSIS
	  Clear run history
	  .DESCRIPTION
	  Clear run history
	  .EXAMPLE
	  Run-PurgeExecHistory 30
	  .PARAMETER DaysBack
	  Number of day back that are not cleard
	#>
	param
	(
		[int]$DaysBack,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	$MMSWebService.PurgeExecHistory([datetime]::Now.AddDays(-$DaysBack).ToUniversalTime().ToShortDateString())
}

function Run-Preview{
	<#
	  .SYNOPSIS
	  Run preview for CS object in MIM from array of CSguids result is return as XML
	  .DESCRIPTION
	  Run preview for CS object in MIM from array of CSguids result is return as XML
	  .EXAMPLE
	  Run-Preview -maName AD -csGuids <guid array> -commit
	  .EXAMPLE
	  Run-Preview -maName AD -csGuids "e5e73bbe-00b5-e711-80cd-00155df29506" -delta -commit
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER csGuids
	  Array of CSGuids
	  .PARAMETER delta
	  Run as delta sync, def. is run as full sync (false)
	  .PARAMETER commit
	  Commit sync, def. if false
	#>
  [CmdletBinding()]
	param(
		[Guid]$maGuid,
		[String]$maName,
		[Parameter(Mandatory = $true)]
		[Guid[]]$csGuids,
		[switch]$delta,
		[switch]$commit,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName -MMSWebService $MMSWebService
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		$XMLString = New-Object System.Text.StringBuilder
		$ErrorString = New-Object System.Text.StringBuilder
	
		[void]$XMLString.Append("<Previews>")
		$PrCount=0
		foreach($csGuid in $csGuids){
			
			#[void]$XMLString.Append($MMSWebService.Preview("{$maGuid}".ToUpper(),"{$csGuid}".ToUpper(),$delta,$commit))
			[xml]$Preview = $MMSWebService.Preview("{$maGuid}".ToUpper(),"{$csGuid}".ToUpper(),$delta,$commit)
			if($Preview.InnerXml.Length -eq 0){
				[void]$ErrorString.Append("Not preview data return for $csGuid")
			}
			else{
				if($Preview.preview.error -ne $null){
					[void]$ErrorString.Append("Error $csGuid `n" + $Preview.preview.error.InnerXml)
				}
				[void]$XMLString.Append($Preview.InnerXml)
			}
			$PrCount++
			Write-Progress -Activity "Run Preview" -Status "Preview done for $csGuid" -PercentComplete ($PrCount/$csGuids.Length*100)
		}
		[void]$XMLString.Append("</Previews>")
		[xml]$XMLString.ToString()
		$ErrorString
	}
}

function Get-CSXml{
	<#
	  .SYNOPSIS
	  Get XML for CS object by Guids result in XML
	  .DESCRIPTION
	  Get XML for CS object by Guids result in XML, CSElementBitMask and CSEntryBitMask
	  .EXAMPLE
	  Get-CSXml -csGuid "166cc497-4b0e-4030-9b03-8f81cfbb7052" -flag 18446744073709551615
	  .EXAMPLE
	  Get-CSXml -csGuid @("166cc497-4b0e-4030-9b03-8f81cfbb7052","166cc498-4b0e-4030-9b03-8f81cfbb7052") -flag 17
	  .PARAMETER csGuid
	  Guid of CS object
	#>
  [CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[Guid[]]$csGuids,
		[uint64]$CSElementBitMask = [uint64]::MaxValue,
		[uint64]$CSEntryBitMask = 17,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)	
	process{
		$GuidStringArray = new-object string[] $csGuids.Length
		for($i = 0; $i -lt $csGuids.Length; $i++){
			$GuidStringArray[$i] = "{{{0}}}" -f $csGuids[$i].ToString().ToUpper()
		}
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_LASTEXPORT + 
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_LASTIMPORT +
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_MAID +
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_UNAPPLIEDEXPORT
		# 402784257
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_LASTEXPORT + 
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_LASTIMPORT +
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_MAID +
		#[Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase.CSElementBitMask]::CS_ELEMENT_PENDINGIMPORT
		# 402784264

		#[xml]("<?xml-stylesheet type=`"text/xsl`" href=`"CsExport.xslt`"?>`n" + $MMSWebService.GetCSObjects($GuidStringArray,$GuidStringArray.Length,$CSElementBitMask,$CSEntryBitMask,0,0))
		[xml]$MMSWebService.GetCSObjects($GuidStringArray,$GuidStringArray.Length,$CSElementBitMask,$CSEntryBitMask,0,0)
	}
}

function Get-CSGuid{
	<#
	  .SYNOPSIS
	  Get array of guids from MA CS by select object in MV
	  .DESCRIPTION
	  Get array of guids from MA CS by select object in MV  
	  .EXAMPLE
	  Get-CSGuid -maName AD -SQLServerInstans SQLserveralias\FIM -MVSQLWhereQuery "accountName = 'anase'"
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER InitialCatalog
	  Name of FIM SynchronizationService database, def.name FIMSynchronizationService
	  .PARAMETER SQLServerInstans
	  MSSQL server alias and/or instans
	  .PARAMETER MVSQLWhereQuery
	  whare SQL query for select MV object(s)
	#>
  [CmdletBinding()]
	param(
		[String]$maName,
		[Guid]$maGuid,
		[String]$InitialCatalog,
		[String]$SQLServerInstans,
		[string]$MVSQLWhereQuery = "object_type = 'person'",
		[switch]$GridView
	)
	
	begin{		
		if(-NOT $InitialCatalog -or -NOT $SQLServerInstans){
			$MIMParameters = Get-MIMParameters
			if(-NOT $InitialCatalog){
				$InitialCatalog = $MIMParameters.DBName
			}
			
			if(-NOT $SQLServerInstans){
				$SQLServerInstans = $MIMParameters.SQLServerInstans
			}
		}
		
		$ConnectionString = "Data Source=$SQLServerInstans;Initial Catalog=$InitialCatalog;Integrated Security=SSPI;"
		$Connection = New-Object System.Data.SqlClient.SqlConnection ($ConnectionString)
		$Connection.Open()
	}
	
	process{
		if(-NOT $maGuid -AND $maName){
			$maGuid = Get-maguid -maName $maName
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
	
		$sqlcommand  = "SELECT DISTINCT cs.rdn,ma.ma_name,cs.object_id FROM mms_csmv_link csmv (nolock) "
		$sqlcommand += "join mms_connectorspace cs (nolock) on csmv.cs_object_id = cs.object_id "
		$sqlcommand += "join mms_management_agent ma (nolock) on ma.ma_id = cs.ma_id "
		if($maGuid){
			$sqlcommand += "where cs.ma_id = '$maGuid' "
		}
		$sqlcommand += "and csmv.mv_object_id in (select object_id from mms_metaverse (nolock) where $MVSQLWhereQuery)"

		
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($sqlcommand,$Connection)
		$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout

		$DataTable = New-Object system.Data.DataTable "csGuids"
		$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
		$RowCount = $Adapter.Fill($DataTable)
		
		$Adapter.Dispose()
		$SqlCmd.Dispose()
		
		if($GridView){
			$SelectData = $DataTable|Out-GridView -Title "CS object select" -OutputMode Multiple
		}else{
			$SelectData = $DataTable.Rows
		}
		
		#$CSGuids = New-Object System.Collections.Generic.HashSet[string]
		$CSGuids = New-Object System.Collections.ArrayList
		foreach($row in $SelectData){
			[void]$CSGuids.Add($row["object_id"])
		}
		
		
		(,$CSGuids.ToArray())
	}
	end{
		$Connection.Close()
	}
}

function Disconnect-CSobject{
	<#
	  .SYNOPSIS
	  Disconnect CS objects from MV
	  .DESCRIPTION
	  Disconnect CS objects from MV
	  .EXAMPLE
	  Disconnect-CS -maName AD -csGuid "166cc497-4b0e-4030-9b03-8f81cfbb7052"
	  .EXAMPLE
	  Disconnect-CS -maName AD -csGuid @("166cc497-4b0e-4030-9b03-8f81cfbb7052","166cc498-4b0e-4030-9b03-8f81cfbb7052")
	  .PARAMETER MMSWebService
	  MMSWebService object
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER csGuids
	  List CSGuids
	#>
  [CmdletBinding()]
	param(
		[String]$maName,
		[Guid]$maGuid,
		[Parameter(Mandatory = $true)]
		[Guid[]]$csGuids,
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName -MMSWebService $MMSWebService
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		foreach($cs in  $csGuids){
			$string = $MMSWebService.Disconnect("{$maGuid}".ToUpper(),"{$cs}".ToUpper())
			if($string.Length -ne 0){
				$string += "$maGuid $cs"
				Write-Error $string
			}
		}
	}
}

function Join-CS-MV{
	<#
	  .SYNOPSIS
	  Join CS object to MV object
	  .DESCRIPTION
	  Join CS object to MV object
	  .EXAMPLE
	  Join-CS-MV -maGuid "5e2bcd35-d191-4ad5-ba25-795bf5fabff4" -csGuid "166cc497-4b0e-4030-9b03-8f81cfbb7052" -mvObjectType "person" -mvGuid "dfd138ce-4eb2-e711-80c9-9d6a50da7060"
	  .EXAMPLE
	  Join-CS-MV -maName AD -csGuid "166cc497-4b0e-4030-9b03-8f81cfbb7052" -mvObjectType "person" -mvGuid "dfd138ce-4eb2-e711-80c9-9d6a50da7060"
	  .PARAMETER MMSWebService
	  MMSWebService object
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER csGuid
	  Guid of CS object
	  .PARAMETER mvObjectType
	  String name of type from MV object 
	  .PARAMETER mvGuid
	  Guid of MV object
	#>
  [CmdletBinding()]
	param(
		[String]$maName,
		[Guid]$maGuid,
		[Parameter(Mandatory = $true)]
		[Guid]$csGuid,
		[Parameter(Mandatory = $true)]
		[Guid]$mvGuid,
		[String]$mvObjectType = "person",
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName -MMSWebService $MMSWebService
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		$string = $MMSWebService.Join("{$maGuid}".ToUpper(),"{$csGuid}".ToUpper(),$mvObjectType,"{$mvGuid}".ToUpper())
		if($string.Length -ne 0){
			$string += "`n maGuid: $maGuid`n csGuid: $csGuid`n mvObjectType: $mvObjectType`n mvGuid: $mvGuid"
			Write-Error $string
		}
	}
}

function refresch-synchronizationRule{
	
	$synchronizationRuleCSGuids = Get-CSGuid -maName "MIMService-MA" -MVSQLWhereQuery "cs.object_type = 'synchronizationRule'"
	#Get precedence?
	
	#Disconnect-CS -maName "MIMService-MA" -csGuid $synchronizationRuleCSGuids
	#First may fail?
	$previewRun = Run-Preview -maName "MIMService-MA" -csGuid $synchronizationRuleCSGuids[0] -commit
	#$previewRun.Previews.preview.error
	Run-Preview -maName "MIMService-MA" -csGuid $synchronizationRuleCSGuids -commit
	
	#set precedence
}
