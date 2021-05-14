
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
	process{
		$GuidList = $null
		$NameList = $null
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
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
		[String]$maName
	)
	process{
		#[xml]$MAList = $MMSWebService.GetMAList()
		#$Node = $MAList.SelectSingleNode("/ma_list/ma[@name='$maName']")
		#[Guid]$Node.guid
		$GuidList = $null
		$NameList = $null
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		[void]$MMSWebService.GetMAGuidList([ref]$GuidList,[ref]$NameList)
		
		if($maName){
			$list = $NameList.Tolower()
			$index = $list.IndexOf($maName.Tolower())
			if($index -gt -1){
				return [Guid]$GuidList[$index]
			}
		}else{
			$list = @{}
			for($i = 0; $i -lt $GuidList.Count; $i++){
				$list.Add($NameList[$i],$GuidList[$i])
			}
			return $list
		}
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
		[Guid]$maGuid
	)
	process{
		#[xml]$MAList = $MMSWebService.GetMAList()
		#$GuidString = "{$maGuid}".ToUpper()
		#$Node = $MAList.SelectSingleNode("/ma_list/ma[@guid='$GuidString']")
		#[string]$Node.name
		$GuidList = $null
		$NameList = $null
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
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
		[switch]$NoDisconnector,
		[switch]$RunStatus
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}

		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
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
			if($NoDisconnector){
				[int]$disconnector = 0
			}
			else{
				[int]$disconnector = $MAStatistics.'total-summary'.'filtered-disconnector'
			}
			if($GetMAState.'mas-state'.ma.'run-history'.'run-details'.'step-details'.Count -eq 1){
			$stagecount = [int]([int]$stagingcounters.'stage-add'.InnerText+
								[int]$stagingcounters.'stage-update'.InnerText+
								[int]$stagingcounters.'stage-rename'.InnerText+
								[int]$stagingcounters.'stage-delete'.InnerText+
								[int]$stagingcounters.'stage-delete-add'.InnerText)
			}
			return @{
				stagecount = $stagecount
				all = [int]$MAStatistics.'total-summary'.'all'
				totalconnector = [int]$MAStatistics.'total-summary'.'total-connector'
				importcount = [int]([int]$MAStatistics.'total-summary'.'import-add'+
								 [int]$MAStatistics.'total-summary'.'import-update'+
								 [int]$MAStatistics.'total-summary'.'import-delete')-
								 $disconnector
				importdelete = [int]$MAStatistics.'total-summary'.'import-delete'
				exportcount = [int]([int]$MAStatistics.'total-summary'.'export-add'+
								 [int]$MAStatistics.'total-summary'.'export-update'+
								 [int]$MAStatistics.'total-summary'.'export-delete') 
				exportdelete = [int]$MAStatistics.'total-summary'.'export-delete'
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
	#>
	process{
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		$MAList = ([xml]$MMSWebService.GetMAList()).ma_list.ma

		for($i=0;$i -lt $MAList.Length;$i++){
			$stat = Get-MAStatistics -maGuid ($MAList[$i].guid) -RunStatus
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
		$InstallMIMVersions = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |  
		? {$_.DisplayName -like "*Identity Manager*Service*"} |
		%{@{DisplayName = $_.DisplayName;Version = $_.DisplayVersion}}
		
		$PortalParameters = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Forefront Identity Manager\2010\Portal" -ErrorAction SilentlyContinue

		#$ServiceParameters = Get-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\FIMService
		$SynchronizationServiceParameters = Get-ItemProperty HKLM:\SYSTEM\CurrentControlSet\services\FIMSynchronizationService\Parameters -ErrorAction SilentlyContinue
		
		@{ 
			InstallVersions = $InstallMIMVersions
			PortalUrl = ("",$PortalParameters.BaseSiteCollectionURL + "/identitymanagement/")[$PortalParameters -ne $null]
			SQLServerInstans = (("","localhost")[$SynchronizationServiceParameters -ne $null],$SynchronizationServiceParameters.Server)[$SynchronizationServiceParameters.Server.Length -gt 0]+("",("\"+$SynchronizationServiceParameters.SQLInstance))[$SynchronizationServiceParameters.SQLInstance.Length -gt 0];
			DBName = $SynchronizationServiceParameters.DBName
			Path = $SynchronizationServiceParameters.Path
		}
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
		[uint64]$CSEntryBitMask = 17
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
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		[xml]$MMSWebService.GetCSObjects($GuidStringArray,$GuidStringArray.Length,$CSElementBitMask,$CSEntryBitMask,0,0)
	}
}

function Get-CSGuidBySQL{
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

function Get-CSGuid{
	<#
	  .SYNOPSIS
	  Get array of guid from CS object by MV object
	  .DESCRIPTION
	  Get array of guid from CS object by MV object
	  .EXAMPLE
	  Get-CSGuid -maName AD -mvGuids "166cc497-4b0e-4030-9b03-8f81cfbb7052"
	  .EXAMPLE
	  Get-CSGuid -maName AD -mvGuids @("166cc497-4b0e-4030-9b03-8f81cfbb7052","166cc498-4b0e-4030-9b03-8f81cfbb7052") -GridView
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER mvGuids
	  Guids of mv object id
	  .PARAMETER GridView
	  Display and select CS object to list
	#>
  [CmdletBinding()]
	param(
		[String]$maName,
		[Guid]$maGuid,
		[Guid[]]$mvGuids,
		[switch]$GridView
	)
	process{
		if(-NOT $maGuid -AND $maName){
			$maGuid = Get-maguid -maName $maName
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		
		$DataTable = New-Object system.Data.DataTable "csGuids"
		$DataTable.Columns.Add("cs-dn")
		$DataTable.Columns.Add("cs-guid")
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		
		foreach($mvGuid in $mvGuids){
			[xml]$Connectors = $MMSWebService.GetMVConnectors($mvGuid)
			foreach($csmvlink in $connectors.'cs-mv-links'.'cs-mv-value'){
				if($csmvlink.'ma-guid' -eq $maguid){
					[void]$CSGuids.Add($csmvlink.'cs-dn',$csmvlink.'cs-guid')
				}
			}
		}
		
		if($GridView){
			$SelectData = $DataTable|Out-GridView -Title "CS object select" -OutputMode Multiple
		}else{
			$SelectData = $DataTable.Rows
		}
		
		$CSGuids = New-Object System.Collections.ArrayList
		foreach($row in $SelectData){
			[void]$CSGuids.Add($row["cs-guid"])
		}
		
		
		(,$CSGuids.ToArray())
	}
}

function Get-SearchMV{
	<#
	  .SYNOPSIS
	  Search MV data and get mv object
	  .DESCRIPTION
	  Search MV data and get mv object
	  .EXAMPLE
	  Get-SearchMV -searchAttrs @{name="accountName";value="anase";searchtype="exact";type="string"} -objecttype person
	  .EXAMPLE
	  Get-SearchMV -searchAttrs @(@{name="accountName";value="anase";searchtype="exact";type="string"},@{name="accountName";value="anase2";searchtype="exact";type="string"})
	  .PARAMETER objecttype
	  MV object type
	  .PARAMETER searchAttrs
	  Array of search object (Hashtable) most contains  name, value, searchtype, type(data type)
	  ex. @{name="accountName";value="anase";searchtype="exact";type="string"}
	  ex. @{name="displayName";searchtype="value-exists";type="string"}
	  searchtype values: exact, starts, ends, contains, not-contains, value-exists, no-value
	  type values: string, integer, binary, bit
	#>
  [CmdletBinding()]
	param(
		[Object[]]$searchAttrs,
		[string]$objecttype = "person"
	)
	process{

		
		$attFilter = ""
		foreach($searchAttr in $searchAttrs){
			$ValueXml = ""
			if($searchAttr.value){
				$ValueXml = "<value>$($searchAttr.value)</value>"
			}
			#$attFilter += "<mv-attr name=`""+$searchAttr.Name+"`" type=`""+$searchAttr.type+"`" search-type=`""+$searchAttr.searchtype+"`">$Value</mv-attr>"
			$attFilter += "<mv-attr name=`"$($searchAttr.Name)`" type=`"$($searchAttr.type)`" search-type=`"$($searchAttr.searchtype)`">$($ValueXml)</mv-attr>"
			#$attFilter += "<mv-attr name=`"accountName2`" type=`"string`" search-type=`"exact`"><value>value2</value></mv-attr>"
		}
		$objectfilter = "<mv-object-type>$objecttype</mv-object-type>"
		$mvfilter = "<mv-filter collation-order=`"Latin1_General_CI_AS`">{0}{1}</mv-filter>" -f $attFilter,$objectfilter
		
		#$mvfilter
		#InnerXml	"<mv-attr name=\"displayName\" type=\"string\" search-type=\"value-exists\" /><mv-object-type>synchronizationRule</mv-object-type>"	string

		#"<mv-filter collation-order=\"Latin1_General_CI_AS\"><mv-object-type>synchronizationRule</mv-object-type></mv-filter>"
		#$filter = "<mv-filter collation-order=`"Latin1_General_CI_AS`"><mv-object-type>synchronizationRule</mv-object-type></mv-filter>"
		#$filter = "<mv-filter collation-order=`"Latin1_General_CI_AS`"><mv-attr name=`"accountName`" type=`"string`" search-type=`"exact`"><value>value</value></mv-attr><mv-object-type>person</mv-object-type></mv-filter>"
		#$filter = "<mv-filter collation-order=`"Latin1_General_CI_AS`"><mv-attr name=`"accountName`" type=`"string`" search-type=`"exact`"><value>value</value></mv-attr><mv-attr name=`"accountName2`" type=`"string2`" search-type=`"exact`"><value>value2</value></mv-attr><mv-object-type>person</mv-object-type></mv-filter>"
		#$MMSWebService.SearchMV("<mv-filter collation-order=`"Latin1_General_CI_AS`"><mv-object-type>synchronizationRule</mv-object-type></mv-filter>")
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		[xml]$SearchMVXml = "<mv-objects>" + ($MMSWebService.SearchMV($mvfilter)).Replace("<mv-objects>","").Replace("</mv-objects>","") + "</mv-objects>"
		
		#[xml]$Connectors = $MMSWebService.GetMVConnectors("MVGuid")
		
		$SearchMVXml
	}
}

function Get-ExecutionHistory{
	<#
	  .SYNOPSIS
	  Get execution history xml form run-number
	  .DESCRIPTION
	  Get execution history xml from run-number
	  .EXAMPLE
	  Get-ExecutionHistory -maName "AD" -runNumber 245
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER runNumber
	  Number for run
	#>
  [CmdletBinding()]
	param(
		[string]$maName,
		[guid]$maGuid,
		[Parameter(Mandatory = $true)]
		[int]$runNumber
	)	
	process{
		#Get MA guid
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-MAguid -maName $maName
			if(-NOT $maGuid){
				$logger.Error("Missing MA '$maName'")
				Throw "Missing MA '$maName'"
				return
			}
		}
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		[xml]$MMSWebService.GetExecutionHistory(("<execution-history-req ma=`"{0}`"><run-number>{1}</run-number><errors-summary>true</errors-summary></execution-history-req>" -f "{$maGuid}".ToUpper(),$runNumber))
	}
}

function Get-StepObjects{
	<#
	  .SYNOPSIS
	  Get step  details CS objects
	  .DESCRIPTION
	  Get step  details CS objects
	  .EXAMPLE
	  Get-StepObjects -stepId "166cc497-4b0e-4030-9b03-8f81cfbb7052" -statisticsType ""
	  .PARAMETER stepId
	  Step id guid from Execution History
	  .PARAMETER statisticsType
	  Type of statistics
	  .PARAMETER PageSize
	  Page size for reualt
	#>
  [CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[guid]$stepId,
		[Parameter(Mandatory = $true)]
		[string]$statisticsType,
		[int]$PageSize = 1000
	)	
	process{
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		$tokenGuid = $MMSWebService.ExecuteStepObjectDetailsSearch(("<step-object-details-filter step-id='{0}'><statistics type='{1}' /></step-object-details-filter>" -f"{$stepId}".ToUpper(), $statisticsType))
		
		$csobjects = New-Object System.Collections.ArrayList
		$Count = 1
		while($Count -gt 0){
			[xml]$StepObjectResult = $MMSWebService.GetStepObjectResults($tokenGuid, $PageSize)
			$Count = $StepObjectResult.'step-object-details'.ChildNodes.Count
			if($Count -gt 0){
				[void]$csobjects.AddRange((,$StepObjectResult.'step-object-details'.ChildNodes))
			}
		}
		$csobjects.ToArray()
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
		[int]$DaysBack
	)
	
	$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
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
		[switch]$commit
	)
	process{
		#foreach($mv in $mvs.'mv-objects'.'mv-object'){
		#	$index = [array]::IndexOf($mv.'cs-mv-links'.'cs-mv-value'.'ma-name',"AD-MA")
		#	Run-Preview -maName "AD-MA" -csGuids $mv.'cs-mv-links'.'cs-mv-value'[$index].'cs-guid' -commit
		#}
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName
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
				[void]$ErrorString.Append("<Error>Not preview data return for $csGuid</Error>")
			}
			else{
				if($Preview.preview.error -ne $null){
					[void]$ErrorString.Append("<Error>Error $csGuid `n $($Preview.preview.error.InnerXml)</Error>")
				}
				[void]$XMLString.Append($Preview.InnerXml)
			}
			$PrCount++
			Write-Progress -Activity "Run Preview" -Status "Preview done for $csGuid" -PercentComplete ($PrCount/$csGuids.Length*100)
		}
		[void]$XMLString.Append("<Errors>$($ErrorString.ToString())</Errors></Previews>")

		[xml]$XMLString.ToString()
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
		[Guid[]]$csGuids	
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
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
		[String]$mvObjectType = "person"
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		$string = $MMSWebService.Join("{$maGuid}".ToUpper(),"{$csGuid}".ToUpper(),$mvObjectType,"{$mvGuid}".ToUpper())
		if($string.Length -ne 0){
			$string += "`n maGuid: $maGuid`n csGuid: $csGuid`n mvObjectType: $mvObjectType`n mvGuid: $mvGuid"
			Write-Error $string
		}
	}
}

function Set-ConnectorState{
	<#
	  .SYNOPSIS
	  Change Connector state on CS object
	  .DESCRIPTION
	  Change Connector state on CS object
	  .EXAMPLE
	  Set-ConnectorState -maGuid "5e2bcd35-d191-4ad5-ba25-795bf5fabff4" -csGuids "166cc497-4b0e-4030-9b03-8f81cfbb7052" -ConnectorState "CONNECTORSTATE_NORMAL"
	  .EXAMPLE
	  Set-ConnectorState -maName AD -csGuids "166cc497-4b0e-4030-9b03-8f81cfbb7052" -ConnectorState "CONNECTORSTATE_STAY"
	  .PARAMETER MMSWebService
	  MMSWebService object
	  .PARAMETER maGuid
	  Guid of MA
	  .PARAMETER maName
	  Name of MA
	  .PARAMETER csGuids
	  Guids of CS objects
	  .PARAMETER ConnectorState
	  State to set on CS object
	#>
  [CmdletBinding()]
	param(
		[String]$maName,
		[Guid]$maGuid,
		[Parameter(Mandatory = $true)]
		[Guid[]]$csGuids,
		[Parameter(Mandatory = $true)]
		[Guid]$mvGuid,
		[Parameter(Mandatory = $true)]
		[ValidateSet("CONNECTORSTATE_EXPLICIT","CONNECTORSTATE_NORMAL","CONNECTORSTATE_STAY")]
		[String]$ConnectorState = "CONNECTORSTATE_NORMAL"
	)
	process{
		if(-NOT $maGuid){
			[Guid]$maGuid = Get-maguid -maName $maName
			if(-NOT $maGuid){ Throw "Missing MA '$maName'" }
		}
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		foreach($csGuid in $csGuids){
			$string = $MMSWebService.SetConnectorState("{$maGuid}".ToUpper(),"{$csGuid}".ToUpper(),$ConnectorState)
			if($string.Length -ne 0){
				$string += "`n maGuid: $maGuid`n csGuid: $csGuid`n ConnectorState: $ConnectorState`n"
				Write-Error $string
			}
		}
	}
	#MMSWebService.SetConnectorState("maGuid","csguid", [Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.CONNECTORSTATE]::CONNECTORSTATE_EXPLICIT)
}

function Refresch-synchronizationRule{
	<#
	  .SYNOPSIS
	  Run full preview and commit of all synchronization rules from servcie
	  .DESCRIPTION
	  Change Connector state on CS object
	  .EXAMPLE
	  Refresch-synchronizationRule
	  .EXAMPLE
	  Refresch-synchronizationRule
	#>
	process{

		$MVsynchronizationRules = Get-SearchMV -objecttype "synchronizationRule"
		Run-Preview -maGuid $MVsynchronizationRules.'mv-objects'.'mv-object'[0].'cs-mv-links'.'cs-mv-value'.'ma-guid' -commit -csGuid ($MVsynchronizationRules.'mv-objects'.'mv-object'.'cs-mv-links'.'cs-mv-value'.'cs-guid')
	}
}

function Update-MaSchema{
	<#
	  .SYNOPSIS
	  Update mv schema from xml data
	  .DESCRIPTION
	  Update metaverse schema in FIM/MIM fom xml string
	  User ex. $MMSWebService.GetMVData(511) or GUI to get management data
	  .EXAMPLE
	  Update-Schema -mvdata "...xml..."
	  .PARAMETER mvdata
	  metaverse xml data
	#>
  [CmdletBinding()]
	param(
		[xml]$mvdata
		)
	process{
		if($mvdata -ne $null){
			
			$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
			$Oldmvdata = [xml]$MMSWebService.GetMVData(511)
			$CurrentVersion = $Oldmvdata.'mv-data'.version 
			$NewSchemaVersion = $mvdata.'mv-data'.version
			
			if($CurrentVersion -ne $NewSchemaVersion ){
				Write-Error "New schema xml version $NewSchemaVersion is not eq current version $CurrentVersion`nPlease update xml data to lates version!"
				#Write-Error ""
				return
			}
		
			$return = $MMSWebService.ModifyMVData($mvdata.InnerXml)
			if($return.Length > 0){
				Write-Error $return
				return $false
			}
			return $true
		}
	}
}

function Add-AttributeToClass{
	<#
	  .SYNOPSIS
	  Add attribute to mv class return mv xml data
	  .DESCRIPTION
	  Add attribute to metaverse class object, crate attribute if needed
	  Reuser xml data to add more then one attribute
	  if no xml data is provided data is loaded from FIM/MIM
	  .EXAMPLE
	  Add-AttributeToClass classname "Person" -attributename "newattributename" -syntax "string" -singlevalue
	  .EXAMPLE
	  Add-AttributeToClass -mvdata "...xml data..." classname "Person" -attributename "newattributename" -syntax "string" -singlevalue
	  .PARAMETER mvdata
	  XML schema mv data
	  .PARAMETER classname
	  Class(object) name i FIM/MIM
	  .PARAMETER requiredInClass
	  If required in class
	  .PARAMETER attributename
	  Name of attribute
	  .PARAMETER syntax
	  Syntax(type) of attribute
	  .PARAMETER singlevalue
	  Singel value or multie value attribute
	  .PARAMETER indexable
	  Indexable or not in mv data
	#>
  [CmdletBinding()]
	param(
		[xml]$mvdata,
		[String]$classname,
		[String]$attributename,
		[ValidateSet("Reference","String","Binary","Bit","Integer")]
		[String]$syntax,
		[switch]$requiredInClass,
		[switch]$singlevalue,
		[switch]$indexable	
		)
	begin{
		$OIDTabel = 
		@{
			"Reference" = "1.3.6.1.4.1.1466.115.121.1.12";
			"String" = "1.3.6.1.4.1.1466.115.121.1.15";
			"Binary" = "1.3.6.1.4.1.1466.115.121.1.5";
			"Bit" = "1.3.6.1.4.1.1466.115.121.1.7";
			"Integer" = "1.3.6.1.4.1.1466.115.121.1.27";
		}
	}
	process{
		if($mvdata -eq $null){
			$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
			$mvdata = [xml]$MMSWebService.GetMVData(511)
		}
		
		#class to lower
		$classname = $classname.ToLower()
		
		$ns = New-Object System.Xml.XmlNamespaceManager($mvdata.NameTable)
		$ns.AddNamespace("dsml","http://www.dsml.org/DSML")
		$ns.AddNamespace("ms-dsml","http://www.microsoft.com/MMS/DSML")
		
		$schemaNode = $mvdata.DocumentElement.SelectSingleNode("//dsml:dsml/dsml:directory-schema", $ns)
		
		#Add if not exist to attribute-type
		if($mvdata.DocumentElement.SelectSingleNode("//dsml:dsml/dsml:directory-schema/dsml:attribute-type[@id='$attributename']", $ns) -eq $null){
		
			if(-NOT $syntax){
				Write-Error "Syntax value missing"
				return
			}
		
			$attributeElement = $mvdata.CreateElement("dsml:attribute-type", "http://www.dsml.org/DSML")
			$attributeElement.SetAttribute("id", $attributename)
			if($singlevalue){ $attributeElement.SetAttribute("single-value", "true") }
			if($indexable){ $attributeElement.SetAttribute("indexable", "true") }

			$NameElement = $mvdata.CreateElement("dsml:name", "http://www.dsml.org/DSML")
			$NameElement.InnerText = $attributename
			[void]$attributeElement.AppendChild($NameElement)
			
			$syntaxElement = $mvdata.CreateElement("dsml:syntax", "http://www.dsml.org/DSML")
			$syntaxElement.InnerText = $OIDTabel[$syntax]
			[void]$attributeElement.AppendChild($syntaxElement)
			
			[void]$schemaNode.AppendChild($attributeElement)
		}
		
		$ClassNode = $mvdata.DocumentElement.SelectSingleNode("//dsml:dsml/dsml:directory-schema/dsml:class[@id='$classname']", $ns)
		#Add if dont exist
		if($ClassNode -eq $null){
			$ClassNode = $mvdata.CreateElement("dsml:class", "http://www.dsml.org/DSML")
			$ClassNode.SetAttribute("id", $classname)
			$ClassNode.SetAttribute("type", "structural")
			$NameElement = $mvdata.CreateElement("dsml:name", "http://www.dsml.org/DSML")
			$NameElement.InnerText = $classname
			[void]$ClassNode.AppendChild($NameElement)
			
			$refNode = $schemaNode.Class[$schemaNode.Class.Count-1]
			
			[void]$schemaNode.InsertAfter($ClassNode,$refNode)
		}
		if($ClassNode -ne $null){
			if($ClassNode.SelectSingleNode("./dsml:attribute[@ref='#$attributename']", $ns) -eq $null){
				$ClassattributeElement = $mvdata.CreateElement("dsml:attribute", "http://www.dsml.org/DSML");
				$ClassattributeElement.SetAttribute("ref", "#$attributename")
				$ClassattributeElement.SetAttribute("required", $requiredInClass.ToString().ToLower())
				#after dsml:name class
				[void]$ClassNode.AppendChild($ClassattributeElement)
			}
		}
		
		return $mvdata
	}
}

function Add-AttributesToClass{
	<#
	  .SYNOPSIS
	  Add attributes to class
	  .DESCRIPTION
	  Add attributes to class
	  .EXAMPLE
	  Add-AttributesToClass -mvdata "...XML data..." -classname "Person" -attributes @( @{requiredInClass=$false;attributename="newattributename";syntax="string";singlevalue=$true;indexable=$true} )
	  .PARAMETER mvdata
	  Management agnet xml data
	  .PARAMETER classname
	  Name of class(type) in FIM/MIM
	  .PARAMETER attributes
	  Array of attribute objects that contains attribute propertes see, Add-AttributeToClass
	#>
  [CmdletBinding()]
	param(
		[xml]$mvdata,
		[String]$classname,
		[object[]]$attributes
		)
	process{		
		if($mvdata -eq $null){
			$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
			$mvdata = [xml]$MMSWebService.GetMVData(511)
		}
		
		foreach($attribute in $attributes){
			#$attribute
			#$mvdata = Add-AttributeToClass -mvdata $mvdata -classname $classname -requiredInClass:$attribute.requiredInClass -attributename $attribute.attributename -syntax $attribute.syntax -singlevalue:$attribute.singlevalue -indexable:$attribute.indexable
			$AttributeArgs = @{}
			$AttributeArgs.Add("mvdata",$mvdata)
			$AttributeArgs.Add("classname",$classname)
			foreach($attributevalue in $attribute){
				$AttributeArgs.Add($attributevalue.Name,$attributevalue.Value)
			}
			$mvdata = Add-AttributeToClass @AttributeArgs
			
		}
		return $mvdata
	}
}

function get-MIMSyncBackup{
	<#
	  .SYNOPSIS
	  Copy schema, MA config and Extensions directory to zip yyyyMMdd
	  .DESCRIPTION
	  Add attributes to class
	  .EXAMPLE
	  get-MIMSyncBackup "C:\backup\"
	  .PARAMETER DestPath
	  Destination path
	#>
  [CmdletBinding()]
	param(
		[String]$DestPath,
		[switch]$EncryptionkeysExport
	)
	process{
		$tempPath = $env:TEMP + "\$([datetime]::Now.ToString("yyyyMMdd"))"
		mkdir $tempPath 
		
		#Set-Alias svrexport "$FIM\Synchronization Service\Bin\svrexport.exe"
		
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		
		#key export
		if($EncryptionkeysExport){
			$par = Get-MIMParameters
			#Set-Alias miiskmu "$($par.Path)\Bin\miiskmu.exe"
			#miiskmu /q /e $DirDate\Encryption_keys\keyback.bin /u:$user $password
			#miiskmu /e "$tempPath\Encryption_keys.bin" /u:$user $password

			$StartName = (Get-WmiObject win32_service|?{$_.name -eq "FIMSynchronizationService"}).StartName
			write-host "Enter password for '$StartName'"
			$password = Read-Host
			Start-Process "$($par.Path)\Bin\miiskmu.exe" -Args "/e $tempPath\Encryption_keys.bin","/u:$StartName $password" -Verb runas -Wait
		}
		
		#511
		$MVXMLdata = [xml]$MMSWebService.GetMVData(511)
		"<export-mv-schema server='{0}' export-date='{1}'>{2}</export-mv-schema>" -f $env:computername, [datetime]::Now.ToString("s"), $MVXMLdata.'mv-data'.schema.InnerXml | Out-File -Encoding utf8 (join-path $tempPath ("mv-schema.xml"))

		#MA Data
		$maGuid = $null
		$maName = $null
		[void]$MMSWebService.GetMAGuidList([ref] $maGuid,[ref] $maName)
		
		#Set-Alias maexport "$FIM\Synchronization Service\Bin\maexport.exe"
		
		for($i=0;$i -lt $maGuid.Count;$i++){
			write-progress -id 1 -activity "Management Agent" -status ($maName[$i]) -percentComplete ($i/$maGuid.Count*100)
			
			#maexport $maName[$i] (join-path $tempPath ($maName[$i]+".xml"))
			
			$MAXmldata = [xml]$MMSWebService.ExportManagementAgent($maName[$i],$false,$true,[datetime]::Now.ToString("s"))
			$MAXmldata.Save((join-path $tempPath ($maName[$i]+".xml")))
		}
		
		#Extensions copy dlls
		$ExtensionsPath = get-MIMParameters | ? { ($_.Path -ne $null) } | % { $_.Path + "Extensions" }
		Add-Type -Assembly System.IO.Compression.FileSystem
		$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
		[System.IO.Compression.ZipFile]::CreateFromDirectory($ExtensionsPath,($tempPath+"\Extensions.zip"), $compressionLevel, $false)
		
		[System.IO.Compression.ZipFile]::CreateFromDirectory($tempPath,($DestPath + "\$([datetime]::Now.ToString("yyyyMMdd")).zip"), $compressionLevel, $false)
		rm $tempPath -Recurse
	}
}

function Get-MIMAdminGroups{
	
	begin{		

		$MIMParameters = Get-MIMParameters
		if(-NOT $InitialCatalog){
			$InitialCatalog = $MIMParameters.DBName
		}
		
		if(-NOT $SQLServerInstans){
			$SQLServerInstans = $MIMParameters.SQLServerInstans
		}
		
		
		$ConnectionString = "Data Source=$SQLServerInstans;Initial Catalog=$InitialCatalog;Integrated Security=SSPI;"
		$Connection = New-Object System.Data.SqlClient.SqlConnection ($ConnectionString)
		$Connection.Open()
		
		$AdminGroups = @("administrators_sid","operators_sid","account_joiners_sid","browse_sid","passwordset_sid")

	}
	
	process{
		$sqlcommand  = "SELECT {0} FROM mms_server_configuration (nolock)" -f ($AdminGroups -join ",")

		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($sqlcommand,$Connection)
		$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout

		$DataTable = New-Object system.Data.DataTable "Configsid"
		$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
		$RowCount = $Adapter.Fill($DataTable)
		
		$Adapter.Dispose()
		$SqlCmd.Dispose()
		
		foreach($groupname in $AdminGroups){
			$sid = New-Object System.Security.Principal.SecurityIdentifier -ArgumentList ($DataTable.Rows[0].($groupname)),0
			$AccountName = $sid.Translate([System.Security.Principal.NTAccount])
			@($AccountName,$groupname,($sid.ToString()))
		}
	}
	end{
		$Connection.Close()
	}
}

function Get-FIMServiceHosturl{
	begin{
		$MMSWebService = (new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService)
		
		$maGuid = $null
		$maName = $null
		[void]$MMSWebService.GetMAGuidList([ref] $maGuid,[ref] $maName)
	}
	process{
		$serviceHost = ""
		for($i=0;$i -lt $maGuid.Count;$i++){
			$MAXmldata = [xml]$MMSWebService.GetMaData($maGuid[$i],[uint32]::MaxValue,[uint32]::MaxValue,[uint32]::MaxValue)
			$MAdata = $MAXmldata.'ma-data'
			
			if($MAdata.category -eq "FIM"){
				$serviceHost = $MAdata.'private-configuration'.'fimma-configuration'.'connection-info'.serviceHost
				break
			}
		}
	}
	end{
		$serviceHost
	}
}

#Load PropertySheetBase from MIM install path
get-MIMParameters | ? { ($_.Path -ne $null) } |% { add-type -Path ($_.Path + "\UIShell\PropertySheetBase.dll") }
