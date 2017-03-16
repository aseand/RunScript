
add-type -Path "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\PropertySheetBase.dll"

#NLog
Add-Type -Path ("C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Extensions\NLog.dll")
([NLog.LogManager]::Configuration) = new-object NLog.Config.XmlLoggingConfiguration("C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\conf\NLog.config.xml")
$global:logger = [NLog.LogManager]::GetLogger("script-run")

#build list of agents
$AgentsList = @{}
$Agents = get-wmiobject -Namespace root\MicrosoftIdentityIntegrationServer -class MIIS_ManagementAgent
0..($Agents.Count-1) | % {$AgentsList.Add($Agents[$_].Name,$_)}


function AnyInProgress{
	$returnValue = $false
	foreach($Agent in $Agents){
		if($Agent.RunStatus().ReturnValue -eq "in-progress") { 
			$logger.info("Agent {0} is in progress",$Agent.Name)
			
			$returnValue = $true 
			break
		}
	}
	
	return $returnValue
}

function Clearrunhistory{
	Param
	(
		[string]$DaysToKeepRunHistory = "15"
	)  
	$ScriptstartTime = get-date
	$logger.Info("Clear run history start")
	try
	{
		# Calculate the date to clear runs against
		$ClearRunsDate = [DateTime]::Now.AddDays(-$DaysToKeepRunHistory)

		# Get the WMI Object for MIIS_Server
		$miiserver = get-wmiobject -Namespace root\MicrosoftIdentityIntegrationServer -class MIIS_SERVER

		$Return = $miiserver.ClearRuns($ClearRunsDate)
		if($Return.ReturnValue -eq "success"){
			$logger.Info("Clear run history success")
		}else{
			$logger.Error("Clear run history error {0}",$ReturnValue)
		}
	}
	catch
	{
		$logger.Error("{0} {1}",$_.Exception.Message,$_.Exception.ItemName)
	}
	$logger.Info("Clear run history end execTime:{0}s",((get-date)-$ScriptstartTime).TotalSeconds )
}



Function MSSQLExecute{
	Param
	(
		[string]$ConnectionString,
		[string]$sqlcommand,
		[Switch]$ExecuteReader
	)

	try
	{
	if($sqlcommand.Length -eq 0 -OR $ConnectionString.Length -eq 0){
		$logger.Error("Error missing connectio or sommand -string")
		return
	}
	
	$Connection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
	$Connection.Open()

	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($sqlcommand,$Connection)
	$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout
	if($ExecuteReader)
	{
		$reader = $SqlCmd.ExecuteReader()
		[void]$reader.read()
		$reader.GetInt32(0)
	}
	else
	{
		$CountRows = $SqlCmd.ExecuteNonQuery()
		$CountRows
		$logger.Info("Row effected: {0}",$CountRows)
	}
	
	$SqlCmd.Dispose()
	$Connection.Close()
	$Connection.Dispose()
	}
	catch
	{
		$logger.Error("{0} {1}",$_.Exception.Message,$_.Exception.ItemName)
	}
}

function export-count{
	Param
    (
		[String]
		$Name
	)
	
	$countvalue = 0
	#some add value cant cast? IO to 
	try{
		$countvalue = [int]::Parse($Agents[$AgentsList[$Name]].NumExportAdd().ReturnValue)
		$countvalue += [int]$Agents[$AgentsList[$Name]].NumExportUpdate().ReturnValue
		$countvalue += [int]$Agents[$AgentsList[$Name]].NumExportDelete().ReturnValue
	}
	catch{}

	#$logger.Debug("{0} count {1}",$Name,$countvalue)
	return $countvalue
}

function import-count{
	Param
    (
		[String]
		$Name
	)
	
	$countvalue = [int]::Parse($Agents[$AgentsList[$Name]].NumImportAdd().ReturnValue)
	$countvalue += [int]::Parse($Agents[$AgentsList[$Name]].NumImportUpdate().ReturnValue)
	$countvalue += [int]::Parse($Agents[$AgentsList[$Name]].NumImportDelete().ReturnValue)
	return $countvalue
}

function stage-count{
	Param
    (
		[String]
		$Name
	)
	$xmlData = [xml]$Agents[$AgentsList[$Name]].RunDetails().ReturnValue

	$countvalue = 0
	try{
		$countvalue =  [int]::Parse(($xmlData.SelectNodes("//stage-add") | % { $_.InnerText })[0])
		$countvalue += [int]::Parse(($xmlData.SelectNodes("//stage-update") | % { $_.InnerText })[0])
		$countvalue += [int]::Parse(($xmlData.SelectNodes("//stage-rename") | % { $_.InnerText })[0])
		$countvalue += [int]::Parse(($xmlData.SelectNodes("//stage-delete") | % { $_.InnerText })[0])
		$countvalue += [int]::Parse(($xmlData.SelectNodes("//stage-delete-add") | % { $_.InnerText })[0])
	}
	catch{
		return -1
	}
	return $countvalue
}

function start-agent{
	Param
    (
		[String]
		$Name,
		
		[String]
		$Profile,
		
		[Switch]
        $ExportCount,
		
		[Switch]
        $ImportCount,
		
		[Switch]
        $StageCount,
		
		[Switch]
        $ReRunOnError,
		
		[Switch]
        $AsJob
	)
	
	if(!$AgentsList.Contains($Name)){
		$logger.Error("Missing agent {0} '{1}' AsJob:{3}",$Name,$Profile,$AsJob)
		throw("Missing agent {0} '{1}' AsJob:{3}" -f $Name,$Profile,$AsJob)
	}
	
	if ( ($ExportCount) -and ((export-count $Name) -le 0) ){
		$logger.Info("{0} not run '{1}' ExportCount: {2}",$Name,$Profile,(export-count $Name))
		return $false
	}
	
	if ( ($ImportCount) -and ((import-count $Name) -le 0) ){
		$logger.Info("{0} not run '{1}' ImportCount: {2}",$Name,$Profile,(import-count $Name))
		return $false
	}
	
	if ( ($StageCount) -and ((stage-count $Name) -le 0) ){
		$logger.Info("{0} not run '{1}' StageCount: {2}",$Name,$Profile,(stage-count $Name))
		return $false
	}
	
	if ($AsJob){
		if(-NOT $runspacesHandels){
			$Global:runspacesHandels = New-Object System.Collections.ArrayList
			$Global:RunspacePool = [RunspaceFactory]::CreateRunspacePool(1,20)
			$RunspacePool.Open()
		}

		$PowerShell = [PowerShell]::Create().AddScript($Function:RunAgent).AddArgument($Agents[$AgentsList[$Name]]).AddArgument($Name).AddArgument($Profile).AddArgument($AsJob).AddArgument($logger)#.AddArgument($Function:getRunHistory)
		$PowerShell.RunspacePool = $runspacepool
		$temp = New-Object -TypeName PSObject -Property @{
			PowerShell = $PowerShell 
			Runspace = $PowerShell.BeginInvoke()
		}
		[void]$runspacesHandels.Add($temp)
	}
	else{
		RunAgent $Agents[$AgentsList[$Name]] $Name $Profile $AsJob $logger #$Function:getRunHistory
	}
}

function RunAgent{
	Param(
		$Agent,
		$Name,
		$Profile,
		$AsJob,
		$logger
		#$getRunHistoryFunc
	)
	#$logger.Info("Start jobb {0} '{1}'",$Name,$Profile)
	$StartDate = Get-Date
	#$Agent = Get-WmiObject -Class MIIS_ManagementAgent -Namespace root/MicrosoftIdentityIntegrationServer -Filter ("Name='{0}'" -f $Name)
	$ReturnValue = $Agent.Execute($Profile).ReturnValue
	
	#need full import if transient-objects (only WaHSAn)
	if( $ReturnValue -eq "completed-transient-objects"){
		#$logger.error("{0} completed-transient-objects",$Name)
		if($Name -eq "WaHSAn" -AND $Profile -eq "Delta import"){
			$logger.warn("{0} completed-transient-objects change to Full import",$Name)
			$ReturnValue = $Agent.Execute("Full import").ReturnValue
		}
	}
	
	#stopped-database-connection-lost
	if($ReturnValue -eq 'stopped-database-connection-lost'){
		$logger.error("{0} stopped-database-connection-lost retry in 60s...",$Name)
		sleep 60
	
		$StartDate = Get-Date
		$ReturnValue = $Agent.Execute($Profile).ReturnValue
	}

	$logger.Info("{0} '{1}' {2} execTime: {3}s {4}",$Name,$Profile,$ReturnValue,((Get-Date)-$StartDate).TotalSeconds,("","AsJob")[[bool]$AsJob])
	if( $ReturnValue -ne 'success' ){
		
		#By wmiobject
		#$filter = "MaName = '{0}' and RunStartTime >= '{1}'" -F $Name,$StartDate.ToUniversalTime().ToString()
		#$RunHistory = get-wmiobject -Namespace root\MicrosoftIdentityIntegrationServer -class MIIS_RunHistory -filter $filter

		#By SQL
		$Connection = New-Object System.Data.SqlClient.SqlConnection "Data Source=server.name\FIM;Initial Catalog=FIMSynchronizationService;Integrated Security=SSPI;Connection Timeout=120;"
		$Connection.Open()

		$sqlcommand = ""
		$sqlcommand += "select top 1 rh.run_result,rh.start_date,rh.end_date,rh.current_step_number,rh.total_steps,sh.ma_connection_information_xml,sh.ma_discovery_errors_xml,sh.ma_counters_xml,sh.sync_errors_xml,sh.step_xml,sh.mv_retry_errors_xml,sh.flow_counters_xml"
		$sqlcommand += " ,STUFF(("
		$sqlcommand += " 			select CAST(',' + RTRIM(sod.cs_dn) AS VARCHAR(MAX))"
		$sqlcommand += " 			from [FIMSynchronizationService].[dbo].[mms_step_object_details] sod (nolock)"
		$sqlcommand += " 			where sod.step_history_id = sh.step_history_id"
		$sqlcommand += " 			FOR XML PATH (''))"
		$sqlcommand += " 			, 1, 1, '' "
		$sqlcommand += " ) as cs_dn"
		$sqlcommand += " from mms_run_history rh (nolock)"
		$sqlcommand += " join mms_step_history sh on rh.run_history_id = sh.run_history_id"
		$sqlcommand += " where rh.ma_id = (select ma_id from mms_management_agent where ma_name = '{0}')" -f $Name
		$sqlcommand += " and rh.start_date >= '{0}'" -f $StartDate.ToUniversalTime().ToString()
		$sqlcommand += " order by sh.start_date desc"
		

		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($sqlcommand,$Connection)
		$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout
		
		$DataTable = New-Object system.Data.DataTable "run_history"
		$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
		[void]$Adapter.Fill($DataTable)
		
		$Adapter.Dispose()
		$SqlCmd.Dispose()
		$Connection.Dispose()
		
		$extraLog = New-Object System.Text.StringBuilder
		foreach($row in $DataTable.Rows){
			[void]$extraLog.AppendFormat("`tResualt: {0}`n",$row.run_result)
			[void]$extraLog.AppendFormat("`tStart date: {0}`n",$row.start_date)
			[void]$extraLog.AppendFormat("`tEnd date: {0}`n",$row.end_date)
			[void]$extraLog.AppendFormat("`tStep: {0}/{1}`n",$row.current_step_number,$row.total_steps)
			
			[xml]$XMLDoc = "<run-history>"+$row.ma_connection_information_xml+$row.ma_discovery_errors_xml+$row.ma_counters_xml+$row.sync_errors_xml+$row.mv_retry_errors_xml+$row.flow_counters_xml+"</run-history>"
			$logger.Debug($XMLDoc.InnerXml)

			if($XMLDoc.'run-history'.'connection-result'.Count -gt 0){
				[void]$extraLog.AppendFormat("`tConnection information: {0} {1}`n",$XMLDoc.'run-history'.'connection-result',$XMLDoc.'run-history'.server)
			}
			
			if($XMLDoc.'run-history'.'ma-object-error'.ChildNodes.Count -gt 0){
				[void]$extraLog.AppendFormat("`tDiscovery errors({0}):`n",(,$XMLDoc.'run-history'.'ma-object-error'.Count))
				[void]$extraLog.AppendFormat("`t`t{0}`t{1}`t{2}`t{3}`n","error-type","entry-number","dn","anchor")
				$XMLDoc.'run-history'.'ma-object-error' | % { [void]$extraLog.AppendFormat("`t`t{0}`t{1}`t{2}`t{3}`n",$_.'error-type',$_.'entry-number',$_.dn,$_.anchor.'#text') }
			}
			
			if($XMLDoc.'run-history'.'synchronization-errors'.ChildNodes.Count -gt 0){
				[void]$extraLog.AppendFormat("`tMA synchronization errors({0}):`n",(,$XMLDoc.'run-history'.'synchronization-errors'.ChildNodes.Count))
				foreach($node in $XMLDoc.'run-history'.'synchronization-errors'.ChildNodes){
					[void]$extraLog.AppendFormat("`t`t{0}:`n",$node.Name)        
					if($node.'cs-guid')            	{[void]$extraLog.AppendFormat("`t`tcs-guid              :{0}`n",$node.'cs-guid')}           
					if($node.'dn')                 	{[void]$extraLog.AppendFormat("`t`tdn                   :{0}`n",$node.'dn')}   
					if($node.'first-occurred')      {[void]$extraLog.AppendFormat("`t`tfirst-occurred       :{0}`n",$node.'first-occurred')}
					if($node.'retry-count')         {[void]$extraLog.AppendFormat("`t`tretry-count          :{0}`n",$node.'retry-count')} 
					if($node.'date-occurred')      	{[void]$extraLog.AppendFormat("`t`tdate-occurred        :{0}`n",$node.'date-occurred')}     
					if($node.'error-type')    		{[void]$extraLog.AppendFormat("`t`terror-type           :{0}`n",$node.'error-type')}
					if($node.'algorithm-step'.'#text')		{[void]$extraLog.AppendFormat("`t`talgorithm-step       :{0}`n",$node.'algorithm-step'.'#text')}
					
					
					if($node.'cd-error'){
						[void]$extraLog.AppendFormat("`t`terror-code           :{0}`n",$node.'cd-error'.'error-code')
						[void]$extraLog.AppendFormat("`t`terror-literal        :{0}`n",$node.'cd-error'.'error-literal')
					}
					
					if($node.'change-not-reimported'){
						[void]$extraLog.Append("`t`tchange-not-reimported:`n")
						if($node.'change-not-reimported'.'delta'){
							[void]$extraLog.AppendFormat("`t`t{0} {1} {2}`n",$node.'change-not-reimported'.'delta'.'operation',$node.'change-not-reimported'.'delta'.'dn',$node.'change-not-reimported'.'delta'.'attr'.'name')
						}
					}
					if($node.'extension-error-info'){
						[void]$extraLog.AppendFormat("`t`textension-error-info:{0} {1} `n",$node.'extension-error-info'.'extension-name',$node.'extension-error-info'.'extension-callsite')
					}
				}
			}
			
			if($XMLDoc.'run-history'.'mv-retry-errors'.ChildNodes.Count -gt 0){
				[void]$extraLog.AppendFormat("`tMA MV retry errors: {0} {1}`n",(,$XMLDoc.'run-history'.'mv-retry-errors'.ChildNodes.Count))
				foreach($node in $XMLDoc.'run-history'.'mv-retry-errors'.ChildNodes){
					[void]$extraLog.AppendFormat("`t`{0}: {1}`n`n",$node.Name,$node.innertext)
				}
			}
		}
		$DataTable.Dispose()
		
		
		$logger.Error("{0} '{1}' {2} {3} {4}`nDetails`n{5}`n see RunDetails (debug)",$Name,$Profile,$ReturnValue,$StartDate,(Get-Date),$extraLog.ToString())
	}
} 

function getRunHistory{
	Param(
		$Name,
		$StartDate,
		$EndDate
	)
	
	$Connection = New-Object System.Data.SqlClient.SqlConnection "Data Source=server.name\FIM;Initial Catalog=FIMSynchronizationService;Integrated Security=SSPI;Connection Timeout=120;"
	$Connection.Open()

	$sqlcommand = ""
	$sqlcommand += "select top 1 rh.run_result,rh.start_date,rh.end_date,rh.current_step_number,rh.total_steps,sh.ma_connection_information_xml,sh.ma_discovery_errors_xml,sh.ma_counters_xml,sh.sync_errors_xml,sh.step_xml,sh.mv_retry_errors_xml,sh.flow_counters_xml"
	$sqlcommand += " ,STUFF(("
	$sqlcommand += " 			select CAST(',' + RTRIM(sod.cs_dn) AS VARCHAR(MAX))"
	$sqlcommand += " 			from [FIMSynchronizationService].[dbo].[mms_step_object_details] sod (nolock)"
	$sqlcommand += " 			where sod.step_history_id = sh.step_history_id"
	$sqlcommand += " 			FOR XML PATH (''))"
	$sqlcommand += " 			, 1, 1, '' "
	$sqlcommand += " ) as cs_dn"
	$sqlcommand += " from mms_run_history rh (nolock)"
	$sqlcommand += " join mms_step_history sh on rh.run_history_id = sh.run_history_id"
	$sqlcommand += " where rh.ma_id = (select ma_id from mms_management_agent where ma_name = '{0}')" -f $Name
	$sqlcommand += " and rh.start_date >= '{0}'" -f $StartDate.ToUniversalTime().ToString()
	if($EndDate){
		$sqlcommand += " and rh.end_date <= '{0}'" -f $EndDate.ToUniversalTime().ToString()
	}
	$sqlcommand += " order by sh.start_date desc"
	

	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($sqlcommand,$Connection)
	$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout
	
	$DataTable = New-Object system.Data.DataTable "run_history"
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
	[void]$Adapter.Fill($DataTable)
	
	$Adapter.Dispose()
	$SqlCmd.Dispose()
	$Connection.Dispose()
	
	$extraLog = New-Object System.Text.StringBuilder
	foreach($row in $DataTable.Rows){
		[void]$extraLog.AppendFormat("`tResualt: {0}`n",$row.run_result)
		[void]$extraLog.AppendFormat("`tStart date: {0}`n",$row.start_date)
		[void]$extraLog.AppendFormat("`tEnd date: {0}`n",$row.end_date)
		[void]$extraLog.AppendFormat("`tStep: {0}/{1}`n",$row.current_step_number,$row.total_steps)
		
		[xml]$XMLDoc = "<run-history>"+$row.ma_connection_information_xml+$row.ma_discovery_errors_xml+$row.ma_counters_xml+$row.sync_errors_xml+$row.mv_retry_errors_xml+$row.flow_counters_xml+"</run-history>"
		$logger.Debug($XMLDoc.InnerXml)
		
		if($XMLDoc.'run-history'.'connection-result'.Count -gt 0){
			[void]$extraLog.AppendFormat("`tConnection information: {0} {1}`n",$XMLDoc.'run-history'.'connection-result',$XMLDoc.'run-history'.server)
		}

		if($XMLDoc.'run-history'.'ma-object-error'.ChildNodes.Count -gt 0){
			[void]$extraLog.AppendFormat("`tDiscovery errors({0}):`n",(,$XMLDoc.'run-history'.'ma-object-error'.Count))
			[void]$extraLog.AppendFormat("`t`t{0}`t{1}`t{2}`t{3}`n","error-type","entry-number","dn","anchor")
			$XMLDoc.'run-history'.'ma-object-error' | % { [void]$extraLog.AppendFormat("`t`t{0}`t{1}`t{2}`t{3}`n",$_.'error-type',$_.'entry-number',$_.dn,$_.anchor.'#text') }
		}

		if($XMLDoc.'run-history'.'synchronization-errors'.ChildNodes.Count -gt 0){
			[void]$extraLog.AppendFormat("`tMA synchronization errors({0}):`n",(,$XMLDoc.'run-history'.'synchronization-errors'.Count))
			foreach($node in $XMLDoc.'run-history'.'synchronization-errors'.ChildNodes){
				[void]$extraLog.AppendFormat("`t`t{0}:`n",$node.Name)        
				if($node.'cs-guid')            	{[void]$extraLog.AppendFormat("`t`tcs-guid              :{0}`n",$node.'cs-guid')}           
				if($node.'dn')                 	{[void]$extraLog.AppendFormat("`t`tdn                   :{0}`n",$node.'dn')}   
				if($node.'first-occurred')      {[void]$extraLog.AppendFormat("`t`tfirst-occurred       :{0}`n",$node.'first-occurred')}
				if($node.'retry-count')         {[void]$extraLog.AppendFormat("`t`tretry-count          :{0}`n",$node.'retry-count')} 
				if($node.'date-occurred')      	{[void]$extraLog.AppendFormat("`t`tdate-occurred        :{0}`n",$node.'date-occurred')}     
				if($node.'error-type')    		{[void]$extraLog.AppendFormat("`t`terror-type           :{0}`n",$node.'error-type')}
				if($node.'algorithm-step'.'#text')		{[void]$extraLog.AppendFormat("`t`talgorithm-step       :{0}`n",$node.'algorithm-step'.'#text')}
				
				
				if($node.'cd-error'){
					[void]$extraLog.AppendFormat("`t`terror-code           :{0}`n",$node.'cd-error'.'error-code')
					[void]$extraLog.AppendFormat("`t`terror-literal        :{0}`n",$node.'cd-error'.'error-literal')
				}
				
				if($node.'change-not-reimported'){
					[void]$extraLog.Append("`t`tchange-not-reimported:`n")
					if($node.'change-not-reimported'.'delta'){
						[void]$extraLog.AppendFormat("`t`t{0} {1} {2}`n",$node.'change-not-reimported'.'delta'.'operation',$node.'change-not-reimported'.'delta'.'dn',$node.'change-not-reimported'.'delta'.'attr'.'name')
					}
				}
				if($node.'extension-error-info'){
					[void]$extraLog.AppendFormat("`t`textension-error-info:{0} {1} `n",$node.'extension-error-info'.'extension-name',$node.'extension-error-info'.'extension-callsite')
				}
			}
		}

		if($XMLDoc.'run-history'.'mv-retry-errors'.ChildNodes.Count -gt 0){
			[void]$extraLog.AppendFormat("`tMA MV retry errors: {0} {1}`n",(,$XMLDoc.'run-history'.'mv-retry-errors'.Count))
			foreach($node in $XMLDoc.'run-history'.'mv-retry-errors'.ChildNodes){
				[void]$extraLog.AppendFormat("`t`{0}: {1}`n`n",$node.Name,$node.innertext)
			}
		}
	}
	$DataTable.Dispose()
	$extraLog.ToString()
}

function getRunHistoryold{
	Param(
		$Name,
		$StartDate,
		$EndDate
	)
	if(-NOT $EndDate){
		$filter = "MaName = '{0}' and RunStartTime >= '{1}'" -F $Name,$StartDate.ToUniversalTime().ToString()
	}
	else{
		 $filter = "MaName = '{0}' and RunStartTime >= '{1}' and RunEndTime <= '{2}'" -F $Name,$StartDate.ToUniversalTime().ToString(),$EndDate.ToUniversalTime().ToString()
	}
	#$RunHistory = get-wmiobject -Namespace root\MicrosoftIdentityIntegrationServer -class MIIS_RunHistory -filter ("MaName = '{0}' and RunStartTime >= '{1}'" -F $Name,([datetime]::ParseExact($StartDate,"yyyy-MM-dd hh:mm:ss",$null).ToUniversalTime()))
	
	$RunHistory = get-wmiobject -Namespace root\MicrosoftIdentityIntegrationServer -class MIIS_RunHistory -filter $filter
	
	if( $RunHistory -eq $null ){
		"Null " + $filter
	}
	else{
		$List = New-Object System.Collections.ArrayList
		$RunHistory | % {[void]$List.add([xml]$_.RunDetails().ReturnValue)}
		#($RunHistory | % {[xml]$_.RunDetails().ReturnValue})
		#,$List
		
		$extraLog = New-Object System.Text.StringBuilder
		
		#List XML to string
		$List | % {
			$logger.Debug( $_.InnerXml  )

			#write-host $_.SelectNodes("/run-history/run-details/step-details").Count
			foreach($stepdetails in $_.SelectNodes("/run-history/run-details/step-details"))
			{
				#$logger.Debug( $_  );
				#$stepdetails | gm
				#write-host "step-number "$stepdetails."step-number"
				[void]$extraLog.Append("`tStep-number: "+$stepdetails."step-number")
				
				#step-result
				[void]$extraLog.Append("`n`t"+$stepdetails["step-result"].InnerText)
				
				#ma-discovery-errors
				foreach($ChildNode in $stepdetails["ma-discovery-errors"]["ma-object-error"].ChildNodes)
				{
					if($ChildNode -ne $null){
						[void]$extraLog.Append("`n`t" + $ChildNode.Name + " : " + $ChildNode.InnerText )
					}
				}
				#synchronization-errors
				#import-error
				foreach($ChildNode in $stepdetails["synchronization-errors"].ChildNodes)
				{
					if($ChildNode -ne $null){
					
						#if($ChildNode.dn -ne $null){
						#	[void]$extraLog.Append("`n`t" + $ChildNode.dn + "`n")
						#}
					
						if($ChildNode["error-type"] -ne $null){
							[void]$extraLog.Append("`n`t" + $ChildNode["error-type"].InnerText + "`n")
						}
						
						#rules-error-info
						if($ChildNode["rules-error-info"] -ne $null -AND $ChildNode["rules-error-info"]["context"] -ne $null){
							[void]$extraLog.Append("`n`t" + $ChildNode["rules-error-info"]["context"].'ma-name' + " " + $ChildNode["rules-error-info"]["context"].dn)
						}
						
						#extension-error-info
						if($ChildNode["extension-error-info"] -ne $null){
							[void]$extraLog.Append("`t"+$ChildNode["extension-error-info"]["extension-name"].InnerText + "`n`t" + $ChildNode["extension-error-info"]["call-stack"].InnerText)
						}
						
						#change-not-reimported
						if($ChildNode["change-not-reimported"] -ne $null){
							if($ChildNode["change-not-reimported"]["delta"] -ne $null){
								[void]$extraLog.Append("`t"+$ChildNode["change-not-reimported"]["delta"].operation)
								[void]$extraLog.Append(" "+$ChildNode["change-not-reimported"]["delta"].dn)
								[void]$extraLog.Append(" "+$ChildNode["change-not-reimported"]["delta"]["attr"].name+"`n")
							}
						}
						

					}
				}
				#export-errors
				foreach($ChildNode in $stepdetails["synchronization-errors"]["export-error"].ChildNodes)
				{
					if($ChildNode -ne $null){
						[void]$extraLog.Append("`n`t" + $ChildNode.dn)
					
						if($ChildNode["error-type"] -ne $null){
							[void]$extraLog.Append("`n`t" + $ChildNode["error-type"].InnerText + "`n")
						}
						#cd-errors
						foreach($cderrors in $ChildNode["cd-error"].ChildNodes){
							if($cderrors -ne $null){
								[void]$extraLog.Append("`n`t" + $cderrors.Name + " : " + $cderrors.InnerText)
							}
						}

						[void]$extraLog.Append("`n`t" + $ChildNode.Name + " : " + $ChildNode.InnerText)
					}
				}
				
				#mv-retry-errors
				foreach($ChildNode in $stepdetails["mv-retry-errors"].ChildNodes)
				{
					if($ChildNode -ne $null){
						[void]$extraLog.Append("`n`t" + $ChildNode.Name + " : " + $ChildNode.InnerText)
					}
				}
				[void]$extraLog.Append("`n")
			}
		}
		$extraLog.ToString()
	}
}

function WriteXmlToScreen{
	Param
    (
		[xml]$xml
	)
	
    $StringWriter = New-Object System.IO.StringWriter
    $XmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter
    $XmlWriter.Formatting = [System.Xml.Formatting]::Indented
    $xml.WriteTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()
	
	#Return string
    $StringWriter.ToString()
	
	#Clean up
	$StringWriter.Dispose()
	$XmlWriter.Dispose()
}

function SaveChangeCS{
	Param
    (
		[DateTime]$StartTime,
		[switch]$export
	)
	
	#Save all change row (hologram and deltas) from CS 
	$Connection = New-Object System.Data.SqlClient.SqlConnection "Data Source=server.name\FIM;Initial Catalog=FIMSynchronizationService;Integrated Security=SSPI;Connection Timeout=120;"
	$Connection.Open()
	#Microsoft.DirectoryServices.MetadirectoryServices.UI.PropertySheetBase
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
	
	$exportType = 402784264
	if($export){
		$sqlcommand = "select * from mms_connectorspace (nolock) where last_export_modification_date > '{0}' and export_operation = 2" -f $StartTime.ToUniversalTime()
		$exportType = 402784257
		#$exportType = [uint64]::MaxValue
	}else{
		$sqlcommand = "select * from mms_connectorspace (nolock) where last_import_modification_date > '{0}' and import_operation = 2" -f $StartTime.ToUniversalTime()
	}
	#$logger.debug($sqlcommand)
	
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($sqlcommand,$Connection)
	$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout

	$DataTable = New-Object system.Data.DataTable "mms_connectorspace"
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
	$Count = $Adapter.Fill($DataTable)
	$Adapter.Dispose()
	$SqlCmd.Dispose()

	#Write to disk
	#$DataTable.WriteXml("E:\MIM-change-data\mms_connectorspace.{0}.xml" -f [datetime]::Now.ToString("yyyyMMdd.HHmmss"))
	#$List = New-Object System.Collections.ArrayList
	#foreach($row in $DataTable.Rows){
	#	[void]$List.Add($row["object_id"])
	#}
	#if($List.Count -gt 0){
	#	#add-type -Path "C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\PropertySheetBase.dll"
	#	$MMSWebService = new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService
	#	[xml]$obj = $MMSWebService.GetCSObjects($List.ToArray(),$List.Count,$exportType,17,0,0)
	#	$obj.Save("E:\MIM-change-data\mms_connectorspace.{0}.xml" -f [datetime]::Now.ToString("yyyyMMdd.HHmmss"))
	#	#
	#}
	
	#add Column
	[void]$DataTable.Columns.Add("delta_gzip",[System.Byte[]])
	
	$MMSWebService = new-object Microsoft.DirectoryServices.MetadirectoryServices.UI.WebServices.MMSWebService
	foreach($row in $DataTable.Rows){
		[xml]$obj = $MMSWebService.GetCSObjects($row["object_id"].ToString(),1,$exportType,17,0,0)
		#$row["hologram"] = $null

		if($export){ $Data = [System.Text.Encoding]::Default.GetBytes($obj.'cs-objects'.'cs-object'.'unapplied-export'.InnerXml) }
		else{ $Data = [System.Text.Encoding]::Default.GetBytes($obj.'cs-objects'.'cs-object'.'pending-import'.InnerXml) }
		
		if($Data.Length -gt 0){
			$MemoryStream = New-Object System.IO.MemoryStream
			$GZipStream = New-Object System.IO.Compression.GZipStream($MemoryStream, [System.IO.Compression.CompressionMode]::Compress)
			$GZipStream.Write($Data,0,$Data.Length)
			$GZipStream.Close()
			$row["delta_gzip"] = $MemoryStream.ToArray()
		
			$GZipStream.Dispose()
			$MemoryStream.Dispose()
		}
		else{
			$row["delta_gzip"] = [System.DBNull]::Value
			
			[xml]$obj = $MMSWebService.GetCSObjects($row["object_id"].ToString(),1,[System.Int64]::MaxValue,17,0,0)
			$obj.Save("E:\MIM-change-data\mms_connectorspace.{0}.{1}.xml" -f $row["object_id"].ToString(),[datetime]::Now.ToString("yyyyMMdd.HHmmss"))
		}
	}
	
	$Connection = New-Object System.Data.SqlClient.SqlConnection "Data Source=server.name\FIM;Initial Catalog=miisExtraFunctions;Integrated Security=SSPI;"
	$Connection.Open()
	
	$SqlBulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy $Connection
	$SqlBulkCopy.BulkCopyTimeout = $Connection.ConnectionTimeout
	$SqlBulkCopy.DestinationTableName = "mms_connectorspace_history"

	foreach ($colum in $DataTable.Columns){
		[void]$SqlBulkCopy.ColumnMappings.Add($colum.ColumnName, $colum.ColumnName)
	}
	
	$SqlBulkCopy.WriteToServer($DataTable)
	$SqlBulkCopy.Close()
	
	$DataTable.Dispose()
	
	$logger.Info("mms_connectorspace change rows {0} export:{1}",$Count,$export)
}

function GetWaitLogRemove-jobs{
	if($runspacesHandels.Count -gt 0){
		$runspacesHandels | % {
			while(-NOT $_.Runspace.IsCompleted){ 
				sleep -Milliseconds 100
			}
			$error = $_.powershell.EndInvoke($_.Runspace)
			if($error){
				$logger.Error("EndInvoke error {0} {1}",$error.count, $error)
				$error | % {
					if($_){
						$logger.Error($_.BaseObject.GetType())
						#$global:logger.Error($_.Exception.Message)
						#$global:logger.Error($_.Exception.Source)
						#$global:logger.Error($_.Exception.StackTrace)
						#$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
					}
				}
			}
			$_.powershell.dispose()
			$_.powershell = $null
			$_.Runspace = $null
		}
		$runspacesHandels.Clear()
		rv runspacesHandels
	}
}