
Add-PSSnapin FIMAutomation

function get-FIMConfig{
	<#
	  .SYNOPSIS
	  Get FIM config(s) object by ObjectID or DisplayName
	  .DESCRIPTION
	  Get FIM config(s) object by ObjectID or DisplayName
	  .EXAMPLE
	  get-FIMConfig -DisplayName "1 Generate E-mail" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  .EXAMPLE
	  get-FIMConfig -DisplayName "1 Generate" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  .EXAMPLE
	  get-FIMConfig -DisplayName "%Generate E-mail" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  .EXAMPLE
	  get-FIMConfig -DisplayName "1 Generate E-mail" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  .PARAMETER ObjectID
	  ObjectID of object
	  .PARAMETER DisplayName
	  Name of object
	  .PARAMETER ObjectType
	  Object type
	  .PARAMETER ExportFileName
	  Filepath for export file
	  .PARAMETER OnlyBaseResources
	  Toogel only export base object
	  .PARAMETER SelectOnMulit
	  Get option to select objects from grid list
	#>
  [CmdletBinding()]
	param(
		[parameter(Mandatory=$true)]
		[ValidateSet(
		"ManagementPolicyRule",
		"WorkflowDefinition",
		"Set",
		"SynchronizationRule",
		"ActivityInformationConfiguration",
		"ObjectVisualizationConfiguration",
		"Person",
		"Group",
		"AttributeTypeDescription",
		"BindingDescription",
		"ObjectTypeDescription",
		"Resource",
		"ma-data"
		)]
		[String]$ObjectType,
		[String]$XpathFilter,
		[String]$DisplayName,
		[Guid]$ObjectID,
		[switch]$ExportToFile,
		[String]$ExportFilePath = (pwd),
		[switch]$OnlyBaseResources,
		[switch]$SelectOnMulit
	)
	process{
		if([string]::IsNullOrEmpty($XpathFilter)){
			$XpathFilter = "/$ObjectType"
			if($ObjectID){
				$XpathFilter += "[ObjectID='$ObjectID']"
			}elseif($DisplayName){
				#$XpathFilter += "[DisplayName='$DisplayName']"
				$XpathFilter += "[(starts-with(DisplayName,'$DisplayName'))]"
			}
		}
		
		$FIMConfigParam = @{
			OnlyBaseResources = $OnlyBaseResources
			customConfig = $XpathFilter
		}

		$Objects = Export-FIMConfig @FIMConfigParam
		
		if($Objects.Count -gt 1 -AND $SelectOnMulit){
			$List = New-Object System.Collections.ArrayList
			foreach($obj in $Objects){

				[void]$List.Add((New-Object PSObject -Property @{
					ObjectIdentifier = $obj.ResourceManagementObject.ObjectIdentifier
					ObjectType = $obj.ResourceManagementObject.ObjectType
					DisplayName = ($obj.ResourceManagementObject.ResourceManagementAttributes[($obj.ResourceManagementObject.ResourceManagementAttributes.AttributeName.IndexOf("DisplayName"))].Value)
					Object = $obj
				}))
			}
			$Objects = ($List|Out-GridView -Title "select object(s)" -OutputMode Multiple).Object
		}
		if($ExportToFile){
			foreach($Object in $Objects){
				$DisplayName = ($Object.ResourceManagementObject.ResourceManagementAttributes[($Object.ResourceManagementObject.ResourceManagementAttributes.AttributeName.IndexOf("DisplayName"))].Value)
				if($DisplayName -eq $null){
					$DisplayName = $Object.ResourceManagementObject.ObjectIdentifier
				}
				$FilePath = (join-path $ExportFilePath "$DisplayName.xml")
				$Object | ConvertFrom-FIMResource -file $FilePath
			}

		}else{
			return $Objects
		}
	}
}

function set-FIMConfig{
	<#
	  .SYNOPSIS
	  set FIM config(s) from file or object
	  .DESCRIPTION
	  set FIM config(s) from file or object
	  .EXAMPLE
	  set-FIMConfig -Objects $testobj
	  .EXAMPLE
	  set-FIMConfig -Objects $testobj -NewGuids
	  .EXAMPLE
	  set-FIMConfig -Objects $testobj -NewGuids -ImportConfig
	  .PARAMETER Objects
	  Set of object(s) for import
	  .PARAMETER ImportFileName
	  Filenamepath for import file
	  .PARAMETER ExportFileName
	  Filenamepath for export file
	  .PARAMETER ReplaceAttributeValue
	  Replace dictionary for attribytes values
	  .PARAMETER NewGuids
	  Generate new guid for objects
	  .PARAMETER ImportConfig
	  Import config into portal
	#>
  [CmdletBinding()]
	param(
		[object[]]$Objects,
		[String]$ImportFileName,
		[String]$ExportFileName,
		[System.Collections.Hashtable]$ReplaceAttributeValue,
		[switch]$NewGuids,
		[switch]$ImportConfig
	)
	begin{
		if(-NOT $Objects){
			if(test-path $ImportFileName){
				$Objects = ConvertTo-FIMResource -File $ImportFileName
			}else{
				throw("Error file not found")
			}
		}
		
		#Add to list
		$GuidTranslate = New-Object 'system.collections.generic.dictionary[string,string]'
		$ObjectsList = New-Object System.Collections.ArrayList
		foreach($object in $Objects){
			if($NewGuids){
				[void]$GuidTranslate.Add($object.ResourceManagementObject.ObjectIdentifier,[System.Guid]::NewGuid().ToString())
			}else{
				[void]$GuidTranslate.Add($object.ResourceManagementObject.ObjectIdentifier,$object.ResourceManagementObject.ObjectIdentifier.Replace("urn:uuid:",""))
			}
		
			if($object.ResourceManagementObject.ObjectType -eq "ManagementPolicyRule"){
				[void]$ObjectsList.Add($object)
			}else{
				[void]$ObjectsList.Insert(0,$object)
			}
		}
		$NotAllowAttributes = "ObjectID CreatedTime Creator ObjectType"
	}
	
	process{
		$ImportObjects = New-Object System.Collections.ArrayList
		
		foreach($object in $ObjectsList){
			$ImportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
			$ImportObject.ObjectType = $object.ResourceManagementObject.ObjectType
			$ImportObject.SourceObjectIdentifier = $GuidTranslate[$object.ResourceManagementObject.ObjectIdentifier]
			
			#Replace not create
			if($object.ResourceManagementObject.ObjectIdentifier.Contains($ImportObject.SourceObjectIdentifier)){
				$ImportObject.TargetObjectIdentifier = $ImportObject.SourceObjectIdentifier
				$ImportObject.State = "Put"
			}
			
			$ImportObjectChanges = New-Object System.Collections.ArrayList
			
			foreach($Attribute in $object.ResourceManagementObject.ResourceManagementAttributes){
				if($NotAllowAttributes.IndexOf($Attribute.AttributeName) -eq -1){
					$ValueList = New-Object System.Collections.ArrayList
					if($Attribute.IsMultiValue){
						foreach($Value in $Attribute.Values){
							if($NewGuids -AND $Value.StartsWith("urn:uuid:")){
								$GuidString = ""
								if($GuidTranslate.TryGetValue($Value,[ref]$GuidString)){
									[void]$ValueList.Add($GuidString)
								}else{
									Write-Warning ("Replace guid missing for attribute: {0} Value: {1} - Using orginal value" -f $Attribute.AttributeName,$Value)
									[void]$ValueList.Add($Value.Replace("urn:uuid:",""))
								}
							}else{
								[void]$ValueList.Add($Value.Replace("urn:uuid:",""))
							}
						}
					}else{
						if($NewGuids -AND $Attribute.Value.StartsWith("urn:uuid:")){
							$GuidString = ""
							if($GuidTranslate.TryGetValue($Attribute.Value,[ref]$GuidString)){
								[void]$ValueList.Add($GuidString)
							}else{
								Write-Warning ("Replace guid missing for attribute: {0} Value: {1} - Using orginal value" -f $Attribute.AttributeName,$Attribute.Value)
								[void]$ValueList.Add($Attribute.Value.Replace("urn:uuid:",""))
							}
						}else{
							[void]$ValueList.Add($Attribute.Value.Replace("urn:uuid:",""))
						}
					}
					
					foreach($Value in $ValueList){
						$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
						$importChange.Operation = "Replace"
						$importChange.FullyResolved = $true
						$importChange.Locale = "Invariant"
						$importChange.AttributeName = $Attribute.AttributeName
						$TmpValue = ""
						#if($ReplaceAttributeValue -AND $ReplaceAttributeValue.TryGetValue($Attribute.AttributeName,[ref]$TmpValue)){
						if($ReplaceAttributeValue -AND $ReplaceAttributeValue.ContainsKey($Attribute.AttributeName)){
							$importChange.AttributeValue = $ReplaceAttributeValue[$Attribute.AttributeName]
						}else{
							$importChange.AttributeValue = $Value
						}
						[void]$ImportObjectChanges.Add($importChange)
					}
				}
			}
			$ImportObject.Changes = $ImportObjectChanges
			[void]$ImportObjects.Add($ImportObject)
		}
		if($ExportFileName){
			$ImportObjects | ConvertFrom-FIMResource -file $ExportFileName
		}
		if($ImportConfig){
			$ImportObjects | Import-FIMConfig
		}
	}
}

function get-BindingDescriptionConfig{
	param(
		$ObjectTypeName,
		$AttributeName,
		$Description = $AttributeName,
		$StringRegex
	)
	#get ObjectTypeDescription by DisplayName
	$ObjectTypeDescription = get-FIMConfig -ObjectType ObjectTypeDescription -XpathFilter "/ObjectTypeDescription[DisplayName='$ObjectTypeName']" -OnlyBaseResources
	
	if($ObjectTypeDescription -ne $null -AND $ObjectTypeDescription.Length -gt 0){
	
		$ObjectTypeGuid = $ObjectTypeDescription.ResourceManagementObject.ObjectIdentifier.Replace("urn:uuid:","")
		
		#get AttributeTypeDescription by name
		$AttributeTypeDescription = get-FIMConfig -ObjectType AttributeTypeDescription -XpathFilter "/AttributeTypeDescription[Name='$AttributeName']" -OnlyBaseResources
		
		if($AttributeTypeDescription -ne $null -AND $AttributeTypeDescription.Length -gt 0){
		
			$AttributeTypeGuid = $AttributeTypeDescription.ResourceManagementObject.ObjectIdentifier.Replace("urn:uuid:","")
			$AttributeDisplayName = ($AttributeTypeDescription.ResourceManagementObject.ResourceManagementAttributes[($AttributeTypeDescription.ResourceManagementObject.ResourceManagementAttributes.AttributeName.IndexOf("DisplayName"))].Value)

			#Create new BindingDescription
			$ImportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
			$ImportObject.ObjectType = "BindingDescription"
			$ImportObject.SourceObjectIdentifier = [System.Guid]::NewGuid().ToString()

			$ImportObjectChanges = New-Object System.Collections.ArrayList

			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "DisplayName"
			$importChange.AttributeValue = $AttributeDisplayName
			[void]$ImportObjectChanges.Add($importChange)

			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Description"
			$importChange.AttributeValue = $Description
			[void]$ImportObjectChanges.Add($importChange)			

			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Required"
			$importChange.AttributeValue = "False"
			[void]$ImportObjectChanges.Add($importChange)

			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "BoundObjectType"
			$importChange.AttributeValue = $ObjectTypeGuid
			[void]$ImportObjectChanges.Add($importChange)

			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "BoundAttributeType"
			$importChange.AttributeValue = $AttributeTypeGuid
			[void]$ImportObjectChanges.Add($importChange)
			
			if(-NOT [string]::IsNullOrEmpty($StringRegex)){
				$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
				$importChange.Operation = "Add"
				$importChange.FullyResolved = $true
				$importChange.Locale = "Invariant"
				$importChange.AttributeName = "StringRegex"
				$importChange.AttributeValue = $StringRegex
				[void]$ImportObjectChanges.Add($importChange)
			}

			$ImportObject.Changes = $ImportObjectChanges

			$ImportObject
		}
		else{
			throw "AttributeTypeDescription ($AttributeName) not found"
		}
	}else{
		throw "ObjectTypeDescription ($ObjectTypeName) not found"
	}
}

function get-AttributeTypeDescriptionConfig{
	param(
		$Name,
		$DisplayName = $Name,
		$Description = $Name,
		$AttributeType,
		$StringRegex,
		[switch]$Multivalued
	)

	#get AttributeTypeDescription by name
	$AttributeTypeDescription = get-FIMConfig -ObjectType AttributeTypeDescription -XpathFilter "/AttributeTypeDescription[Name='$Name']" -OnlyBaseResources
	
	if($AttributeTypeDescription -eq $null){

		#Create new AttributeTypeDescription
		$ImportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
		$ImportObject.ObjectType = "AttributeTypeDescription"
		$ImportObject.SourceObjectIdentifier = [System.Guid]::NewGuid().ToString()

		$ImportObjectChanges = New-Object System.Collections.ArrayList

		$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
		$importChange.Operation = "Add"
		$importChange.FullyResolved = $true
		$importChange.Locale = "Invariant"
		$importChange.AttributeName = "DisplayName"
		$importChange.AttributeValue = $DisplayName
		[void]$ImportObjectChanges.Add($importChange)
		
		$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
		$importChange.Operation = "Add"
		$importChange.FullyResolved = $true
		$importChange.Locale = "Invariant"
		$importChange.AttributeName = "Name"
		$importChange.AttributeValue = $Name
		[void]$ImportObjectChanges.Add($importChange)

		$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
		$importChange.Operation = "Add"
		$importChange.FullyResolved = $true
		$importChange.Locale = "Invariant"
		$importChange.AttributeName = "Description"
		$importChange.AttributeValue = $Description
		[void]$ImportObjectChanges.Add($importChange)			

		$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
		$importChange.Operation = "Add"
		$importChange.FullyResolved = $true
		$importChange.Locale = "Invariant"
		$importChange.AttributeName = "Multivalued"
		$importChange.AttributeValue = $Multivalued.ToString()
		[void]$ImportObjectChanges.Add($importChange)
	
		if(-NOT [string]::IsNullOrEmpty($StringRegex)){
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "StringRegex"
			$importChange.AttributeValue = $StringRegex
			[void]$ImportObjectChanges.Add($importChange)
		}

		$ImportObject.Changes = $ImportObjectChanges
		
		$ImportObject
	}
	else{
		throw "AttributeTypeDescription ($AttributeName) exist"
	}
}

function add-Admin{
	<#
	  .SYNOPSIS
	  Get FIM config(s) object by ObjectID or DisplayName
	  .DESCRIPTION
	  Get FIM config(s) object by ObjectID or DisplayName
	  .EXAMPLE
	  get-FIMConfig -DisplayName "1 Generate E-mail" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  get-FIMConfig -DisplayName "1 Generate" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  get-FIMConfig -DisplayName "%Generate E-mail" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  .EXAMPLE
	  get-FIMConfig -DisplayName "1 Generate E-mail" -ObjectType ManagementPolicyRule -SelectOnMulit -OnlyBaseResources
	  .PARAMETER ObjectID
	  ObjectID of object
	  .PARAMETER DisplayName
	  Name of object
	  .PARAMETER ObjectType
	  Object type
	  .PARAMETER ExportFileName
	  Filepath for export file
	  .PARAMETER OnlyBaseResources
	  Toogel only export base object
	  .PARAMETER SelectOnMulit
	  Get option to select objects from grid list
	#>
  [CmdletBinding()]
	param(
		[parameter(Mandatory=$true)]
		[String]$username,
		[String]$Server,
		[switch]$OverWriteSID
	)
	begin{
		#Get AD user
		$ADParam = @{
			Identity = $username
			Properties = "DisplayName"
		}
		if($Server){
			$ADParam.Add("Server",$Server)
		}
		
		$User = Get-ADUser @ADParam
		if(-NOT $User){ throw("User $username missing in AD") }
		
	}

	process{
	
		#Get AD Stuff
		$SidBinary = New-Object byte[] ($User.SID.BinaryLength)
		$User.SID.GetBinaryForm($SidBinary,0)
		$SidBase64 = [System.Convert]::ToBase64String($SidBinary, 0, $SidBinary.Length)
		$Domain = ($User.SID.Translate([System.Security.Principal.NTAccount])).Value.Split("\")[0]
		
		#Get FIM portal if exist
		$FIMUser = Export-FIMConfig -OnlyBaseResources -CustomConfig ("/Person[AccountName='{0}']" -f $User.SamAccountName)
		if(-NOT $FIMUser){
			#Create new user in Portal
			$ImportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
			$ImportObject.ObjectType = "Person"
			$ImportObject.SourceObjectIdentifier = [System.Guid]::NewGuid().ToString()			
			
			$ImportObjectChanges = New-Object System.Collections.ArrayList
			#AccountName
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "AccountName"
			$importChange.AttributeValue = $User.SamAccountName 
			[void]$ImportObjectChanges.Add($importChange)
			#DisplayName
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "DisplayName"
			$importChange.AttributeValue = $User.DisplayName 
			[void]$ImportObjectChanges.Add($importChange)
			#ObjectSID
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "ObjectSID"
			$importChange.AttributeValue = $SidBase64
			[void]$ImportObjectChanges.Add($importChange)
			#Domain
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Domain"
			$importChange.AttributeValue = $Domain
			[void]$ImportObjectChanges.Add($importChange)
			
			$ImportObject.Changes = $ImportObjectChanges
			$ImportObject | Import-FIMConfig
			
			$FIMUserObjectIdentifier = $ImportObject.TargetObjectIdentifier
			
		}else{
			if($FIMUser.Count -gt 1){ throw(("Mulit AccountName i FIM for {0}" -f $User.SamAccountName)) }
			$FIMObjectSID = ($FIMUser[0].ResourceManagementObject.ResourceManagementAttributes[($FIMUser[0].ResourceManagementObject.ResourceManagementAttributes.AttributeName.IndexOf("ObjectSID"))].Value)
			$FIMUserObjectIdentifier = $FIMUser[0].ResourceManagementObject.ObjectIdentifier.Replace("urn:uuid:","")
			
			#Cheack Domain?
			if($SidBase64 -ne $FIMObjectSID){
				if(-not $OverWriteSID){
					throw("SID dont match curent user")
				}else{
					$ImportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
					$ImportObject.ObjectType = "Person"
					$ImportObject.SourceObjectIdentifier = $FIMUserObjectIdentifier
					$ImportObject.TargetObjectIdentifier = $FIMUserObjectIdentifier
					$ImportObject.State = "Put"
					
					$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
					$importChange.Operation = "Replace"
					$importChange.FullyResolved = $true
					$importChange.Locale = "Invariant"
					$importChange.AttributeName = "ObjectSID"
					$importChange.AttributeValue = $SidBase64
					$ImportObject.Changes += $importChange
					
					$ImportObject | Import-FIMConfig
				}
			}
		}
		
		#Get FIM portal Admins group
		$Administrators = Export-FIMConfig -OnlyBaseResources -CustomConfig "/Set[DisplayName='Administrators']"
		if($Administrators -AND $Administrators.Count -eq 1){
			$indexExplicitMember = $Administrators[0].ResourceManagementObject.ResourceManagementAttributes.AttributeName.IndexOf("ExplicitMember")
			$AddUser = $true
			
			foreach($Value in $Administrators[0].ResourceManagementObject.ResourceManagementAttributes[$indexExplicitMember].Values){
				if($Value.Contains($FIMUserObjectIdentifier)){
					$AddUser = $false
					break
				}
			}
			
			if($AddUser){
			
				$ImportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
				$ImportObject.ObjectType = "Set"
				$ImportObject.SourceObjectIdentifier = $Administrators[0].ResourceManagementObject.ObjectIdentifier.Replace("urn:uuid:","")
				$ImportObject.TargetObjectIdentifier = $ImportObject.SourceObjectIdentifier
				$ImportObject.State = "Put"
				
				$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
				$importChange.Operation = "Add"
				$importChange.FullyResolved = $true
				$importChange.Locale = "Invariant"
				$importChange.AttributeName = "ExplicitMember"
				$importChange.AttributeValue = $FIMUserObjectIdentifier
				$ImportObject.Changes += $importChange
				
				$ImportObject | Import-FIMConfig
			}else{
				Write-Host -BackgroundColor Black -ForegroundColor Yellow "User exist in Administrators"
			}
		}
	}
}

function CheckMembershipLockedGroups-RemoveExplicitMember{

	$groups = Export-FIMConfig -only -custom "/Group[MembershipLocked = 'true' and ExplicitMember = /Person]"
	 
	if ($groups -ne $null) {
		foreach ($group in $groups)
		{
			$importObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
			$importObject.ObjectType = $group.ResourceManagementObject.ObjectType
			$importObject.TargetObjectIdentifier = $group.ResourceManagementObject.ObjectIdentifier
			$importObject.SourceObjectIdentifier = $group.ResourceManagementObject.ObjectIdentifier
			$importObject.State = "Put"
			
			$index = $group.ResourceManagementObject.ResourceManagementAttributes.AttributeName.IndexOf("ExplicitMember")
			foreach ($member in $group.ResourceManagementObject.ResourceManagementAttributes[$index].Values)
			{
				$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
				$importChange.Operation = "Delete"
				$importChange.AttributeName = "ExplicitMember"
				$importChange.AttributeValue = $member.Replace("urn:uuid:","")
				$importChange.FullyResolved = 1
				$importChange.Locale = "Invariant"
				$importObject.Changes += $importChange
			}
			$importObject | Import-FIMConfig
		}
	}
}

function new-GroupByCreteria{
	<#
	  .SYNOPSIS
	  Create new group in MIM portal
	  .DESCRIPTION
	  Create new security or distribution group in MIM portal
	  .EXAMPLE
	  new-GroupByCreteria -AccountName Security1 -OwnerAccountName andase -Type Security
	  .EXAMPLE
	  new-GroupByCreteria -AccountName Security1 -OwnerAccountName andase -Type Security -Domain domnain -Scope Global -Filter "/Person[(starts-with(AccountName,'A'))]" 
	  new-GroupByCreteria -AccountName Security1 -OwnerAccountName andase -Type Security -Domain domnain -Scope Global -Filter "/Person[GroupList-index = 'a']" 
	  .PARAMETER AccountName
	  AccountName of group (MailNickname for distribution group)
	  .PARAMETER OwnerAccountName
	  AccountName for Owner of group
	  .PARAMETER Domain
	  Domain for group
	  .PARAMETER Type
	  Type of group security or distribution
	  .PARAMETER Scope
	  Scope for group in domnain
	  .PARAMETER MembershipAddWorkflow
	  MembershipAddWorkflow is is None
	  .PARAMETER Filter
	  Creteria filter

	#>
  [CmdletBinding()]
	param(
		[parameter(Mandatory=$true)]
		[String]$AccountName,
		[parameter(Mandatory=$true)]
		[String]$OwnerAccountName,
		[String]$Domain = (Get-ADDomain).NetBIOSName,
		[ValidateSet("Security","Distribution")]
		[String]$Type = "Security",
		[ValidateSet("DomainLocal","Global","Universal")]
		[String]$Scope = "Global",
		[String]$MembershipAddWorkflow = "None",
		#[Switch]$MembershipLocked,
		[String]$Filter = "/Person[AccountName = '$OwnerAccountName']",
		[System.Collections.Hashtable]$SyncAttribute
	)
	begin{
		$MembershipLocked = $true
		$FilterXmlstring = @"
<Filter xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
Dialect="http://schemas.microsoft.com/2006/11/XPathFilterDialect" 
xmlns="http://schemas.xmlsoap.org/ws/2004/09/enumeration">{0}</Filter>
"@
		
		
		if($Type -eq "Security"){
			$MIMexist = Export-FIMConfig -OnlyBaseResources -CustomConfig "/Group[AccountName='$AccountName']"
			}else{
			$MIMexist = Export-FIMConfig -OnlyBaseResources -CustomConfig "/Group[MailNickname='$AccountName']"
		}
		try{$ADexist = Get-ADGroup $AccountName}Catch{}
		
		if($MIMexist -OR $ADexist){ throw("Group $AccountName exist") }
		
		$Owndexist = Export-FIMConfig -OnlyBaseResources -CustomConfig "/Person[AccountName='$OwnerAccountName']"
		if(-NOT $OwnerAccountName -AND $Owndexist.count -ne 1){ throw("Ownder $OwnerAccountName dont exist") }
		
		$DisplayedOwnerGuid = $Owndexist.ResourceManagementObject.ObjectIdentifier
	}
	process{
		
			#Create new Group in Portal
			$ImportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
			$ImportObject.ObjectType = "Group"
			$ImportObject.SourceObjectIdentifier = [System.Guid]::NewGuid().ToString()
			$FIMUserObjectIdentifier = $ImportObject.SourceObjectIdentifier
			
			$ImportObjectChanges = New-Object System.Collections.ArrayList
			if($Type -eq "Security"){
				#AccountName
				$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
				$importChange.Operation = "Add"
				$importChange.FullyResolved = $true
				$importChange.Locale = "Invariant"
				$importChange.AttributeName = "AccountName"
				$importChange.AttributeValue = $AccountName
				[void]$ImportObjectChanges.Add($importChange)
			}else{
				#MailNickname
				$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
				$importChange.Operation = "Add"
				$importChange.FullyResolved = $true
				$importChange.Locale = "Invariant"
				$importChange.AttributeName = "MailNickname"
				$importChange.AttributeValue = $AccountName
				[void]$ImportObjectChanges.Add($importChange)
			}
			#DisplayName
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "DisplayName"
			$importChange.AttributeValue = $AccountName 
			[void]$ImportObjectChanges.Add($importChange)
			
			#Filter
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Filter"
			$importChange.AttributeValue = $FilterXmlstring -f $filter
			[void]$ImportObjectChanges.Add($importChange)
			
			#MembershipLocked
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "MembershipLocked"
			$importChange.AttributeValue = $MembershipLocked
			[void]$ImportObjectChanges.Add($importChange)
			
			#DisplayedOwner
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "DisplayedOwner"
			$importChange.AttributeValue = $DisplayedOwnerGuid
			[void]$ImportObjectChanges.Add($importChange)
			
			#Owner
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Owner"
			$importChange.AttributeValue = $DisplayedOwnerGuid
			[void]$ImportObjectChanges.Add($importChange)
			
			#Type
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Type"
			$importChange.AttributeValue = $Type
			[void]$ImportObjectChanges.Add($importChange)
			
			#Scope
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Scope"
			$importChange.AttributeValue = $Scope
			[void]$ImportObjectChanges.Add($importChange)
			
			#Domain
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "Domain"
			$importChange.AttributeValue = $Domain
			[void]$ImportObjectChanges.Add($importChange)
			
			#MembershipAddWorkflow
			$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
			$importChange.Operation = "Add"
			$importChange.FullyResolved = $true
			$importChange.Locale = "Invariant"
			$importChange.AttributeName = "MembershipAddWorkflow"
			$importChange.AttributeValue = $MembershipAddWorkflow
			[void]$ImportObjectChanges.Add($importChange)
			
			#Sync attributes
			foreach($pair in $SyncAttribute.GetEnumerator()){
				$importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
				$importChange.Operation = "Add"
				$importChange.FullyResolved = $true
				$importChange.Locale = "Invariant"
				$importChange.AttributeName = $pair.Name
				$importChange.AttributeValue = $pair.Value
				[void]$ImportObjectChanges.Add($importChange)
			}
			
			$ImportObject.Changes = $ImportObjectChanges
			$ImportObject | Import-FIMConfig
	}
}

function new-NewGroupByCreteria{
	<#
	  .SYNOPSIS
	  Get count of Post Processing Request from MIM portal
	  .DESCRIPTION
	  Get count of Post Processing Request from MIM portal
	  .EXAMPLE
	  get-PostProcessingCount
	#>
	param(
		[parameter(Mandatory=$true)]
		[String]$AccountName,
		
		[parameter(Mandatory=$true)]
		[String]$DisplayName,
		
		[parameter(Mandatory=$true)]
		[String]$Filter,
		
		[String]$Domain,
		
		[ValidateSet("None","Custom","Owner Approval")]
		[String]$MembershipAddWorkflow = "None",
		
		[bool]$MembershipLocked = $True,
		
		[ValidateSet("DomainLocal","Global","Universal")]
		$Scope = "Universal",
		
		[parameter(Mandatory=$true)]
		[Guid]$Owner,
		
		[switch]$Commit
	)
	begin{		
		if(-NOT (Test-Path (join-path ($PSScriptRoot) Lithnet.ResourceManagement.Client.dll)))
		{
			$FileName = "Lithnet.nuget.zip"
			Invoke-WebRequest "https://www.nuget.org/api/v2/package/Lithnet.ResourceManagement.Client/" -OutFile $FileName
			Add-Type -AssemblyName System.IO.Compression.FileSystem
			$zip = [System.IO.Compression.ZipFile]::OpenRead((join-path ($PSScriptRoot) $FileName))
			$zip.Entries|?{$_.FullName.StartsWith("lib/net40/")}|%{[System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, (join-path ($PSScriptRoot) $_.Name), $true)}
			$zip.Dispose()
			rm $FileName
		}

		Add-Type -Path (join-path ($PSScriptRoot) Lithnet.ResourceManagement.Client.dll)
		$client = new-object Lithnet.ResourceManagement.Client.ResourceManagementClient
		
		$FilterXml = "<Filter xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" Dialect=`"http://schemas.microsoft.com/2006/11/XPathFilterDialect`" xmlns=`"http://schemas.xmlsoap.org/ws/2004/09/enumeration`">{0}</Filter>"
	}
	
	process{

		$Group = $client.CreateResource("Group")
		$Group.Attributes["AccountName"].Value = $AccountName
		$Group.Attributes["DisplayName"].Value = $DisplayName
		$Group.Attributes["Filter"].Value = $FilterXml -f $Filter

		$Group.Attributes["Domain"].Value = $Domain
		$Group.Attributes["MembershipAddWorkflow"].Value = $MembershipAddWorkflow
		$Group.Attributes["MembershipLocked"].Value = $MembershipLocked
		$Group.Attributes["Scope"].Value = $Scope
		$Group.Attributes["Type"].Value = $Type

		$Group.Attributes["Owner"].Value = $owner
		$Group.Attributes["DisplayedOwner"].Value = $owner
		
		if($Commit){ $Group.Save() }else{ $Group }

	}
}

function get-PostProcessingCount{
	<#
	  .SYNOPSIS
	  Get count of Post Processing Request from MIM portal
	  .DESCRIPTION
	  Get count of Post Processing Request from MIM portal
	  .EXAMPLE
	  get-PostProcessingCount
	#>
	param(
		[DateTime]$Date = [datetime]::Now.AddDays(-1),
		[String]$Address,
		[switch]$DetectAddress
		
	)
	begin{		
		if(-NOT (Test-Path (join-path ($PSScriptRoot) Lithnet.ResourceManagement.Client.dll)))
		{
			$FileName = "Lithnet.nuget.zip"
			Invoke-WebRequest "https://www.nuget.org/api/v2/package/Lithnet.ResourceManagement.Client/" -OutFile $FileName
			Add-Type -AssemblyName System.IO.Compression.FileSystem
			$zip = [System.IO.Compression.ZipFile]::OpenRead((join-path ($PSScriptRoot) $FileName))
			$zip.Entries|?{$_.FullName.StartsWith("lib/net40/")}|%{[System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, (join-path ($PSScriptRoot) $_.Name), $true)}
			$zip.Dispose()
			rm $FileName
		}

		Add-Type -Path (join-path ($PSScriptRoot) Lithnet.ResourceManagement.Client.dll)
		#$client = new-object Lithnet.ResourceManagement.Client.ResourceManagementClient
		
		if($DetectAddress){
			if(-NOT $Global:DetectedAddress){
				. (join-path (PWD) "MIM.syncservice.funtions.OP.ps1")
				$Global:DetectedAddress = Get-FIMServiceHosturl
			}
			$Address = $Global:DetectedAddress
		}

		if($Address){ $client = new-object Lithnet.ResourceManagement.Client.ResourceManagementClient $Address }else{ $client = new-object Lithnet.ResourceManagement.Client.ResourceManagementClient }
	}
	
	process{
		#$Request = $client.GetResources("/Request[RequestStatus = 'PostProcessing']")
		#$Request = $client.GetResources("/Request[RequestStatus = 'PostProcessing' or RequestStatus = 'Validating' or RequestStatus = 'Committed']")
		#$Request = $client.GetResources("/Request[(CreatedTime > op:subtract-dayTimeDuration-from-dateTime(fn:current-dateTime(), xs:dayTimeDuration('P1D'))) and ((RequestStatus = 'Validating') or (RequestStatus = 'PostProcessing') or (RequestStatus = 'Committed'))]")
		$DateString = $Date.ToUniversalTime().ToString("s")
		$Request = $client.GetResources("/Request[(CreatedTime > '$DateString') and ((RequestStatus = 'Validating') or (RequestStatus = 'PostProcessing') or (RequestStatus = 'Committed'))]")
		$Request.Count
	}
}

function start-SQLJob{
	<#
	  .SYNOPSIS
	  Start sql job
	  .DESCRIPTION
	  Start sql job by using Microsoft.SqlServer.SMO, need SQLAgentUserRole SQLAgentReaderRole SQLAgentOperatorRole permissions for msdb db
	  .EXAMPLE
	  start-SQLJob -ServerName "sql.server.namn" -JobName "sql job name" -StepName "step1"
	  .EXAMPLE
	  start-SQLJob -ServerName "SQL01.adm.namn.se" -JobName "FIM_TemporalEventsJob" -StepName "step1"
	#>
	param 
	( 
		[string]$ServerName,
		[string]$JobName,
		[string]$StepName,
		[switch]$DontWait
	)

	begin{	
		[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
	}
	
	process{
		try{
			$srv = New-Object Microsoft.SqlServer.Management.SMO.Server($ServerName)
			$job = $srv.jobserver.jobs[$JobName] 
			$step = $job.JobSteps[$StepName]
			
			$job.Start()

			if($DontWait){ 
				sleep 2
				$job.Refresh()
				return $job.CurrentRunStatus
			}else{

				do 
				{ 
					sleep 5
					$job.Refresh()
					
				}while($job.CurrentRunStatus.ToString() -ne "Idle")
				
				return $job.LastRunOutcome
			}
			
			return $job.LastRunOutcome
		}Catch{
			return $_
		}
	}
}