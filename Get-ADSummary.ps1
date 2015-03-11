<#
    NAME: Get-ADStats.ps1
 	AUTHOR: Sean Metcalf	
 	AUTHOR EMAIL: SeanMetcalf@MetcorpConsulting.com
 	CREATION DATE: 2/27/2013
	LAST MODIFIED DATE: 3/01/2013
 	LAST MODIFIED BY: Sean Metcalf
 	INTERNAL VERSION: 01.13.03.01.21
	RELEASE VERSION: 0.1.1
#>

Param 
    (
		
	[alias("UserLA","ULA")]
	[int] $UserLogonAge = 365, 
    
	[alias("UserPA","UPA")]
	[int] $UserPasswordAge,  
	
	[alias("ComputerLA","CLA")]
	[int] $ComputerLogonAge = 90, 
	
	[alias("ComputerPA","CPA")]
	[int] $ComputerPasswordAge = 90
    
    )

###############################
# Set Environmental Variables #
###############################

$DomainDNS = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name #Get AD Domain (lightweight & fast method)    
$CurrentUserName = $env:UserName
$LogDir = "C:\temp\Logs\"
IF (!(Test-Path $LogDir)) {new-item -type Directory -path $LogDir}  
$LogFileName = "GetADStats-$DomainDNS-$TimeVal.log"
$LogFile = $LogDir + $LogFileName

$DateTime = Get-Date #Get date/time
$ShortDate = Get-Date -format d
$UserStaleDate = $DateTime.AddDays(-$UserPasswordAge)
$NeverLoggedOnDate = $DateTime.AddDays(-1000)

$1Days = $DateTime.AddDays(-1)
$2Days = $DateTime.AddDays(-2)
$3Days = $DateTime.AddDays(-3)
$4Days = $DateTime.AddDays(-4)
$5Days = $DateTime.AddDays(-5)
$6Days = $DateTime.AddDays(-6)
$7Days = $DateTime.AddDays(-7)
$30Days = $DateTime.AddDays(-30)
$45Days = $DateTime.AddDays(-45)
$60Days = $DateTime.AddDays(-60)
$90Days = $DateTime.AddDays(-90)
$120Days = $DateTime.AddDays(-120)
$180Days = $DateTime.AddDays(-180)

###########################################
# Import Powershell Elements #
###########################################
import-module ActiveDirectory
import-module GroupPolicy 

##################
# Start Logging  #
##################
# Log all configuration changes shown on the screen during run-time in a transcript file.  This
# inforamtion can be used for troubleshooting if necessary
Write-Verbose "Start Logging to $LogFile  `r "

# Start-Transcript $LogFile -force

## Process Start Time
$ProcessStartTime = Get-Date
Write-Verbose " `r "
write-Verbose "Script initialized by $CurrentUserName and started processing at $ProcessStartTime `r "
Write-Verbose " `r "

#############################################
# Get Active Directory Forest & Domain Info #  20120201-15
#############################################
# Get Forest Info
write-output "Gathering Active Directory Forest Information..." `r
Write-Verbose "Running Get-ADForest Powershell command `r"
$ADForestInfo =  Get-ADForest

$ADForestApplicationPartitions = $ADForestInfo.ApplicationPartitions
$ADForestCrossForestReferences = $ADForestInfo.CrossForestReferences
$ADForestDomainNamingMaster = $ADForestInfo.DomainNamingMaster
$ADForestDomains = $ADForestInfo.Domains
$ADForestForestMode = $ADForestInfo.ForestMode
$ADForestGlobalCatalogs = $ADForestInfo.GlobalCatalogs
$ADForestName = $ADForestInfo.Name
$ADForestPartitionsContainer = $ADForestInfo.PartitionsContainer
$ADForestRootDomain = $ADForestInfo.RootDomain
$ADForestSchemaMaster = $ADForestInfo.SchemaMaster
$ADForestSites = $ADForestInfo.Sites
$ADForestSPNSuffixes = $ADForestInfo.SPNSuffixes
$ADForestUPNSuffixes = $ADForestInfo.UPNSuffixes

# Get Domain Info
write-output "Gathering Active Directory Domain Information..." `r
Write-Verbose "Performing Get-ADDomain powershell command `r"
$ADDomainInfo = Get-ADDomain

$ADDomainAllowedDNSSuffixes = $ADDomainInfo.ADDomainAllowedDNSSuffixes
$ADDomainChildDomains = $ADDomainInfo.ChildDomains
$ADDomainComputersContainer = $ADDomainInfo.ComputersContainer
$ADDomainDeletedObjectsContainer = $ADDomainInfo.DeletedObjectsContainer
$ADDomainDistinguishedName = $ADDomainInfo.DistinguishedName
$ADDomainDNSRoot = $ADDomainInfo.DNSRoot
$ADDomainDomainControllersContainer = $ADDomainInfo.DomainControllersContainer
$ADDomainDomainMode = $ADDomainInfo.DomainMode
$ADDomainDomainSID = $ADDomainInfo.DomainSID
$ADDomainForeignSecurityPrincipalsContainer = $ADDomainInfo.ForeignSecurityPrincipalsContainer
$ADDomainForest = $ADDomainInfo.Forest
$ADDomainInfrastructureMaster = $ADDomainInfo.InfrastructureMaster
$ADDomainLastLogonReplicationInterval = $ADDomainInfo.LastLogonReplicationInterval
$ADDomainLinkedGroupPolicyObjects = $ADDomainInfo.LinkedGroupPolicyObjects
$ADDomainLostAndFoundContainer = $ADDomainInfo.LostAndFoundContainer
$ADDomainName = $ADDomainInfo.Name
$ADDomainNetBIOSName = $ADDomainInfo.NetBIOSName
$ADDomainObjectClass = $ADDomainInfo.ObjectClass
$ADDomainObjectGUID = $ADDomainInfo.ObjectGUID
$ADDomainParentDomain = $ADDomainInfo.ParentDomain
$ADDomainPDCEmulator = $ADDomainInfo.PDCEmulator
$ADDomainQuotasContainer = $ADDomainInfo.QuotasContainer
$ADDomainReadOnlyReplicaDirectoryServers = $ADDomainInfo.ReadOnlyReplicaDirectoryServers
$ADDomainReplicaDirectoryServers = $ADDomainInfo.ReplicaDirectoryServers
$ADDomainRIDMaster = $ADDomainInfo.RIDMaster
$ADDomainSubordinateReferences = $ADDomainInfo.SubordinateReferences
$ADDomainSystemsContainer = $ADDomainInfo.SystemsContainer
$ADDomainUsersContainer = $ADDomainInfo.UsersContainer			
$DomainDNS = $ADDomainDNSRoot

$ForestDNSZoneNC = $ADForestApplicationPartitions[0]
$DomainDNSZoneNC = $ADForestApplicationPartitions[0]
$SchemaNC = "CN=Schema,CN=Configuration,$ADDomainDistinguishedName"
$ConfigurationNC = "CN=Configuration,$ADDomainDistinguishedName"

###################################
# Create Schema Version Hashtable # 20130215-13
###################################
Write-Verbose "Create Schema Version hashtable `r "
$SchemaVersionTable = 
@{ 
    "13" = "Windows 2000 Schema" ; 
    "30" = "Windows 2003 Schema"; 
    "31" = "Windows 2003 R2 Schema" ;
    "44" = "Windows 2008 Schema" ; 
    "47" = "Windows 2008 R2 Schema" ; 
    "51" = "Windows Server 8 BETA Schema" ;

    "4397"  = "Exchange 2000 RTM Schema" ; 
    "4406"  = "Exchange 2000 SP3 Schema" ;
    "6870"  = "Exchange 2003 RTM Schema" ; 
    "6936"  = "Exchange 2003 SP3 Schema" ; 
    "10637"  = "Exchange 2007 RTM Schema" ;
    "11116"  = "Exchange 2007 RTM Schema" ; 
    "14622"  = "Exchange 2007 SP2 & Exchange 2010 RTM Schema" ; 
    "14625"  = "Exchange 2007 SP3" ;
    "14726" = "Exchange 2010 SP1 Schema" 
 }

# Use Add method to add updates to Schema Version Hashtable (SVH)
$SchemaVersionTable.Add("14732", "Exchange 2010 SP2 Schema")
$SchemaVersionTable.Add("56", "Windows Server 2012 Forest Functional Mode")
Write-Verbose " `r "

################################
# Get AD Schema Version Number # 20111029-14
################################
write-Output "Checking Schema version on the PDC Emulator ($ADDomainPDCEmulator) `r "
$ADSchemaInfo = Get-ADObject "cn=schema,cn=configuration,$ADDomainDistinguishedName" -properties objectversion
$ADSchemaVersion = $ADSchemaInfo.objectversion
$ADSchemaVersionName = $SchemaVersionTable.Get_Item("$ADSchemaVersion")
Write-Output "The current AD Schema Version is $ADSchemaVersion which is $ADSchemaVersionName `r "
Write-Output "  `r "

######################################
# Get Exchange Schema Version Number #
######################################
write-Output "Checking Exchange Schema version on the PDC" `r
$ExchangeSchemaInfo = Get-ADObject "cn=ms-exch-schema-version-pt,cn=Schema,cn=Configuration,$ADDomainDistinguishedName" -properties rangeUpper
$ExchangeSchemaVersion = $ExchangeSchemaInfo.rangeUpper
$ExchangeSchemaVersionName = $SchemaVersionTable.Get_Item("$ExchangeSchemaVersion")
Write-Output "The current Exchange Schema Version is $ExchangeSchemaVersion which is $ExchangeSchemaVersionName `r "

#############################
# Get AD Instantiation Date #
#############################
# Code from: http://blogs.technet.com/b/heyscriptingguy/archive/2012/01/05/how-to-find-active-directory-schema-update-history-by-using-powershell.aspx
write-output "Checking Active Directory Creation Date... " `r
write-output "Displaying AD partition creation information " `r

$ADInstatiationObject = Get-ADObject -SearchBase (Get-ADForest).PartitionsContainer `
-LDAPFilter "(&(objectClass=crossRef)(systemFlags=3))" `
-Property dnsRoot, nETBIOSName, whenCreated | Sort-Object whenCreated 
$ADInstatiationObject |  Format-Table DNSRoot, NETBIOSName, WhenCreated -AutoSize
$ADInstatiationObjectDNSRoot = $ADInstatiationObject.DNSRoot
$ADInstatiationObjectWhenCreated = $ADInstatiationObject.WhenCreated

##############################
# Get Domain Password Policy # 20111027-20 
##############################
Write-Verbose "Get Domain Password Policy `r"
$DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
[int]$UserMaxPasswordAge = ($DomainPasswordPolicy.MaxPasswordAge).Days


#########################
# Get Tombstone Setting # 20111027-20
#########################
Write-Verbose "Get Tombstone Setting `r"
$DirectoryServicesConfigPartition = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$ADDomainDistinguishedName" `
 -Partition "CN=Configuration,$ADDomainDistinguishedName" -Properties *
$TombstoneLifetime = $DirectoryServicesConfigPartition.tombstoneLifetime
Write-Verbose "Active Directory's Tombstone Lifetime is set to $TombstoneLifetime days `r "

###########################
# Get AD Last Backup Date #
###########################
## Code based on: http://blogs.technet.com/b/heyscriptingguy/archive/2013/01/18/use-a-powershell-script-to-show-active-directory-backup-status-info.aspx?Redirected=true
Write-Verbose "Checking when the directory data was last backed up. `r "
IF (!$BackupDate) {$BackupAge = ($TombstoneLifetime -1)}
IF ($ADBAckupStatus) { Clear-Variable ADBAckupStatus }
$BackupDate = (Get-Date) - (New-TimeSpan -days $BackupAge)                  
[string]$DNSRoot = (Get-ADDomain).DNSRoot
[string[]]$ForestNCs = (Get-ADRootDSE).NamingContexts
$ContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain
$Context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext($ContextType,$DNSRoot)
$DomainController = [System.DirectoryServices.ActiveDirectory.DomainController]::findOne($Context)
    
ForEach ($NC in $ForestNCs)
    {  ## OPEN ForEach ($NC in $ForestNCs) 
        $DomainControllerMetadata = $DomainController.GetReplicationMetadata($NC) 
        $DSASignature = $DomainControllerMetadata.Item("dsaSignature")
        $LastOriginatingChangeTime = $DSASignature.LastOriginatingChangeTime
        
        IF ($LastOriginatingChangeTime -lt $BackupDate)  
            {  ## OPEN IF ($LastOriginatingChangeTime -lt $BackupDate)    
                [array]$ADBAckupStatus += " $NC last backed up on $LastOriginatingChangeTime " 
                Write-Host "$NC has NOT been backed up in over $BackupAge Days ($LastOriginatingChangeTime) " -fore RED 
            }  ## OPEN IF ($LastOriginatingChangeTime -lt $BackupDate) 
          ELSE { Write-host "$NC was last backed up on $LastOriginatingChangeTime" -fore GREEN }
    }  ## CLOSE ForEach ($NC in $ForestNCs)    
    

#######################
# Get Domain RID Info # 20111027-20 
#######################
## Based on code From https://blogs.technet.com/b/askds/archive/2011/09/12/managing-rid-pool-depletion.aspx
TRY
  {  ## OPEN TRY Get RID Information
    Write-Output "Get RID Information from AD including the number of RIDs issued and remaining `r "
    $RIDManagerProperty = Get-ADObject "cn=rid manager$,cn=system,$ADDomainDistinguishedName" -property RIDAvailablePool -server $ADDomainRIDMaster
    $RIDInfo = $RIDManagerProperty.RIDAvailablePool   
    [int32]$TotalSIDS = $RIDInfo / ([math]::Pow(2,32))
    [int64]$Temp64val = $TotalSIDS * ([math]::Pow(2,32))
    [int32]$CurrentRIDPoolCount = $RIDInfo - $Temp64val
    $RIDsRemaining = $TotalSIDS - $CurrentRIDPoolCount

    $RIDsIssuedPcntOfTotal = ( $CurrentRIDPoolCount / $TotalSIDS )
    $RIDsIssuedPercentofTotal = "{0:P2}" -f $RIDsIssuedPcntOfTotal
    $RIDsRemainingPcntOfTotal = ( $RIDsRemaining / $TotalSIDS )
    $RIDsRemainingPercentofTotal = "{0:P2}" -f $RIDsRemainingPcntOfTotal

    Write-Output "RIDs Issued: $CurrentRIDPoolCount ($RIDsIssuedPercentofTotal of total) `r "
    Write-Output "RIDs Remaining: $RIDsRemaining ($RIDsRemainingPercentofTotal of total) `r "
  }  ## CLOSE TRY Get RID Information
    
CATCH
    { Write-Warning "Unable to gather RID information `r " }
    
    
######################
# Get AD Object Count #  20120201-15
######################
Write-Output "Getting a list of all Active Directory objects in the NC: $SchemaNC `r "
$AllADSchemaObjects = Get-ADObject -filter * -SearchBase $SchemaNC -properties DistinguishedName,isDeleted,isRecycled,whenChanged -IncludeDeletedObjects
$AllADSchemaDeletedObjects = $AllADSchemaObjects | where { $_.isDeleted -eq $True }
$AllADSchemaRecycledObjects = $AllADSchemaObjects | where { $_.isRecycled -eq $True }
[int]$AllADSchemaObjectsCount = $AllADSchemaObjects.Count

Write-Output "Getting a list of all Active Directory objects in the NC: $ConfigurationNC `r "
$AllADConfigurationObjects = Get-ADObject -filter * -SearchBase $ConfigurationNC -properties DistinguishedName,isDeleted,isRecycled,whenChanged -IncludeDeletedObjects
$AllADConfigurationDeletedObjects = $AllADConfigurationObjects | where { $_.isDeleted -eq $True }
$AllADConfigurationRecycledObjects = $AllADConfigurationObjects | where { $_.isRecycled -eq $True }
$AllADConfigurationLostFoundObjects = $AllADConfigurationObjects | where { $_.DistinguishedName -like "*CN=LostAndFound*" }

[int]$AllADConfigurationObjectsCount = $AllADConfigurationObjects.Count
[int]$AllADConfigurationDeletedObjectsCount = $AllADConfigurationDeletedObjects.Count
[int]$AllADConfigurationRecycledObjectsCount = $AllADConfigurationRecycledObjects.Count
[int]$AllADConfigurationLostFoundObjectsCount = $AllADConfigurationLostFoundObjects.Count

$AllADConfigurationDeletedObjectsPCTofTotal = $AllADConfigurationDeletedObjectsCount / $AllADConfigurationObjectsCount
$AllADConfigurationDeletedObjectsPCTofTotal = "{0:p2}"                      -f $AllADConfigurationDeletedObjectsPCTofTotal

$AllADConfigurationRecycledObjectsCountPCTofTotal = $AllADConfigurationRecycledObjectsCount / $AllADConfigurationObjectsCount
$AllADConfigurationRecycledObjectsCountPCTofTotal = "{0:p2}"                      -f $AllADConfigurationRecycledObjectsCountPCTofTotal

$AllADConfigurationLostFoundObjectsPCTofTotal = $AllADConfigurationLostFoundObjectsCount / $AllADConfigurationObjectsCount
$AllADConfigurationLostFoundObjectsObjectsPCTofTotal = "{0:p2}"                      -f $AllADConfigurationLostFoundObjectsPCTofTotal

Write-Output "Getting a list of all Active Directory objects in the NC: $ADDomainDistinguishedName `r "
$AllADDomainObjects = Get-ADObject -filter * -SearchBase $ADDomainDistinguishedName -properties DistinguishedName,isDeleted,isRecycled,whenChanged -IncludeDeletedObjects
$AllADDomainDeletedObjects = $AllADDomainObjects | where { $_.isDeleted -eq $True }
$AllADDomainRecycledObjects = $AllADDomainObjects | where { $_.isRecycled -eq $True }
$AllADDomainLostFoundObjects = $AllADDomainObjects | where { $_.DistinguishedName -like "*CN=LostAndFound*" }

[int]$AllADDomainObjectsCount = $AllADDomainObjects.Count
[int]$AllADDomainDeletedObjectsCount = $AllADDomainDeletedObjects.Count
[int]$AllADDomainRecycledObjectsCount = $AllADDomainRecycledObjects.Count
[int]$AllADDomainLostFoundObjectsCount = $AllADDomainLostFoundObjects.Count

$AllADDomainDeletedObjectsPCTofTotal = $AllADDomainDeletedObjectsCount / $AllADDomainObjectsCount
$AllADDomainDeletedObjectsPCTofTotal = "{0:p2}"                      -f $AllADDomainDeletedObjectsPCTofTotal

$AllADDomainRecycledObjectsPCTofTotal = $AllADDomainRecycledObjectsCount / $AllADDomainObjectsCount
$AllADDomainRecycledObjectsPCTofTotal = "{0:p2}"                      -f $AllADDomainRecycledObjectsPCTofTotal

$AllADDomainLostFoundObjectsPCTofTotal = $AllADDomainLostFoundObjectsCount / $AllADDomainObjectsCount
$AllADDomainLostFoundObjectsPCTofTotal = "{0:p2}"                      -f $AllADDomainLostFoundObjectsPCTofTotal

Write-Output "Getting a list of all Active Directory objects in the NC: $ForestDNSZoneNC `r "
TRY
    {  ## OPEN TRY
        $AllADForestDNSZoneObjects = Get-ADObject -filter * -SearchBase $ForestDNSZoneNC -properties DistinguishedName,isDeleted,isRecycled,whenChanged -IncludeDeletedObjects
        $AllADForestDNSZoneTombstonedObjects = $AllADForestDNSZoneObjects | where { $_.dNSTombstoned -eq $True }
        $AllADForestDNSZoneDeletedObjects = $AllADForestDNSZoneObjects | where { $_.isDeleted -eq $True }
        $AllADForestDNSZoneRecycledObjects = $AllADForestDNSZoneObjects | where { $_.isRecycled -eq $True }
        $AllADForestDNSZoneLostFoundObjects = $AllADForestDNSZoneObjects | where { $_.DistinguishedName -like "*CN=LostAndFound*" }

        $AllADForestDNSZoneTombstonedObjectsCount = $AllADForestDNSZoneTombstonedObjects.Count
        $AllADForestDNSZoneObjectsCount = $AllADForestDNSZoneObjects.Count
        $AllADForestDNSZoneDeletedObjectsCount = $AllADForestDNSZoneDeletedObjects.Count
        $AllADForestDNSZoneRecycledObjectsCount = $AllADForestDNSZoneRecycledObjects.Count
        $AllADForestDNSZoneLostFoundObjectsCount = $AllADForestDNSZoneLostFoundObjects.Count
    }  ## CLOSE TRY
 CATCH { Write-Verbose "Could not connect to Forest DNS partition. It may not exist. `r " }

IF ($AllADForestDNSZoneObjects -gt 0)
{
$AllADForestDNSZoneTombstonedObjectsPCTofTotal = $AllADForestDNSZoneTombstonedObjectsCount / $AllADForestDNSZoneObjects
$AllADForestDNSZoneTombstonedObjectsPCTofTotal = "{0:p2}"                      -f $AllADForestDNSZoneTombstonedObjectsPCTofTotal

$AllADForestDNSZoneDeletedObjectsPCTofTotal = $AllADForestDNSZoneDeletedObjectsCount / $AllADForestDNSZoneObjectsCount
$AllADForestDNSZoneDeletedObjectsPCTofTotal = "{0:p2}"                      -f $AllADForestDNSZoneDeletedObjectsPCTofTotal

$AllADForestDNSZoneRecycledObjectsPCTofTotal = $AllADForestDNSZoneRecycledObjectsCount / $AllADForestDNSZoneObjectsCount
$AllADForestDNSZoneRecycledObjectsPCTofTotal = "{0:p2}"                      -f $AllADForestDNSZoneRecycledObjectsPCTofTotal

$AllADForestDNSZoneLostFoundObjectsPCTofTotal = $AllADForestDNSZoneLostFoundObjectsCount / $AllADForestDNSZoneObjectsCount
$AllADForestDNSZoneLostFoundObjectsPCTofTotal = "{0:p2}"                      -f $AllADForestDNSZoneRecycledObjectsPCTofTotal  
}

Write-Output "Getting a list of all Active Directory objects in the NC: $DomainDNSZoneNC `r "
TRY
    {  ## OPEN TRY
        $AllADDomainDNSZoneZoneObjects = Get-ADObject -filter * -SearchBase $DomainDNSZoneNC -properties DistinguishedName,isDeleted,isRecycled,whenChanged -IncludeDeletedObjects
        $AllADDDomainDNSZonedNSTombstonedObjects = $AllADDomainDNSZoneZoneObjects | where { $_.dNSTombstoned -eq $True }
        $AllADDDomainDNSZoneDeletedObjects = $AllADDomainDNSZoneZoneObjects | where { $_.isDeleted -eq $True }
        $AllADDDomainDNSZoneRecycledObjects = $AllADDomainDNSZoneZoneObjects | where { $_.isRecycled -eq $True }
        $AllADDDomainDNSZoneLostFoundObjects = $AllADDomainDNSZoneZoneObjects | where { $_.DistinguishedName -like "*CN=LostAndFound*" }

        [int]$AllADDDomainDNSZoneObjectsCount = $AllADDomainDNSZoneZoneObjects.Count
        [int]$AllADDDomainDNSZoneTombstonedObjectsCount = $AllADDDomainDNSZonedNSTombstonedObjects.Count
        [int]$AllADDDomainDNSZoneDeletedObjectsCount = $AllADDDomainDNSZoneDeletedObjects.Count
        [int]$AllADDDomainDNSZoneRecycledObjectsCount = $AllADDDomainDNSZoneRecycledObjects.Count
        [int]$AllADDDomainDNSZoneLostFoundObjectsCount = $AllADDDomainDNSZoneLostFoundObjects.Count
    }  ## CLOSE TRY
 CATCH { Write-Verbose "Could not connect to Forest DNS partition. It may not exist. `r " }


IF ($AllADDDomainDNSZoneObjects -gt 0)
{
$AllADDDomainDNSZoneTombstonedObjectsPCTofTotal = $AllADDDomainDNSZoneTombstonedObjectsCount / $AllADDDomainDNSZoneObjectsCount
$AllADDDomainDNSZoneTombstonedObjectsPCTofTotal = "{0:p2}"                      -f $AllADDDomainDNSZoneTombstonedObjectsPCTofTotal

$AllADDDomainDNSZoneDeletedObjectsPCTofTotal = $AllADDDomainDNSZoneDeletedObjectsCount / $AllADDDomainDNSZoneObjectsCount
$AllADDDomainDNSZoneDeletedObjectsPCTofTotal = "{0:p2}"                      -f $AllADDDomainDNSZoneDeletedObjectsPCTofTotal

$AllADDDomainDNSZoneRecycledObjectsPCTofTotal = $AllADDDomainDNSZoneRecycledObjectsCount / $AllADDDomainDNSZoneObjectsCount
$AllADDDomainDNSZoneRecycledObjectsPCTofTotal = "{0:p2}"                      -f $AllADDDomainDNSZoneRecycledObjectsPCTofTotal

$AllADDDomainDNSZoneLostFoundObjectsPCTofTotal = $AllADDDomainDNSZoneLostFoundObjectsCount / $AllADDDomainDNSZoneObjectsCount
$AllADDDomainDNSZoneLostFoundObjectsPCTofTotal = "{0:p2}"                      -f $AllADDDomainDNSZoneLostFoundObjectsPCTofTotal  
}

Write-Output "  `r "
Write-Output "Active Directory Object Snapshot `r "
Write-Output "-------------------------------- `r "
Write-Output "`r "
Write-Output "Active Directory Schema Partition Stats:  `r "
Write-Output "   $AllADSchemaObjectsCount objects in $SchemaNC  `r "
Write-Output "  `r "
Write-Output "Active Directory Configuration Partition Stats:  `r "
Write-Output "   $AllADSchemaObjectsCount objects in $ConfigurationNC  `r "
Write-Output "   $AllADConfigurationDeletedObjectsCount Deleted (tombstone) objects in $ConfigurationNC . This is $AllADConfigurationDeletedObjectsPCTofTotal of the total.   `r "
Write-Output "   $AllADConfigurationRecycledObjectsCount Recycled objects in $ConfigurationNC . This is $AllADConfigurationRecycledObjectsCountPCTofTotal of the total.  `r "
Write-Output "   $AllADConfigurationLostFoundObjectsCount Lost & Found objects in $ConfigurationNC . This is $AllADConfigurationLostFoundObjectsPCTofTotal of the total.  `r "
Write-Output "  `r "
Write-Output "Active Directory Domain Partition Stats:  `r "
Write-Output "   $AllADDomainObjectsCount objects in $ADDomainDistinguishedName  `r "
Write-Output "   $AllADDomainDeletedObjectsCount Deleted (tombstone) objects in $ADDomainDistinguishedName . This is $AllADDomainDeletedObjectsPCTofTotal of the total.   `r "
Write-Output "   $AllADDomainRecycledObjectsCount Recycled objects in $ADDomainDistinguishedName . This is $AllADDomainRecycledObjectsPCTofTotal of the total.   `r "
Write-Output "   $AllADDomainLostFoundObjectsCount Lost & Found objects in $ADDomainDistinguishedName . This is $AllADDomainLostFoundObjectsPCTofTotal of the total.   `r "
Write-Output "  `r "
Write-Output "Active Directory ForestDNS Partition Stats:  `r "
Write-Output "   $AllADForestDNSZoneObjectsCount objects in $ForestDNSZoneNC  `r "
Write-Output "   $AllADForestDNSZoneDeletedObjectsCount Deleted objects in $ForestDNSZoneNC . This is $AllADForestDNSZoneDeletedObjectsPCTofTotal of the total.   `r "
Write-Output "   $AllADForestDNSZoneRecycledObjectsCount Recycled objects in $ForestDNSZoneNC . This is $AllADForestDNSZoneRecycledObjectsPCTofTotal of the total.   `r "
Write-Output "   $AllForestDNSZoneLostFoundObjectsCount Lost & Found objects in $ForestDNSZoneNC . This is $AllADForestDNSZoneLostFoundObjectsPCTofTotal of the total.   `r "
Write-Output "  `r "
Write-Output "Active Directory DomainDNS Partition Stats:  `r "
Write-Output "   $AllADDomainDNSZoneObjectsCount objects in $DomainDNSZoneNC  `r "
Write-Output "   $AllADDomainDNSZoneDeletedObjectsCount Deleted objects in $DomainDNSZoneNC . This is $AllADDomainDNSZoneDeletedObjectsPCTofTotal of the total.   `r "
Write-Output "   $AllADDomainDNSZoneRecycledObjectsCount Recycled objects in $DomainDNSZoneNC . This is $AllADDomainDNSZoneRecycledObjectsPCTofTotal of the total.   `r "
Write-Output "   $AllADDomainDNSZoneLostFoundObjectsCount Lost & Found objects in $DomainDNSZoneNC . This is $AllADDomainDNSZoneLostFoundObjectsPCTofTotal of the total.   `r "
Write-Output "  `r "

#################################################
# Get AD LastLogonTimeStamp Replication Setting # 20111027-20 
#################################################
Write-Verbose "Get AD LastLogonTimeStamp Replication Setting (Default is 14 which means the attribute value is blank) `r "
Write-Verbose "Naming Context Attribute that defines this is 'ms-DS-Logon-Time-Sync-Interval' and the system GUID is 'ad7940f8-e43a-4a42-83bc-d688e59ea605' `r "
$DirectoryServicesNamingContext = Get-ADObject -Identity "$ADDomainDistinguishedName" -Properties *
$LLTReplicationValue = $DirectoryServicesNamingContext."msDS-LogonTimeSyncInterval"

IF ($LLTReplicationValue -ge 1) 
    { Write-Verbose "The msDS-LogonTimeSyncInterval attribute value on $DomainDNS was changed from the default value of 14 to $LLTReplicationValue which means the LastLogonTimeStamp attribute will replication about every $LLTReplicationValue days `r " }
 ELSE 
    { $LLTReplicationValue = 14 ; Write-Verbose "The msDS-LogonTimeSyncInterval attribute value on $DomainDNS is configured with the default value of 14 (value is blank) `r " }  

########################        
# Get Domain GPO Stats #
########################
Write-Output "Discovering all $DomainDNS GPOs `r "
$DomainGPOs = Get-GPO -All -Domain $DomainDNS
[int]$DomainGPOsCount = $DomainGPOs.Count
Write-Output "There are $DomainGPOsCount GPOs found `r "
Write-Output "  `r "

####################
# Get AD Site Data # 20120203-207
####################
Write-Output "Get AD Site List `r"
$ADSites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites 
[int]$ADSitesCount = $ADSites.Count
Write-Output "There are $ADSitesCount AD Sites `r"

# List all sites & site DCs
Write-Output "List all sites & site DCs `r"

ForEach ($Site in $ADSites)
    {  ## OPEN ForEach DC in DomainDCs
      [array]$ADSiteList += $Site.Name 
    }  ## CLOSE ForEach DC in DomainDCs

Write-Output "Discover DCs in each AD Site `r"
ForEach ($Site in $ADSiteList)
    {  ## OPEN ForEach Site in ADSites
       IF ($SiteDCs) { Clear-Variable SiteDCs }
       $SiteDCs = Get-ADDomainController -Filter {Site -eq $Site}
       $SiteDCsCount = $SiteDCs.Count
       IF ($SiteDCsCount -gt 1) 
        {  ## OPEN IF there is more than 1 DC in the site
          $SiteDCList = ""
          [int]$SiteGC = 0
          
          ForEach ($SiteDC in $SiteDCs)
            {  ## OPEN ForEach SiteDC in SiteDCs
                [array]$SiteDCList += $SiteDC.Name
                $SiteDCInfo = Get-ADDomainController -Identity $SiteDC
                IF ($SiteDCInfo.IsGlobalCatalog -eq $True) {$SiteGC++ ; Write-Verbose "$SiteDC is a Global Catalog `r"}
                 ELSE {Write-Verbose "$SiteDC is NOT a Global Catalog `r"}
            }  ## CLOSE ForEach SiteDC in SiteDCs
            
            IF ($SiteGC -eq 0) {$SitesWithNoGCs += $Site ; Write-Verbose "There are no Global Catalogs in the site $Site `r" }
            $SiteDCListCount = $SiteDCList.Count
            $SiteDC = $SiteDCs.Name
            Write-Output "$Site has the following $SiteDCListCount DCs: $SiteDCList  `r"
        }  ## CLOSE IF there is more than 1 DC in the site
       
       IF ($SiteDCsCount -eq 1) 
        {  ## OPEN IF there is only 1 DC in the site
            $SiteDC = $SiteDCs.Name
            Write-Output "$Site has one DC: $SiteDC  `r"
        }  ## CLOSE IF there is only 1 DC in the site
       
       IF ($SiteDCsCount -eq 0) 
        {  ## OPEN IF there are no DCs in the site
            [array]$SitesWithNoGCs += $Site
            Write-Output "$Site has no DCs `r"
        }  ## CLOSE IF there are no DCs in the site
    }  ## CLOSE ForEach Site in ADSites

#####################################
# Get Domain Controller Information # 20111028-09
#####################################            
# Get Domain DCs
Write-Output "Get Domain Controller Information `r"
[array] $DomainDCs = $ADDomainReplicaDirectoryServers
[int]$DomainWritableDCsCount =  $DomainDCs.Count
Write-Output "There are $DomainWritableDCsCount Writable DCs in $DomainDNS.  `r "
Write-Output "  `r "

# Get Domain RODCs
Write-Output "Get Domain RODCs  `r "
[array] $DomainRODCs = $ADDomainReadOnlyReplicaDirectoryServers
[int]$DomainRODCsCount =  $DomainRODCs.Count
Write-Verbose "There are $DomainRODCsCount Read-Only DCs in $DomainDNS. `r "
Write-Output "  `r "

# Discover Domain Controllers in AD
[array]$DomainControllers =  $ADDomainReadOnlyReplicaDirectoryServers 
[array]$DomainControllers +=  $ADDomainReplicaDirectoryServers
$DomainControllers = $DomainControllers -Replace(".$DomainDNS","")

$DomainControllers = $DomainControllers | Sort-Object | Get-Unique
[int]$DomainControllersCount = $DomainControllers.Count

Write-Output "Processing the following $DomainControllersCount Domain Controllers in $DomainDNS ...  `r " 

# Get DC Operating System Versions

Write-Verbose "Discovering all Windows Server 2012 DCs in the domain $DomainDNS ... `r "
[array] $2012DCs = Get-ADDomainController -Filter  { OperatingSystem -like "Windows Server 2012*" }

ForEach ($DC in $2012DCs)
    {  ## OPEN ForEach DC in $2012DCs
        [array] $Domain2012DCs += $DC.HostName
    }  ## CLOSE ForEach DC in $2012DCs
	 
[int] $2012DCCount = $2012DCs.Count
$Domain2012DCs = $Domain2012DCs | sort
Write-Verbose "Found $2012DCCount Windows Server 2012 DCs in the domain $DomainDNS. `r "
#

Write-Verbose "Discovering all Windows Server 2008 R2 Service Pack 1 DCs in the domain $DomainDNS ... `r "
[array] $2008R2SP1Dcs = Get-ADDomainController -Filter  { (OperatingSystem -like "Windows Server 2008 R2*") -and (OperatingSystemServicePack -eq "Service Pack 1")}

ForEach ($DC in $2008R2SP1Dcs)
    {  ## OPEN ForEach DC in 2008Dcs
        [array] $Domain2008R2Sp1DCs += $DC.HostName
    }  ## CLOSE ForEach DC in 2008Dcs
	 
[int] $2008R2Sp1DCCount = $Domain2008R2Sp1DCs.Count
$Domain2008R2Sp1DCs = $Domain2008R2Sp1DCs | sort
Write-Verbose "Found $2008R2Sp1DCCount Windows Server 2008 R2 Service Pack 1 DCs in $DomainDNS. `r "
#
Write-Verbose "Discovering all Windows Server 2008 R2 (No Service Pack) DCs in the domain $DomainDNS ... `r "
[array] $2008R2Dcs = Get-ADDomainController -Filter  { (OperatingSystem -like "Windows Server 2008 R2*") -and (OperatingSystemServicePack -notlike "Service Pack*")}

ForEach ($DC in $2008R2Dcs)
    {  ## OPEN ForEach DC in 2008Dcs
        [array] $Domain2008R2DCs += $DC.HostName
    }  ## CLOSE ForEach DC in 2008Dcs
	
[int] $2008R2DCCount = $Domain2008R2DCs.Count
$Domain2008R2DCs = $Domain2008R2DCs | sort
Write-Verbose "Found $2008R2DCCount Windows Server 2008 R2 (No Service Pack) DCs in $DomainDNS. `r "
#
Write-Verbose "Discovering all Windows Server 2008 (Any Service Pack) DCs in the domain $DomainDNS ... `r "
[array] $2008Dcs = Get-ADDomainController -Filter  { (OperatingSystem -like "Windows Server 2008*") -and (OperatingSystem -notlike "Windows Server 2008 R2*") }

ForEach ($DC in $2008Dcs)
    {  ## OPEN ForEach DC in 2008Dcs
        [array] $Domain2008DCs += $DC.HostName
    }  ## CLOSE ForEach DC in 2008Dcs
	
[int] $2008DCCount = $Domain2008DCs.Count
$Domain2008DCs = $Domain2008DCs | sort
Write-Verbose "Found $2008DCCount Windows Server 2008 (Any Service Pack) DCs in $DomainDNS. `r "
#
Write-Verbose "Discovering all Windows Server 2003 R2 (Any Service Pack) DCs in the domain $DomainDNS ... `r "
[array] $2003R2DCsInSiteArray = Get-ADDomainController -Filter  { (OperatingSystem -like "Windows Server 2003 R2*") }
ForEach ($DC in $2003R2DCsInSiteArray)
    {  ## OPEN ForEach DC in 2003DCsInSiteArray
      [array] $Domain2003R2DCs += $DC.HostName
    }  ## CLOSE ForEach DC in 2003DCsInSiteArray
	
[int] $2003R2DcCount = $Domain2003R2DCs.Count
$Domain2003R2DCs = $Domain2003R2DCs | sort
Write-Verbose "Found $2003DcCount Windows Server 2003 (Any Service Pack) DCs in $DomainDNS.  `r "
#
Write-Verbose "Discovering all Windows Server 2003 (Any Service Pack) DCs in the domain $DomainDNS ... `r "
[array] $2003DCsInSiteArray = Get-ADDomainController -Filter  { (OperatingSystem -like "Windows Server 2003*") -and (OperatingSystem -notlike "Windows Server 2003 R2*") }
ForEach ($DC in $2003DCsInSiteArray)
    {  ## OPEN ForEach DC in 2003DCsInSiteArray
      [array] $Domain2003DCs += $DC.HostName
    }  ## CLOSE ForEach DC in 2003DCsInSiteArray
	
[int] $2003DcCount = $Domain2003DCs.Count
$Domain2003DCs = $Domain2003DCs | sort
Write-Verbose "Found $2003DcCount Windows Server 2003 (Any Service Pack) DCs in $DomainDNS. `r "
#
Write-Verbose "Discovering all Windows 2000 Server (Any Service Pack) DCs in the domain $DomainDNS ... `r "
[array] $2000DCsInSiteArray = Get-ADDomainController -Filter  { OperatingSystem -like "Windows 2000 Server*"}
ForEach ($DC in $2000DCsInSiteArray)
    {  ## OPEN ForEach DC in 2000DCsInSiteArray
      [array] $Domain2000DCs += $DC.HostName
    }  ## CLOSE ForEach DC in 2000DCsInSiteArray
	
[int] $2000DcCount = $Domain2000DCs.Count
$Domain2000DCs = $Domain2000DCs | sort
Write-Output "Found $2000DcCount Windows 2000 Server (Any Service Pack) DCs in $DomainDNS. `r "
##
write-Output "Out of the $DomainControllersCount DCs in $DomainDNS : `r "
Write-Output "$DomainRODCsCount are RODCs & $DomainWritableDCsCount are writable DCs `r "
write-Output "$2012DCCount are running Windows Server 2012 `r "
write-Output "$2008R2Sp1DCCount are running Windows Server 2008 R2 SP1 `r "
write-Output "$2008R2DCCount are running Windows Server 2008 R2 (No Service Pack) `r "
write-Output "$2008DCCount are running Windows Server 2008 (Any Service Pack)  `r "
write-Output "$2003R2DcCount are running Windows Server 2003 R2 (Any Service Pack)  `r "
write-Output "$2003DcCount are running Windows Server 2003 (Any Service Pack)  `r "
write-Output "$2000DcCount are running Windows 2000 Server (Any Service Pack) `r "

[int]$2008DCCount = $2008R2Sp1DCCount + $2008R2DCCount + $2008DCCount
[int]$2003DCCount = $2003DcCount + $2003R2DcCount

write-Output " `r "
write-Output " `r "
write-Output "$DomainControllersCount TOTAL DCs `r "
write-Output "$2012DCCount DCs running Windows Server 2012 (any version, any service pack) `r "
write-Output "$2008DCCount DCs running Windows Server 2008 (any version, any service pack) `r "
write-Output "$2003DCCount DCs running Windows Server 2003 (any version, any service pack) `r "
write-Output " `r "
write-Output " `r "


###################
# User Statistics # 20120119-15
###################
Write-Output "Generating AD User Statistics `r "
$LastLoggedOnDate = $DateTime - (New-TimeSpan -days $UserLogonAge)
$PasswordStaleDate = $DateTime - (New-TimeSpan -days $UserPasswordAge)
$DateTime = Get-Date #Get date/time
$UserStaleDate = $DateTime.AddDays(-$UserPasswordAge)
$NeverLoggedOnDate = $DateTime.AddDays(-365) 
$Yesterday = $DateTime.AddDays(-1) 
$Yesterday = $Yesterday.ToShortDateString()
$Today = $DateTime.ToShortDateString()
[string]$Today = $Today

Write-Output "Discovering all users in $ADDomainDNSRoot ... `r "
# Gather a list of all users in AD including necessary attributes
[array]$AllUsers = Get-ADUser -filter * -properties Name,DistinguishedName,Enabled,LastLogonDate,LastLogonTimeStamp,LockedOut,msExchHomeServerName,PasswordLastSet,SAMAccountName,Certificates,userCertificate,userSMIMECertificate
[int]$AllUsersCount = $AllUsers.Count
Write-Verbose "There are $AllUsersCount user objects discovered in $ADDomainDNSRoot ... `r "
Write-Verbose " `r "

Write-Verbose "Count Users with SIDHistory `r "
[array] $SIDHistoryUsers = $AllUsers | Where-Object { $_.SIDHistory -ne $NULL }
[int]$SIDHistoryUsersCount = $SIDHistoryUsers.Count
Write-Verbose "There are $SIDHistoryUsersCount users with SIDHistory in $DomainDNS `r "

Write-Verbose "Count Disabled & Enabled Users `r "
[array] $DisabledUsers = $AllUsers | Where-Object { $_.Enabled -eq $False }
[int]$DisabledUsersCount = $DisabledUsers.Count
[array] $EnabledUsers = $AllUsers | Where-Object { $_.Enabled -eq $True }
[int]$EnabledUsersCount = $EnabledUsers.Count
Write-Verbose "There are $EnabledUsersCount Enabled users and there are $DisabledUsersCount Disabled users in $DomainDNS `r "

Write-Verbose "Count Inactive Users that are Enabled `r "
[array] $InactiveUsers = $AllUsers | Where-Object { ($_.PasswordLastSet -lt $UserStaleDate) -and ($_.Enabled -eq $True) }
[int]$InactiveUsersCount = $InactiveUsers.Count
Write-Verbose "There are $InactiveUsersCount users identified as Inactive (with passwords older than $UserPasswordAge days in $DomainDNS `r "

Write-Verbose "Count AdminSDHolder protected Admin Accounts `r "
$ADPropUsers =  Get-ADObject -filter {(ObjectClass -eq "User") -AND (AdminCount -eq 1)}
[int]$ADPropUsersCount = $ADPropUsers.Count

$ADPropGroups =  Get-ADObject -filter {(ObjectClass -eq "Group") -AND (AdminCount -eq 1)}
[int]$ADPropGroupsCount = $ADPropGroups.Count
Write-Verbose "There are $ADPropUsersCount AdminSDHolder protected Admin Accounts & $ADPropGroupsCount AdminSDHolder protected Admin Groups. `r "

Write-Verbose "Count Service Accounts (accounts with SVC in the name) `r "
[array] $AllServiceAccounts = $AllUsers | Where-Object { $_.SAMAccountName -like "*SVC*" }
[int]$AllServiceAccountsCount = $AllServiceAccounts.Count
[array] $EnabledServiceAccounts = $AllServiceAccounts | Where-Object  { $_.Enabled -eq $True }
[int]$EnabledServiceAccountsCount = $EnabledServiceAccounts.Count
Write-Verbose "There are $EnabledServiceAccountsCount Enabled Service accounts in $DomainDNS (out of a total $AllServiceAccountsCount Service accounts) `r "

Write-Verbose "Count users with an Exchange Mailbox `r "
[array] $MailboxUsers = $AllUsers | Where-Object { $_.msExchHomeServerName -notlike $NULL }
[int]$MailboxUsersCount = $MailboxUsers.Count
[array] $MailboxEnabledUsers = $MailboxUsers | Where-Object { $_.msExchHomeServerName -notlike $NULL }
[int]$MailboxEnabledUsersCount = $MailboxEnabledUsers.Count
Write-Verbose "There are $MailboxUsersCount users in $DomainDNS with an Exchange Mailbox. `r "
Write-Verbose "There are $MailboxEnabledUsersCount Enabled users in $DomainDNS with an Exchange Mailbox. `r "

Write-Verbose "Count Enabled users who have logged on in the last 30 days `r "
[array] $LastLogon30 = $EnabledUsers | Where-Object { $_.LastLogonDate -ge $30Days }
[int]$LastLogon30Count = $LastLogon30.Count
Write-Verbose "Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon30Count have logged on in the last 30 days (there may be up to a 14 day margin of error for this count) `r "

Write-Verbose "Count Enabled users who have logged on in the last 45 days `r "
[array] $LastLogon45 = $EnabledUsers | Where-Object { $_.LastLogonDate -ge $45Days }
[int]$LastLogon45Count = $LastLogon45.Count
Write-Verbose "Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon45Count have logged on in the last 45 days `r"

Write-Verbose "Count Enabled users who have logged on in the last 60 days `r "
[array] $LastLogon60 = $EnabledUsers | Where-Object { $_.LastLogonDate -ge $60Days }
[int]$LastLogon60Count = $LastLogon60.Count
Write-Verbose "Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon60Count have logged on in the last 60 days `r"

Write-Verbose "Count Enabled users who have logged on in the last 90 days `r"
[array] $LastLogon90 = $EnabledUsers | Where-Object { $_.LastLogonDate -ge $90Days }
[int]$LastLogon90Count = $LastLogon90.Count
Write-Verbose "Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon90Count have logged on in the last 90 days `r"

Write-Verbose "Count Enabled users who have logged on in the last 120 days `r"
[array] $LastLogon120 = $EnabledUsers | Where-Object { $_.LastLogonDate -ge $120Days }
[int]$LastLogon120Count = $LastLogon120.Count
Write-Verbose "Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon120Count have logged on in the last 120 days `r"

Write-Verbose "Count Enabled users who have logged on in the last 180 days `r"
[array] $LastLogon180 = $EnabledUsers | Where-Object { $_.LastLogonDate -ge $180Days }
$LastLogon180Count = $LastLogon180.Count
Write-Verbose "Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon180Count have logged on in the last 180 days   `r"

Write-Verbose "Count All users who have NEVER logged on `r"
[array] $LastLogonNever = $AllUsers | Where-Object { ($_.LastLogonDate -eq $NULL) -and ($_.PasswordLastSet -gt $NeverLoggedOnDate) }
[int]$LastLogonNeverCount = $LastLogonNever.Count
Write-Verbose "Out of All $AllUsersCount users in $DomainDNS $LastLogonNeverCount have NEVER logged on (no logon date associated with account) `r"

Write-Verbose "Count users who logged in during the last week `r"
[array] $LastLogon1 = $EnabledUsers | Where-Object { $_.LastLogonDate -ge $7days }
[int]$LastLogon1Count = $LastLogon1.Count
Write-Verbose "$LastLogon1Count Enabled users logged in within the last week  `r"

Write-Verbose "Count users who logged in yesterday `r"
[array] $LastLogonYesterday = $EnabledUsers | Where-Object { $_.LastLogonDate -like "*$Yesterday*" }
[int]$LastLogonYesterdayCount = $LastLogonYesterday.Count
Write-Verbose "$LastLogonYesterdayCount Enabled users logged in yesterday  `r"

Write-Verbose "Count users who logged in today `r"
[array] $LastLogontoday = $EnabledUsers | Where-Object { $_.LastLogonDate -like "*$Today*" }
$LastLogontodayCount = $LastLogontoday.Count
Write-Verbose "$LastLogontodayCount Enabled users logged in today (so far) `r"

Write-Verbose "Count enabled users who have accounts locked out `r"
[array] $EnabledLockedUsers = $EnabledUsers | Where-Object { $_.LockedOut -eq $True }
$EnabledLockedUsersCount = $EnabledLockedUsers.Count
Write-Verbose "$EnabledLockedUsersCount Enabled users are currently locked out `r"

## PKI User Logon Stats
Write-Verbose "Discovering all PKI-enabled users in $ADDomainDNSRoot ... `r "
# Get a list of PKI-enabled users
[array] $AllPKIEnabledADUsers = $AllUsers | Where { ($_.Certificates -ne $NULL) -OR ($_.userCertificate -ne $NULL) -OR ($_.userSMIMECertificate -ne $NULL) }
[int]$AllPKIEnabledADUsersCount = $AllPKIEnabledADUsers.Count
Write-Verbose "There are $AllPKIEnabledADUsersCount PKI-enabled users discovered in $ADDomainDNSRoot ... `r "

Write-Verbose "Count Enabled users that have a PKI certificate `r "
# Get a list of PKI-enabled (& Enabled AD) users
[array] $PKIEnabledADUsers = $EnabledUsers | Where { ($_.Certificates -ne $NULL) -OR ($_.userCertificate -ne $NULL) -OR ($_.userSMIMECertificate -ne $NULL) }
[int]$PKIEnabledADUsersCount = $PKIEnabledADUsers.Count
Write-Verbose "There are $PKIEnabledADUsersCount Enabled users that have a PKI certificate discovered in $ADDomainDNSRoot ... `r "

Write-Verbose "Count PKI-enabled (& Enabled AD) Users configured to logon with a Smart Card (required) `r "
# Get a list of PKI-enabled (& Enabled AD) users configured to logon with a Smart Card (required)
[array] $PKIEnabledLogonADUsers = $AllUsers | Where { $_.SmartcardLogonRequired -eq $True }
[int]$PKIEnabledLogonADUsersCount = $PKIEnabledLogonADUsers.Count
Write-Verbose "There are $PKIEnabledLogonADUsersCount PKI-enabled (& Enabled AD) Users configured to logon with a Smart Card (required) discovered in $ADDomainDNSRoot ... `r "

Write-Verbose "Count PKI-enabled (& Enabled AD) users who have logged on in the last 30 days `r "
[array] $PKILastLogon30 = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -ge $30Days }
[int]$PKILastLogon30Count = $PKILastLogon30.Count
Write-Verbose "Out of $EnabledUsersCount PKI-enabled (& Enabled AD) users in $DomainDNS only $PKILastLogon30Count have logged on in the last 30 days (there may be up to a 14 day margin of error for this count) `r "

Write-Verbose "Count PKI-enabled (& Enabled AD) users who have logged on in the last 45 days `r "
[array] $PKILastLogon45 = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -ge $45Days }
[int]$PKILastLogon45Count = $PKILastLogon45.Count
Write-Verbose "Out of $EnabledUsersCount PKI-enabled (& Enabled AD) users in $DomainDNS only $PKILastLogon45Count have logged on in the last 45 days `r"

Write-Verbose "Count PKI-enabled (& Enabled AD) users who have logged on in the last 60 days `r "
[array] $PKILastLogon60 = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -ge $60Days }
[int]$PKILastLogon60Count = $PKILastLogon60.Count
Write-Verbose "Out of $EnabledUsersCount PKI-enabled (& Enabled AD) users in $DomainDNS only $PKILastLogon60Count have logged on in the last 60 days `r"

Write-Verbose "Count PKI-enabled (& Enabled AD) users who have logged on in the last 90 days `r"
[array] $PKILastLogon90 = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -ge $90Days }
[int]$PKILastLogon90Count = $PKILastLogon90.Count
Write-Verbose "Out of $EnabledUsersCount PKI-enabled (& Enabled AD) users in $DomainDNS only $PKILastLogon90Count have logged on in the last 90 days `r"

Write-Verbose "Count PKI-enabled (& Enabled AD) users who have logged on in the last 120 days `r"
[array] $PKILastLogon120 = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -ge $120Days }
[int]$PKILastLogon120Count = $PKILastLogon120.Count
Write-Verbose "Out of $EnabledUsersCount PKI-enabled (& Enabled AD) users in $DomainDNS only $PKILastLogon120Count have logged on in the last 120 days `r"

Write-Verbose "Count PKI-enabled (& Enabled AD) users who have logged on in the last 180 days `r"
[array] $PKILastLogon180 = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -ge $180Days }
[int]$PKILastLogon180Count = $PKILastLogon180.Count
Write-Verbose "Out of $EnabledUsersCount PKI-enabled (& Enabled AD) users in $DomainDNS only $PKILastLogon180Count have logged on in the last 180 days   `r"

Write-Verbose "Count All PKI-enabled (& Enabled AD) users who have NEVER logged on `r"
[array] $PKILastLogonNever = $PKIEnabledLogonADUsers | Where-Object { ($_.LastLogonDate -eq $NULL) -and ($_.PasswordLastSet -gt $NeverLoggedOnDate) }
[int]$PKILastLogonNeverCount = $PKILastLogonNever.Count
Write-Verbose "Out of All $AllUsersCount PKI-enabled (& Enabled AD) in $DomainDNS $PKILastLogonNeverCount have NEVER logged on (no logon date associated with account) `r"

Write-Verbose "Count PKI-enabled (& Enabled AD) users who logged in during the last week `r"
[array] $PKILastLogon1 = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -le $7days }
[int]$PKILastLogon1Count = $PKILastLogon1.Count
Write-Verbose "$PKILastLogon1Count PKI-enabled (& Enabled AD) users logged in within the last week  `r"

Write-Verbose "Count PKI-enabled (& Enabled AD) users who logged in yesterday `r"
[array] $PKILastLogonYesterday = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -like "*$Yesterday*" }
$PKILastLogonYesterdayCount = $PKILastLogonYesterday.Count
Write-Verbose "$PKILastLogonYesterdayCount PKI-enabled (& Enabled AD) users logged in yesterday  `r"

Write-Verbose "Count PKI-enabled (& Enabled AD) users who logged in today `r"
[array] $PKILastLogontoday = $PKIEnabledLogonADUsers | Where-Object { $_.LastLogonDate -like "*$Today*" }
[int]$PKILastLogontodayCount = $PKILastLogontoday.Count
Write-Verbose "$PKILastLogontodayCount PKI-enabled (& Enabled AD) users logged in today (so far) `r"

##############################
# Get AD Group Statistics # 20111130-30
##############################
Write-Output "Generating AD Group Statistics  `r "

Write-Output "Getting list of AD Groups `r "
[array]$AllADGroups = Get-ADGroup -Filter * -Properties *
[int]$AllADGroupsCount = $AllADGroups.Count
Write-Verbose "There are $AllADGroupsCount Total groups in AD `r "

[array]$ADUniversalGroups = $AllADGroups | Where {$_.GroupScope -eq "Universal" }
[int]$ADUniversalGroupsCount = $ADUniversalGroups.Count
$ADUniversalGroupsPCTofTotal = $ADUniversalGroupsCount / $AllADGroupsCount
$ADUniversalGroupsPCTofTotal = "{0:p2}"                      -f $ADUniversalGroupsPCTofTotal
Write-Verbose "There are $ADUniversalGroupsCount Universal groups in AD ($ADUniversalGroupsPCTofTotal of all groups) `r "

[array]$ADGlobalGroups = $AllADGroups | Where {$_.GroupScope -eq "Global" }
[int]$ADGlobalGroupsCount = $ADGlobalGroups.Count
$ADGlobalGroupsPCTofTotal = $ADGlobalGroupsCount / $AllADGroupsCount
$ADGlobalGroupsPCTofTotal = "{0:p2}"                      -f $ADGlobalGroupsPCTofTotal
Write-Verbose "There are $ADGlobalGroupsCount Global groups in AD ($ADGlobalGroupsPCTofTotal of all groups)  `r "

[array]$ADDomainLocalGroups = $AllADGroups | Where {$_.GroupScope -eq "DomainLocal" }
[int]$ADDomainLocalGroupsCount = $ADDomainLocalGroups.Count
$ADDomainLocalGroupsPCTofTotal = $ADDomainLocalGroupsCount / $AllADGroupsCount
$ADDomainLocalGroupsPCTofTotal = "{0:p2}"                      -f $ADDomainLocalGroupsPCTofTotal
Write-Verbose "There are $ADDomainLocalGroupsCount Domain Local groups in AD ($ADDomainLocalGroupsPCTofTotal of all groups)  `r "

[array]$ADSecurityGroups = $AllADGroups | Where {$_.GroupCategory -eq "Security" }
[int]$ADSecurityGroupsCount = $ADSecurityGroups.Count
$ADSecurityGroupsPCTofTotal = $ADSecurityGroupsCount / $AllADGroupsCount
$ADSecurityGroupsPCTofTotal = "{0:p2}"                      -f $ADSecurityGroupsPCTofTotal
Write-Verbose "There are $ADSecurityGroupsCount Security groups in AD ($ADSecurityGroupsPCTofTotal of all groups)  `r "

[array]$ADMESecurityGroups = $ADSecurityGroups | Where {$_.Mail -ne $Null }
[int]$ADMESecurityGroupsCount = $ADMESecurityGroups.Count
$ADMESecurityGroupsPCTofTotal = $ADMESecurityGroupsCount / $AllADGroupsCount
$ADMESecurityGroupsPCTofTotal = "{0:p2}"                      -f $ADMESecurityGroupsPCTofTotal
Write-Verbose "There are $ADSecurityGroupsCount Mail-Enabled Security groups in AD ($ADMESecurityGroupsPCTofTotal of all groups)  `r "

[array]$ADDistributionGroups = $AllADGroups | Where {$_.GroupCategory -eq "Distribution" }
[int]$ADDistributionGroupsCount = $ADDistributionGroups.Count
$ADDistributionGroupsPCTofTotal = $ADDistributionGroupsCount / $AllADGroupsCount
$ADDistributionGroupsPCTofTotal = "{0:p2}"                      -f $ADDistributionGroupsPCTofTotal
Write-Verbose "There are $ADDistributionGroupsCount Distribution groups in AD ($ADDistributionGroupsPCTofTotal of all groups)  `r "

[array]$ADMEnotUniGroups = $AllADGroups | Where { ($_.GroupCategory -eq "Distribution") -AND ($_.GroupScope -ne "Universal" ) }
[int]$ADMEnotUniGroupsCount = $ADMEnotUniGroups.Count
$ADMEnotUniGroupsPCTofTotal = $ADMEnotUniGroupsCount / $AllADGroupsCount
$ADMEnotUniGroupsPCTofTotal = "{0:p2}"                      -f $ADMEnotUniGroupsPCTofTotal
Write-Verbose "There are $ADMEnotUniGroupsCount Distribution groups that are not Universal groups in AD ($ADMEnotUniGroupsPCTofTotal of all groups)  `r "


##############################
# Get AD Computer Statistics # 20111027-20
##############################
$LastLoggedOnDate = $DateTime - (New-TimeSpan -days $ComputerLogonAge)
$PasswordStaleDate = $DateTime - (New-TimeSpan -days $ComputerPasswordAge)

Write-Output "Generating AD Computer Statistics  `r "

Write-Output "Getting list of AD Computers `r "
$Time = (Measure-Command `
	{[array] $AllComputers = Get-ADComputer -filter * -properties Name,CanonicalName,DNSHostName, Enabled,passwordLastSet,SAMAccountName,LastLogonTimeStamp,DistinguishedName,OperatingSystem }).TotalMinutes  
[int]$AllComputersCount = $AllComputers.Count
Write-Verbose "There are $AllComputersCount Computers discovered in $DomainDNS in $Time minutes...  `r "

Write-Verbose "Count Disabled Computers `r"
[array] $DisabledComputers = $AllComputers | Where-Object { $_.Enabled -eq $False }
[int] $DisabledComputersCount = $DisabledComputers.Count
Write-Verbose "Count Enabled Computers `r"
[array] $EnabledComputers = $AllComputers | Where-Object { $_.Enabled -eq $True }
[int] $EnabledComputersCount = $EnabledComputers.Count
Write-Verbose "There are $EnabledComputersCount Enabled Computers and there are $DisabledComputersCount Disabled Computers in $DomainDNS `r "

Write-Verbose "Count Inactive Computers... `r"
[array] $InactiveComputers = $AllComputers | Where-Object { $_.PasswordLastSet -lt $PasswordStaleDate }
[int] $InactiveComputersCount = $InactiveComputers.Count
Write-Verbose "There are $InactiveComputersCount Computers identified as Inactive (with passwords older than $ComputerPasswordAge days in $DomainDNS  `r "

# Count Workstations
Write-Verbose "Identifying all Workstation computers... `r"
[array]$WorkstationComputers = $AllComputers |  Where { ($_.OperatingSystem -notlike "*Server*") -and ($_.OperatingSystem -like "*Windows*") } 
[int] $WorkstationComputersCount = $WorkstationComputers.Count
Write-Verbose "After filtering server objects, There are $WorkstationComputersCount Windows workstations Total discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Enabled Workstation computers... `r"
[array]$EnabledWorkstationComputers = $EnabledComputers |  Where { ($_.OperatingSystem -notlike "*Server*") -and ($_.OperatingSystem -like "*Windows*") } 
[int] $EnabledWorkstationComputersCount = $EnabledWorkstationComputers.Count
Write-Verbose "There are $EnabledWorkstationComputersCount Enabled Windows workstations discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Workstation computers... `r"
[array]$ActiveEnabledWorkstationComputers = $EnabledComputers |  Where { ($_.OperatingSystem -notlike "*Server*") -and ($_.OperatingSystem -like "*Windows*") -and ($_.PasswordLastSet -gt $PasswordStaleDate)  } 
[int] $ActiveEnabledWorkstationComputersCount = $ActiveEnabledWorkstationComputers.Count
Write-Verbose " There are $ActiveEnabledWorkstationComputersCount Active Enabled Windows workstations discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Count Active Workstation computers who have logged on in the last 60 days `r "
[array] $ActiveEnabledWorkstationComputers60 = $WorkstationComputers | Where-Object { $_.PasswordLastSet -ge $60Days }
[int]$ActiveEnabledWorkstationComputers60Count = $ActiveEnabledWorkstationComputers60.Count
Write-Verbose " There are $ActiveEnabledWorkstationComputers60Count Active Enabled Windows workstations discovered in $ADDomainDNSRoot with a changed password in the last 60 days `r"

Write-Verbose "Count Active Workstation computers who have logged on in the last 90 days `r"
[array] $ActiveEnabledWorkstationComputers90 = $WorkstationComputers | Where-Object { $_.PasswordLastSet -ge $90Days }
[int]$ActiveEnabledWorkstationComputers90Count = $ActiveEnabledWorkstationComputers90.Count
Write-Verbose " There are $ActiveEnabledWorkstationComputers90Count Active Enabled Windows workstations discovered in $ADDomainDNSRoot with a changed password in the last 90 days `r"

Write-Verbose "Count Active Workstation computers who have logged on in the last 120 days `r"
[array] $ActiveEnabledWorkstationComputers120 = $WorkstationComputers | Where-Object { $_.PasswordLastSet -ge $120Days }
[int]$ActiveEnabledWorkstationComputers120Count = $ActiveEnabledWorkstationComputers120.Count
Write-Verbose " There are $ActiveEnabledWorkstationComputers120Count Active Enabled Windows workstations discovered in $ADDomainDNSRoot with a changed password in the last 120 days `r"

Write-Verbose "Count Active Workstation computers who have logged on in the last 180 days `r"
[array] $ActiveEnabledWorkstationComputers180 = $WorkstationComputers | Where-Object { $_.PasswordLastSet -ge $180Days }
[int]$ActiveEnabledWorkstationComputers180Count = $PKILastLogon180.Count
Write-Verbose " There are $ActiveEnabledWorkstationComputers180Count Active Enabled Windows workstations discovered in $ADDomainDNSRoot with a changed password in the last 180 days `r"

Write-Verbose "Identifying Active Workstation computers running Windows XP... `r"
[array]$WinXPActiveEnabledWorkstationComputers = $ActiveEnabledWorkstationComputers |  Where { ($_.OperatingSystem -like "*Windows XP*") } 
[int] $WinXPActiveEnabledWorkstationComputersCount = $WinXPActiveEnabledWorkstationComputers.Count
Write-Verbose " There are $WinXPActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows XP discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Workstation computers running Windows Vista... `r"
[array]$WinVistaActiveEnabledWorkstationComputers = $ActiveEnabledWorkstationComputers |  Where { ($_.OperatingSystem -like "*Windows Vista*") } 
[int]$WinVistaActiveEnabledWorkstationComputersCount = $WinVistaActiveEnabledWorkstationComputers.Count
Write-Verbose " There are $WinVistaActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows Vista discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Workstation computers running Windows 7... `r"
[array]$Win7ActiveEnabledWorkstationComputers = $ActiveEnabledWorkstationComputers |  Where { ($_.OperatingSystem -like "*Windows 7*") } 
[int] $Win7ActiveEnabledWorkstationComputersCount = $Win7ActiveEnabledWorkstationComputers.Count
Write-Verbose " There are $Win7ActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows 7 discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Workstation computers running Windows 8... `r"
[array]$Win8ActiveEnabledWorkstationComputers = $ActiveEnabledWorkstationComputers |  Where { ($_.OperatingSystem -like "*Windows 8*") } 
[int] $Win8ActiveEnabledWorkstationComputersCount = $Win8ActiveEnabledWorkstationComputers.Count
Write-Verbose " There are $Win8ActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows 8 discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying enabled workstations missing a DNS Host Name attribute... `r"
$BlankDNSHostNames = $EnabledWorkstationComputers | where { $_.DNSHostName -eq $NULL }
[int] $BlankDNSHostNamesCount = $BlankDNSHostNames.Count
Write-Verbose " There are $BlankDNSHostNamesCount Active Enabled Windows workstations missing a DNS Host Name attribute ...  `r " 

Write-Verbose "Identifying enabled workstations with a subdomain (NE)... `r"
$SDDNSHostNames = $EnabledWorkstationComputers | where { $_.DNSHostName -like "*.ne.*" }
[int] $SDDNSHostNamesCount = $SDDNSHostNames.Count
Write-Verbose " There are $SDDNSHostNamesCount Active Enabled Windows workstations with a subdomain (NE) ...  `r " 

# Count Servers
Write-Verbose "Identifying all Server computers...  `r "
[array]$ServerComputers = $AllComputers | Where {$_.OperatingSystem -like "*Server*"} 
[int] $ServerComputersCount = $ServerComputers.Count
Write-Verbose "After filtering workstation objects, There are $ServerComputersCount Windows Servers Total discovered in $ADDomainDNSRoot ...  `r "

Write-Verbose "Identifying Enabled Server computers...  `r "
[array]$EnabledServerComputers = $EnabledComputers | Where { ($_.Enabled -eq $True) -and ($_.OperatingSystem -like "*Server*") -and ($_.OperatingSystem -like "*Windows*") } 
[int] $EnabledServerComputersCount = $EnabledServerComputers.Count
Write-Verbose "There are $EnabledServerComputersCount Enabled Windows Servers discovered in $ADDomainDNSRoot ...  `r "

Write-Verbose "Identifying Active Enabled Server computers... `r"
[array]$ActiveEnabledServerComputers = $EnabledServerComputers |  Where { ($_.OperatingSystem -like "*Server*") -and ($_.OperatingSystem -like "*Windows*") -and ($_.PasswordLastSet -lt $PasswordStaleDate)  -and ($_.LastLogonDate -gt $LastLoggedOnDate)  } 
[int] $ActiveEnabledServerComputersCount = $ActiveEnabledServerComputers.Count
Write-Verbose " There are $ActiveEnabledServerComputersCount Active Enabled Windows workstations discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Enabled Server computers running Windows NT... `r"
[array]$NTActiveEnabledServerComputers = $EnabledServerComputers |  Where { ($_.OperatingSystem -like "*NT*") } 
[int] $NTActiveEnabledServerComputersCount = $NTActiveEnabledServerComputers.Count
Write-Verbose " There are $NTActiveEnabledServerComputersCount Active Enabled Windows workstations discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Enabled Server computers running Windows 2000... `r"
[array]$W2kActiveEnabledServerComputers = $EnabledServerComputers |  Where { ($_.OperatingSystem -like "*2000*") } 
[int] $W2kActiveEnabledServerComputersCount = $W2kActiveEnabledServerComputers.Count
Write-Verbose " There are $W2kActiveEnabledServerComputersCount Active Enabled Windows workstations discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Enabled Server computers running Windows 2003 (or Windows 2003 R2)... `r"
[array]$W2k3ActiveEnabledServerComputers = $EnabledServerComputers |  Where { ($_.OperatingSystem -like "*2003*") } 
[int] $W2k3ActiveEnabledServerComputersCount = $W2k3ActiveEnabledServerComputers.Count
Write-Verbose " There are $W2k3ActiveEnabledServerComputersCount Active Enabled Windows workstations discovered in $ADDomainDNSRoot ...  `r " 

Write-Verbose "Identifying Active Enabled Server computers running Windows 2008 (or Windows 2008 R2)... `r"
[array]$W2k8ActiveEnabledServerComputers = $EnabledServerComputers |  Where { ($_.OperatingSystem -like "*2008*") } 
[int] $W2k8ActiveEnabledServerComputersCount = $W2k8ActiveEnabledServerComputers.Count
Write-Verbose " There are $W2k8ActiveEnabledServerComputersCount Active Enabled Windows workstations discovered in $ADDomainDNSRoot ...  `r " 

ForEach ($ADBAckupStatusItem in $ADBackupStatus)
	{
		[array]$ADBackupStatusArray += $ADBAckupStatusItem
	}

##################
# Display Report # 20120120-10
##################
Write-output " `r "
Write-output " `r ";
$ReportData = 
@"

TODAY'S REPORT
==============
Generated: $ProcessStartTime
-----------------------------

ACTIVE DIRECTORY DETAILS:
-------------------------
Active Directory Forest Name : $ADForestName  `r 
Active Directory Domains in the Forest: $ADForestDomains  `r
Active Directory Current Domain Name : $ADDomainName  `r

Active Directory Instatiation Date: $ADInstatiationObjectWhenCreated `r

The AD Schema Version is $ADSchemaVersion which is $ADSchemaVersionName   `r
The Exchange Schema Version is $ExchangeSchemaVersion which is $ExchangeSchemaVersionName  `r

Active Directory's Tombstone Lifetime is set to $TombstoneLifetime days  `r
The LastLogonTimestamp attribute is configured to replicate every $LLTReplicationValue days.   `r
 This affects the accuracy of the User & Computer logon stats below (they may be off by ~ $LLTReplicationValue days).  `r

Domain Relative IDentifiers (RIDs) determine how many Security IDentifiers (SIDs) can be created.   `r
RIDs Issued: $CurrentRIDPoolCount ($RIDsIssuedPercentofTotal of total)  `r
RIDs Remaining: $RIDsRemaining ($RIDsRemainingPercentofTotal of total)  `r


ACTIVE DIRECTORY BACKUP STATUS:
-------------------------------
$ADBackupStatusArray


ACTIVE DIRECTORY OBJECT SNAPSHOT:
---------------------------------- 
Active Directory Schema Partition Stats:
$AllADSchemaObjectsCount objects in $SchemaNC 

Active Directory Configuration Partition Stats:  
$AllADSchemaObjectsCount objects in $ConfigurationNC
$AllADConfigurationDeletedObjectsCount Deleted objects in $ConfigurationNC  
$AllADConfigurationRecycledObjectsCount Recycled objects in $ConfigurationNC 

Active Directory Domain Partition Stats:  
$AllADDomainObjectsCount objects in $ADDomainDistinguishedName 
$AllADDomainDeletedObjectsCount Deleted objects in $ADDomainDistinguishedName 
$AllADDomainRecycledObjectsCount Recycled objects in $ADDomainDistinguishedName 

There are $ADSitesCount Active Directory Sites in the AD Forest.  `r
There are $DomainGPOsCount GPOs in the current Active Directory Domain.  `r


ACTIVE DIRECTORY FSMOs:  `r
-----------------------
AD Forest Naming Master : $ADForestDomainNamingMaster  `r
AD Forest Schema Master : $ADForestSchemaMaster   `r
AD Domain PDC Master : $ADDomainPDCEmulator  `r
AD Domain RID Master : $ADDomainRIDMaster  `r
AD Domain Infrastructure Master : $ADDomainInfrastructureMaster     `r

DOMAIN CONTROLLER INFORMATION:
------------------------------
Out of the $DomainControllersCount DCs in $DomainDNS :   `r
$DomainRODCsCount are RODCs & $DomainWritableDCsCount are writable DCs   `r
$2012DCCount are running Windows Server 2012   `r
$2008R2Sp1DCCount are running Windows Server 2008 R2 SP1   `r
$2008R2DCCount are running Windows Server 2008 R2 (No Service Pack)   `r
$2008DCCount are running Windows Server 2008 (Any Service Pack)    `r
$2003R2DcCount are running Windows Server 2003 R2 (Any Service Pack)   `r
$2003DcCount are running Windows Server 2003 (Any Service Pack)   `r
$2000DcCount are running Windows 2000 Server (Any Service Pack)  `r

DOMAIN USER STATISTICS:
-----------------------
There are $AllUsersCount user objects discovered in $ADDomainDNSRoot  `r
There are $EnabledUsersCount Enabled users and there are $DisabledUsersCount Disabled users in $DomainDNS   `r
There are $InactiveUsersCount users identified as Inactive (with passwords older than $ComputerPasswordAge days in $DomainDNS   `r
There are $EnabledServiceAccountsCount Enabled Service accounts in $DomainDNS (out of a total $AllServiceAccountsCount Service accounts)  `r
There are $MailboxUsersCount users in $DomainDNS with an Exchange Mailbox.   `r
There are $MailboxEnabledUsersCount Enabled users in $DomainDNS with an Exchange Mailbox.   `r
Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon30Count have logged on in the last 30 days (there may be up to a 14 day margin of error for this count)   `r
Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon45Count have logged on in the last 45 days   `r
Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon60Count have logged on in the last 60 days   `r
Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon90Count have logged on in the last 90 days   `r
Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon120Count have logged on in the last 120 days   `r
Out of $EnabledUsersCount Enabled users in $DomainDNS only $LastLogon180Count have logged on in the last 180 days    `r 
Out of All $AllUsersCount users in $DomainDNS $LastLogonNeverCount have NEVER logged on (no logon date associated with account)   `r
$LastLogon1Count Enabled users logged in within the last week  `r
$LastLogonYesterdayCount Enabled users logged in yesterday  `r
$LastLogontodayCount Enabled users logged in today (so far)  `r
$EnabledLockedUsersCount Enabled users are currently locked out  `r

DOMAIN GROUP STATISTICS:
------------------------
There are $ADUniversalGroupsCount Universal groups in AD ($ADUniversalGroupsPCTofTotal of all groups)  `r
There are $ADGlobalGroupsCount Global groups in AD ($ADGlobalGroupsPCTofTotal of all groups)  `r
There are $ADDomainLocalGroupsCount Domain Local groups in AD ($ADDomainLocalGroupsPCTofTotal of all groups)  `r
There are $ADSecurityGroupsCount Security groups in AD ($ADSecurityGroupsPCTofTotal of all groups)  `r
There are $ADSecurityGroupsCount Mail-Enabled Security groups in AD ($ADMESecurityGroupsPCTofTotal of all groups)  `r
There are $ADDistributionGroupsCount Distribution groups in AD ($ADDistributionGroupsPCTofTotal of all groups)  `r
There are $ADMEnotUniGroupsCount Distribution groups that are not Universal groups in AD ($ADMEnotUniGroupsPCTofTotal of all groups)  `r

DOMAIN COMPUTER STATISTICS:
---------------------------
There are $AllComputersCount Computers discovered in $DomainDNS  `r
There are $EnabledComputersCount Enabled Computers and there are $DisabledComputersCount Disabled Computers in $DomainDNS   `r
There are $InactiveComputersCount Computers identified as Inactive (with passwords older than $ComputerPasswordAge days in $DomainDNS  `r

WORKSTATIONS:
After filtering server objects, There are $WorkstationComputersCount Windows workstations discovered in $DomainDNS   `r
There are $EnabledWorkstationComputersCount Enabled Windows workstations discovered in $DomainDNS   `r
There are $ActiveEnabledWorkstationComputersCount Active Enabled Windows workstations discovered in $DomainDNS   `r
There are $WinXPActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows XP discovered in $DomainDNS   `r
There are $WinVistaActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows Vista discovered in $DomainDNS   `r
There are $Win7ActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows 7 discovered in $DomainDNS   `r
There are $Win8ActiveEnabledWorkstationComputersCount Active Enabled Windows workstations running Windows 8 discovered in $DomainDNS   `r
There are $BlankDNSHostNamesCount enabled workstations that have a blank DNS host name attribute   `r
There are $SDDNSHostNamesCount enabled workstations that have a DNS subdomain   `r

SERVERS: 
After filtering workstation objects, There are $ServerComputersCount SERVERS discovered in $DomainDNS  `r
There are $EnabledServerComputersCount Enabled Windows SERVERS discovered in $DomainDNS   `r
There are $ActiveEnabledServerComputersCount Active Enabled Windows SERVERS discovered in $DomainDNS   `r
There are $NTActiveEnabledServerComputersCount Active Enabled Windows SERVERS running Windows NT discovered in $DomainDNS   `r
There are $W2kActiveEnabledServerComputersCount Active Enabled Windows SERVERS running Windows 2000 discovered in $DomainDNS   `r
There are $W2k3ActiveEnabledServerComputersCount Active Enabled Windows SERVERS running Windows 2003 discovered in $DomainDNS   `r
There are $W2k8ActiveEnabledServerComputersCount Active Enabled Windows SERVERS running Windows 2008 discovered in $DomainDNS   `r

"@

$ReportData
Write-output " `r "
Write-output " `r "
                    
########################################
# Provide Script Processing Statistics #
########################################
$ProcessEndTime = Get-Date
Write-output "Script started processing at $ProcessStartTime and completed at $ProcessEndTime." `r 
$TotalProcessTimeCalc = $ProcessEndTime - $ProcessStartTime
$TotalProcessTime = "{0:HH:mm}" -f $TotalProcessTimeCalc
Write-output "" `r 
Write-output "The script completed in $TotalProcessTime." `r
Write-Output " `r "

#################
# Stop Logging  #
#################

#Stop logging the configuration changes in a transript file
#Stop-Transcript

Write-output "Review the logfile $LogFile for script operation information." `r                    
