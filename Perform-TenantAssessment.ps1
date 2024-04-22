<##Author: Sean McAvinue
##Details: Graph / PowerShell Script to assess a Microsoft 365 tenant for migration of Exchange, Teams, SharePoint and OneDrive, 
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Reports on multiple factors of a Microsoft 365 tenant to help with migration preparation. Exports results to Excel

        .DESCRIPTION
        Gathers information using Microsoft Graph API and Exchange Online Management Shell and Exports to CSV

        .PARAMETER ClientID
        Required - Application (Client) ID of the App Registration

        .PARAMETER TenantID
        Required - Directory (Tenant) ID of the Azure AD Tenant

        .PARAMETER certificateThumbprint
        Required - Thumbprint of the certificate generated from the prepare-tenantassessment.ps1 script    

        .Notes
        For similar scripts check out the links below
        
            Blog: https://seanmcavinue.net
            GitHub: https://github.com/smcavinue
            Twitter: @Sean_McAvinue
            Linkedin: https://www.linkedin.com/in/sean-mcavinue-4a058874/


    #>
Param(
    [parameter(Mandatory = $true)]
    $clientId,
    [parameter(Mandatory = $true)]
    $tenantId,
    [parameter(Mandatory = $true)]
    $certificateThumbprint,
    [parameter(Mandatory = $false)]
    [switch]$IncludeGroupMembership = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludeMailboxPermissions = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludeDocumentLibraries = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludeLists = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludePlans = $false
)

function UpdateProgress {
    Write-Progress -Activity "Tenant Assessment in Progress" -Status "Processing Task $ProgressTracker of $($TotalProgressTasks): $ProgressStatus" -PercentComplete (($ProgressTracker / $TotalProgressTasks) * 100)
}
$ProgressTracker = 1
$TotalProgressTasks = 27
$ProgressStatus = $null

if ($IncludeGroupMembership) {
    $TotalProgressTasks++
}

if ($IncludeMailboxPermissions) {
    $TotalProgressTasks++
}
if ($IncludeDocumentLibraries) {
    $TotalProgressTasks++
}
if ($IncludeLists) {
    $TotalProgressTasks++
}
if ($IncludePlans) {
    $TotalProgressTasks++
}

$ProgressStatus = "Checking Modules..."
UpdateProgress
$ProgressTracker++
##Import Modules
##Check if Microsoft.graph module is installed
$GraphModule = Get-Module -Name Microsoft.Graph -ListAvailable
$ExchangeModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable
$ImportExcelModule = Get-Module -Name ImportExcel -ListAvailable

if (!$GraphModule) {
    write-host "Microsoft.Graph module not installed, please install and re-run script" -ForegroundColor Red
    pause
    #exit
}
If (!$ExchangeModule) {
    write-host "ExchangeOnlineManagement module not installed, please install and re-run script" -ForegroundColor Red
    pause
    #exit
}
If (!$ImportExcelModule) {
    write-host "ImportExcel module not installed, please install and re-run script" -ForegroundColor Red
    pause
    #exit
}
$ProgressStatus = "Connecting to Microsoft Graph..."
UpdateProgress
$ProgressTracker++
##Attempt to get an Access Token
Try {
    $CertificatePath = "cert:\currentuser\my\$CertificateThumbprint"
    $Certificate = Get-Item $certificatePath
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $Certificate -NoWelcome
}
Catch {
    write-host "Unable to acquire access token, check the parameters are correct`n$($Error[0])"
    exit
}

$ProgressStatus = "Preparing environment..."
UpdateProgress
$ProgressTracker++

##Report File Name
$Filename = "TenantAssessment.xlsx"
##File Location
$FilePath = "C:\TenantAssessment"
Try {
    if (!(test-path -Path $FilePath)) {
        New-Item -Path $FilePath -ItemType Directory
    }
}
catch {
    write-host "Could not create folder at c:\temp - check you have appropriate permissions" -ForegroundColor red
    exit
}

##Check if cover page is present
$TemplatePath = $null
$TemplatePresent = $null
$TemplatePath = "TenantAssessment-Template.xlsx"
$TemplatePresent = Test-Path $TemplatePath


$ProgressStatus = "Getting users..."
UpdateProgress
$ProgressTracker++

##List All Tenant Users
$users = Get-MgUser -All -Property id, userprincipalname, mail, displayname, givenname, surname, licenseAssignmentStates, proxyaddresses, usagelocation, usertype, accountenabled, onPremisesSyncEnabled

$ProgressStatus = "Getting groups..."
UpdateProgress
$ProgressTracker++

##List all Tenant Groups
$Groups = get-mggroup -all

$ProgressStatus = "Getting Teams..."
UpdateProgress
$ProgressTracker++

##Get Teams details
$TeamGroups = $Groups | ? { ($_.grouptypes -Contains "unified") -and ($_.additionalproperties.resourceProvisioningOptions -contains "Team") }

$i = 1

foreach ($teamgroup in $TeamGroups) {

    $ProgressStatus = "Processing Team $i of $($Teamgroups.count)..."
    UpdateProgress
    $i++
    $ApiUri = "https://graph.microsoft.com/beta/teams/$($Teamgroup.id)/channels"
    $Teamchannels = ((Invoke-MgGraphRequest -Uri $ApiUri -Method Get).value)
    $standardchannels = ($teamchannels | ? { $_.membershipType -eq "standard" })
    $privatechannels = ($teamchannels | ? { $_.membershipType -eq "private" })
    $outgoingsharedchannels = ($teamchannels | ? { ($_.membershipType -eq "shared") -and (($_.WebUrl) -like "*$($teamgroup.id)*") })
    $incomingsharedchannels = ($teamchannels | ? { ($_.membershipType -eq "shared") -and ($_.WebURL -notlike "*$($teamgroup.id)*") })
    $teamgroup | Add-Member -MemberType NoteProperty -Name "StandardChannels" -Value $standardchannels.id.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannels" -Value $privatechannels.id.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannels" -Value $outgoingsharedchannels.id.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "IncomingSharedChannels" -Value $incomingsharedchannels.id.count -Force
    $privatechannelSize = 0
    
    foreach ($Privatechannel in $privatechannels) {
        $PrivateChannelObject = $null
        Try {
            $PrivatechannelObject = Get-MgTeamChannelFileFolder -TeamId $teamgroup.id -ChannelId $Privatechannel.id
            $Privatechannelsize += $PrivateChannelObject.size

        }
        Catch {
            $Privatechannelsize += 0
        }
    }

    $sharedchannelSize = 0
    
    foreach ($sharedchannel in $outgoingsharedchannels) {
        $sharedChannelObject = $null
        Try {
            $SharedChannelObject = Get-MgTeamChannelFileFolder -TeamId $teamgroup.id -ChannelId $sharedChannel.id
            $Sharedchannelsize += $SharedChannelObject.size

        }
        Catch {
            $Sharedchannelsize += 0
        }
    }

    $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannelsSize" -Value $privatechannelSize -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannelsSize" -Value $sharedchannelSize -Force
    

    $TeamDetails = $null
    [array]$TeamDetails = Get-MgGroupDrive -GroupId $teamgroup.id
    $teamgroup | Add-Member -MemberType NoteProperty -Name "DocumentLibraries" -Value $TeamDetails.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "DataSize" -Value ($TeamDetails.quota.used | measure-object -sum).sum -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "URL" -Value $TeamDetails[0].webUrl.replace("/Shared%20Documents", "") -Force

}

$ProgressStatus = "Getting licenses..."
UpdateProgress
$ProgressTracker++

##Get All License SKUs
$SKUs = Get-MgSubscribedSku -all

$ProgressStatus = "Getting organization details..."
UpdateProgress
$ProgressTracker++

##Get Org Details
$OrgDetails = Get-MgOrganization -All

$ProgressStatus = "Getting apps..."
UpdateProgress
$ProgressTracker++

##List All Azure AD Service Principals
[array]$AADApps = Get-MgServicePrincipal -All

foreach ($user in $users) {
    $user | Add-Member -MemberType NoteProperty -Name "License SKUs" -Value ($user.licenseAssignmentStates.skuid -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Group License Assignments" -Value ($user.licenseAssignmentStates.assignedByGroup -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Disabled Plan IDs" -Value ($user.licenseAssignmentStates.disabledplans -join ";") -Force
}

##Translate License SKUs and groups
foreach ($user in $users) {

    foreach ($Group in $Groups) {
        $user.'Group License Assignments' = $user.'Group License Assignments'.replace($group.id, $group.displayName) 
    }
    foreach ($SKU in $SKUs) {
        $user.'License SKUs' = $user.'License SKUs'.replace($SKU.skuid, $SKU.skuPartNumber)
    }
    foreach ($SKUplan in $SKUs.servicePlans) {
        $user.'Disabled Plan IDs' = $user.'Disabled Plan IDs'.replace($SKUplan.servicePlanId, $SKUplan.servicePlanName)
    }

}

$ProgressStatus = "Getting Conditional Access policies..."
UpdateProgress
$ProgressTracker++

##Get Conditional Access Policies
[array]$ConditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -All

##Get Directory Roles
[array]$DirectoryRoleTemplates = Get-MgDirectoryRoleTemplate

##Get Trusted Locations
[array]$NamedLocations = Get-MgIdentityConditionalAccessNamedLocation

##Tidy GUIDs to names
$ConditionalAccessPoliciesJSON = $ConditionalAccessPolicies | ConvertTo-Json -Depth 5
if ($ConditionalAccessPoliciesJSON -ne $null) {
    ##TidyUsers
    foreach ($User in $Users) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($user.id, ("$($user.displayname) - $($user.userPrincipalName)"))
    }

    ##Tidy Groups
    foreach ($Group in $Groups) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($group.id, ("$($group.displayname) - $($group.id)"))
    }

    ##Tidy Roles
    foreach ($DirectoryRoleTemplate in $DirectoryRoleTemplates) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($DirectoryRoleTemplate.Id, $DirectoryRoleTemplate.displayname)
    }

    ##Tidy Apps
    foreach ($AADApp in $AADApps) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.appid, $AADApp.displayname)
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.id, $AADApp.displayname)
    }

    ##Tidy Locations
    foreach ($NamedLocation in $NamedLocations) {

        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($NamedLocation.id, $NamedLocation.displayname)
    }


    $ConditionalAccessPolicies = $ConditionalAccessPoliciesJSON | ConvertFrom-Json


    $CAOutput = @()
    $CAHeadings = @(
        "displayName",
        "createdDateTime",
        "modifiedDateTime",
        "state",
        "Conditions.users.includeusers",
        "Conditions.users.excludeusers",
        "Conditions.users.includegroups",
        "Conditions.users.excludegroups",
        "Conditions.users.includeroles",
        "Conditions.users.excluderoles",
        "Conditions.clientApplications.includeServicePrincipals",
        "Conditions.clientApplications.excludeServicePrincipals",
        "Conditions.applications.includeApplications",
        "Conditions.applications.excludeApplications",
        "Conditions.applications.includeUserActions",
        "Conditions.applications.includeAuthenticationContextClassReferences",
        "Conditions.userRiskLevels",
        "Conditions.signInRiskLevels",
        "Conditions.platforms.includePlatforms",
        "Conditions.platforms.excludePlatforms",
        "Conditions.locations.includLocations",
        "Conditions.locations.excludeLocations"
        "Conditions.clientAppTypes",
        "Conditions.devices.deviceFilter.mode",
        "Conditions.devices.deviceFilter.rule",
        "GrantControls.operator",
        "grantcontrols.builtInControls",
        "grantcontrols.customAuthenticationFactors",
        "grantcontrols.termsOfUse",
        "SessionControls.disableResilienceDefaults",
        "SessionControls.applicationEnforcedRestrictions",
        "SessionControls.persistentBrowser",
        "SessionControls.cloudAppSecurity",
        "SessionControls.signInFrequency"

    )

    Foreach ($Heading in $CAHeadings) {
        $Row = $null
        $Row = New-Object psobject -Property @{
            PolicyName = $Heading
        }
    
        foreach ($CAPolicy in $ConditionalAccessPolicies) {
            $Nestingcheck = ($Heading.split('.').count)

            if ($Nestingcheck -eq 1) {
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value $CAPolicy.$Heading -Force
            }
            elseif ($Nestingcheck -eq 2) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()) -join ';' )-Force
            }
            elseif ($Nestingcheck -eq 3) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()).($SplitHeading[2].ToString()) -join ';' )-Force
            }
            elseif ($Nestingcheck -eq 4) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()).($SplitHeading[2].ToString()).($SplitHeading[3].ToString()) -join ';' )-Force       
            }
        }

        $CAOutput += $Row

    }

}
$ProgressStatus = "Getting OneDrive report..."
UpdateProgress
$ProgressTracker++

##Get OneDrive Report##
Get-MgReportOneDriveUsageAccountDetail -Period "D30" -OutFile "$($FilePath)\OneDriveReport.csv"
$OneDrive = import-csv "$($FilePath)\OneDriveReport.csv"
Remove-Item "$($FilePath)\OneDriveReport.csv"

$ProgressStatus = "Getting SharePoint report..."
UpdateProgress
$ProgressTracker++

##Get SharePoint Report##
Get-MgReportSharePointSiteUsageDetail -Period "D30" -OutFile "$($FilePath)\SharePointReport.csv"
$SharePoint = import-csv "$($FilePath)\SharePointReport.csv"
Remove-Item "$($FilePath)\SharePointReport.csv"
$SharePoint | Add-Member -MemberType NoteProperty -Name "TeamID" -Value "" -force
foreach ($Site in $Sharepoint) {
    $DriveLookup = ((Get-MgSiteDrive -siteId $Site.'Site Id' -ErrorAction SilentlyContinue | ? { $_.name -eq "Documents" }).weburl)
    If ($DriveLookup) {
        $Site.'Site URL' = $DriveLookup.replace('/Shared%20Documents', '')
    }
    $Site.TeamID = ($TeamGroups | ? { $_.url -contains $site.'site url' }).id

}

$ProgressStatus = "Getting Mailbox Usage report..."
UpdateProgress
$ProgressTracker++

##Get Mailbox Report##
Get-MgReportMailboxUsageDetail -Period "D30" -OutFile "$($FilePath)\MailboxReport.csv"
$MailboxStatsReport = import-csv "$($FilePath)\MailboxReport.csv"
Remove-Item "$($FilePath)\MailboxReport.csv"

##Get M365 Apps usage report
Get-MgReportOffice365ServiceUserCount -Period "D30" -OutFile "$($FilePath)\M365AppsUsage.csv"
$M365AppsUsage = import-csv "$($FilePath)\M365AppsUsage.csv"
Remove-Item "$($FilePath)\M365AppsUsage.csv"

##Process Group Membership
If ($IncludeGroupMembership) {
    $ProgressStatus = "Enumerating Group Membership - This may take some time..."
    UpdateProgress
    $GroupMembersObject = @()
    $i = 1
    foreach ($group in $groups) {
        $ProgressStatus = "Enumerating Group Membership - This may take some time... Processing Group $i of $($Groups.count)"
        UpdateProgress
        $i++
        $Members = get-mggroupmember -groupid $group.id -all
        foreach ($member in $members) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.AdditionalProperties["displayName"]
                MemberUserPrincipalName = $member.AdditionalProperties["userPrincipalName"]
                MemberType              = "Member"
                MemberObjectType        = $member.AdditionalProperties["@odata.type"].replace('#microsoft.graph.', '')

            }

            $GroupMembersObject += $memberEntry

        }

        $Owners = Get-MgGroupOwner -GroupId $group.id -All
        foreach ($member in $Owners) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.AdditionalProperties["displayName"]
                MemberUserPrincipalName = $member.AdditionalProperties["userPrincipalName"]
                MemberType              = "Owner"
                MemberObjectType        = $member.AdditionalProperties["@odata.type"].replace('#microsoft.graph.', '')

            }

            $GroupMembersObject += $memberEntry

        }
    }


    $ProgressTracker++
}

If ($IncludeDocumentLibraries) {
    $ProgressStatus = "Enumerating Document Libraries - This may take some time..."
    UpdateProgress
    $Sites = Get-MgSite -All | ? { $_.weburl -notlike "*sites/appcatalog*" -and $_.weburl -notlike "*sites/recordscenter*" -and $_.weburl -notlike "*sites/search*" -and $_.weburl -notlike "*sites/CompliancePolicyCenter" -and $_.weburl -notlike "-my.sharepoint.com*" }
    $LibraryOutput = @()
    foreach ($site in $sites) {
        [array]$Drives = Get-MgSiteDrive -SiteId $site.id | ? { $_.Name -eq "Documents" }
        foreach ($drive in $drives) {
            $LibraryObject = [PSCustomObject]@{
                LibraryID    = $Drive.id
                LibraryName  = $Drive.Name
                LibraryURL   = $Drive.WebUrl
                LibraryUsage = $Drive.quota.used
                SiteID       = $Site.id
                SiteName     = $Site.DisplayName
                SiteURL      = $Site.WebURL
            }
            $LibraryOutput += $LibraryObject
        }
    }
    $ProgressTracker++
}

If ($IncludeLists) {
    $ProgressStatus = "Enumerating Lists - This may take some time..."
    UpdateProgress
    $Sites = Get-MgSite -All | ? { $_.weburl -notlike "*sites/appcatalog*" -and $_.weburl -notlike "*sites/recordscenter*" -and $_.weburl -notlike "*sites/search*" -and $_.weburl -notlike "*sites/CompliancePolicyCenter" -and $_.weburl -notlike "-my.sharepoint.com*" }
    $ListOutput = @()
    foreach ($site in $sites) {
        [array]$Lists = Get-MgSiteList -SiteId $site.id | ? { $_.List.template -ne "documentLibrary" }
        foreach ($list in $lists) {
            $ListObject = [PSCustomObject]@{
                ListID   = $list.id
                ListName = $List.DisplayName
                ListURL  = $List.webUrl
                SiteID   = $Site.id
                SiteName = $Site.DisplayName
                SiteURL  = $Site.WebURL
            }
            $ListOutput += $ListObject
        }
    }
    $ProgressTracker++
}

if ($IncludePlans) {
    $ProgressStatus = "Enumerating Planner Plans - This may take some time..."
    UpdateProgress
    $unifiedGroups = $Groups | ? { ($_.grouptypes -Contains "unified") 
        $PlanOutput = @()
        foreach ($unifiedgroup in $unifiedGroups) {
            [array]$Plans = Get-MgGroupPlannerPlan -GroupId $unifiedgroup.id
            foreach ($plan in $plans) {
                $PlanObject = [PSCustomObject]@{
                    PlanID    = $plan.id
                    PlanName  = $plan.title
                    GroupID   = $unifiedgroup.id
                    GroupName = $unifiedgroup.displayName
                }
                $PlanOutput += $PlanObject
            }
        }
    }
    $ProgressTracker++
}
##Tidy up Proxyaddresses
foreach ($user in $users) {
    $user | Add-member -MemberType NoteProperty -Name "Email Addresses" -Value ($user.proxyaddresses -join ';') -Force
}
##Tidy up Proxyaddresses
foreach ($group in $groups) {
    $group | Add-member -MemberType NoteProperty -Name "Email Addresses" -Value ($group.proxyaddresses -join ';') -Force
}

###################EXCHANGE ONLINE############################

$ProgressStatus = "Connecting to Exchange Online..."
UpdateProgress
$ProgressTracker++

Try {
    Connect-ExchangeOnline -Certificate $Certificate -AppID $clientid -Organization ($orgdetails.verifieddomains | ? { $_.isinitial -eq "true" }).name -ShowBanner:$false
}
catch {
    write-host "Error connecting to Exchange Online...Exiting..." -ForegroundColor red
    Pause
    Exit
}

$ProgressStatus = "Getting shared and room mailboxes..."
UpdateProgress
$ProgressTracker++
##Get Shared and Resource Mailboxes

[array]$RoomMailboxes = Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize unlimited
[array]$EquipmentMailboxes = Get-EXOMailbox -RecipientTypeDetails EquipmentMailbox -ResultSize unlimited
[array]$SharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

$ProgressStatus = "Getting room mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 1

##Get Resource Mailbox Sizes
foreach ($room in $RoomMailboxes) {
    $ProgressStatus = "Getting room mailbox statistics $i of $($RoomMailboxes.count)..."
    $i++
    UpdateProgress

    $RoomStats = $null
    $RoomStats = get-EXOmailboxstatistics $room.primarysmtpaddress
    $room | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $RoomStats.TotalItemSize -Force
    $room | Add-Member -MemberType NoteProperty -Name ItemCount -Value $RoomStats.ItemCount -Force

    ##Clean email addresses value
    $room.EmailAddresses = $room.EmailAddresses -join ';'
}

$ProgressStatus = "Getting Equipment mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 1

foreach ($equipment in $EquipmentMailboxes) {
    $ProgressStatus = "Getting Equipment mailbox statistics $i of $($EquipmentMailboxes.count)..."
    $i++
    UpdateProgress

    $EquipmentStats = $null
    $EquipmentStats = get-EXOmailboxstatistics $equipment.primarysmtpaddress
    $equipment | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $EquipmentStats.TotalItemSize -Force
    $equipment | Add-Member -MemberType NoteProperty -Name ItemCount -Value $EquipmentStats.ItemCount -Force

    ##Clean email addresses value
    $equipment.EmailAddresses = $equipment.EmailAddresses -join ';'
}


$ProgressStatus = "Getting shared mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 1

##Get Shared Mailbox Sizes
foreach ($SharedMailbox in $SharedMailboxes) {
    $ProgressStatus = "Getting shared mailbox statistics $i of $($SharedMailboxes.count)..."
    $i++
    UpdateProgress

    $SharedStats = $null
    $SharedStats = get-EXOmailboxstatistics $SharedMailbox.primarysmtpaddress
    $SharedMailbox | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $SharedStats.TotalItemSize -Force
    $SharedMailbox | Add-Member -MemberType NoteProperty -Name ItemCount -Value $SharedStats.ItemCount -Force
    
    ##Clean email addresses value
    $SharedMailbox.EmailAddresses = $SharedMailbox.EmailAddresses -join ';'
}

$ProgressStatus = "Getting user mailbox statistics..."
UpdateProgress
$ProgressTracker++

##Collect Mailbox statistics
$MailboxStats = @()
foreach ($user in ($users | ? { ($_.mail -ne $null ) -and ($_.userType -eq "Member") })) {
    $stats = $null
    $stats = $MailboxStatsReport | ? { $_.'User Principal Name' -eq $user.userprincipalname }
    $stats | Add-Member -MemberType NoteProperty -Name ObjectID -Value $user.id -Force
    $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $user.mail -Force
    $MailboxStats += $stats
    

}

$ProgressStatus = "Getting archive mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 0

##Collect Archive Statistics
$ArchiveStats = @()
[array]$ArchiveMailboxes = get-EXOmailbox -Archive -ResultSize unlimited
foreach ($archive in $ArchiveMailboxes) {
    $ProgressStatus = "Getting archive mailbox statistics $i of $($ArchiveMailboxes.count)..."
    $i++
    UpdateProgress
    $stats = $null
    $stats = get-EXOmailboxstatistics $archive.PrimarySmtpAddress -Archive #-erroraction SilentlyContinue
    $stats | Add-Member -MemberType NoteProperty -Name ObjectID -Value $archive.ExternalDirectoryObjectId -Force
    $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $archive.primarysmtpaddress -Force
    $ArchiveStats += $stats
    
}

$ProgressStatus = "Getting mail contacts..."
UpdateProgress
$ProgressTracker++

##Collect Mail Contacts
##Collect transport rules

$MailContacts = Get-MailContact -ResultSize unlimited | select displayname, alias, externalemailaddress, emailaddresses, HiddenFromAddressListsEnabled
foreach ($mailcontact in $MailContacts) {
    $mailcontact.emailaddresses = $mailcontact.emailaddresses -join ';'
}

$ProgressStatus = "Getting transport rules..."
UpdateProgress
$ProgressTracker++

##Collect transport rules

$Rules = $null
[array]$Rules = Get-TransportRule -ResultSize unlimited | select name, state, mode, priority, description, comments
$RulesOutput = @()
##Output rules to variable
foreach ($Rule in $Rules) {

    $RulesOutput += $Rule

}


#######Optional Items - EXO#######

##Process Mailbox Permissions
If ($IncludeMailboxPermissions) {
    $ProgressStatus = "Fetching Mailbox Permissions - This may take some time..."
    UpdateProgress
    $PermissionOutput = @()
    ##Get all mailboxes
    $MailboxList = Get-EXOMailbox -ResultSize unlimited
    $PermissionProgress = 1
    foreach ($mailbox in $MailboxList) {
        $ProgressStatus = "Fetching Mailbox Permissions for mailbox $PermissionProgress of $($Mailboxlist.count) - This may take some time..."
        UpdateProgress

        

        [array]$Permissions = Get-EXOMailboxPermission -UserPrincipalName $mailbox.UserPrincipalName | ? { $_.User -ne "NT AUTHORITY\SELF" }

        foreach ($permission in $Permissions) {

            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.user

            }
            
            $PermissionOutput += $PermissionObject
        }

        [array]$RecipientPermissions = Get-EXORecipientPermission $mailbox.UserPrincipalName |  ? { $_.Trustee -ne "NT AUTHORITY\SELF" }

        foreach ($permission in $RecipientPermissions) {

            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.trustee

            }
            
            $PermissionOutput += $PermissionObject
        }

        $PermissionProgress++
    }
    $ProgressTracker++

}

#######Report Export#######

$ProgressStatus = "Getting mail connectors..."
UpdateProgress
$ProgressTracker++

##Collect Mailflow Connectors

$InboundConnectors = Get-InboundConnector | select enabled, name, connectortype, connectorsource, SenderIPAddresses, SenderDomains, RequireTLS, RestrictDomainsToIPAddresses, RestrictDomainsToCertificate, CloudServicesMailEnabled, TreatMessagesAsInternal, TlsSenderCertificateName, EFTestMode, Comment 
foreach ($inboundconnector in $InboundConnectors) {
    $inboundconnector.senderipaddresses = $inboundconnector.senderipaddresses -join ';'
    $inboundconnector.senderdomains = $inboundconnector.senderdomains -join ';'
}
$OutboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors:$true | select enabled, name, connectortype, connectorsource, TLSSettings, RecipientDomains, UseMXRecord, SmartHosts, Comment
foreach ($OutboundConnector in $OutboundConnectors) {
    $OutboundConnector.RecipientDomains = $OutboundConnector.RecipientDomains -join ';'
    $OutboundConnector.SmartHosts = $OutboundConnector.SmartHosts -join ';'
}
$ProgressStatus = "Getting MX records..."
UpdateProgress
$ProgressTracker++

##MX Record Check
$MXRecordsObject = @()
foreach ($domain in $orgdetails.verifieddomains) {
    Try {
        [array]$MXRecords = Resolve-DnsName -Name $domain.name -Type mx -ErrorAction SilentlyContinue
    }
    catch {
        write-host "Error obtaining MX Record for $($domain.name)"
    }
    foreach ($MXRecord in $MXRecords) {
        $MXRecordsObject += $MXRecord
    }
}

$ProgressStatus = "Updating references..."
UpdateProgress
$ProgressTracker++

##Update users tab with Values
$users | Add-Member -MemberType NoteProperty -Name MailboxSizeGB -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name MailboxItemCount -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name OneDriveSizeGB -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name OneDriveFileCount -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name ArchiveSizeGB -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name Mailboxtype -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name ArchiveItemCount -Value "" -Force

foreach ($user in ($users | ? { $_.usertype -ne "Guest" })) {
    ##Set Mailbox Type
    if ($roommailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Room"
    }    
    elseif ($EquipmentMailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Equipment"
    }
    elseif ($sharedmailboxes.ExternalDirectoryObjectId -contains $user.id) {
        $user.Mailboxtype = "Shared"
    }
    else {
        $user.Mailboxtype = "User"
    }

    ##Set Mailbox Size and count
    If ($MailboxStats | ? { $_.objectID -eq $user.id }) {
        $user.MailboxSizeGB = (((($MailboxStats | ? { $_.objectID -eq $user.id }).'Storage Used (Byte)' / 1024) / 1024) / 1024) 
        $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
        $user.MailboxItemCount = ($MailboxStats | ? { $_.objectID -eq $user.id }).'item count'
    }

    ##Set Shared Mailbox size and count
    If ($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($SharedMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }

    ##Set Equipment Mailbox size and count
    If ($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($EquipmentMailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }


    ##Set Room Mailbox size and count
    If ($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }) {
        if (($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize) {
            $user.MailboxSizeGB = (((($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).mailboxsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
            $user.MailboxSizeGB = [math]::Round($user.MailboxSizeGB, 2)
            $user.MailboxItemCount = ($roommailboxes | ? { $_.ExternalDirectoryObjectId -eq $user.id }).ItemCount
        }
    }

    ##Set archive size and count
    If ($ArchiveStats | ? { $_.objectID -eq $user.id }) {
        $user.ArchiveSizeGB = (((($ArchiveStats | ? { $_.objectID -eq $user.id }).totalitemsize.value.tostring().replace(',', '').replace(' ', '').split('b')[0].split('(')[1] / 1024) / 1024) / 1024) 
        $user.ArchiveSizeGB = [math]::Round($user.ArchiveSizeGB, 2)
        $user.ArchiveItemCount = ($ArchiveStats | ? { $_.objectID -eq $user.id }).ItemCount
    }

    ##Set OneDrive Size and count
    if ($OneDrive | ? { $_.'Owner Principal Name' -eq $user.userPrincipalName }) {
        if (($OneDrive | ? { $_.'Owner Principal Name' -eq $user.userPrincipalName }).'Storage Used (Byte)') {
            $user.OneDriveSizeGB = (((($OneDrive | ? { $_.'Owner Principal Name' -eq $user.UserPrincipalName }).'Storage Used (Byte)' / 1024) / 1024) / 1024)
            $user.OneDriveSizeGB = [math]::Round($user.OneDriveSizeGB, 2)
            $user.OneDriveFileCount = ($OneDrive | ? { $_.'Owner Principal Name' -eq $user.UserPrincipalName }).'file count'
        }
    }
}




$ProgressStatus = "Exporting report..."
UpdateProgress
$ProgressTracker++
Try {
    IF ($TemplatePresent) {
        ##Add cover sheet
        Copy-ExcelWorksheet -SourceObject TenantAssessment-Template.xlsx -SourceWorksheet "High-Level" -DestinationWorkbook "$FilePath\$Filename" -DestinationWorksheet "High-Level"
        
    }
    $users | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force
    $SharePoint | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force
    $TeamGroups | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force
    $Groups | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force

    ##Export Data File##
    ##Export User Accounts tab
    $users | ? { ($_.usertype -ne "Guest") -and ($_.mailboxtype -eq "User") } | Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, OneDriveSizeGB, OneDriveFileCount, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, 'Email addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype, onPremisesSyncEnabled  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "User Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export Shared Mailboxes tab
    $users | ? { ($_.usertype -ne "Guest") -and ($_.mailboxtype -eq "shared") } | Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, 'Email Addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype, onPremisesSyncEnabled  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Shared Mailboxes" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export Resource Accounts tab
    $users | ? { ($_.usertype -ne "Guest") -and (($_.mailboxtype -eq "Room") -or ($_.mailboxtype -eq "Equipment")) } | Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, 'Email Addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype, onPremisesSyncEnabled  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Resource Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export SharePoint Tab
    $SharePoint | ? { ($_.teamid -eq $null) -and ($_.'Root Web Template' -ne "Team Channel") } | select Migrate, 'Site ID', 'Site URL', 'Owner Display Name', 'Is Deleted', 'Last Activity Date', 'File Count', 'Active File Count', 'Page View Count', 'Storage Used (Byte)', 'Root Web Template', 'Owner Principal Name' | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "SharePoint Sites" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Teams Tab
    $TeamGroups | select Migrate, id, displayname, standardchannels, privatechannels, SharedChannels, Datasize, PrivateChannelsSize, SharedChannelsSize, IncomingSharedChannels, mail, URL, description, createdDateTime, mailEnabled, securityenabled, mailNickname, 'Email Addresses', visibility | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Teams"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Unified Groups tab
    $Groups | ? { ($_.grouptypes -Contains "unified") -and ($_.resourceProvisioningOptions -notcontains "Team") } | select Migrate, id, displayname, mail, description, createdDateTime, mailEnabled, securityenabled, mailNickname, 'Email Addresses', visibility, membershipRule | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Unified Groups"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Standard Groups tab
    $Groups | ? { $_.grouptypes -notContains "unified" } | select Migrate, id, displayname, mail, description, createdDateTime, mailEnabled, securityenabled, mailNickname, 'Email Addresses', visibility, membershipRule | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Standard Groups"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Guest Accounts tab
    $users | ? { $_.usertype -eq "Guest" } | Select-Object id, accountenabled, userPrincipalName, mail, displayName, givenName, surname, 'Email Addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usertype | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Guest Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 
    ##Export AAD Apps Tab
    $AADApps | ? { $_.publishername -notlike "Microsoft*" } | select createddatetime, displayname, publisherName, signinaudience | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "AAD Apps" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Conditional Access Tab
    $CAOutput   | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Conditional Access" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export M365 Apps Usage
    $M365AppsUsage  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "M365 Apps Usage" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Mail Contacts tab
    $MailContacts | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "MailContacts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export MX Records tab
    $MXRecordsObject | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "MX Records"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Verified Domains tab
    $orgdetails.verifieddomains | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Verified Domains"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Transport Rules tab
    $RulesOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Transport Rules" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Receive Connectors Tab
    $InboundConnectors  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Receive Connectors" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export Send Connectors Tab
    $OutboundConnectors  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Send Connectors" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    ##Export OneDrive Tab
    $OneDrive  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "OneDrive Sites" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    If ($IncludeMailboxPermissions) {
        ##Export Mailbox Permissions Tab
        $PermissionOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Mailbox Permissions" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    If ($IncludeGroupMembership) {
        ##Export Group Membership Tab
        $GroupMembersObject | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Group Membership" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    If ($IncludeDocumentLibraries) {
        ##Export Document Libraries Tab
        $LibraryOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Document Libraries" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    
    }
    If ($IncludeLists) {
        ##Export Lists Tab
        $ListOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Lists" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
}
catch {
    write-host "Error exporting report, check permissions and make sure the file is not open! $_"
    pause

}

$ProgressStatus = "Finalizing..."
UpdateProgress
$ProgressTracker++

