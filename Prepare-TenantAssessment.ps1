<##Author: Sean McAvinue
##Details: PowerShell Script to Configure an Application Registration with the appropriate permissions to run Perform-TenantAssessment.ps1
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Creates an app reg with the appropriate permissions to run the tenant assessment script and uploads a self signed certificate

        .DESCRIPTION
        Connects to Microsoft Graph API and provisions an app reg with the appropriate permissions

        .Notes
        For similar scripts check out the links below
        
            Blog: https://seanmcavinue.net
            GitHub: https://github.com/smcavinue
            Twitter: @Sean_McAvinue
            Linkedin: https://www.linkedin.com/in/sean-mcavinue-4a058874/

            For full instructions on how to use this script, please visit the blog posts below:
            https://practical365.com/office-365-migration-plan-assessment/
            https://practical365.com/microsoft-365-tenant-to-tenant-migration-assessment-version-2/

        .EXAMPLE
        .\Prepare-TenantAssessment.ps1
        
    #>

##
function New-AadApplicationCertificate {
    [CmdletBinding(DefaultParameterSetName = 'DefaultSet')]
    Param(
        [Parameter(mandatory = $true, ParameterSetName = 'ClientIdSet')]
        [string]$ClientId,

        [string]$CertificateName,

        [Parameter(mandatory = $false, ParameterSetName = 'ClientIdSet')]
        [switch]$AddToApplication
    )
    ##Function source: https://www.powershellgallery.com/packages/AadSupportPreview/0.3.8/Content/functions%5CNew-AadApplicationCertificate.ps1

    # Create self-signed Cert
    $notAfter = (Get-Date).AddYears(2)

    try {
        $cert = (New-SelfSignedCertificate -DnsName "TenantAssessment" -CertStoreLocation "cert:\currentuser\My" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter)
        
    }

    catch {
        Write-Error "ERROR. Probably need to run as Administrator."
        Write-host $_
        return
    }

    if ($AddToApplication) {
        $Key = @{
            Type  = "AsymmetricX509Cert";
            Usage = "Verify";
            key   = $cert.RawData
        }
        Update-MgApplication -ApplicationId $ClientId -KeyCredentials $Key
    }
    Return $cert.Thumbprint
}

write-host "Provisioning Entra App Registration for Tenant Migration Assessment Tool" -ForegroundColor Green
##Name of the app
$appName = "Tenant Assessment Tool"
##Consent URL
$ConsentURl = "https://login.microsoftonline.com/{tenant-id}/adminconsent?client_id={client-id}"

##Attempt Graph API connection until successful
$Context = get-mgcontext 
while (!$Context) {
    Try {
        Connect-MgGraph -NoWelcome -Scopes "Application.ReadWrite.All RoleManagement.ReadWrite.Directory"
        $Context = get-mgcontext
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: `n$($error[0])`n Try again..." -ForegroundColor Red
        $Context = $null
    }
}

##Create Resource Access Variable
Try {
    $params = @{
        RequiredResourceAccess = @()
    }
    ##Create Resource Access Variable
    ##Get the EXO Api
    $EXOapi = (Get-MgServicePrincipal -Filter "AppID eq '00000002-0000-0ff1-ce00-000000000000'")
    ## Get the Exchange Online API permission ID
    $EXOpermission = $EXOapi.AppRoles | Where-Object { $_.Value -eq 'Exchange.ManageAsApp' }
    $params.RequiredResourceAccess = @{
        ResourceAppId  = $EXOapi.AppId
        ResourceAccess = @(
            @{
                Id   = $EXOpermission.id
                Type = "Role"
            }
        )
    }
    $params.RequiredResourceAccess = @{
        ResourceAppId  = "00000003-0000-0000-c000-000000000000"
        ResourceAccess = @(
            @{
                Id   = "332a536c-c7ef-4017-ab91-336970924f0d"
                Type = "Role"
            },
            @{
                Id   = "246dd0d5-5bd0-4def-940b-0421030a5b68"
                Type = "Role"
            },
            @{
                Id   = "01d4889c-1287-42c6-ac1f-5d1e02578ef6"
                Type = "Role"
            },
            @{
                Id   = "5b567255-7703-4780-807c-7be8301ae99b"
                Type = "Role"
            },
            @{
                Id   = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"
                Type = "Role"
            },
            @{
                Id   = "2280dda6-0bfd-44ee-a2f4-cb867cfc4c1e"
                Type = "Role"
            }
            
            @{
                Id   = "230c1aed-a721-4c5d-9cb4-a90514e508ef"
                Type = "Role"
            },
            @{
                Id   = "37730810-e9ba-4e46-b07e-8ca78d182097"
                Type = "Role"
            },
            @{
                Id   = "59a6b24b-4225-4393-8165-ebaec5f55d7a"
                Type = "Role"
            },
            @{
                Id   = "f10e1f91-74ed-437f-a6fd-d6ae88e26c1f"
                Type = "Role"
            }
            
        )
    },
    @{
        ResourceAppId  = $EXOapi.AppId
        ResourceAccess = @(
            @{
                Id   = $EXOpermission.id
                Type = "Role"
            }
        )
    }
    

    
}
Catch {

    Write-Host "Error preparing script: `n$($error[0])`nCheck Prerequisites`nExiting..." -ForegroundColor Red
    pause
    exit

}


##Check for existing app reg with the same name
$AppReg = Get-MgApplication -Filter "DisplayName eq '$($appName)'"  -ErrorAction SilentlyContinue

##If the app reg already exists, do nothing
if ($appReg) {
    write-host "App already exists - Please delete the existing 'Tenant Assessment Tool' app from Entra and rerun the preparation script to recreate, exiting" -ForegroundColor yellow
    Pause
    exit
}
else {

    Try {
        ##Create the new App Reg
        $appReg = New-MgApplication -DisplayName $appName -Web @{ RedirectUris = "http://localhost"; } -RequiredResourceAccess $params.RequiredResourceAccess -ErrorAction Stop
        Write-Host "Waiting for app to provision..."
        start-sleep -Seconds 20

        ##Enable Service Principal
        $SP = New-MgServicePrincipal -AppId $appReg.AppId
        ##Thanks to: https://adamtheautomator.com/exchange-online-v2/
        ##Add the Global Reader to the app service principal
        $directoryRole = 'Global Reader'
        ## Find the ObjectID of 'Global Reader'
        $RoleId = (Get-MgDirectoryRole | Where-Object { $_.displayname -eq $directoryRole }).Id
        ##If Role is not activated, activate it
        if (!$RoleId) {
            $RoleId = (Get-MgDirectoryRoletemplate | Where-Object { $_.displayname -eq $directoryRole }).ID
            $RoleId = (New-MgDirectoryRole -RoleTemplateId $RoleId).id
        }
        ## Add the service principal to the directory role
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $RoleId  -BodyParameter @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)"}
        $directoryRole = 'Exchange Administrator'
        ## Find the ObjectID of 'Global Reader'
        $RoleId = $null
        $RoleId = (Get-MgDirectoryRole | Where-Object { $_.displayname -eq $directoryRole }).ID
        if (!$RoleId) {
            $RoleId = (Get-MgDirectoryRoletemplate | Where-Object { $_.displayname -eq $directoryRole }).ID
            $RoleId = (New-MgDirectoryRole -RoleTemplateId $RoleId).id
        }
        ## Add the service principal to the directory role
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $RoleId  -BodyParameter @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)"}
    }
    catch {
        Write-Host "Error creating new app reg: `n$($error[0])`n Exiting..." -ForegroundColor Red
        pause
        exit
    }

}

$Thumbprint = New-AadApplicationCertificate -ClientId $appReg.Id -AddToApplication -certificatename "Tenant Assessment Certificate"

##Update Consent URL
$ConsentURl = $ConsentURl.replace('{tenant-id}', $context.TenantID)
$ConsentURl = $ConsentURl.replace('{client-id}', $appReg.AppId)

write-host "Consent page will appear, don't forget to log in as admin to grant consent!" -ForegroundColor Yellow
Start-Process $ConsentURl

Write-Host "The below details can be used to run the assessment, take note of them and press any button to clear the window.`nTenant ID: $($context.TenantID)`nClient ID: $($appReg.appID)`nCertificate Thumbprint: $thumbprint" -ForegroundColor Green
Pause
clear
