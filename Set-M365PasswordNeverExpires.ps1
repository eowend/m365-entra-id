<#
.SYNOPSIS
    Sets Microsoft 365 user accounts so their passwords never expire.

.DESCRIPTION
    This script connects to Microsoft Graph PowerShell, reads all Microsoft 365 user accounts,
    and sets the PasswordPolicies property to include DisablePasswordExpiration.

    The script preserves any existing password policy values already assigned to the user.
    For example, if a user already has DisableStrongPassword, the script keeps it and adds
    DisablePasswordExpiration.

    By default, the script targets internal tenant users only, meaning UserType = Member.
    Guest accounts are excluded unless the -IncludeGuests switch is used.

    When no OutputDirectory is provided, the script opens a graphical folder picker so the
    operator can select where the CSV report will be saved.

.REQUIREMENTS
    - Windows PowerShell 5.1 or PowerShell 7+
    - Microsoft.Graph PowerShell module
    - Microsoft 365 account with permission to update users
    - Required Microsoft Graph delegated permission:
        User.ReadWrite.All

.PERMISSIONS
    The account running this script should have an administrative role that can update users,
    such as User Administrator or Global Administrator.

.EXAMPLES
    Open a GUI folder picker and test without making changes:

        .\Set-M365PasswordNeverExpires.ps1 -WhatIf

    Open a GUI folder picker and apply the change to all internal users:

        .\Set-M365PasswordNeverExpires.ps1

    Apply the change to internal users and guest users:

        .\Set-M365PasswordNeverExpires.ps1 -IncludeGuests

    Run without GUI folder picker by passing the output folder directly:

        .\Set-M365PasswordNeverExpires.ps1 -OutputDirectory "C:\Reports\M365"

.OUTPUT
    CSV report with:
    - DisplayName
    - UserPrincipalName
    - UserType
    - AccountEnabled
    - OnPremisesSyncEnabled
    - PasswordPoliciesBefore
    - PasswordPoliciesAfter
    - Status
    - ErrorMessage

.NOTES
    Author: Juan Ortiz / To Connect, LLC
    Purpose: Microsoft 365 tenant administration
    Change Type: User account password policy update
#>

#requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeGuests,

    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $false)]
    [switch]$DisconnectWhenDone
)

$ErrorActionPreference = "Stop"

function Get-ReportOutputDirectory {
    <#
    .SYNOPSIS
        Gets the folder where the report will be saved.

    .DESCRIPTION
        If OutputDirectory is provided, the script uses it.
        If OutputDirectory is not provided, the script opens a graphical folder picker.

        This allows the script to be used both interactively and in automation.
    #>

    param(
        [string]$PathFromParameter
    )

    if (-not [string]::IsNullOrWhiteSpace($PathFromParameter)) {
        if (-not (Test-Path -Path $PathFromParameter)) {
            New-Item -Path $PathFromParameter -ItemType Directory -Force | Out-Null
        }

        return $PathFromParameter
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select the folder where the Microsoft 365 password policy report will be saved."
    $folderDialog.ShowNewFolderButton = $true

    $defaultPath = [Environment]::GetFolderPath("MyDocuments")

    if (-not [string]::IsNullOrWhiteSpace($defaultPath) -and (Test-Path -Path $defaultPath)) {
        $folderDialog.SelectedPath = $defaultPath
    }

    $dialogResult = $folderDialog.ShowDialog()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "No output folder was selected. Script cancelled."
    }

    if ([string]::IsNullOrWhiteSpace($folderDialog.SelectedPath)) {
        throw "No output folder was selected. Script cancelled."
    }

    return $folderDialog.SelectedPath
}

function Ensure-MicrosoftGraphModule {
    <#
    .SYNOPSIS
        Verifies that the required Microsoft Graph commands are available.

    .DESCRIPTION
        If the Microsoft.Graph module is missing, the script installs it for the current user.
        Then it imports the required Graph modules.
    #>

    $connectCommand = Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue
    $updateCommand  = Get-Command -Name Update-MgUser -ErrorAction SilentlyContinue

    if ($null -eq $connectCommand -or $null -eq $updateCommand) {
        Write-Host "Microsoft Graph PowerShell module was not found. Installing Microsoft.Graph..." -ForegroundColor Yellow

        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
}

function Connect-ToMicrosoftGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph with the required scope.

    .DESCRIPTION
        The script checks whether an existing Microsoft Graph session already has the required
        scope. If not, it prompts for authentication.
    #>

    $requiredScopes = @(
        "User.ReadWrite.All"
    )

    $context = Get-MgContext -ErrorAction SilentlyContinue

    $needsConnection = $true

    if ($null -ne $context) {
        $currentScopes = @($context.Scopes)
        $missingScopes = $requiredScopes | Where-Object { $_ -notin $currentScopes }

        if ($missingScopes.Count -eq 0) {
            $needsConnection = $false
            Write-Host "Existing Microsoft Graph session detected." -ForegroundColor Green
        }
    }

    if ($needsConnection) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredScopes
    }
}

function Convert-PasswordPoliciesToArray {
    <#
    .SYNOPSIS
        Converts the PasswordPolicies string into an array.

    .DESCRIPTION
        Microsoft Graph stores PasswordPolicies as a comma-separated string.
        This function normalizes that value so the script can safely add a policy
        without removing existing ones.
    #>

    param(
        [AllowNull()]
        [string]$PasswordPolicies
    )

    if ([string]::IsNullOrWhiteSpace($PasswordPolicies) -or $PasswordPolicies -eq "None") {
        return @()
    }

    return @(
        $PasswordPolicies -split "," |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

function Convert-PasswordPoliciesToText {
    <#
    .SYNOPSIS
        Converts a password policy array back into Microsoft Graph text format.
    #>

    param(
        [array]$Policies
    )

    if ($null -eq $Policies -or $Policies.Count -eq 0) {
        return "None"
    }

    return ($Policies | Select-Object -Unique) -join ", "
}

try {
    $resolvedOutputDirectory = Get-ReportOutputDirectory -PathFromParameter $OutputDirectory

    Ensure-MicrosoftGraphModule
    Connect-ToMicrosoftGraph

    Write-Host "Reading Microsoft 365 users..." -ForegroundColor Cyan

    $userProperties = @(
        "id",
        "displayName",
        "userPrincipalName",
        "passwordPolicies",
        "userType",
        "accountEnabled",
        "onPremisesSyncEnabled"
    )

    $users = Get-MgUser -All -Property $userProperties

    if (-not $IncludeGuests) {
        $users = $users | Where-Object { $_.UserType -eq "Member" }
    }

    $users = @($users)

    Write-Host "Users selected for processing: $($users.Count)" -ForegroundColor Cyan

    $results = foreach ($user in $users) {
        $errorMessage = $null

        $policiesBeforeArray = Convert-PasswordPoliciesToArray -PasswordPolicies $user.PasswordPolicies
        $policiesBeforeText  = Convert-PasswordPoliciesToText -Policies $policiesBeforeArray

        $alreadyConfigured = $policiesBeforeArray -contains "DisablePasswordExpiration"

        if ($alreadyConfigured) {
            [PSCustomObject]@{
                DisplayName              = $user.DisplayName
                UserPrincipalName        = $user.UserPrincipalName
                UserType                 = $user.UserType
                AccountEnabled           = $user.AccountEnabled
                OnPremisesSyncEnabled    = $user.OnPremisesSyncEnabled
                PasswordPoliciesBefore   = $policiesBeforeText
                PasswordPoliciesAfter    = $policiesBeforeText
                Status                   = "AlreadyConfigured"
                ErrorMessage             = $null
            }

            continue
        }

        $policiesAfterArray = @(
            $policiesBeforeArray + "DisablePasswordExpiration" |
                Select-Object -Unique
        )

        $policiesAfterText = Convert-PasswordPoliciesToText -Policies $policiesAfterArray

        try {
            if ($PSCmdlet.ShouldProcess($user.UserPrincipalName, "Set PasswordPolicies to '$policiesAfterText'")) {
                Update-MgUser -UserId $user.Id -PasswordPolicies $policiesAfterText

                $status = "Updated"
            }
            else {
                $status = "WhatIf"
            }
        }
        catch {
            $status = "Failed"
            $errorMessage = $_.Exception.Message
        }

        [PSCustomObject]@{
            DisplayName              = $user.DisplayName
            UserPrincipalName        = $user.UserPrincipalName
            UserType                 = $user.UserType
            AccountEnabled           = $user.AccountEnabled
            OnPremisesSyncEnabled    = $user.OnPremisesSyncEnabled
            PasswordPoliciesBefore   = $policiesBeforeText
            PasswordPoliciesAfter    = $policiesAfterText
            Status                   = $status
            ErrorMessage             = $errorMessage
        }
    }

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $reportPath = Join-Path -Path $resolvedOutputDirectory -ChildPath "M365-PasswordNeverExpires-$timestamp.csv"

    $results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Completed." -ForegroundColor Green
    Write-Host "Report saved to: $reportPath" -ForegroundColor Green
    Write-Host ""

    $results |
        Group-Object -Property Status |
        Select-Object Name, Count |
        Format-Table -AutoSize

    if ($DisconnectWhenDone) {
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
    }
}
catch {
    Write-Error "Script failed. Error: $($_.Exception.Message)"
    throw
}