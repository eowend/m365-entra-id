<#
.SYNOPSIS
    Exports Microsoft 365 groups, distribution lists, and security groups with their members.

.DESCRIPTION
    Connects to a Microsoft 365 tenant through the Microsoft Graph PowerShell SDK, retrieves
    all supported groups, classifies them by group type, reads their membership, and exports
    the results to three timestamped CSV files.

    The script exports:
      - Microsoft 365 groups
      - Distribution lists
      - Security groups, including mail-enabled security groups

    The destination folder is selected through a Windows graphical folder picker.

    By default, the script exports direct membership. Use -MembershipMode Transitive to return
    a flattened list that includes members found through nested group membership.

.PARAMETER TenantId
    Optional Microsoft Entra tenant ID or verified tenant domain.

    When omitted, Microsoft Graph uses the tenant selected during interactive authentication.

.PARAMETER MembershipMode
    Controls how group membership is read.

    Direct:
        Exports the group's direct members. Nested groups appear as member objects.

    Transitive:
        Exports direct and nested members as a flattened membership list.

    The default value is Direct.

.PARAMETER GraphModuleVersion
    Exact version used for Microsoft.Graph.Authentication and Microsoft.Graph.Groups.

    The default is 2.33.0. This version is pinned because later releases can raise
    a GetTokenAsync TypeLoadException under Windows PowerShell 5.1.

.EXAMPLE
    .\Export-M365GroupMembership.ps1

    Prompts for Microsoft Graph authentication, asks the user to select an export folder,
    and exports direct membership.

.EXAMPLE
    .\Export-M365GroupMembership.ps1 -TenantId "contoso.onmicrosoft.com"

    Connects to the specified tenant and exports direct membership.

.EXAMPLE
    .\Export-M365GroupMembership.ps1 -TenantId "00000000-0000-0000-0000-000000000000" -MembershipMode Transitive

    Connects to the specified tenant and exports flattened membership, including nested members.

.INPUTS
    None.

.OUTPUTS
    Three CSV files in the folder selected by the user.

.NOTES
    File Name : Export-M365GroupMembership.ps1
    Author    : Juan Ortiz
    Version   : 1.2.0
    Created   : 2026-07-17
    Updated   : 2026-07-17

    Requirements:
      - PowerShell 7 on Windows is preferred; Windows PowerShell 5.1 is supported
      - Internet access to PowerShell Gallery and Microsoft Graph
      - Permission to install modules under the current user profile when required
      - Microsoft Graph delegated permissions Directory.Read.All and Member.Read.Hidden
      - An account with sufficient Microsoft Entra directory access

    Limitations:
      - Microsoft Graph does not return dynamic distribution groups.
      - Service-principal membership has a documented Microsoft Graph v1.0 limitation.
      - The graphical folder picker requires Windows.
      - The script relaunches itself in a clean, no-profile PowerShell process.
        This isolates Microsoft Graph from DLLs retained by the VS Code PowerShell
        Extension Host and from Exchange Online or PnP PowerShell assemblies.

.LINK
    https://learn.microsoft.com/graph/api/resources/groups-overview

.LINK
    https://learn.microsoft.com/powershell/microsoftgraph/installation

.LINK
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3479

.LINK
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3576
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Direct', 'Transitive')]
    [string]$MembershipMode = 'Direct',

    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [version]$GraphModuleVersion = [version]'2.33.0',

    [Parameter(Mandatory = $false, DontShow = $true)]
    [switch]$IsolatedProcess
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Process isolation bootstrap

# VS Code runs PowerShell scripts inside a persistent PowerShell Extension Host and
# commonly invokes them through dot-sourcing. DLLs loaded by earlier commands remain
# resident in that process and cannot be unloaded. The script therefore runs the actual
# Microsoft Graph work in a new no-profile process with a clean .NET assembly context.
if (-not $IsolatedProcess) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'The script path could not be determined. Save the script as a .ps1 file before running it.'
    }

    $powerShellEngine = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if (-not $powerShellEngine) {
        $powerShellEngine = Get-Command -Name 'powershell.exe' -ErrorAction SilentlyContinue |
            Select-Object -First 1
    }

    if (-not $powerShellEngine) {
        throw 'Neither PowerShell 7 (pwsh.exe) nor Windows PowerShell (powershell.exe) could be located.'
    }

    $childArguments = @(
        '-NoLogo'
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        $PSCommandPath
        '-IsolatedProcess'
        '-MembershipMode'
        $MembershipMode
        '-GraphModuleVersion'
        $GraphModuleVersion.ToString()
    )

    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $childArguments += @(
            '-TenantId'
            $TenantId
        )
    }

    Write-Host 'Starting Microsoft Graph export in a clean PowerShell process...' -ForegroundColor Cyan
    Write-Host "PowerShell engine : $($powerShellEngine.Source)" -ForegroundColor DarkGray
    Write-Host "Graph SDK version : $GraphModuleVersion" -ForegroundColor DarkGray

    & $powerShellEngine.Source @childArguments
    $childExitCode = $LASTEXITCODE

    if ($childExitCode -ne 0) {
        throw "The isolated PowerShell process ended with exit code $childExitCode."
    }

    # Return instead of Exit because VS Code may have dot-sourced this script.
    return
}

Write-Host "Isolated process  : PID $PID" -ForegroundColor DarkGray
Write-Host "PowerShell host   : $($Host.Name)" -ForegroundColor DarkGray
Write-Host "PowerShell version: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" -ForegroundColor DarkGray

#endregion Process isolation bootstrap

#region Helper functions

function Test-IsWindowsPlatform {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    if ($PSVersionTable.PSVersion.Major -ge 6) {
        return [bool]$IsWindows
    }

    return ($env:OS -eq 'Windows_NT')
}

function Select-ExportFolder {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    if (-not (Test-IsWindowsPlatform)) {
        throw 'The graphical folder picker is available only on Windows.'
    }

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog

    try {
        $dialog.Description = 'Select the folder where the Microsoft 365 group exports will be saved.'
        $dialog.ShowNewFolderButton = $true
        $dialog.RootFolder = [System.Environment+SpecialFolder]::Desktop

        $result = $dialog.ShowDialog()

        if ($result -ne [System.Windows.Forms.DialogResult]::OK -or
            [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
            throw [System.OperationCanceledException]::new('Folder selection was canceled by the user.')
        }

        return $dialog.SelectedPath
    }
    finally {
        $dialog.Dispose()
    }
}

function Enable-Tls12ForWindowsPowerShell {
    [CmdletBinding()]
    param()

    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor
            [Net.SecurityProtocolType]::Tls12
    }
}

function Import-RequiredGraphModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [version]$RequiredVersion
    )

    $loadedModule = Get-Module -Name $Name -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($loadedModule) {
        if ($loadedModule.Version -ne $RequiredVersion) {
            throw "Module '$Name' version $($loadedModule.Version) is already loaded. Version $RequiredVersion is required."
        }

        Write-Host "Loaded $($loadedModule.Name) version $($loadedModule.Version)." -ForegroundColor DarkGray
        return
    }

    $installedModule = Get-Module -ListAvailable -Name $Name |
        Where-Object { $_.Version -eq $RequiredVersion } |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $installedModule) {
        Write-Host "Microsoft Graph module '$Name' version $RequiredVersion was not found. Installing it for the current user..." -ForegroundColor Yellow

        if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
            throw 'Install-Module is not available. Install or update PowerShellGet and run the script again.'
        }

        $nuGetProvider = Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if (-not $nuGetProvider) {
            Write-Host 'Installing the NuGet package provider...' -ForegroundColor Yellow
            Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }

        Install-Module `
            -Name $Name `
            -RequiredVersion $RequiredVersion `
            -Scope CurrentUser `
            -Repository PSGallery `
            -Force `
            -AllowClobber `
            -ErrorAction Stop

        $installedModule = Get-Module -ListAvailable -Name $Name |
            Where-Object { $_.Version -eq $RequiredVersion } |
            Sort-Object Version -Descending |
            Select-Object -First 1
    }

    if (-not $installedModule) {
        throw "Module '$Name' version $RequiredVersion could not be located after installation."
    }

    # Import the exact manifest path. This prevents PowerShell module auto-loading from
    # selecting a newer version that may already be installed beside the pinned release.
    Import-Module `
        -Name $installedModule.Path `
        -RequiredVersion $RequiredVersion `
        -Force `
        -ErrorAction Stop

    $loadedModule = Get-Module -Name $Name -ErrorAction SilentlyContinue |
        Where-Object { $_.Version -eq $RequiredVersion } |
        Select-Object -First 1

    if (-not $loadedModule) {
        throw "Module '$Name' version $RequiredVersion could not be imported."
    }

    Write-Host "Loaded $($loadedModule.Name) version $($loadedModule.Version)." -ForegroundColor DarkGray
}

function Connect-ToMicrosoftGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$RequestedTenantId
    )

    $requiredScopes = @(
        'Directory.Read.All'
        'Member.Read.Hidden'
    )

    $connectParameters = @{
        Scopes       = $requiredScopes
        ContextScope = 'Process'
        ErrorAction  = 'Stop'
    }

    if (-not [string]::IsNullOrWhiteSpace($RequestedTenantId)) {
        $connectParameters['TenantId'] = $RequestedTenantId
    }

    Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan

    $connectCommand = Get-Command Connect-MgGraph -ErrorAction Stop
    if ($connectCommand.Parameters.ContainsKey('NoWelcome')) {
        $connectParameters['NoWelcome'] = $true
    }

    Connect-MgGraph @connectParameters | Out-Null

    $context = Get-MgContext

    if (-not $context -or [string]::IsNullOrWhiteSpace($context.TenantId)) {
        throw 'Microsoft Graph authentication completed without a valid tenant context.'
    }

    $missingScopes = @(
        $requiredScopes | Where-Object { $context.Scopes -notcontains $_ }
    )

    if ($missingScopes.Count -gt 0) {
        throw "Microsoft Graph did not grant the required delegated scopes: $($missingScopes -join ', ')."
    }

    Write-Host "Connected account : $($context.Account)" -ForegroundColor Green
    Write-Host "Connected tenant  : $($context.TenantId)" -ForegroundColor Green

    return $context
}

function Get-GraphPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$PropertyNames
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($propertyName in $PropertyNames) {
        $directProperty = $InputObject.PSObject.Properties[$propertyName]

        if ($null -ne $directProperty -and $null -ne $directProperty.Value) {
            if ($directProperty.Value -isnot [string] -or
                -not [string]::IsNullOrWhiteSpace([string]$directProperty.Value)) {
                return $directProperty.Value
            }
        }

        $additionalPropertiesProperty = $InputObject.PSObject.Properties['AdditionalProperties']

        if ($null -ne $additionalPropertiesProperty -and
            $null -ne $additionalPropertiesProperty.Value -and
            $additionalPropertiesProperty.Value -is [System.Collections.IDictionary]) {

            foreach ($key in $additionalPropertiesProperty.Value.Keys) {
                if ([string]::Equals(
                        [string]$key,
                        $propertyName,
                        [System.StringComparison]::OrdinalIgnoreCase)) {

                    $value = $additionalPropertiesProperty.Value[$key]

                    if ($null -ne $value) {
                        if ($value -isnot [string] -or
                            -not [string]::IsNullOrWhiteSpace([string]$value)) {
                            return $value
                        }
                    }
                }
            }
        }
    }

    return $null
}

function Convert-ToCsvText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [datetimeoffset]) {
        return $Value.ToString('o')
    }

    if ($Value -is [datetime]) {
        return $Value.ToString('o')
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return (@($Value) -join ';')
    }

    return [string]$Value
}

function Get-NormalizedMemberType {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Member
    )

    if ($null -eq $Member) {
        return $null
    }

    $odataType = [string](Get-GraphPropertyValue -InputObject $Member -PropertyNames @('@odata.type', 'OdataType'))

    switch -Regex ($odataType) {
        'microsoft\.graph\.user$'             { return 'User' }
        'microsoft\.graph\.group$'            { return 'Group' }
        'microsoft\.graph\.orgContact$'       { return 'OrganizationalContact' }
        'microsoft\.graph\.device$'           { return 'Device' }
        'microsoft\.graph\.servicePrincipal$' { return 'ServicePrincipal' }
        'microsoft\.graph\.application$'      { return 'Application' }
        default {
            if (-not [string]::IsNullOrWhiteSpace($odataType)) {
                return ($odataType -replace '^#microsoft\.graph\.', '')
            }

            return 'Unknown'
        }
    }
}

function Get-GroupClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Group
    )

    $groupTypes = @($Group.GroupTypes)

    if ($groupTypes -contains 'Unified') {
        return [pscustomobject]@{
            FileKey  = 'Microsoft365'
            GroupType = 'Microsoft 365 Group'
            SubType   = if ($groupTypes -contains 'DynamicMembership') {
                'Dynamic Microsoft 365 Group'
            }
            else {
                'Assigned Microsoft 365 Group'
            }
        }
    }

    if ([bool]$Group.MailEnabled -and -not [bool]$Group.SecurityEnabled) {
        return [pscustomobject]@{
            FileKey   = 'Distribution'
            GroupType = 'Distribution List'
            SubType   = 'Distribution List'
        }
    }

    if ([bool]$Group.SecurityEnabled) {
        return [pscustomobject]@{
            FileKey   = 'Security'
            GroupType = 'Security Group'
            SubType   = if ([bool]$Group.MailEnabled) {
                'Mail-Enabled Security Group'
            }
            elseif ($groupTypes -contains 'DynamicMembership') {
                'Dynamic Security Group'
            }
            else {
                'Assigned Security Group'
            }
        }
    }

    return $null
}

function Get-GroupMembersForExport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Direct', 'Transitive')]
        [string]$Mode
    )

    if ($Mode -eq 'Transitive') {
        return @(
            Get-MgGroupTransitiveMember -GroupId $GroupId -All -ErrorAction Stop
        )
    }

    return @(
        Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
    )
}

function New-GroupMemberExportRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Group,

        [Parameter(Mandatory = $true)]
        [object]$Classification,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Member,

        [Parameter(Mandatory = $true)]
        [int]$MemberCount,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Success', 'Empty', 'Failed')]
        [string]$LookupStatus,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$LookupError,

        [Parameter(Mandatory = $true)]
        [object]$GraphContext,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Direct', 'Transitive')]
        [string]$Mode
    )

    $groupTypes = @($Group.GroupTypes)
    $membershipType = if ($groupTypes -contains 'DynamicMembership') {
        'Dynamic'
    }
    else {
        'Assigned'
    }

    [pscustomobject][ordered]@{
        TenantId                   = $GraphContext.TenantId
        ConnectedAccount           = $GraphContext.Account
        MembershipMode             = $Mode
        GroupType                  = $Classification.GroupType
        GroupSubType               = $Classification.SubType
        GroupId                    = $Group.Id
        GroupDisplayName           = $Group.DisplayName
        GroupMail                  = $Group.Mail
        GroupDescription           = $Group.Description
        GroupTypes                 = Convert-ToCsvText -Value $groupTypes
        GroupMembershipType        = $membershipType
        GroupMembershipRule        = $Group.MembershipRule
        GroupMailEnabled           = Convert-ToCsvText -Value $Group.MailEnabled
        GroupSecurityEnabled       = Convert-ToCsvText -Value $Group.SecurityEnabled
        GroupRoleAssignable        = Convert-ToCsvText -Value $Group.IsAssignableToRole
        GroupOnPremisesSyncEnabled = Convert-ToCsvText -Value $Group.OnPremisesSyncEnabled
        GroupVisibility            = $Group.Visibility
        GroupCreatedDateTime       = Convert-ToCsvText -Value $Group.CreatedDateTime
        MemberCount                = $MemberCount
        MemberId                   = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('Id', 'id'))
        MemberType                 = Get-NormalizedMemberType -Member $Member
        MemberDisplayName          = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('DisplayName', 'displayName'))
        MemberUserPrincipalName    = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('UserPrincipalName', 'userPrincipalName'))
        MemberMail                 = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('Mail', 'mail'))
        MemberUserType             = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('UserType', 'userType'))
        MemberAccountEnabled       = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('AccountEnabled', 'accountEnabled'))
        MemberDeviceId             = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('DeviceId', 'deviceId'))
        MemberAppId                = Convert-ToCsvText -Value (Get-GraphPropertyValue -InputObject $Member -PropertyNames @('AppId', 'appId'))
        MemberLookupStatus         = $LookupStatus
        MemberLookupError          = $LookupError
    }
}

function Export-CsvWithHeaders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Data,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Columns,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $rows = @($Data)

    if ($rows.Count -gt 0) {
        $rows |
            Select-Object -Property $Columns |
            Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Force

        return
    }

    $emptyRow = [ordered]@{}

    foreach ($column in $Columns) {
        $emptyRow[$column] = $null
    }

    $headerLine = [pscustomobject]$emptyRow |
        ConvertTo-Csv -NoTypeInformation |
        Select-Object -First 1

    Set-Content -LiteralPath $Path -Value $headerLine -Encoding UTF8 -Force
}

#endregion Helper functions

#region Main execution

$graphConnected = $false

try {
    $exportFolder = Select-ExportFolder
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    Enable-Tls12ForWindowsPowerShell

    Write-Host "Required Microsoft Graph SDK version: $GraphModuleVersion" -ForegroundColor DarkGray
    Import-RequiredGraphModule -Name 'Microsoft.Graph.Authentication' -RequiredVersion $GraphModuleVersion
    Import-RequiredGraphModule -Name 'Microsoft.Graph.Groups' -RequiredVersion $GraphModuleVersion

    $graphContext = Connect-ToMicrosoftGraph -RequestedTenantId $TenantId
    $graphConnected = $true

    Write-Warning 'Microsoft Graph does not return dynamic distribution groups. Those groups require Exchange Online management tools.'

    $groupProperties = @(
        'id'
        'displayName'
        'description'
        'groupTypes'
        'mail'
        'mailEnabled'
        'securityEnabled'
        'membershipRule'
        'onPremisesSyncEnabled'
        'visibility'
        'createdDateTime'
        'isAssignableToRole'
    )

    Write-Host 'Reading all supported groups from Microsoft Graph...' -ForegroundColor Cyan
    $groups = @(
        Get-MgGroup -All -Property $groupProperties -ErrorAction Stop
    )

    Write-Host "Groups returned by Microsoft Graph: $($groups.Count)" -ForegroundColor Green

    $exportColumns = @(
        'TenantId'
        'ConnectedAccount'
        'MembershipMode'
        'GroupType'
        'GroupSubType'
        'GroupId'
        'GroupDisplayName'
        'GroupMail'
        'GroupDescription'
        'GroupTypes'
        'GroupMembershipType'
        'GroupMembershipRule'
        'GroupMailEnabled'
        'GroupSecurityEnabled'
        'GroupRoleAssignable'
        'GroupOnPremisesSyncEnabled'
        'GroupVisibility'
        'GroupCreatedDateTime'
        'MemberCount'
        'MemberId'
        'MemberType'
        'MemberDisplayName'
        'MemberUserPrincipalName'
        'MemberMail'
        'MemberUserType'
        'MemberAccountEnabled'
        'MemberDeviceId'
        'MemberAppId'
        'MemberLookupStatus'
        'MemberLookupError'
    )

    $exportData = @{
        Microsoft365 = New-Object 'System.Collections.Generic.List[object]'
        Distribution = New-Object 'System.Collections.Generic.List[object]'
        Security     = New-Object 'System.Collections.Generic.List[object]'
    }

    $unsupportedGroupCount = 0
    $processedGroupCount = 0

    foreach ($group in $groups) {
        $processedGroupCount++

        Write-Progress `
            -Activity 'Exporting Microsoft 365 group membership' `
            -Status "Processing group $processedGroupCount of $($groups.Count): $($group.DisplayName)" `
            -PercentComplete (($processedGroupCount / [math]::Max($groups.Count, 1)) * 100)

        $classification = Get-GroupClassification -Group $group

        if ($null -eq $classification) {
            $unsupportedGroupCount++
            Write-Warning "Skipping unsupported group object '$($group.DisplayName)' with ID '$($group.Id)'."
            continue
        }

        try {
            $members = @(Get-GroupMembersForExport -GroupId $group.Id -Mode $MembershipMode)

            if ($members.Count -eq 0) {
                $row = New-GroupMemberExportRow `
                    -Group $group `
                    -Classification $classification `
                    -Member $null `
                    -MemberCount 0 `
                    -LookupStatus 'Empty' `
                    -LookupError $null `
                    -GraphContext $graphContext `
                    -Mode $MembershipMode

                [void]$exportData[$classification.FileKey].Add($row)
                continue
            }

            foreach ($member in $members) {
                $row = New-GroupMemberExportRow `
                    -Group $group `
                    -Classification $classification `
                    -Member $member `
                    -MemberCount $members.Count `
                    -LookupStatus 'Success' `
                    -LookupError $null `
                    -GraphContext $graphContext `
                    -Mode $MembershipMode

                [void]$exportData[$classification.FileKey].Add($row)
            }
        }
        catch {
            $errorText = ($_.Exception.Message -replace '[\r\n]+', ' ').Trim()

            Write-Warning "Could not read members for '$($group.DisplayName)': $errorText"

            $row = New-GroupMemberExportRow `
                -Group $group `
                -Classification $classification `
                -Member $null `
                -MemberCount 0 `
                -LookupStatus 'Failed' `
                -LookupError $errorText `
                -GraphContext $graphContext `
                -Mode $MembershipMode

            [void]$exportData[$classification.FileKey].Add($row)
        }
    }

    Write-Progress -Activity 'Exporting Microsoft 365 group membership' -Completed

    $microsoft365Path = Join-Path $exportFolder "Microsoft365_Groups_Members_$timestamp.csv"
    $distributionPath = Join-Path $exportFolder "Distribution_Lists_Members_$timestamp.csv"
    $securityPath = Join-Path $exportFolder "Security_Groups_Members_$timestamp.csv"

    Export-CsvWithHeaders -Data $exportData.Microsoft365 -Columns $exportColumns -Path $microsoft365Path
    Export-CsvWithHeaders -Data $exportData.Distribution -Columns $exportColumns -Path $distributionPath
    Export-CsvWithHeaders -Data $exportData.Security -Columns $exportColumns -Path $securityPath

    Write-Host ''
    Write-Host 'Export completed successfully.' -ForegroundColor Green
    Write-Host "Microsoft 365 groups : $microsoft365Path"
    Write-Host "Distribution lists   : $distributionPath"
    Write-Host "Security groups      : $securityPath"

    if ($unsupportedGroupCount -gt 0) {
        Write-Warning "$unsupportedGroupCount group object(s) did not match the requested export categories."
    }
}
catch [System.OperationCanceledException] {
    Write-Host 'Operation canceled. No files were created.' -ForegroundColor Yellow
    exit 0
}
catch {
    $failureMessage = $_.Exception.Message

    if ($failureMessage -match 'GetTokenAsync' -and
        $failureMessage -match 'UserProvidedTokenCredential') {

        Write-Error (
            "The export failed because an incompatible Microsoft Graph authentication assembly was loaded. " +
            "The script attempted to run in an isolated process with Microsoft Graph SDK version $GraphModuleVersion. " +
            "Original error: $failureMessage"
        )
    }
    else {
        Write-Error "The export failed: $failureMessage"
    }

    exit 1
}
finally {
    if ($graphConnected) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

#endregion Main execution
