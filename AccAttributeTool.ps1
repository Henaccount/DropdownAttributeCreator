<#
ACC/BIM 360 Docs Attribute Tool
Windows onboard version: PowerShell 5.1 + built-in .NET only.

Modes:
  1) create-dropdown-definitions
     Creates dropdown custom attribute definitions from JSON.

  2) populate-file-attributes-from-folders
     Walks customer/machine folder structure and writes text custom attribute values
     to every file in matching machine folders.

Authentication:
  - Preferred: set APS_CLIENT_ID and APS_CLIENT_SECRET environment variables.
  - Or run the script and paste credentials when prompted.
  - Optional override for testing: set APS_ACCESS_TOKEN to use an existing token.
#>

param(
    [string]$InputPath,
    [switch]$DryRun,
    [string]$ClientId = $env:APS_CLIENT_ID,
    [string]$ClientSecret = $env:APS_CLIENT_SECRET,
    [string]$AccessToken = $env:APS_ACCESS_TOKEN,
    [string[]]$Scopes = @("data:read", "data:write"),
    [string]$LogPath,
    [ValidateSet("auto", "create-dropdown-definitions", "populate-file-attributes-from-folders")]
    [string]$Mode = "auto"
)

# Set-StrictMode intentionally not enabled to keep JSON property handling tolerant across API response shapes.
$ErrorActionPreference = "Stop"

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Warning "Could not force TLS 1.2: $($_.Exception.Message)"
}

function ConvertTo-PlainText {
    param([Parameter(Mandatory=$true)][Security.SecureString]$SecureString)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Select-JsonFile {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select ACC attribute JSON file"
    $dialog.Filter = "JSON files (*.json)|*.json|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $dialog.Multiselect = $false

    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "No input file selected."
    }

    return $dialog.FileName
}

function Test-JsonProperty {
    param([Parameter(Mandatory=$true)][object]$Object, [Parameter(Mandatory=$true)][string]$Name)
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function Get-JsonPropertyValue {
    param([Parameter(Mandatory=$true)][object]$Object, [Parameter(Mandatory=$true)][string]$Name, [object]$Default = $null)
    if (Test-JsonProperty -Object $Object -Name $Name) { return $Object.PSObject.Properties[$Name].Value }
    return $Default
}

function Get-Aps2LeggedToken {
    param(
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][Security.SecureString]$ClientSecret,
        [Parameter(Mandatory=$true)][string[]]$Scopes
    )

    $plainSecret = ConvertTo-PlainText -SecureString $ClientSecret
    try {
        $clientPair = "{0}:{1}" -f $ClientId, $plainSecret
        $basicValue = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($clientPair))
        $scopeValue = [Uri]::EscapeDataString(($Scopes -join " "))

        $headers = @{ Authorization = "Basic $basicValue"; Accept = "application/json" }
        $body = "grant_type=client_credentials&scope=$scopeValue"

        Write-Host "Requesting APS 2-legged access token for scopes: $($Scopes -join ', ')"

        $tokenResponse = Invoke-RestMethod `
            -Method Post `
            -Uri "https://developer.api.autodesk.com/authentication/v2/token" `
            -Headers $headers `
            -ContentType "application/x-www-form-urlencoded" `
            -Body $body

        if (-not $tokenResponse.access_token) { throw "APS token response did not contain access_token." }
        return $tokenResponse.access_token
    }
    finally { $plainSecret = $null }
}

function Get-ApsAccessToken {
    param([string]$AccessToken, [string]$ClientId, [string]$ClientSecret, [string[]]$Scopes)

    if (-not [string]::IsNullOrWhiteSpace($AccessToken)) {
        Write-Host "Using APS_ACCESS_TOKEN / provided access token."
        return $AccessToken
    }

    if ([string]::IsNullOrWhiteSpace($ClientId)) { $ClientId = Read-Host "APS Client ID" }

    if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
        $secureSecret = Read-Host "APS Client Secret" -AsSecureString
    }
    else {
        $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    }

    return Get-Aps2LeggedToken -ClientId $ClientId -ClientSecret $secureSecret -Scopes $Scopes
}

function Get-ErrorResponseText {
    param([object]$ErrorRecord)

    try {
        $response = $ErrorRecord.Exception.Response
        if ($null -eq $response) { return $ErrorRecord.Exception.Message }
        $stream = $response.GetResponseStream()
        if ($null -eq $stream) { return $ErrorRecord.Exception.Message }
        $reader = New-Object IO.StreamReader($stream)
        try { return $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
    catch { return $ErrorRecord.Exception.Message }
}

function Get-DistinctNonEmptyStrings {
    param([object[]]$Values)

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $result = New-Object 'System.Collections.Generic.List[string]'

    foreach ($value in $Values) {
        if ($null -eq $value) { continue }
        $text = ([string]$value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        if ($seen.Add($text)) { [void]$result.Add($text) }
    }

    return ,$result.ToArray()
}

function Resolve-FolderId {
    param([Parameter(Mandatory=$true)][string]$FolderId, [string]$FolderIdTemplate)

    $trimmed = $FolderId.Trim()
    if ($trimmed.StartsWith("urn:")) { return $trimmed }
    if (-not [string]::IsNullOrWhiteSpace($FolderIdTemplate)) { return ($FolderIdTemplate -f $trimmed) }
    return $trimmed
}

function Get-ProjectIdForDocs {
    param([Parameter(Mandatory=$true)][string]$ProjectId)
    $text = $ProjectId.Trim()
    if ($text.StartsWith("b.")) { return $text.Substring(2) }
    return $text
}

function Get-ProjectIdForData {
    param([Parameter(Mandatory=$true)][string]$ProjectId)
    $text = $ProjectId.Trim()
    if ($text.StartsWith("b.")) { return $text }
    return "b.$text"
}

function Invoke-ApsGet {
    param([Parameter(Mandatory=$true)][string]$Url, [Parameter(Mandatory=$true)][string]$Token)

    $headers = @{
        Authorization = "Bearer $Token"
        Accept = "application/vnd.api+json, application/json"
    }

    return Invoke-RestMethod -Method Get -Uri $Url -Headers $headers
}

function Invoke-ApsPostJson {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$Token,
        [Parameter(Mandatory=$true)][string]$Body
    )

    $headers = @{
        Authorization = "Bearer $Token"
        Accept = "application/json"
    }

    return Invoke-RestMethod -Method Post -Uri $Url -Headers $headers -ContentType "application/json" -Body $Body
}

function New-CustomAttributePayloadJson {
    param([Parameter(Mandatory=$true)][string]$Name, [Parameter(Mandatory=$true)][string[]]$Values)
    $payload = @{ name = $Name; type = "array"; arrayValues = $Values }
    return ($payload | ConvertTo-Json -Depth 20)
}

function Invoke-CreateDropdownAttribute {
    param(
        [Parameter(Mandatory=$true)][string]$ProjectId,
        [Parameter(Mandatory=$true)][string]$FolderId,
        [Parameter(Mandatory=$true)][string]$AttributeName,
        [Parameter(Mandatory=$true)][string[]]$Values,
        [Parameter(Mandatory=$true)][string]$Token,
        [switch]$DryRun
    )

    $docsProjectId = Get-ProjectIdForDocs -ProjectId $ProjectId
    $encodedProjectId = [Uri]::EscapeDataString($docsProjectId)
    $encodedFolderId = [Uri]::EscapeDataString($FolderId)
    $url = "https://developer.api.autodesk.com/bim360/docs/v1/projects/$encodedProjectId/folders/$encodedFolderId/custom-attribute-definitions"
    $body = New-CustomAttributePayloadJson -Name $AttributeName -Values $Values

    if ($DryRun) {
        Write-Host "DRY RUN: POST $url"
        Write-Host $body
        return @{ ok = $true; status = "DRY_RUN"; details = "Not sent" }
    }

    try {
        $response = Invoke-ApsPostJson -Url $url -Token $Token -Body $body
        return @{ ok = $true; status = "OK"; details = ($response | ConvertTo-Json -Depth 20 -Compress) }
    }
    catch {
        $details = Get-ErrorResponseText -ErrorRecord $_
        return @{ ok = $false; status = "FAILED"; details = $details }
    }
}

function Get-DataFolderContents {
    param(
        [Parameter(Mandatory=$true)][string]$ProjectId,
        [Parameter(Mandatory=$true)][string]$FolderId,
        [Parameter(Mandatory=$true)][string]$Token,
        [ValidateSet("folders", "items", "all")][string]$Type = "all"
    )

    $projectForData = Get-ProjectIdForData -ProjectId $ProjectId
    $encodedProjectId = [Uri]::EscapeDataString($projectForData)
    $encodedFolderId = [Uri]::EscapeDataString($FolderId)
    $url = "https://developer.api.autodesk.com/data/v1/projects/$encodedProjectId/folders/$encodedFolderId/contents"
    if ($Type -ne "all") { $url = "$url`?filter[type]=$Type" }

    $items = New-Object 'System.Collections.Generic.List[object]'
    while (-not [string]::IsNullOrWhiteSpace($url)) {
        $response = Invoke-ApsGet -Url $url -Token $Token
        if ($response.data) {
            foreach ($entry in @($response.data)) { [void]$items.Add($entry) }
        }

        $nextUrl = $null
        if ($response.links -and $response.links.next -and $response.links.next.href) {
            $nextUrl = [string]$response.links.next.href
        }
        $url = $nextUrl
    }

    return ,$items.ToArray()
}

function Get-DataEntryName {
    param([Parameter(Mandatory=$true)][object]$Entry)

    $attrs = $Entry.attributes
    if ($attrs) {
        if ($attrs.displayName) { return [string]$attrs.displayName }
        if ($attrs.name) { return [string]$attrs.name }
        if ($attrs.extension -and $attrs.extension.data -and $attrs.extension.data.sourceFileName) {
            return [string]$attrs.extension.data.sourceFileName
        }
    }
    return [string]$Entry.id
}

function Get-ItemTipVersionId {
    param([Parameter(Mandatory=$true)][object]$Item)

    if ($Item.relationships -and $Item.relationships.tip -and $Item.relationships.tip.data -and $Item.relationships.tip.data.id) {
        return [string]$Item.relationships.tip.data.id
    }
    return $null
}

function Get-CustomAttributeDefinitions {
    param(
        [Parameter(Mandatory=$true)][string]$ProjectId,
        [Parameter(Mandatory=$true)][string]$FolderId,
        [Parameter(Mandatory=$true)][string]$Token
    )

    $docsProjectId = Get-ProjectIdForDocs -ProjectId $ProjectId
    $encodedProjectId = [Uri]::EscapeDataString($docsProjectId)
    $encodedFolderId = [Uri]::EscapeDataString($FolderId)
    $url = "https://developer.api.autodesk.com/bim360/docs/v1/projects/$encodedProjectId/folders/$encodedFolderId/custom-attribute-definitions"

    $definitions = New-Object 'System.Collections.Generic.List[object]'
    $nextUrl = $url
    while (-not [string]::IsNullOrWhiteSpace($nextUrl)) {
        $response = Invoke-ApsGet -Url $nextUrl -Token $Token
        if ($response.results) {
            foreach ($definition in @($response.results)) { [void]$definitions.Add($definition) }
        }
        elseif ($response.data) {
            foreach ($definition in @($response.data)) { [void]$definitions.Add($definition) }
        }

        $nextUrl = $null
        if ($response.pagination -and $response.pagination.nextUrl) { $nextUrl = [string]$response.pagination.nextUrl }
        elseif ($response.links -and $response.links.next -and $response.links.next.href) { $nextUrl = [string]$response.links.next.href }
    }

    return ,$definitions.ToArray()
}

function Find-TextAttributeDefinition {
    param(
        [Parameter(Mandatory=$true)][object[]]$Definitions,
        [Parameter(Mandatory=$true)][string]$Name
    )

    foreach ($definition in $Definitions) {
        $defName = $null
        $defType = $null
        if ($definition.name) { $defName = [string]$definition.name }
        elseif ($definition.attributes -and $definition.attributes.name) { $defName = [string]$definition.attributes.name }
        if ($definition.type) { $defType = [string]$definition.type }
        elseif ($definition.attributes -and $definition.attributes.type) { $defType = [string]$definition.attributes.type }

        if ($defName -and ($defName -ieq $Name) -and ($defType -ieq "string")) {
            return $definition
        }
    }

    return $null
}

function Get-DefinitionId {
    param([Parameter(Mandatory=$true)][object]$Definition)
    if ($Definition.id) { return $Definition.id }
    if ($Definition.attributes -and $Definition.attributes.id) { return $Definition.attributes.id }
    return $null
}

function Invoke-UpdateFileTextAttributes {
    param(
        [Parameter(Mandatory=$true)][string]$ProjectId,
        [Parameter(Mandatory=$true)][string]$VersionId,
        [Parameter(Mandatory=$true)][object]$CustomerAttributeId,
        [Parameter(Mandatory=$true)][string]$CustomerValue,
        [Parameter(Mandatory=$true)][object]$MachineAttributeId,
        [Parameter(Mandatory=$true)][string]$MachineValue,
        [Parameter(Mandatory=$true)][string]$Token,
        [switch]$DryRun
    )

    $docsProjectId = Get-ProjectIdForDocs -ProjectId $ProjectId
    $encodedProjectId = [Uri]::EscapeDataString($docsProjectId)
    $encodedVersionId = [Uri]::EscapeDataString($VersionId)
    $url = "https://developer.api.autodesk.com/bim360/docs/v1/projects/$encodedProjectId/versions/$encodedVersionId/custom-attributes:batch-update"

    $payload = @(
        @{ id = $CustomerAttributeId; value = $CustomerValue },
        @{ id = $MachineAttributeId; value = $MachineValue }
    )
    $body = $payload | ConvertTo-Json -Depth 20

    if ($DryRun) {
        Write-Host "DRY RUN: POST $url"
        Write-Host $body
        return @{ ok = $true; status = "DRY_RUN"; details = "Not sent" }
    }

    try {
        $response = Invoke-ApsPostJson -Url $url -Token $Token -Body $body
        return @{ ok = $true; status = "OK"; details = ($response | ConvertTo-Json -Depth 20 -Compress) }
    }
    catch {
        $details = Get-ErrorResponseText -ErrorRecord $_
        return @{ ok = $false; status = "FAILED"; details = $details }
    }
}

function Invoke-DropdownDefinitionMode {
    param([Parameter(Mandatory=$true)][object]$Config, [Parameter(Mandatory=$true)][string]$Token, [switch]$DryRun)

    if (-not (Test-JsonProperty -Object $Config -Name "projectId")) { throw "Input JSON is missing projectId." }
    if (-not (Test-JsonProperty -Object $Config -Name "folders")) { throw "Input JSON is missing folders array." }

    $folderIdTemplate = [string](Get-JsonPropertyValue -Object $Config -Name "folderIdTemplate" -Default "")
    $logRows = New-Object 'System.Collections.Generic.List[object]'
    $projectId = [string]$Config.projectId

    foreach ($folder in @($Config.folders)) {
        if (-not (Test-JsonProperty -Object $folder -Name "folderId")) {
            Write-Warning "Skipping folder entry without folderId."
            continue
        }

        $folderId = Resolve-FolderId -FolderId ([string]$folder.folderId) -FolderIdTemplate $folderIdTemplate

        if (-not (Test-JsonProperty -Object $folder -Name "attributes")) {
            Write-Warning "Skipping folder '$folderId' because it has no attributes array."
            continue
        }

        foreach ($attribute in @($folder.attributes)) {
            if (-not (Test-JsonProperty -Object $attribute -Name "name")) {
                Write-Warning "Skipping attribute without name in folder '$folderId'."
                continue
            }

            $attributeName = ([string]$attribute.name).Trim()
            $values = Get-DistinctNonEmptyStrings -Values @($attribute.values)
            if ($values.Count -eq 0) {
                Write-Warning "Skipping '$attributeName' in folder '$folderId' because it has no dropdown values."
                continue
            }

            Write-Host "Creating dropdown '$attributeName' in folder '$folderId' with $($values.Count) values..."
            $result = Invoke-CreateDropdownAttribute -ProjectId $projectId -FolderId $folderId -AttributeName $attributeName -Values $values -Token $Token -DryRun:$DryRun

            [void]$logRows.Add([pscustomobject]@{
                Time = (Get-Date).ToString("s")
                Mode = "create-dropdown-definitions"
                ProjectId = $projectId
                FolderId = $folderId
                AttributeName = $attributeName
                ValueCount = $values.Count
                Status = $result.status
                Details = $result.details
            })

            if ($result.ok) { Write-Host "  OK" } else { Write-Host "  FAILED: $($result.details)" -ForegroundColor Red }
        }
    }

    return ,$logRows.ToArray()
}

function Invoke-PopulateFileAttributesMode {
    param([Parameter(Mandatory=$true)][object]$Config, [Parameter(Mandatory=$true)][string]$Token, [switch]$DryRun)

    foreach ($requiredName in @("projectId", "rootFolderId", "customerAttributeName", "machineAttributeName")) {
        if (-not (Test-JsonProperty -Object $Config -Name $requiredName)) { throw "Input JSON is missing $requiredName." }
    }

    $projectId = [string]$Config.projectId
    $folderIdTemplate = [string](Get-JsonPropertyValue -Object $Config -Name "folderIdTemplate" -Default "")
    $rootFolderId = Resolve-FolderId -FolderId ([string]$Config.rootFolderId) -FolderIdTemplate $folderIdTemplate
    $customerAttributeName = ([string]$Config.customerAttributeName).Trim()
    $machineAttributeName = ([string]$Config.machineAttributeName).Trim()
    $machineFolderNamePattern = [string](Get-JsonPropertyValue -Object $Config -Name "machineFolderNamePattern" -Default "^\d{3}\.\d{3}\.\d{3}$")

    $logRows = New-Object 'System.Collections.Generic.List[object]'
    Write-Host "Reading custom attribute definitions from root folder '$rootFolderId'..."
    $rootDefinitions = Get-CustomAttributeDefinitions -ProjectId $projectId -FolderId $rootFolderId -Token $Token
    $customerDef = Find-TextAttributeDefinition -Definitions $rootDefinitions -Name $customerAttributeName
    $machineDef = Find-TextAttributeDefinition -Definitions $rootDefinitions -Name $machineAttributeName

    if ($null -eq $customerDef -or $null -eq $machineDef) {
        throw "Could not find both text custom attributes '$customerAttributeName' and '$machineAttributeName' on the specified root folder. Make sure both are text fields and attached to that folder."
    }

    $customerAttributeId = Get-DefinitionId -Definition $customerDef
    $machineAttributeId = Get-DefinitionId -Definition $machineDef
    if ($null -eq $customerAttributeId -or $null -eq $machineAttributeId) { throw "One of the custom attribute definitions did not contain an id." }

    Write-Host "Customer attribute: '$customerAttributeName' id=$customerAttributeId"
    Write-Host "Machine attribute:  '$machineAttributeName' id=$machineAttributeId"
    Write-Host "Reading customer folders below '$rootFolderId'..."

    $customerFolders = Get-DataFolderContents -ProjectId $projectId -FolderId $rootFolderId -Token $Token -Type folders

    foreach ($customerFolder in @($customerFolders)) {
        $customerName = (Get-DataEntryName -Entry $customerFolder).Trim()
        $customerFolderId = [string]$customerFolder.id
        if ([string]::IsNullOrWhiteSpace($customerName) -or [string]::IsNullOrWhiteSpace($customerFolderId)) { continue }

        Write-Host "Customer folder: $customerName"
        $machineFolders = Get-DataFolderContents -ProjectId $projectId -FolderId $customerFolderId -Token $Token -Type folders

        foreach ($machineFolder in @($machineFolders)) {
            $machineName = (Get-DataEntryName -Entry $machineFolder).Trim()
            $machineFolderId = [string]$machineFolder.id
            if ($machineName -notmatch $machineFolderNamePattern) {
                Write-Host "  Skipping non-machine folder: $machineName"
                continue
            }

            Write-Host "  Machine folder: $machineName"
            $files = Get-DataFolderContents -ProjectId $projectId -FolderId $machineFolderId -Token $Token -Type items

            foreach ($file in @($files)) {
                $fileName = Get-DataEntryName -Entry $file
                $versionId = Get-ItemTipVersionId -Item $file

                if ([string]::IsNullOrWhiteSpace($versionId)) {
                    Write-Warning "    Skipping file '$fileName' because no tip version id was returned."
                    [void]$logRows.Add([pscustomobject]@{
                        Time = (Get-Date).ToString("s"); Mode = "populate-file-attributes-from-folders"; ProjectId = $projectId
                        Customer = $customerName; Machine = $machineName; FileName = $fileName; VersionId = ""; Status = "SKIPPED"; Details = "Missing tip version id"
                    })
                    continue
                }

                Write-Host "    Updating file: $fileName"
                $result = Invoke-UpdateFileTextAttributes -ProjectId $projectId -VersionId $versionId -CustomerAttributeId $customerAttributeId -CustomerValue $customerName -MachineAttributeId $machineAttributeId -MachineValue $machineName -Token $Token -DryRun:$DryRun

                [void]$logRows.Add([pscustomobject]@{
                    Time = (Get-Date).ToString("s")
                    Mode = "populate-file-attributes-from-folders"
                    ProjectId = $projectId
                    RootFolderId = $rootFolderId
                    CustomerFolderId = $customerFolderId
                    MachineFolderId = $machineFolderId
                    Customer = $customerName
                    Machine = $machineName
                    FileName = $fileName
                    ItemId = [string]$file.id
                    VersionId = $versionId
                    CustomerAttributeName = $customerAttributeName
                    MachineAttributeName = $machineAttributeName
                    Status = $result.status
                    Details = $result.details
                })

                if ($result.ok) { Write-Host "      OK" } else { Write-Host "      FAILED: $($result.details)" -ForegroundColor Red }
            }
        }
    }

    return ,$logRows.ToArray()
}

try {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { $InputPath = Select-JsonFile }
    if (-not (Test-Path -LiteralPath $InputPath)) { throw "Input file not found: $InputPath" }

    $config = Get-Content -LiteralPath $InputPath -Raw | ConvertFrom-Json

    if (Test-JsonProperty -Object $config -Name "scopes") {
        $Scopes = @($config.scopes | ForEach-Object { [string]$_ })
    }

    if ($Mode -eq "auto") {
        if (Test-JsonProperty -Object $config -Name "mode") { $Mode = [string]$config.mode }
        elseif (Test-JsonProperty -Object $config -Name "folders") { $Mode = "create-dropdown-definitions" }
        elseif (Test-JsonProperty -Object $config -Name "rootFolderId") { $Mode = "populate-file-attributes-from-folders" }
        else { throw "Could not infer mode. Add 'mode' to the JSON file." }
    }

    if ($Mode -ne "create-dropdown-definitions" -and $Mode -ne "populate-file-attributes-from-folders") {
        throw "Unsupported mode '$Mode'."
    }

    $token = Get-ApsAccessToken -AccessToken $AccessToken -ClientId $ClientId -ClientSecret $ClientSecret -Scopes $Scopes

    if ($Mode -eq "create-dropdown-definitions") {
        $logRows = Invoke-DropdownDefinitionMode -Config $config -Token $token -DryRun:$DryRun
    }
    else {
        $logRows = Invoke-PopulateFileAttributesMode -Config $config -Token $token -DryRun:$DryRun
    }

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $baseName = [IO.Path]::GetFileNameWithoutExtension($InputPath)
        $folder = [IO.Path]::GetDirectoryName($InputPath)
        $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $LogPath = Join-Path $folder "$baseName-acc-attribute-log-$stamp.csv"
    }

    $logRows | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "Log written to: $LogPath"
    Write-Host "Done."
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
