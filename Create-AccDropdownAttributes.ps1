<#
Create ACC/BIM 360 Docs dropdown custom attribute definitions from a JSON text file.
Windows onboard version: PowerShell 5.1 + built-in .NET only.

Authentication:
  - Preferred: set APS_CLIENT_ID and APS_CLIENT_SECRET environment variables.
  - Or run the script and paste credentials when prompted.
  - Optional override for testing: set APS_ACCESS_TOKEN to use an existing token.

Input JSON structure:
{
  "projectId": "b.YOUR_PROJECT_ID_OR_GUID",
  "folderIdTemplate": "urn:adsk.wipprod:fs.folder:co.{0}",
  "scopes": ["data:read", "data:write"],
  "folders": [
    {
      "folderId": "urn:adsk.wipprod:fs.folder:co.FOLDER_ID",
      "attributes": [
        { "name": "Status", "values": ["Draft", "For Review", "Approved"] }
      ]
    }
  ]
}
#>

param(
    [string]$InputPath,
    [switch]$DryRun,
    [string]$ClientId = $env:APS_CLIENT_ID,
    [string]$ClientSecret = $env:APS_CLIENT_SECRET,
    [string]$AccessToken = $env:APS_ACCESS_TOKEN,
    [string[]]$Scopes = @("data:read", "data:write"),
    [string]$LogPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# Older Windows PowerShell builds may otherwise negotiate outdated TLS versions.
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Warning "Could not force TLS 1.2: $($_.Exception.Message)"
}

function ConvertTo-PlainText {
    param([Parameter(Mandatory=$true)][Security.SecureString]$SecureString)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
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
    $dialog.Title = "Select ACC dropdown attribute JSON file"
    $dialog.Filter = "JSON files (*.json)|*.json|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $dialog.Multiselect = $false

    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "No input file selected."
    }

    return $dialog.FileName
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

        $headers = @{
            Authorization = "Basic $basicValue"
            Accept = "application/json"
        }

        $body = "grant_type=client_credentials&scope=$scopeValue"

        Write-Host "Requesting APS 2-legged access token for scopes: $($Scopes -join ', ')"

        $tokenResponse = Invoke-RestMethod `
            -Method Post `
            -Uri "https://developer.api.autodesk.com/authentication/v2/token" `
            -Headers $headers `
            -ContentType "application/x-www-form-urlencoded" `
            -Body $body

        if (-not $tokenResponse.access_token) {
            throw "APS token response did not contain access_token."
        }

        return $tokenResponse.access_token
    }
    finally {
        # Remove plaintext secret reference as soon as possible.
        $plainSecret = $null
    }
}

function Get-ApsAccessToken {
    param(
        [string]$AccessToken,
        [string]$ClientId,
        [string]$ClientSecret,
        [string[]]$Scopes
    )

    if (-not [string]::IsNullOrWhiteSpace($AccessToken)) {
        Write-Host "Using APS_ACCESS_TOKEN / provided access token."
        return $AccessToken
    }

    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        $ClientId = Read-Host "APS Client ID"
    }

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
    catch {
        return $ErrorRecord.Exception.Message
    }
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
    param(
        [Parameter(Mandatory=$true)][string]$FolderId,
        [string]$FolderIdTemplate
    )

    $trimmed = $FolderId.Trim()
    if ($trimmed.StartsWith("urn:")) { return $trimmed }

    if (-not [string]::IsNullOrWhiteSpace($FolderIdTemplate)) {
        return ($FolderIdTemplate -f $trimmed)
    }

    return $trimmed
}

function New-CustomAttributePayloadJson {
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string[]]$Values
    )

    $payload = @{
        name = $Name
        type = "array"
        arrayValues = $Values
    }

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

    $encodedProjectId = [Uri]::EscapeDataString($ProjectId)
    $encodedFolderId = [Uri]::EscapeDataString($FolderId)
    $url = "https://developer.api.autodesk.com/bim360/docs/v1/projects/$encodedProjectId/folders/$encodedFolderId/custom-attribute-definitions"
    $body = New-CustomAttributePayloadJson -Name $AttributeName -Values $Values

    if ($DryRun) {
        Write-Host "DRY RUN: POST $url"
        Write-Host $body
        return @{ ok = $true; status = "DRY_RUN"; details = "Not sent" }
    }

    $headers = @{
        Authorization = "Bearer $Token"
        Accept = "application/json"
    }

    try {
        $response = Invoke-RestMethod `
            -Method Post `
            -Uri $url `
            -Headers $headers `
            -ContentType "application/json" `
            -Body $body

        return @{ ok = $true; status = "OK"; details = ($response | ConvertTo-Json -Depth 20 -Compress) }
    }
    catch {
        $details = Get-ErrorResponseText -ErrorRecord $_
        return @{ ok = $false; status = "FAILED"; details = $details }
    }
}

try {
    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        $InputPath = Select-JsonFile
    }

    if (-not (Test-Path -LiteralPath $InputPath)) {
        throw "Input file not found: $InputPath"
    }

    $config = Get-Content -LiteralPath $InputPath -Raw | ConvertFrom-Json

    if (-not $config.projectId) { throw "Input JSON is missing projectId." }
    if (-not $config.folders) { throw "Input JSON is missing folders array." }

    if ($config.scopes) {
        $Scopes = @($config.scopes | ForEach-Object { [string]$_ })
    }

    $folderIdTemplate = $null
    if ($config.folderIdTemplate) { $folderIdTemplate = [string]$config.folderIdTemplate }

    $token = Get-ApsAccessToken -AccessToken $AccessToken -ClientId $ClientId -ClientSecret $ClientSecret -Scopes $Scopes

    $logRows = New-Object 'System.Collections.Generic.List[object]'
    $projectId = [string]$config.projectId

    foreach ($folder in $config.folders) {
        if (-not $folder.folderId) {
            Write-Warning "Skipping folder entry without folderId."
            continue
        }

        $folderId = Resolve-FolderId -FolderId ([string]$folder.folderId) -FolderIdTemplate $folderIdTemplate

        if (-not $folder.attributes) {
            Write-Warning "Skipping folder '$folderId' because it has no attributes array."
            continue
        }

        foreach ($attribute in $folder.attributes) {
            if (-not $attribute.name) {
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

            $result = Invoke-CreateDropdownAttribute `
                -ProjectId $projectId `
                -FolderId $folderId `
                -AttributeName $attributeName `
                -Values $values `
                -Token $token `
                -DryRun:$DryRun

            $row = [pscustomobject]@{
                Time = (Get-Date).ToString("s")
                ProjectId = $projectId
                FolderId = $folderId
                AttributeName = $attributeName
                ValueCount = $values.Count
                Status = $result.status
                Details = $result.details
            }
            [void]$logRows.Add($row)

            if ($result.ok) {
                Write-Host "  OK"
            }
            else {
                Write-Host "  FAILED: $($result.details)" -ForegroundColor Red
            }
        }
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
