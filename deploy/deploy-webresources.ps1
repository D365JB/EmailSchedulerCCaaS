<#
.SYNOPSIS
  Uploads the files in src/ to Dataverse as web resources and publishes them.

.DESCRIPTION
  Each file becomes a web resource named "<Prefix>_/<filename>" (folder-style name so the
  relative references in index.html keep working). Existing web resources with the same name
  are updated in place. Uses the Dataverse Web API with an az CLI access token.

  The prefix MUST match a publisher prefix that exists in your environment. The default 'osa'
  assumes a publisher/solution with customization prefix 'osa'. Pass -Prefix to use your own.

.PARAMETER OrgUrl
  Dataverse org URL, e.g. https://org.crm.dynamics.com

.EXAMPLE
  az login
  ./deploy-webresources.ps1 -OrgUrl "https://org.crm.dynamics.com" -Prefix "osa" -SolutionUniqueName "OutlookScheduler"
#>
param(
  [Parameter(Mandatory = $true)][string]$OrgUrl,
  [string]$Prefix = 'osa',
  [string]$SolutionUniqueName,
  [bool]$Publish = $true
)

$ErrorActionPreference = 'Stop'
$OrgUrl = $OrgUrl.TrimEnd('/')
$api = "$OrgUrl/api/data/v9.2"

$token = az account get-access-token --resource $OrgUrl --query accessToken -o tsv
if (-not $token) { throw "Could not get an access token for $OrgUrl. Run 'az login' first." }
$H = @{
  Authorization    = "Bearer $token"
  'Content-Type'   = 'application/json; charset=utf-8'
  'OData-MaxVersion' = '4.0'
  'OData-Version'  = '4.0'
  Accept           = 'application/json'
}

$srcDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
$typeMap = @{ '.html' = 1; '.htm' = 1; '.css' = 2; '.js' = 3; '.xml' = 4; '.png' = 5; '.jpg' = 6; '.gif' = 7; '.ico' = 10; '.svg' = 11; '.json' = 3 }

# Friendly display names. In Customer Service workspace the web-resource tab caption comes from
# the web resource DisplayName, so index.html must not show as "index.html".
$displayNameMap = @{ 'index.html' = 'Outlook Scheduler' }

$script:ids = @()

Get-ChildItem $srcDir -File | ForEach-Object {
  $ext = $_.Extension.ToLower()
  if (-not $typeMap.ContainsKey($ext)) { return }

  $name = "$Prefix" + '_/' + $_.Name
  $display = if ($displayNameMap.ContainsKey($_.Name)) { $displayNameMap[$_.Name] } else { $_.Name }
  $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($_.FullName))
  $payload = @{ name = $name; displayname = $display; webresourcetype = $typeMap[$ext]; content = $b64 } | ConvertTo-Json -Compress

  $filter = "name eq '$name'"
  $getUri = "$api/webresourceset?`$select=webresourceid&`$filter=" + [uri]::EscapeDataString($filter)
  $found = Invoke-RestMethod -Method Get -Headers $H -Uri $getUri

  if ($found.value.Count -gt 0) {
    $id = $found.value[0].webresourceid
    Invoke-RestMethod -Method Patch -Headers $H -Uri "$api/webresourceset($id)" -Body $payload | Out-Null
    Write-Host "Updated  $name"
  }
  else {
    $resp = Invoke-WebRequest -Method Post -Headers $H -Uri "$api/webresourceset" -Body $payload -UseBasicParsing
    $loc = $resp.Headers['OData-EntityId']; if ($loc -is [array]) { $loc = $loc[0] }
    $id = [regex]::Match($loc, '\(([0-9a-fA-F-]{36})\)').Groups[1].Value
    Write-Host "Created  $name"
  }

  $script:ids += $id

  if ($SolutionUniqueName) {
    $addBody = @{ ComponentId = $id; ComponentType = 61; SolutionUniqueName = $SolutionUniqueName; AddRequiredComponents = $false } | ConvertTo-Json
    try { Invoke-RestMethod -Method Post -Headers $H -Uri "$api/AddSolutionComponent" -Body $addBody | Out-Null }
    catch { Write-Warning "Add to solution failed for ${name}: $_" }
  }
}

if ($Publish -and $script:ids.Count -gt 0) {
  $inner = ($script:ids | ForEach-Object { "<webresource>$_</webresource>" }) -join ''
  $xml = "<importexportxml><webresources>$inner</webresources></importexportxml>"
  $pubBody = @{ ParameterXml = $xml } | ConvertTo-Json
  Invoke-RestMethod -Method Post -Headers $H -Uri "$api/PublishXml" -Body $pubBody | Out-Null
  Write-Host "Published $($script:ids.Count) web resources."
}

Write-Host ""
Write-Host ("App URL: {0}/WebResources/{1}_/index.html" -f $OrgUrl, $Prefix)
