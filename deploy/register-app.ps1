<#
.SYNOPSIS
  Registers the Entra ID SPA app registration for the Outlook Scheduler and prints the
  clientId/tenantId to paste into src/config.js.

.DESCRIPTION
  Creates a single-tenant app registration, adds delegated Microsoft Graph permissions
  (User.Read, Calendars.ReadWrite), sets the SPA redirect URIs, and (optionally) grants
  admin consent.

  NOTE: SPA redirect URIs must be set with a Graph PATCH. `az ad app update --web-redirect-uris`
  only populates the WEB platform, not SPA.

.PARAMETER RedirectUri
  One or more SPA redirect URIs. Include the deployed blank.html web-resource URL, e.g.
  https://<org>.crm.dynamics.com/WebResources/osa_/blank.html
  Add http://localhost:5500/src/blank.html too if you want to test locally.

.EXAMPLE
  ./register-app.ps1 -RedirectUri "https://org.crm.dynamics.com/WebResources/osa_/blank.html" -AdminConsent
#>
param(
  [string]$DisplayName = 'Outlook Scheduler (Contact Center)',
  [Parameter(Mandatory = $true)][string[]]$RedirectUri,
  [switch]$AdminConsent
)

$ErrorActionPreference = 'Stop'
$graphPermsFile = Join-Path $PSScriptRoot 'graph-permissions.json'
if (-not (Test-Path $graphPermsFile)) { throw "Missing $graphPermsFile" }

Write-Host "Creating app registration '$DisplayName'..."
$app = az ad app create --display-name $DisplayName --sign-in-audience AzureADMyOrg | ConvertFrom-Json
$appId = $app.appId
$objectId = $app.id
Write-Host "  appId    = $appId"
Write-Host "  objectId = $objectId"

Write-Host "Adding delegated Microsoft Graph permissions (User.Read, Calendars.ReadWrite)..."
az ad app update --id $appId --required-resource-accesses "@$graphPermsFile" | Out-Null

Write-Host "Setting SPA redirect URIs..."
$spa = @{ spa = @{ redirectUris = @($RedirectUri) } } | ConvertTo-Json -Depth 5
$spaFile = Join-Path $env:TEMP 'osa-spa-redirect.json'
$spa | Set-Content -Path $spaFile -Encoding utf8
az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
  --headers 'Content-Type=application/json' `
  --body "@$spaFile" | Out-Null
Remove-Item $spaFile -Force -ErrorAction SilentlyContinue

Write-Host "Ensuring a service principal exists..."
az ad sp create --id $appId 2>$null | Out-Null

if ($AdminConsent) {
  Write-Host "Granting admin consent (requires privileged role)..."
  az ad app permission admin-consent --id $appId
} else {
  Write-Host "Skipping admin consent. A Global Admin can run:"
  Write-Host "  az ad app permission admin-consent --id $appId"
}

$tenantId = az account show --query tenantId -o tsv

Write-Host ""
Write-Host "==================== Paste into src/config.js ===================="
Write-Host "  clientId: '$appId'"
Write-Host "  tenantId: '$tenantId'"
Write-Host "================================================================="
