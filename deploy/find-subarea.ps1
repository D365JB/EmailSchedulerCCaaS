<#
.SYNOPSIS
  Find which active app(s) have a sitemap containing the given text (e.g. a SubArea title).
#>
param(
  [Parameter(Mandatory = $true)][string]$OrgUrl,
  [Parameter(Mandatory = $true)][string]$SearchText
)
$ErrorActionPreference = 'Stop'
$OrgUrl = $OrgUrl.TrimEnd('/'); $api = "$OrgUrl/api/data/v9.2"
$token = az account get-access-token --resource $OrgUrl --query accessToken -o tsv
$H = @{ Authorization = "Bearer $token"; Accept = 'application/json' }
$apps = Invoke-RestMethod -Headers $H -Uri "$api/appmodules?`$select=appmoduleid,appmoduleidunique,name,uniquename&`$filter=statecode eq 0"
foreach ($a in $apps.value) {
  try {
    $ac = Invoke-RestMethod -Headers $H -Uri ("$api/appmodulecomponents?`$select=objectid&`$filter=" + [uri]::EscapeDataString("_appmoduleidunique_value eq $($a.appmoduleidunique) and componenttype eq 62"))
    if ($ac.value.Count -eq 0) { continue }
    $sid = $ac.value[0].objectid
    $sm = Invoke-RestMethod -Headers $H -Uri "$api/sitemaps($sid)?`$select=sitemapxml"
    if ($sm.sitemapxml -match [regex]::Escape($SearchText)) {
      Write-Host "MATCH  app=$($a.uniquename)  name=$($a.name)  sitemapid=$sid"
    }
  }
  catch {}
}
Write-Host "done"
