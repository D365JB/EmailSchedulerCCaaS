<#
.SYNOPSIS
  Inspect ('get') or add a web-resource tab ('addtab') to a model-driven app's sitemap.
  'addtab' backs up the current sitemap XML before changing it.
#>
param(
  [Parameter(Mandatory = $true)][string]$OrgUrl,
  [Parameter(Mandatory = $true)][string]$AppUniqueName,
  [ValidateSet('get', 'addtab')][string]$Action = 'get',
  [string]$WebResourceName = 'jmb_/index.html',
  [string]$Url = '/WebResources/jmb_/index.html',
  [string]$Title = 'Outlook Scheduler',
  [string]$SubAreaId = 'jmb_outlookscheduler',
  [string]$GroupTitle = 'Productivity'
)
$ErrorActionPreference = 'Stop'
$OrgUrl = $OrgUrl.TrimEnd('/'); $api = "$OrgUrl/api/data/v9.2"
$token = az account get-access-token --resource $OrgUrl --query accessToken -o tsv
if (-not $token) { throw "no token" }
$H = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json; charset=utf-8'; 'OData-MaxVersion' = '4.0'; 'OData-Version' = '4.0'; Accept = 'application/json' }

$appFilter = "uniquename eq '$AppUniqueName'"
$app = Invoke-RestMethod -Headers $H -Uri ("$api/appmodules?`$select=appmoduleid,appmoduleidunique,name&`$filter=" + [uri]::EscapeDataString($appFilter))
if ($app.value.Count -eq 0) { throw "App $AppUniqueName not found" }
$appId = $app.value[0].appmoduleid
$appIdUnique = $app.value[0].appmoduleidunique
Write-Host "App: $($app.value[0].name) ($appId)"

$acFilter = "_appmoduleidunique_value eq $appIdUnique and componenttype eq 62"
$ac = Invoke-RestMethod -Headers $H -Uri ("$api/appmodulecomponents?`$select=objectid&`$filter=" + [uri]::EscapeDataString($acFilter))
if ($ac.value.Count -eq 0) { throw "No sitemap component (type 62) for app" }
$sitemapId = $ac.value[0].objectid
$sm = Invoke-RestMethod -Headers $H -Uri "$api/sitemaps($sitemapId)?`$select=sitemapxml,sitemapnameunique"
$xmlText = $sm.sitemapxml
Write-Host "SiteMap: $($sm.sitemapnameunique) ($sitemapId), xml length $($xmlText.Length)"

if ($Action -eq 'get') {
  $xmlText | Set-Content -Path (Join-Path $PSScriptRoot 'sitemap-current.xml') -Encoding utf8
  [xml]$d = $xmlText
  foreach ($a in $d.SiteMap.Area) {
    Write-Host "Area Id=$($a.Id)"
    foreach ($g in $a.Group) { Write-Host "  Group Id=$($g.Id) SubAreas=$(@($g.SubArea).Count)" }
  }
  Write-Host "Web-resource SubArea examples:"
  $d.SelectNodes('//SubArea') | Where-Object { $_.Url -match 'WebResources|webresource' } | Select-Object -First 3 | ForEach-Object { Write-Host "  $($_.OuterXml)" }
  return
}

# addtab
$xmlText | Set-Content -Path (Join-Path $PSScriptRoot 'sitemap-backup.xml') -Encoding utf8
[xml]$d = $xmlText
if ($d.SelectSingleNode("//SubArea[@Id='$SubAreaId']")) {
  Write-Host "SubArea $SubAreaId already present; skipping insert."
}
else {
  $area = @($d.SiteMap.Area)[0]
  if (-not $area) { throw "No Area in sitemap" }
  $grp = $d.CreateElement('Group'); $grp.SetAttribute('Id', 'jmb_group')
  $gt = $d.CreateElement('Titles'); $gtt = $d.CreateElement('Title'); $gtt.SetAttribute('LCID', '1033'); $gtt.SetAttribute('Title', $GroupTitle); $gt.AppendChild($gtt) | Out-Null; $grp.AppendChild($gt) | Out-Null
  $sa = $d.CreateElement('SubArea'); $sa.SetAttribute('Id', $SubAreaId); $sa.SetAttribute('Icon', '/_imgs/imagestrips/transparent_spacer.gif'); $sa.SetAttribute('Url', $Url); $sa.SetAttribute('Client', 'All,Outlook,OutlookLaptopClient,OutlookWorkstationClient,Web'); $sa.SetAttribute('AvailableOffline', 'false'); $sa.SetAttribute('PassParams', 'false')
  $stt2 = $d.CreateElement('Titles'); $t2 = $d.CreateElement('Title'); $t2.SetAttribute('LCID', '1033'); $t2.SetAttribute('Title', $Title); $stt2.AppendChild($t2) | Out-Null; $sa.AppendChild($stt2) | Out-Null
  $grp.AppendChild($sa) | Out-Null
  $area.AppendChild($grp) | Out-Null
  Write-Host "Inserted SubArea $SubAreaId into Area $($area.Id)"
}
$newXml = $d.OuterXml

$patch = @{ sitemapxml = $newXml } | ConvertTo-Json
Invoke-RestMethod -Method Patch -Headers $H -Uri "$api/sitemaps($sitemapId)" -Body ([System.Text.Encoding]::UTF8.GetBytes($patch)) | Out-Null
Write-Host "Sitemap updated"

$xml = "<importexportxml><sitemaps><sitemap>$sitemapId</sitemap></sitemaps><appmodules><appmodule>$appId</appmodule></appmodules></importexportxml>"
$pub = @{ ParameterXml = $xml } | ConvertTo-Json
Invoke-RestMethod -Method Post -Headers $H -Uri "$api/PublishXml" -Body ([System.Text.Encoding]::UTF8.GetBytes($pub)) | Out-Null
Write-Host "Published sitemap + app. DONE"
