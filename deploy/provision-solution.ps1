<#
.SYNOPSIS
  Ensures a publisher (with the given customization prefix) and an unmanaged solution exist,
  so web resources can be named "<prefix>_/..." and packaged. Idempotent.
#>
param(
  [Parameter(Mandatory = $true)][string]$OrgUrl,
  [string]$Prefix = 'jmb',
  [string]$SolutionUniqueName = 'OutlookScheduler',
  [string]$SolutionFriendly = 'Outlook Scheduler'
)
$ErrorActionPreference = 'Stop'
$OrgUrl = $OrgUrl.TrimEnd('/')
$api = "$OrgUrl/api/data/v9.2"
$token = az account get-access-token --resource $OrgUrl --query accessToken -o tsv
if (-not $token) { throw "Could not get token for $OrgUrl" }
$H = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json; charset=utf-8'; 'OData-MaxVersion' = '4.0'; 'OData-Version' = '4.0'; Accept = 'application/json' }

$pubFilter = "customizationprefix eq '$Prefix'"
$pub = Invoke-RestMethod -Headers $H -Uri ("$api/publishers?`$select=publisherid,uniquename&`$filter=" + [uri]::EscapeDataString($pubFilter))
if ($pub.value.Count -gt 0) {
  $pubId = $pub.value[0].publisherid
  Write-Host "Publisher exists: $($pub.value[0].uniquename) (prefix $Prefix)"
}
else {
  $body = @{ uniquename = 'jmbdemopublisher'; friendlyname = 'JMB Demo Publisher'; customizationprefix = $Prefix; customizationoptionvalueprefix = (Get-Random -Minimum 20000 -Maximum 40000) } | ConvertTo-Json
  $resp = Invoke-WebRequest -Method Post -Headers $H -Uri "$api/publishers" -Body $body
  $loc = $resp.Headers['OData-EntityId']; if ($loc -is [array]) { $loc = $loc[0] }
  $pubId = [regex]::Match($loc, '\(([0-9a-fA-F-]{36})\)').Groups[1].Value
  Write-Host "Publisher created (prefix $Prefix)"
}

$solFilter = "uniquename eq '$SolutionUniqueName'"
$sol = Invoke-RestMethod -Headers $H -Uri ("$api/solutions?`$select=solutionid&`$filter=" + [uri]::EscapeDataString($solFilter))
if ($sol.value.Count -gt 0) {
  Write-Host "Solution exists: $SolutionUniqueName"
}
else {
  $body = @{ uniquename = $SolutionUniqueName; friendlyname = $SolutionFriendly; version = '1.0.0.0'; 'publisherid@odata.bind' = "/publishers($pubId)" } | ConvertTo-Json
  Invoke-RestMethod -Method Post -Headers $H -Uri "$api/solutions" -Body $body | Out-Null
  Write-Host "Solution created: $SolutionUniqueName"
}
Write-Host "DONE prefix=$Prefix solution=$SolutionUniqueName"
