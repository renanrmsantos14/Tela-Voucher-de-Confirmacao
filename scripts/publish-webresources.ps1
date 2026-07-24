param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^https://')]
  [string] $EnvironmentUrl,

  [string] $SolutionUniqueName = 'appbetinhos',

  [switch] $DeviceCode
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$base = $EnvironmentUrl.TrimEnd('/')
$api = "$base/api/data/v9.2"

function Write-Step([string] $message) { Write-Host "[voucher:webresource] $message" }

function Get-DataverseHeaders {
  if (-not (Get-Module -ListAvailable MSAL.PS)) {
    throw 'Modulo MSAL.PS nao encontrado. Instale-o com: Install-Module MSAL.PS -Scope CurrentUser'
  }

  Import-Module MSAL.PS -ErrorAction Stop
  $client = New-MsalClientApplication -ClientId '51f81489-12ee-4a9e-aaae-a2591f45987d' -TenantId 'organizations' -RedirectUri ([Uri] 'http://localhost')
  Enable-MsalTokenCacheOnDisk -PublicClientApplication $client

  try { $tokenResult = Get-MsalToken -PublicClientApplication $client -Scopes "$base/user_impersonation" -Silent }
  catch {
    if ($DeviceCode) { $tokenResult = Get-MsalToken -PublicClientApplication $client -Scopes "$base/user_impersonation" -DeviceCode }
    else { $tokenResult = Get-MsalToken -PublicClientApplication $client -Scopes "$base/user_impersonation" -Interactive }
  }

  return @{ Authorization = "Bearer $($tokenResult.AccessToken)"; Accept = 'application/json'; 'OData-MaxVersion' = '4.0'; 'OData-Version' = '4.0' }
}

function Invoke-Dataverse([string] $method, [string] $uri, $headers, $body = $null) {
  $arguments = @{ Method = $method; Uri = $uri; Headers = $headers }
  if ($null -ne $body) {
    $arguments.ContentType = 'application/json; charset=utf-8'
    $arguments.Body = $body | ConvertTo-Json -Depth 10 -Compress
  }
  return Invoke-RestMethod @arguments
}

function Get-WebResourceType([string] $filePath) {
  switch ([IO.Path]::GetExtension($filePath).ToLowerInvariant()) {
    '.html' { return 1 }
    '.css' { return 2 }
    '.js' { return 3 }
    '.png' { return 5 }
    default { throw "Tipo de web resource nao suportado: $filePath" }
  }
}

function Add-Resource($list, [string] $relativePath, [string] $displayName) {
  $filePath = Join-Path $root $relativePath
  if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { throw "Arquivo nao encontrado: $filePath" }
  $list.Add([pscustomobject]@{ FilePath = $filePath; Name = ($relativePath -replace '\\', '/'); DisplayName = $displayName; Type = Get-WebResourceType $filePath })
}

function Ensure-SolutionComponent([string] $componentId, $headers) {
  try {
    Invoke-Dataverse 'Post' "$api/AddSolutionComponent" $headers @{ ComponentId = $componentId; ComponentType = 61; SolutionUniqueName = $SolutionUniqueName; AddRequiredComponents = $false } | Out-Null
  }
  catch {
    if ($_.Exception.Message -notmatch '(?i)already|exists|ja pertence|ja existe') { throw }
  }
}

$resources = New-Object System.Collections.Generic.List[object]
Add-Resource $resources 'cr40f_VoucherConfirmacao.html' 'Voucher de Confirmacao - HTML'
Add-Resource $resources 'cr40f_VoucherConfirmacao.css' 'Voucher de Confirmacao - CSS'
Add-Resource $resources 'cr40f_VoucherConfirmacao.js' 'Voucher de Confirmacao - JavaScript'
Add-Resource $resources 'cr40f_VoucherConfirmacaoCommand.js' 'Voucher de Confirmacao - Command Bar'
Add-Resource $resources 'cr40f_VoucherConfirmacaoLogo.png' 'Betinhos - Logo branca'

$headers = Get-DataverseHeaders
$solutionName = $SolutionUniqueName.Replace("'", "''")
$solution = Invoke-Dataverse 'Get' "$api/solutions?`$select=solutionid,uniquename,friendlyname&`$filter=uniquename eq '$solutionName'" $headers
if (@($solution.value).Count -ne 1) { throw "Solucao nao encontrada ou duplicada: $SolutionUniqueName" }
Write-Step "solucao validada: $($solution.value[0].friendlyname) [$SolutionUniqueName]"

$publishedIds = New-Object System.Collections.Generic.List[string]
foreach ($resource in $resources) {
  Write-Step "enviando $($resource.Name)"
  $escapedName = $resource.Name.Replace("'", "''")
  $existing = Invoke-Dataverse 'Get' "$api/webresourceset?`$select=webresourceid&`$filter=name eq '$escapedName'" $headers
  $items = @($existing.value)
  if ($items.Count -gt 1) { throw "Web resource duplicado: $($resource.Name)" }

  $content = [Convert]::ToBase64String([IO.File]::ReadAllBytes($resource.FilePath))
  if ($items.Count -eq 0) {
    Invoke-Dataverse 'Post' "$api/webresourceset" $headers @{ name = $resource.Name; displayname = $resource.DisplayName; description = 'Voucher de Confirmacao - publicado pelo npm run push'; webresourcetype = $resource.Type; content = $content } | Out-Null
    $created = Invoke-Dataverse 'Get' "$api/webresourceset?`$select=webresourceid&`$filter=name eq '$escapedName'" $headers
    $createdItems = @($created.value)
    if ($createdItems.Count -ne 1) { throw "Nao foi possivel localizar o web resource criado: $($resource.Name)" }
    $resourceId = $createdItems[0].webresourceid
  }
  else {
    $resourceId = $items[0].webresourceid
    Invoke-Dataverse 'Patch' "$api/webresourceset($resourceId)" $headers @{ content = $content } | Out-Null
  }

  Ensure-SolutionComponent $resourceId $headers
  $publishedIds.Add($resourceId)
}

$xml = '<importexportxml><webresources>' + (($publishedIds | ForEach-Object { "<webresource>$_</webresource>" }) -join '') + '</webresources></importexportxml>'
Invoke-Dataverse 'Post' "$api/PublishXml" $headers @{ ParameterXml = $xml } | Out-Null
Write-Step "publicado na solucao ${SolutionUniqueName}: $($resources.Count) web resources"
