[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9]+$')]
    [string]$BitId,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Environment = 'Production',

    [Parameter(Mandatory)]
    [ValidatePattern('^https://')]
    [string]$RunUri,

    [Parameter()]
    [string]$PayloadPath,

    [Parameter()]
    [string]$PayloadJson,

    [Parameter()]
    [int]$TimeoutSec = 60
)

$ErrorActionPreference = 'Stop'
$resourceUrl = 'api://518a59cf-dba0-4cb9-8392-8a328a4831bf'

if ($PayloadPath -and $PayloadJson) {
    throw 'Use either -PayloadPath or -PayloadJson, not both.'
}

if ($PayloadPath) {
    if (-not (Test-Path -LiteralPath $PayloadPath -PathType Leaf)) {
        throw "Payload file '$PayloadPath' was not found."
    }
    $payloadJson = Get-Content -LiteralPath $PayloadPath -Raw
}
elseif ($PayloadJson) {
    $payloadJson = $PayloadJson
}
else {
    $payloadJson = [ordered]@{
        bitId       = $BitId
        environment = $Environment
    } | ConvertTo-Json -Compress
}

try {
    $null = $payloadJson | ConvertFrom-Json
}
catch {
    throw "The request payload is not valid JSON: $($_.Exception.Message)"
}

if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    throw 'Az.Accounts is not installed. Install it with: Install-Module Az.Accounts -Scope CurrentUser'
}

Import-Module Az.Accounts

function ConvertTo-PlainTextToken {
    param([Parameter(Mandatory)]$Token)

    if ($Token -isnot [securestring]) {
        return [string]$Token
    }

    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
}

try {
    $accessToken = Get-AzAccessToken -ResourceUrl $resourceUrl
    $plainToken = ConvertTo-PlainTextToken -Token $accessToken.Token
    $headers = @{
        Authorization = "Bearer $plainToken"
        Accept        = 'application/json'
        'Content-Type' = 'application/json'
    }

    Write-Host "Triggering ORR evaluation for $BitId in $Environment..." -ForegroundColor Cyan
    if ($PSCmdlet.ShouldProcess($RunUri, "POST ORR evaluation for $BitId")) {
        $response = Invoke-RestMethod -Uri $RunUri -Method Post -Headers $headers -Body $payloadJson -TimeoutSec $TimeoutSec
        Write-Host 'Evaluation request accepted.' -ForegroundColor Green
        if ($null -ne $response) {
            $response | ConvertTo-Json -Depth 20
        }
    }
}
catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    if ($statusCode) {
        throw "ORR evaluation request failed with HTTP $statusCode. $($_.Exception.Message)"
    }
    throw "Unable to trigger ORR evaluation. $($_.Exception.Message)"
}
