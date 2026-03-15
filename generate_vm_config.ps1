param(
    [Parameter(Mandatory = $true)]
    [string]$AgentId,
    [Parameter(Mandatory = $true)]
    [string]$BaseUrl,
    [Parameter(Mandatory = $true)]
    [string]$ExePath,
    [string]$AuthToken = "",
    [string]$TemplatePath = ".\game_tool_config.template.json",
    [string]$OutputPath = ".\game_tool_config.json"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $TemplatePath)) {
    throw "???????: $TemplatePath"
}

$content = Get-Content $TemplatePath -Raw -Encoding UTF8
$content = $content.Replace('__AGENT_ID__', $AgentId)
$content = $content.Replace('__BASE_URL__', $BaseUrl)
$content = $content.Replace('__EXE_PATH__', $ExePath.Replace('\', '/'))
$content = $content.Replace('__AUTH_TOKEN__', $AuthToken)

Set-Content -Path $OutputPath -Value $content -Encoding UTF8
Write-Host "DONE -> $OutputPath"
