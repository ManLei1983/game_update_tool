param(
    [string]$PythonExe = "H:\python\3.11.9\python.exe",
    [switch]$IncludeConfig,
    [switch]$Zip
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$distRoot = Join-Path $root "dist"
$buildRoot = Join-Path $root "build"
$releaseName = "game_tool"
$releaseDir = Join-Path $distRoot $releaseName

if (Test-Path $releaseDir) { Remove-Item $releaseDir -Recurse -Force }
if (Test-Path (Join-Path $buildRoot $releaseName)) { Remove-Item (Join-Path $buildRoot $releaseName) -Recurse -Force }

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --name $releaseName `
    "$root\game_tool.py"

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --noconsole `
    --distpath $releaseDir `
    --workpath (Join-Path $buildRoot "game_tool_gui") `
    --specpath $buildRoot `
    --name "game_tool_gui" `
    "$root\game_tool_gui.py"

Copy-Item "$root\game_tool_config.example.json" "$releaseDir\game_tool_config.example.json" -Force
Copy-Item "$root\game_tool_config.template.json" "$releaseDir\game_tool_config.template.json" -Force
Copy-Item "$root\generate_vm_config.ps1" "$releaseDir\generate_vm_config.ps1" -Force
Copy-Item "$root\DEPLOY_GUIDE.md" "$releaseDir\DEPLOY_GUIDE.md" -Force
if ($IncludeConfig -and (Test-Path "$root\game_tool_config.json")) {
    Copy-Item "$root\game_tool_config.json" "$releaseDir\game_tool_config.json" -Force
}

@"
@echo off
cd /d %~dp0
if not exist game_tool_config.json (
    copy /y game_tool_config.example.json game_tool_config.json >nul
)
game_tool.exe status
pause
"@ | Set-Content -Path (Join-Path $releaseDir "run_status.bat") -Encoding ASCII

@"
@echo off
cd /d %~dp0
if not exist game_tool_config.json (
    copy /y game_tool_config.example.json game_tool_config.json >nul
)

echo ========================================
echo [1] Resume today immediately
echo     Use this when this VM should continue today's task.
echo [2] Skip today and wait next schedule
echo     Use this when today's remaining work was finished on another device.
echo ========================================
choice /c 12 /n /m "Select mode [1/2]: "
if errorlevel 2 goto skip_today
goto start_agent

:skip_today
game_tool.exe skip-today
if errorlevel 1 goto end

:start_agent
game_tool.exe agent

:end
pause
"@ | Set-Content -Path (Join-Path $releaseDir "run_agent.bat") -Encoding ASCII


@"
@echo off
cd /d %~dp0
if not exist game_tool_config.json (
    copy /y game_tool_config.example.json game_tool_config.json >nul
)
start "" game_tool_gui.exe
"@ | Set-Content -Path (Join-Path $releaseDir "run_gui.bat") -Encoding ASCII

if ($Zip) {
    $zipPath = Join-Path $distRoot ("{0}_{1}.zip" -f $releaseName, (Get-Date -Format "yyyyMMdd_HHmmss"))
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path "$releaseDir\*" -DestinationPath $zipPath
    Write-Host "ZIP -> $zipPath"
}

Write-Host "DONE -> $releaseDir"
