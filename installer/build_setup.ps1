param(
    [string]$InnoCompiler = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$versionFile = Join-Path $root "version.py"
$versionText = Get-Content -LiteralPath $versionFile -Raw
if ($versionText -notmatch 'APP_VERSION\s*=\s*"([^\"]+)"') {
    throw "Could not read APP_VERSION from version.py"
}
$version = $Matches[1]

if (-not $InnoCompiler) {
    $cmd = Get-Command iscc -ErrorAction SilentlyContinue
    if ($cmd) {
        $InnoCompiler = $cmd.Source
    } else {
        $candidates = @(
            "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
            "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
        )
        $InnoCompiler = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    }
}

if (-not $InnoCompiler -or -not (Test-Path $InnoCompiler)) {
    throw "Inno Setup compiler was not found. Install Inno Setup 6, then rerun this script."
}

$script = Join-Path $PSScriptRoot "ProQuoteSetup.iss"
& $InnoCompiler "/DMyAppVersion=$version" $script

$output = Join-Path $PSScriptRoot "output\ProQuoteSetup-$version.exe"
Write-Host "Built: $output"