param(
    [string]$InnoCompiler = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$pythonVersion = "3.13.14"
$pythonInstallerName = "python-$pythonVersion-amd64.exe"
$pythonInstallerUrl = "https://www.python.org/ftp/python/$pythonVersion/$pythonInstallerName"
$pythonInstallerSha256 = "C54D9B9BBB8A36E6489363DDD01139707FD781D72F1F9E90C7EC65D0061368E0"
$payloadDir = Join-Path $PSScriptRoot "payload"
$pythonInstaller = Join-Path $payloadDir $pythonInstallerName

New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null
$downloadPython = -not (Test-Path -LiteralPath $pythonInstaller)
if (-not $downloadPython) {
    $currentHash = (Get-FileHash -LiteralPath $pythonInstaller -Algorithm SHA256).Hash
    $downloadPython = $currentHash -ne $pythonInstallerSha256
}
if ($downloadPython) {
    Write-Host "Downloading Python $pythonVersion from python.org..."
    Invoke-WebRequest -Uri $pythonInstallerUrl -OutFile $pythonInstaller
}
$verifiedHash = (Get-FileHash -LiteralPath $pythonInstaller -Algorithm SHA256).Hash
if ($verifiedHash -ne $pythonInstallerSha256) {
    throw "Python installer checksum verification failed. Expected $pythonInstallerSha256, got $verifiedHash."
}
Write-Host "Verified Python $pythonVersion installer: $verifiedHash"

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
& $InnoCompiler "/DMyAppVersion=$version" "/DPythonVersion=$pythonVersion" $script

$output = Join-Path $PSScriptRoot "output\ProQuoteSetup-$version.exe"
Write-Host "Built: $output"
