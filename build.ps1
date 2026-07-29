# Build esa-lookup into a standalone Windows .exe using PyInstaller.
#
# Prereqs on the build machine:
#   * Python 3.10 or later installed (python.org, NOT the Windows Store shim).
#     Confirm with `py --version` or `python --version`.
#
# Usage (from this folder):
#   powershell -ExecutionPolicy Bypass -File .\build.ps1
#
# Output:
#   dist\esa-lookup.exe   <- single-file, no console window, ~80-120 MB
#
# The exe is fully self-contained: it does NOT need Python on the target
# machine. It DOES still need SAP GUI for Windows + Excel installed there.
#
# IMPORTANT: this script deliberately puts the build virtualenv,
# PyInstaller's work directory, AND a staged copy of the source .py files
# on the LOCAL disk (%LOCALAPPDATA%\esa-lookup-build). Running PyInstaller
# with the source on Z: and the workpath on C: raises
# "ValueError: path is on mount 'Z:', start on mount 'C:'" in
# os.path.relpath. Only the final .exe is copied back to the source folder.

$ErrorActionPreference = 'Stop'

Set-Location -Path (Split-Path -Parent $MyInvocation.MyCommand.Path)

# ---- 1. Resolve a real Python interpreter -------------------------------
function Resolve-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) { return 'py' }
    $c = Get-Command python -ErrorAction SilentlyContinue
    if ($c -and $c.Source -notlike '*WindowsApps*') { return $c.Source }
    return $null
}

$python = Resolve-Python
if (-not $python) {
    Write-Host "ERROR: No real Python interpreter found." -ForegroundColor Red
    Write-Host ""
    Write-Host "Install Python 3.10+ from https://www.python.org/downloads/ ,"
    Write-Host "tick 'Add python.exe to PATH' during install, then re-run this"
    Write-Host "script. The 'python.exe' that Windows offers by default is a"
    Write-Host "Store shim and cannot be used."
    exit 1
}
Write-Host "Using Python: $python" -ForegroundColor Cyan
& $python --version

# ---- 2. Anchor the build on a LOCAL disk ------------------------------
$buildRoot = Join-Path $env:LOCALAPPDATA 'esa-lookup-build'
$venv      = Join-Path $buildRoot '.venv-build'
$workpath  = Join-Path $buildRoot 'work'
$stageDir  = Join-Path $buildRoot 'src'
$localDist = Join-Path $buildRoot 'dist'
New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null

# Old on-share venv from an earlier version of this script is almost
# certainly broken from a prior SMB pip failure -- wipe it so it does not
# confuse anyone.
$legacyVenv = Join-Path (Get-Location) '.venv-build'
if (Test-Path $legacyVenv) {
    Write-Host "Removing legacy on-share venv (SMB shares cannot host pip installs):" -ForegroundColor Yellow
    Write-Host "  $legacyVenv" -ForegroundColor Yellow
    Remove-Item -Recurse -Force $legacyVenv
}

Write-Host "Local build venv:     $venv"     -ForegroundColor Cyan
Write-Host "Local PyInstaller wd: $workpath" -ForegroundColor Cyan
Write-Host "Local source stage:   $stageDir" -ForegroundColor Cyan

# ---- 3. Create / reuse the local venv ---------------------------------
$py = Join-Path $venv 'Scripts\python.exe'
if (-not (Test-Path $py)) {
    Write-Host "Creating build virtualenv ..." -ForegroundColor Cyan
    if (Test-Path $venv) { Remove-Item -Recurse -Force $venv }
    & $python -m venv $venv
    if (-not (Test-Path $py)) { throw "venv creation failed: $py not found" }
}
& $py -m pip install --upgrade pip | Out-Host

# ---- 4. Install runtime deps + PyInstaller ----------------------------
# If a previous run half-installed a package we may still see EEXIST even on
# the local disk (unlikely but possible). Try once; on failure retry once
# with --force-reinstall which overwrites any half-installed package files.
function Invoke-Pip($packageArgs, $label) {
    Write-Host "Installing $label ..." -ForegroundColor Cyan
    & $py -m pip install --no-cache-dir @packageArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  pip failed; retrying with --force-reinstall ..." -ForegroundColor Yellow
        & $py -m pip install --no-cache-dir --force-reinstall @packageArgs
        if ($LASTEXITCODE -ne 0) { throw "pip install $label failed twice" }
    }
}
Invoke-Pip @('-r', 'requirements.txt') 'runtime dependencies'
Invoke-Pip @('pyinstaller>=6.0')       'PyInstaller'

# ---- 5. Stage source + build ENTIRELY on the local disk ---------------
# PyInstaller uses os.path.relpath internally, which raises ValueError on
# Windows when the source script and the spec/dist directories live on
# different drives. Rather than fight that, copy the .py files to a local
# staging dir and run PyInstaller with CWD there. Only the final .exe is
# copied back to the source folder on Z:.
$srcDir = Get-Location

foreach ($p in @($stageDir, $workpath, $localDist)) {
    if (Test-Path $p) { Remove-Item -Recurse -Force $p }
}
New-Item -ItemType Directory -Force -Path $stageDir, $workpath, $localDist | Out-Null

# Copy the Python sources into the local staging dir. Add any new .py file
# here as the project grows.
$sources = @('esa_lookup.py', 'sap_ops.py', 'excel_ops.py', 'pipeline.py')
foreach ($src in $sources) {
    $srcPath = Join-Path $srcDir $src
    if (-not (Test-Path $srcPath)) { throw "source file missing: $srcPath" }
    Copy-Item -Path $srcPath -Destination $stageDir -Force
}

Write-Host "Staged sources at: $stageDir" -ForegroundColor Cyan
Write-Host "Running PyInstaller (all I/O on local disk) ..." -ForegroundColor Cyan

$pyinstallerArgs = @(
    '-m', 'PyInstaller',
    '--noconfirm',
    '--onefile',
    '--windowed',
    '--name', 'esa-lookup',
    '--distpath', $localDist,
    '--workpath', $workpath,
    '--specpath', $stageDir,
    '--collect-submodules', 'win32com',
    '--collect-submodules', 'pandas',
    'esa_lookup.py'
)

Push-Location $stageDir
try {
    & $py @pyinstallerArgs | Out-Host
    $pyExit = $LASTEXITCODE
} finally {
    Pop-Location
}

# ---- 6. Copy the finished .exe back to the share ---------------------
$localExe = Join-Path $localDist 'esa-lookup.exe'
if ($pyExit -ne 0 -or -not (Test-Path $localExe)) {
    Write-Host "PyInstaller failed (exit $pyExit) or $localExe is missing." -ForegroundColor Red
    Write-Host "Check the output above for errors." -ForegroundColor Red
    exit 1
}

# Wipe the on-share dist folder and copy the fresh binary in. A single-file
# copy on SMB is fine -- the atomic-rename problem was pip's, not ours.
$sharedDist = Join-Path $srcDir 'dist'
if (Test-Path $sharedDist) { Remove-Item -Recurse -Force $sharedDist }
New-Item -ItemType Directory -Path $sharedDist -Force | Out-Null
Copy-Item -Path $localExe -Destination (Join-Path $sharedDist 'esa-lookup.exe') -Force

$sharedExe = Join-Path $sharedDist 'esa-lookup.exe'
$sizeMb = [Math]::Round((Get-Item $sharedExe).Length / 1MB, 1)
Write-Host ""
Write-Host "SUCCESS: built $sharedExe ($sizeMb MB)" -ForegroundColor Green
Write-Host "Double-click it, or run from a shell -- no Python needed on the target."
