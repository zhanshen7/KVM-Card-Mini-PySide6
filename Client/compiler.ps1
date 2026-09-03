[CmdletBinding()]
param(
    [ValidateRange(1, 256)]
    [int]$Jobs = [Math]::Max(1, [Environment]::ProcessorCount),
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$clientRoot = $PSScriptRoot
$projectRoot = Split-Path -Parent $clientRoot
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"
$outputDir = Join-Path $clientRoot "build_console"
$releaseDir = Join-Path $clientRoot "Mini-KVM-Client"
$iconFile = Join-Path $clientRoot "icons\icon.ico"
$dataDir = Join-Path $clientRoot "Data"
$translationFiles = @(
    (Join-Path $clientRoot "trans_cn.qm"),
    (Join-Path $clientRoot "qtbase_cn.qm")
)

function Write-BuildOutput {
    param([AllowEmptyString()][string]$Message)

    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message
    Write-Host $line
}

function Invoke-Nuitka {
    param([string[]]$Arguments)

    Write-BuildOutput "Running: $python -m nuitka $($Arguments -join ' ')"
    & $python -m nuitka @Arguments
    return $LASTEXITCODE
}

Write-BuildOutput "Build started"
$locationPushed = $false

try {
    Write-BuildOutput "Checking Python environment"
    if (-not (Test-Path -LiteralPath $python)) {
        throw "Python environment not found: $python"
    }

    foreach ($path in @($iconFile, $dataDir) + $translationFiles) {
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Required build asset not found: $path"
        }
    }

    Write-BuildOutput "Checking Nuitka installation"
    & $python -c "import nuitka; print('Nuitka module available')"
    if ($LASTEXITCODE -ne 0) {
        throw "Nuitka is not installed in .venv. Run: $python -m pip install Nuitka"
    }

    $hookExtension = Get-ChildItem -LiteralPath (Join-Path $projectRoot ".venv\Lib\site-packages\pyWinhook") -Filter "_cpyHook*.pyd" |
        Select-Object -First 1
    if ($null -eq $hookExtension) {
        throw "The pyWinhook native extension was not found in .venv."
    }

    Push-Location $clientRoot
    $locationPushed = $true
    $nuitkaArgs = @(
        "--windows-console-mode=disable"
        "--show-progress"
        "--standalone"
        "--enable-plugin=pyside6"
        "--experimental=force-dependencies-pefile"
        "--assume-yes-for-downloads"
        "--output-dir=$outputDir"
        "--windows-icon-from-ico=$iconFile"
        "--jobs=$Jobs"
        ".\Mini-KVM.py"
        "--include-data-dir=.\icons=icons"
        "--include-data-dir=.\Data=Data"
        "--include-data-files=$($hookExtension.FullName)=$($hookExtension.Name)"
        "--include-data-files=trans_cn.qm=trans_cn.qm"
        "--include-data-files=qtbase_cn.qm=qtbase_cn.qm"
        "--include-qt-plugins=multimedia"
        "--noinclude-qt-plugins=iconengines"
        "--onefile"
        "--noinclude-qt-translations"
        "--noinclude-dlls=libQt6Charts*"
        "--noinclude-dlls=libQt6Quick3D*"
        "--noinclude-dlls=libQt6Sensors*"
        "--noinclude-dlls=libQt6Test*"
        "--noinclude-dlls=libQt6WebEngine*"
        "--noinclude-dlls=qt6web*"
        "--noinclude-dlls=qt6pdf*"
    )
    if ($Clean) {
        $nuitkaArgs += "--remove-output"
    }

    $nuitkaExitCode = Invoke-Nuitka $nuitkaArgs

    if ($nuitkaExitCode -ne 0) {
        throw "Nuitka build failed with exit code $nuitkaExitCode"
    }

    New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
    Move-Item -LiteralPath (Join-Path $outputDir "Mini-KVM.exe") -Destination (Join-Path $releaseDir "Mini-KVM.exe") -Force
    Write-BuildOutput "Build succeeded: $(Join-Path $releaseDir 'Mini-KVM.exe')"
}
catch {
    Write-BuildOutput "ERROR: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($locationPushed) {
        Pop-Location
    }
}
