$ErrorActionPreference = "Stop"
$Headless = $env:TM_SVN_HEADLESS -eq "1"
$SvnArgs = @()
foreach ($argument in @($args)) {
    if ($argument -in @("--headless", "-Headless")) {
        $Headless = $true
    }
    else {
        $SvnArgs += $argument
    }
}

function Get-TableCorePath {
    if ($env:TABLE_CORE_EXE -and (Test-Path -LiteralPath $env:TABLE_CORE_EXE -PathType Leaf)) {
        return (Resolve-Path -LiteralPath $env:TABLE_CORE_EXE).Path
    }

    $candidates = @(
        (Join-Path $PSScriptRoot "..\target\release\table-core.exe"),
        (Join-Path $PSScriptRoot "..\target\debug\table-core.exe"),
        (Join-Path $PSScriptRoot "..\table-core.exe"),
        (Join-Path $PSScriptRoot "table-core.exe")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    throw "table-core.exe not found. Build it with: cargo build --release --manifest-path rust-core/Cargo.toml"
}

function Test-ExcelWorkbook {
    param([string] $Path)

    if (!(Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $false
    }

    $name = [IO.Path]::GetFileName($Path)
    if ($name -match "(?i)\.(xlsx|xlsm|xlsb|xls)(?:\.(mine|theirs|base|working|r?\d+))?$") {
        return $true
    }

    try {
        $stream = [IO.File]::OpenRead($Path)
        try {
            if ($stream.Length -lt 8) {
                return $false
            }
            $signature = New-Object byte[] 8
            [void] $stream.Read($signature, 0, 8)
            if ($signature[0] -eq 0x50 -and $signature[1] -eq 0x4B) {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                $archive = [IO.Compression.ZipFile]::OpenRead($Path)
                try {
                    return [bool] ($archive.Entries | Where-Object {
                        $_.FullName -like "xl/*"
                    } | Select-Object -First 1)
                }
                finally {
                    $archive.Dispose()
                }
            }
            if (
                $signature[0] -eq 0xD0 -and $signature[1] -eq 0xCF -and
                $signature[2] -eq 0x11 -and $signature[3] -eq 0xE0 -and
                $signature[4] -eq 0xA1 -and $signature[5] -eq 0xB1 -and
                $signature[6] -eq 0x1A -and $signature[7] -eq 0xE1
            ) {
                return $true
            }
            return $false
        }
        finally {
            $stream.Dispose()
        }
    }
    catch {
        return $false
    }
}

function Get-OutputPath {
    param([string[]] $InputPaths)

    if ($env:TM_SVN_OUTPUT) {
        $output = [IO.Path]::GetFullPath($env:TM_SVN_OUTPUT)
    }
    else {
        if ($env:TM_SVN_OUTPUT_DIR) {
            $outputDirectory = [IO.Path]::GetFullPath($env:TM_SVN_OUTPUT_DIR)
        }
        else {
            $workingParent = Split-Path -Parent $InputPaths[-1]
            if ($workingParent -and (Test-Path -LiteralPath $workingParent -PathType Container)) {
                $outputDirectory = Join-Path $workingParent ".temp"
            }
            else {
                $outputDirectory = Join-Path $env:TEMP "TableManager\SVN"
            }
        }
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        $sourceName = [IO.Path]::GetFileNameWithoutExtension($InputPaths[-1])
        $safeName = $sourceName -replace "[^A-Za-z0-9._-]", "_"
        $output = Join-Path $outputDirectory (
            "diff-{0}-{1}-{2}.xlsx" -f $safeName, (Get-Date -Format "yyyyMMdd-HHmmss"), $PID
        )
    }

    $parent = Split-Path -Parent $output
    if ($parent) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    return $output
}

function Open-ReviewFile {
    param([string] $Path)

    if (!$Headless -and $env:TM_SVN_NO_OPEN -ne "1") {
        Start-Process -FilePath $Path | Out-Null
    }
}

function Invoke-XlCompareDiff {
    param([string] $OldPath, [string] $NewPath)

    $xlCompare = Join-Path $env:LOCALAPPDATA "Spreadsheet Tools\xlCompare\xlCompare.exe"
    if (!(Test-Path -LiteralPath $xlCompare -PathType Leaf)) {
        throw "Neither an Excel workbook nor xlCompare was found."
    }
    & $xlCompare $OldPath $NewPath -quit_on_close -titleMine:WORKING -titleBase:BASE
    exit $LASTEXITCODE
}

if ($SvnArgs.Count -lt 2) {
    [Console]::Error.WriteLine("SVN diff requires two file paths.")
    exit 2
}

$oldPath = [IO.Path]::GetFullPath($SvnArgs[-2])
$newPath = [IO.Path]::GetFullPath($SvnArgs[-1])
if (!(Test-ExcelWorkbook $oldPath) -or !(Test-ExcelWorkbook $newPath)) {
    if ($Headless) {
        [Console]::Error.WriteLine(
            "Headless mode only handles Excel workbook families: .xlsx, .xlsm, .xlsb and .xls."
        )
        exit 2
    }
    Invoke-XlCompareDiff $oldPath $newPath
}

$tableCore = Get-TableCorePath
$output = Get-OutputPath @($oldPath, $newPath)
& $tableCore diff $oldPath $newPath --format excel --layout arranged --output $output
$exitCode = $LASTEXITCODE

if (!(Test-Path -LiteralPath $output -PathType Leaf) -or (Get-Item -LiteralPath $output).Length -eq 0) {
    [Console]::Error.WriteLine("Table diff did not generate an Excel report: $output")
    exit 2
}

Open-ReviewFile $output
if ($exitCode -eq 0 -or $exitCode -eq 1) {
    exit 0
}
exit $exitCode
