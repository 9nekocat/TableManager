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

$ShippedConfigPath = Join-Path $PSScriptRoot "..\config\svn-wrapper.json"
$WrapperConfigPath = if ($env:TM_SVN_CONFIG) {
    [IO.Path]::GetFullPath($env:TM_SVN_CONFIG)
}
else {
    Join-Path $env:LOCALAPPDATA "CatsInTheShell\TableManager\config\svn-wrapper.json"
}

function Get-WrapperConfig {
    $config = [pscustomobject]@{
        diff_output_dir = "%TEMP%\TableManager\SVN"
        fallback_diff_output_dir = '${SOURCE_DIR}\.temp'
        remember_output_dir = $true
    }
    $sourceConfigPath = if (Test-Path -LiteralPath $WrapperConfigPath -PathType Leaf) {
        $WrapperConfigPath
    }
    elseif (Test-Path -LiteralPath $ShippedConfigPath -PathType Leaf) {
        $ShippedConfigPath
    }
    else {
        $null
    }
    if ($sourceConfigPath) {
        try {
            $loaded = Get-Content -LiteralPath $sourceConfigPath -Raw | ConvertFrom-Json
            foreach ($name in @("diff_output_dir", "fallback_diff_output_dir", "remember_output_dir")) {
                if ($null -ne $loaded.$name) {
                    $config.$name = $loaded.$name
                }
            }
        }
        catch {
            [Console]::Error.WriteLine(
                "Cannot read SVN wrapper config '$WrapperConfigPath': $($_.Exception.Message)"
            )
        }
    }
    return $config
}

function Save-WrapperConfig {
    param([object] $Config)
    try {
        $parent = Split-Path -Parent $WrapperConfigPath
        if ($parent) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        $json = $Config | ConvertTo-Json -Depth 4
        [IO.File]::WriteAllText(
            $WrapperConfigPath,
            $json,
            (New-Object System.Text.UTF8Encoding($false))
        )
        return $true
    }
    catch {
        [Console]::Error.WriteLine(
            "Cannot save SVN wrapper config '$WrapperConfigPath': $($_.Exception.Message)"
        )
        return $false
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

function Resolve-OutputDirectory {
    param([string] $Template, [string] $SourcePath)

    $sourceDirectory = Split-Path -Parent $SourcePath
    $resolved = $Template.Replace('${SOURCE_DIR}', $sourceDirectory)
    $resolved = [Environment]::ExpandEnvironmentVariables($resolved)
    if (![IO.Path]::IsPathRooted($resolved)) {
        $resolved = Join-Path $sourceDirectory $resolved
    }
    return [IO.Path]::GetFullPath($resolved)
}

function Get-OutputPath {
    param([string] $OutputDirectory, [string] $SourcePath)

    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $sourceName = [IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $safeName = $sourceName -replace "[^A-Za-z0-9._-]", "_"
    return Join-Path $OutputDirectory (
        "diff-{0}-{1}-{2}.xlsx" -f $safeName, (Get-Date -Format "yyyyMMdd-HHmmss"), $PID
    )
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
$config = Get-WrapperConfig
$candidates = @()
if ($env:TM_SVN_OUTPUT) {
    $candidates += [pscustomobject]@{
        Template = $null
        ExactPath = [IO.Path]::GetFullPath($env:TM_SVN_OUTPUT)
        Remember = $false
    }
}
else {
    $primaryTemplate = if ($env:TM_SVN_OUTPUT_DIR) {
        [IO.Path]::GetFullPath($env:TM_SVN_OUTPUT_DIR)
    }
    else {
        [string] $config.diff_output_dir
    }
    $fallbackTemplate = [string] $config.fallback_diff_output_dir
    $candidates += [pscustomobject]@{
        Template = $primaryTemplate
        ExactPath = $null
        Remember = (-not [bool] $env:TM_SVN_OUTPUT_DIR)
    }
    if ($fallbackTemplate -and $fallbackTemplate -ne $primaryTemplate) {
        $candidates += [pscustomobject]@{
            Template = $fallbackTemplate
            ExactPath = $null
            Remember = $true
        }
    }
}

$output = $null
$exitCode = 2
$usedCandidate = $null
foreach ($candidate in $candidates) {
    $output = if ($candidate.ExactPath) {
        $candidate.ExactPath
    }
    else {
        $outputDirectory = Resolve-OutputDirectory $candidate.Template $newPath
        Get-OutputPath $outputDirectory $newPath
    }
    $outputParent = Split-Path -Parent $output
    if ($outputParent) {
        New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
    }
    & $tableCore diff $oldPath $newPath --format excel --layout arranged --output $output
    $exitCode = $LASTEXITCODE
    if (
        (Test-Path -LiteralPath $output -PathType Leaf) -and
        (Get-Item -LiteralPath $output).Length -gt 0
    ) {
        $usedCandidate = $candidate
        break
    }
    [Console]::Error.WriteLine("Diff output failed, trying fallback directory: $output")
}

if (!$usedCandidate) {
    [Console]::Error.WriteLine("Table diff did not generate an Excel report.")
    exit 2
}

if (
    $usedCandidate.Remember -and
    $usedCandidate.Template -and
    [bool] $config.remember_output_dir -and
    $usedCandidate.Template -ne $config.diff_output_dir
) {
    $config.diff_output_dir = $usedCandidate.Template
    [void] (Save-WrapperConfig $config)
}

Open-ReviewFile $output
if ($exitCode -eq 0 -or $exitCode -eq 1) {
    exit 0
}
exit $exitCode
