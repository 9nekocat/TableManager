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
        merge_conflict_default = "theirs"
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
            if ($loaded.merge_conflict_default) {
                $config.merge_conflict_default = ([string] $loaded.merge_conflict_default).ToLowerInvariant()
            }
        }
        catch {
            [Console]::Error.WriteLine(
                "Cannot read SVN wrapper config '$sourceConfigPath': $($_.Exception.Message)"
            )
        }
    }
    if ($config.merge_conflict_default -notin @("ours", "theirs", "empty")) {
        [Console]::Error.WriteLine(
            "Invalid merge_conflict_default '$($config.merge_conflict_default)'; using 'theirs'."
        )
        $config.merge_conflict_default = "theirs"
    }
    return $config
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

function Get-MergeReviewPath {
    param([string] $MergedPath)

    if ($env:TM_SVN_MERGE_REVIEW) {
        $path = [IO.Path]::GetFullPath($env:TM_SVN_MERGE_REVIEW)
        $parent = Split-Path -Parent $path
        if ($parent) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        return $path
    }
    $directory = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.temp"))
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $name = [IO.Path]::GetFileNameWithoutExtension($MergedPath)
    $safeName = $name -replace "[^A-Za-z0-9._-]", "_"
    return Join-Path $directory (
        "merge-{0}-{1}-{2}.xlsx" -f $safeName, (Get-Date -Format "yyyyMMdd-HHmmss"), $PID
    )
}

function Open-ReviewFileReadOnly {
    param([string] $Path)

    if ($Headless -or $env:TM_SVN_NO_OPEN -eq "1") {
        return
    }
    try {
        $item = Get-Item -LiteralPath $Path
        $item.IsReadOnly = $true
    }
    catch {
        Write-Host "[TableManager] Could not mark the review workbook read-only." -ForegroundColor Yellow
    }
    Start-Process -FilePath $Path | Out-Null
}

function Show-MergeTip {
    param(
        [string] $Message,
        [ValidateSet("Information", "Warning")]
        [string] $Kind = "Information"
    )

    if ($env:TM_SVN_TIP_LOG) {
        try {
            [IO.File]::AppendAllText(
                $env:TM_SVN_TIP_LOG,
                $Message + [Environment]::NewLine,
                (New-Object System.Text.UTF8Encoding($false))
            )
        }
        catch {
            Write-Host "[TableManager] Could not write the tip log: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        return
    }
    if (
        $Headless -or
        $env:TM_SVN_NO_OPEN -eq "1" -or
        $env:TM_SVN_NO_TIP -eq "1"
    ) {
        return
    }
    $icon = if ($Kind -eq "Warning") { 48 } else { 64 }
    try {
        $shell = New-Object -ComObject WScript.Shell
        [void] $shell.Popup($Message, 15, "TableManager", $icon)
    }
    catch {
        Write-Host $Message -ForegroundColor Yellow
    }
}

function Complete-SvnResolution {
    param(
        [string] $MergedPath,
        [string] $WorkingPath
    )

    if (
        $WorkingPath -and
        ([IO.Path]::GetFullPath($WorkingPath) -ne [IO.Path]::GetFullPath($MergedPath))
    ) {
        Copy-Item -LiteralPath $MergedPath -Destination $WorkingPath -Force
    }
    $target = if ($WorkingPath) { $WorkingPath } else { $MergedPath }
    $previousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "SilentlyContinue"
    & svn info --show-item wc-root -- $target > $null 2>&1
    $isWorkingCopy = $LASTEXITCODE -eq 0
    $ErrorActionPreference = $previousErrorAction
    if (!$isWorkingCopy) {
        Write-Host "[TableManager] Target is not in an SVN working copy; skipped svn resolve." -ForegroundColor Yellow
        return
    }
    & svn resolve --accept working -- $target
    if ($LASTEXITCODE -ne 0) {
        throw "svn resolve --accept working failed for: $target"
    }
    Write-Host "[TableManager] Marked SVN resolved: $target" -ForegroundColor Green
}

function Invoke-XlCompareMerge {
    param(
        [string] $BasePath,
        [string] $TheirPath,
        [string] $MinePath,
        [string] $MergedPath
    )

    $xlCompare = Join-Path $env:LOCALAPPDATA "Spreadsheet Tools\xlCompare\xlCompare.exe"
    if (!(Test-Path -LiteralPath $xlCompare -PathType Leaf)) {
        throw "Neither an Excel workbook nor xlCompare was found."
    }
    & $xlCompare -merge -titleMine:WORKING -titleBase:BASE -titleTheir:THEIRS `
        $MinePath $BasePath $TheirPath -output:$MergedPath
    if ($LASTEXITCODE -ne 0) {
        throw "xlCompare merge failed with exit code $LASTEXITCODE"
    }
}

trap {
    $message = "TableManager merge failed.`n$($_.Exception.Message)"
    Write-Host "[TableManager] $($_.Exception.Message)" -ForegroundColor Red
    Show-MergeTip $message -Kind Warning
    exit 2
}

if ($SvnArgs.Count -eq 4) {
    # TortoiseSVN: merged theirs mine base
    $mergedPath = [IO.Path]::GetFullPath($SvnArgs[0])
    $theirPath = [IO.Path]::GetFullPath($SvnArgs[1])
    $minePath = [IO.Path]::GetFullPath($SvnArgs[2])
    $basePath = [IO.Path]::GetFullPath($SvnArgs[3])
    $workingPath = $mergedPath
}
elseif ($SvnArgs.Count -ge 5) {
    # svn CLI: base theirs mine merged wcfile
    $basePath = [IO.Path]::GetFullPath($SvnArgs[0])
    $theirPath = [IO.Path]::GetFullPath($SvnArgs[1])
    $minePath = [IO.Path]::GetFullPath($SvnArgs[2])
    $mergedPath = [IO.Path]::GetFullPath($SvnArgs[3])
    $workingPath = [IO.Path]::GetFullPath($SvnArgs[4])
}
else {
    throw "SVN merge requires four or five file paths."
}

if (
    !(Test-ExcelWorkbook $basePath) -or
    !(Test-ExcelWorkbook $theirPath) -or
    !(Test-ExcelWorkbook $minePath)
) {
    if ($Headless) {
        throw "Headless mode only handles Excel workbook families: .xlsx, .xlsm, .xlsb and .xls."
    }
    Invoke-XlCompareMerge $basePath $theirPath $minePath $mergedPath
}

$tableCore = Get-TableCorePath
$config = Get-WrapperConfig
$conflictDefault = [string] $config.merge_conflict_default
$reviewPath = Get-MergeReviewPath $mergedPath
$mergedParent = Split-Path -Parent $mergedPath
if ($mergedParent) {
    New-Item -ItemType Directory -Path $mergedParent -Force | Out-Null
}

& $tableCore merge $basePath $minePath $theirPath `
    --strategy fail --conflict-default $conflictDefault --report text `
    --output $mergedPath --review-output $reviewPath
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    if (!(Test-Path -LiteralPath $mergedPath -PathType Leaf) -or (Get-Item -LiteralPath $mergedPath).Length -eq 0) {
        throw "Table merge did not generate an Excel result: $mergedPath"
    }
    Write-Host "[TableManager] No real conflicts found; merged automatically." -ForegroundColor Green
    Complete-SvnResolution $mergedPath $workingPath
    Write-Host "[TableManager] Merge completed." -ForegroundColor Green
    Show-MergeTip (
        "Merge completed successfully.`n" +
        "No real conflicts were found.`n" +
        "SVN has been marked resolved."
    )
    exit 0
}

if ($exitCode -ne 2) {
    throw "Table merge failed with exit code $exitCode"
}

if (!(Test-Path -LiteralPath $reviewPath -PathType Leaf) -or (Get-Item -LiteralPath $reviewPath).Length -eq 0) {
    throw "Table merge conflict preview was not generated: $reviewPath"
}

Write-Host "[TableManager] Real conflicts found; opened a three-way merge review." -ForegroundColor Yellow
Write-Host "  Default conflict side: $conflictDefault"
Write-Host "  Review file: $reviewPath"

if ($Headless -or $env:TM_SVN_NO_OPEN -eq "1") {
    exit 2
}

Open-ReviewFileReadOnly $reviewPath
Write-Host "The review workbook is read-only. Review it, then return to this window." -ForegroundColor Yellow
Write-Host "You do not need to close Excel before confirming." -ForegroundColor Yellow
$answer = Read-Host "Enter Y to apply the final values and mark resolved; anything else cancels"
if ($answer -notmatch "^\s*(y|yes)\s*$") {
    Write-Host "[TableManager] Cancelled; the file remains conflicted in SVN." -ForegroundColor Yellow
    Show-MergeTip (
        "Merge was not applied.`n" +
        "The file remains conflicted in SVN."
    ) -Kind Warning
    exit 2
}

& $tableCore resolve-review $reviewPath $basePath $minePath $theirPath `
    --strategy fail --conflict-default $conflictDefault --output $mergedPath
if ($LASTEXITCODE -ne 0) {
    throw "Applying the merge review failed with exit code $LASTEXITCODE"
}
if (!(Test-Path -LiteralPath $mergedPath -PathType Leaf) -or (Get-Item -LiteralPath $mergedPath).Length -eq 0) {
    throw "Reviewed merge did not generate an Excel result: $mergedPath"
}

Complete-SvnResolution $mergedPath $workingPath
if (Test-Path -LiteralPath $reviewPath -PathType Leaf) {
    Remove-Item -LiteralPath $reviewPath -Force -ErrorAction SilentlyContinue
}
Write-Host "[TableManager] Applied the reviewed values and completed the merge." -ForegroundColor Green
Show-MergeTip (
    "Reviewed merge completed successfully.`n" +
    "The final values were written and SVN has been marked resolved."
)
exit 0
