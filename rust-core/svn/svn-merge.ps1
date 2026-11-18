$ErrorActionPreference = "Stop"
$SvnArgs = @($args)

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
            if ($stream.Length -lt 4) {
                return $false
            }
            $signature = New-Object byte[] 4
            [void] $stream.Read($signature, 0, 4)
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

function Open-ReviewFile {
    param([string] $Path)

    if ($env:TM_SVN_NO_OPEN -ne "1") {
        Start-Process -FilePath $Path | Out-Null
    }
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
    exit $LASTEXITCODE
}

if ($SvnArgs.Count -eq 4) {
    # TortoiseSVN: merged theirs mine base
    $mergedPath = [IO.Path]::GetFullPath($SvnArgs[0])
    $theirPath = [IO.Path]::GetFullPath($SvnArgs[1])
    $minePath = [IO.Path]::GetFullPath($SvnArgs[2])
    $basePath = [IO.Path]::GetFullPath($SvnArgs[3])
}
elseif ($SvnArgs.Count -ge 5) {
    # svn CLI: base theirs mine merged wcfile
    $basePath = [IO.Path]::GetFullPath($SvnArgs[0])
    $theirPath = [IO.Path]::GetFullPath($SvnArgs[1])
    $minePath = [IO.Path]::GetFullPath($SvnArgs[2])
    $mergedPath = [IO.Path]::GetFullPath($SvnArgs[3])
}
else {
    [Console]::Error.WriteLine("SVN merge requires four or five file paths.")
    exit 2
}

if (
    !(Test-ExcelWorkbook $basePath) -or
    !(Test-ExcelWorkbook $theirPath) -or
    !(Test-ExcelWorkbook $minePath)
) {
    Invoke-XlCompareMerge $basePath $theirPath $minePath $mergedPath
}

$tableCore = Get-TableCorePath
$mergedParent = Split-Path -Parent $mergedPath
if ($mergedParent) {
    New-Item -ItemType Directory -Path $mergedParent -Force | Out-Null
}

& $tableCore merge $basePath $minePath $theirPath `
    --strategy fail --allow-conflicts --report text --output $mergedPath
$exitCode = $LASTEXITCODE

if (!(Test-Path -LiteralPath $mergedPath -PathType Leaf) -or (Get-Item -LiteralPath $mergedPath).Length -eq 0) {
    [Console]::Error.WriteLine("Table merge did not generate an Excel result: $mergedPath")
    exit 2
}

Open-ReviewFile $mergedPath

# Exit code 2 means the review workbook was written but unresolved conflicts remain.
if ($exitCode -eq 0 -or $exitCode -eq 2) {
    exit $exitCode
}
exit $exitCode
