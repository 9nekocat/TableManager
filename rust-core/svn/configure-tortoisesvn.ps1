param(
    [Parameter(Mandatory = $true)]
    [string] $RootDirectory,
    [switch] $DryRun,
    [switch] $Quiet
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path -LiteralPath $RootDirectory).Path
$diffCmd = Join-Path $root "svn\svn-diff.cmd"
$mergeCmd = Join-Path $root "svn\svn-merge.cmd"
foreach ($required in @($diffCmd, $mergeCmd)) {
    if (!(Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Missing TableManager SVN wrapper: $required"
    }
}

$extensions = @(".xlsx", ".xlsm", ".xlsb", ".xls")
$diffValue = '"' + $diffCmd + '" %base %mine'
$mergeValue = '"' + $mergeCmd + '" %merged %theirs %mine %base'
$diffKey = "HKCU:\Software\TortoiseSVN\DiffTools"
$mergeKey = "HKCU:\Software\TortoiseSVN\MergeTools"

if ($DryRun) {
    if (!$Quiet) {
        Write-Output "TortoiseSVN DiffTools:"
        foreach ($extension in $extensions) {
            Write-Output "  $extension = $diffValue"
        }
        Write-Output "TortoiseSVN MergeTools:"
        foreach ($extension in $extensions) {
            Write-Output "  $extension = $mergeValue"
        }
        Write-Output "Dry run only. The registry was not modified."
    }
    exit 0
}

$backupRoot = Join-Path $env:LOCALAPPDATA "CatsInTheShell\TableManager\backup"
New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
foreach ($item in @(
    @{
        RegKey = "HKCU\Software\TortoiseSVN\DiffTools"
        PsKey = $diffKey
        Name = "DiffTools"
    },
    @{
        RegKey = "HKCU\Software\TortoiseSVN\MergeTools"
        PsKey = $mergeKey
        Name = "MergeTools"
    }
)) {
    if (Test-Path -LiteralPath $item.PsKey) {
        $backup = Join-Path $backupRoot ("TortoiseSVN-{0}-{1}.reg" -f $item.Name, $stamp)
        & reg.exe export $item.RegKey $backup /y | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to back up TortoiseSVN registry key: $($item.RegKey)"
        }
    }
}

New-Item -Path $diffKey -Force | Out-Null
New-Item -Path $mergeKey -Force | Out-Null
foreach ($extension in $extensions) {
    New-ItemProperty -Path $diffKey -Name $extension -Value $diffValue -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $mergeKey -Name $extension -Value $mergeValue -PropertyType String -Force | Out-Null
}

if (!$Quiet) {
    Write-Output "TortoiseSVN xlsx-family diff/merge configuration completed."
    Write-Output "Backup directory: $backupRoot"
}
