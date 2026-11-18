param(
    [ValidatePattern('^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$')]
    [string] $Version = "0.1.0"
)

$ErrorActionPreference = "Stop"

$coreRoot = Split-Path $PSScriptRoot -Parent
$manifest = Join-Path $coreRoot "Cargo.toml"
$distRoot = Join-Path $coreRoot "dist"
$packageName = "table-core-v$Version-windows-x64"
$stage = Join-Path $distRoot $packageName
$archive = Join-Path $distRoot "$packageName.zip"
$checksum = "$archive.sha256"

$cargo = Get-Command cargo -ErrorAction SilentlyContinue
if ($cargo) {
    $cargoPath = $cargo.Source
}
else {
    $cargoPath = Join-Path $env:USERPROFILE ".cargo\bin\cargo.exe"
    if (!(Test-Path -LiteralPath $cargoPath -PathType Leaf)) {
        throw "cargo.exe was not found in PATH or %USERPROFILE%\.cargo\bin"
    }
}

& $cargoPath build --release --locked --manifest-path $manifest
if ($LASTEXITCODE -ne 0) {
    throw "cargo build failed with exit code $LASTEXITCODE"
}

$distFull = [IO.Path]::GetFullPath($distRoot)
foreach ($path in @($stage, $archive, $checksum)) {
    $full = [IO.Path]::GetFullPath($path)
    if (!$full.StartsWith($distFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to modify a path outside the distribution directory: $full"
    }
}

New-Item -ItemType Directory -Path $distRoot -Force | Out-Null
if (Test-Path -LiteralPath $stage) {
    Remove-Item -LiteralPath $stage -Recurse -Force
}
if (Test-Path -LiteralPath $archive) {
    Remove-Item -LiteralPath $archive -Force
}
if (Test-Path -LiteralPath $checksum) {
    Remove-Item -LiteralPath $checksum -Force
}

New-Item -ItemType Directory -Path (Join-Path $stage "svn") -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $coreRoot "target\release\table-core.exe") `
    -Destination (Join-Path $stage "table-core.exe")
Copy-Item -LiteralPath (Join-Path $coreRoot "README.md") `
    -Destination (Join-Path $stage "README.md")
Copy-Item -Path (Join-Path $coreRoot "svn\*") `
    -Destination (Join-Path $stage "svn") -Recurse

& (Join-Path $stage "table-core.exe") --version
if ($LASTEXITCODE -ne 0) {
    throw "Packaged table-core.exe failed its smoke test"
}

Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $archive -CompressionLevel Optimal
$hash = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath $checksum -Value "$hash  $packageName.zip" -Encoding ascii

Write-Output $archive
Write-Output $checksum
