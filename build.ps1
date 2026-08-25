Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$versionPath = Join-Path $repoRoot 'VERSION'
$version = (Get-Content -Raw -Encoding UTF8 -LiteralPath $versionPath).Trim()
if ($version -notmatch '^1\.2\.0$') {
    throw "This build recipe is for version 1.2.0; VERSION contains '$version'."
}

$cnPackageLabel = -join ([char[]](0x5DE5, 0x4F5C, 0x8868, 0x79BB, 0x7EBF, 0x89E3, 0x9501, 0x5DE5, 0x5177))
$packageName = 'Excel' + $cnPackageLabel + '_v' + $version
$distDirectory = Join-Path $repoRoot 'dist'
$zipPath = Join-Path $distDirectory ($packageName + '.zip')
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('legacy-xls-build-' + [Guid]::NewGuid().ToString('N'))
$packageDirectory = Join-Path $tempRoot $packageName
$launcherFiles = @(Get-ChildItem -LiteralPath $repoRoot -Filter '*.vbs' -File)
if ($launcherFiles.Count -ne 1) {
    throw 'Expected exactly one VBS launcher in the repository root.'
}
$runtimeFiles = @(
    'ExcelUnlocker.ps1',
    'ExcelUnlocker-GUI.ps1',
    'ExcelUnlocker.cmd',
    $launcherFiles[0].Name,
    'README.md',
    'AUTHORS.md',
    'VERSION'
)

try {
    New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $distDirectory -Force | Out-Null
    foreach ($name in $runtimeFiles) {
        $source = Join-Path $repoRoot $name
        if (-not [IO.File]::Exists($source)) {
            throw "Missing runtime file: $name"
        }
        Copy-Item -LiteralPath $source -Destination (Join-Path $packageDirectory $name)
    }

    $checksumLines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($file in (Get-ChildItem -LiteralPath $packageDirectory -File | Sort-Object Name)) {
        $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $file.FullName).Hash
        $checksumLines.Add(('{0}  {1}' -f $hash, $file.Name))
    }
    $checksumPath = Join-Path $packageDirectory 'CHECKSUMS-SHA256.txt'
    [IO.File]::WriteAllLines($checksumPath, $checksumLines.ToArray(), (New-Object Text.UTF8Encoding($false)))

    if ([IO.File]::Exists($zipPath)) {
        [IO.File]::Delete($zipPath)
    }
    Compress-Archive -Path (Join-Path $packageDirectory '*') -DestinationPath $zipPath -CompressionLevel Optimal
    Write-Output "Created $zipPath"
    Write-Output ((Get-FileHash -Algorithm SHA256 -LiteralPath $zipPath).Hash + '  ' + (Split-Path -Leaf $zipPath))
} finally {
    if ([IO.Directory]::Exists($tempRoot)) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}
