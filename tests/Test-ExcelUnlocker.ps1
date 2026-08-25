Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$backend = Join-Path $repoRoot 'ExcelUnlocker.ps1'
$gui = Join-Path $repoRoot 'ExcelUnlocker-GUI.ps1'

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Set-U16([byte[]]$Data, [int]$Offset, [UInt16]$Value) {
    [Buffer]::BlockCopy([BitConverter]::GetBytes($Value), 0, $Data, $Offset, 2)
}

function Set-I32([byte[]]$Data, [int]$Offset, [Int32]$Value) {
    [Buffer]::BlockCopy([BitConverter]::GetBytes($Value), 0, $Data, $Offset, 4)
}

function Set-I64([byte[]]$Data, [int]$Offset, [Int64]$Value) {
    [Buffer]::BlockCopy([BitConverter]::GetBytes($Value), 0, $Data, $Offset, 8)
}

function Add-BiffRecord([byte[]]$Data, [ref]$Offset, [UInt16]$Id, [byte[]]$Payload) {
    Set-U16 $Data $Offset.Value $Id
    Set-U16 $Data ($Offset.Value + 2) ([UInt16]$Payload.Length)
    if ($Payload.Length -gt 0) {
        [Buffer]::BlockCopy($Payload, 0, $Data, $Offset.Value + 4, $Payload.Length)
    }
    $Offset.Value += 4 + $Payload.Length
}

function Set-DirectoryEntry([byte[]]$Data, [int]$Offset, [string]$Name, [byte]$Type, [int]$StartSector, [Int64]$Size) {
    $nameBytes = [Text.Encoding]::Unicode.GetBytes($Name + [char]0)
    [Buffer]::BlockCopy($nameBytes, 0, $Data, $Offset, $nameBytes.Length)
    Set-U16 $Data ($Offset + 64) ([UInt16]$nameBytes.Length)
    $Data[$Offset + 66] = $Type
    $Data[$Offset + 67] = 1
    Set-I32 $Data ($Offset + 68) -1
    Set-I32 $Data ($Offset + 72) -1
    Set-I32 $Data ($Offset + 76) -1
    Set-I32 $Data ($Offset + 116) $StartSector
    Set-I64 $Data ($Offset + 120) $Size
}

function New-SyntheticXls {
    $sectorSize = 512
    $sectorCount = 10
    $bytes = New-Object byte[] ($sectorSize * ($sectorCount + 1))
    $signature = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    [Buffer]::BlockCopy($signature, 0, $bytes, 0, $signature.Length)
    Set-U16 $bytes 24 0x003E
    Set-U16 $bytes 26 0x0003
    Set-U16 $bytes 28 0xFFFE
    Set-U16 $bytes 30 9
    Set-U16 $bytes 32 6
    Set-I32 $bytes 44 1
    Set-I32 $bytes 48 1
    Set-I32 $bytes 56 4096
    Set-I32 $bytes 60 -2
    Set-I32 $bytes 64 0
    Set-I32 $bytes 68 -2
    Set-I32 $bytes 72 0
    for ($i = 0; $i -lt 109; $i++) { Set-I32 $bytes (76 + ($i * 4)) -1 }
    Set-I32 $bytes 76 0

    $fatOffset = $sectorSize
    for ($i = 0; $i -lt ($sectorSize / 4); $i++) { Set-I32 $bytes ($fatOffset + ($i * 4)) -1 }
    Set-I32 $bytes ($fatOffset + (0 * 4)) -3
    Set-I32 $bytes ($fatOffset + (1 * 4)) -2
    for ($i = 2; $i -lt 10; $i++) {
        Set-I32 $bytes ($fatOffset + ($i * 4)) ($(if ($i -eq 9) { -2 } else { $i + 1 }))
    }

    $directoryOffset = $sectorSize * 2
    Set-DirectoryEntry $bytes $directoryOffset 'Root Entry' 5 -2 0
    Set-DirectoryEntry $bytes ($directoryOffset + 128) 'Workbook' 2 2 4096

    $workbook = New-Object byte[] 4096
    $recordOffset = 0
    $boundSheet = New-Object byte[] 12
    Set-I32 $boundSheet 0 0
    Set-U16 $boundSheet 4 0
    $boundSheet[6] = 4
    $boundSheet[7] = 0
    [Buffer]::BlockCopy([Text.Encoding]::ASCII.GetBytes('Test'), 0, $boundSheet, 8, 4)
    Add-BiffRecord $workbook ([ref]$recordOffset) 0x0085 $boundSheet
    $value = New-Object byte[] 2
    Set-U16 $value 0 1
    Add-BiffRecord $workbook ([ref]$recordOffset) 0x0012 $value
    Set-U16 $value 0 0x1234
    Add-BiffRecord $workbook ([ref]$recordOffset) 0x0013 $value
    Set-U16 $value 0 1
    Add-BiffRecord $workbook ([ref]$recordOffset) 0x0019 $value
    Set-U16 $value 0 0x5678
    Add-BiffRecord $workbook ([ref]$recordOffset) 0x0013 $value
    Add-BiffRecord $workbook ([ref]$recordOffset) 0x000A (New-Object byte[] 0)

    $workbookOffset = $sectorSize * 3
    [Buffer]::BlockCopy($workbook, 0, $bytes, $workbookOffset, $workbook.Length)
    return ,$bytes
}

function Invoke-ToolJson([string[]]$Arguments) {
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'powershell.exe'
    $quoted = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $backend + '"'))
    foreach ($argument in $Arguments) { $quoted += ('"' + $argument + '"') }
    $startInfo.Arguments = $quoted -join ' '
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) { throw 'Could not start PowerShell child process.' }
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    if ($process.ExitCode -ne 0) { throw "Tool failed: $stderr$stdout" }
    return ($stdout | ConvertFrom-Json)
}

$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($backend, [ref]$tokens, [ref]$errors)
Assert-True ($errors.Count -eq 0) 'ExcelUnlocker.ps1 has a PowerShell parse error.'
$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($gui, [ref]$tokens, [ref]$errors)
Assert-True ($errors.Count -eq 0) 'ExcelUnlocker-GUI.ps1 has a PowerShell parse error.'

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('legacy-xls-test-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
$inputPath = Join-Path $testRoot 'synthetic-protected.xls'
$outputPath = $null
try {
    [IO.File]::WriteAllBytes($inputPath, (New-SyntheticXls))
    $originalHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $inputPath).Hash

    $doctor = Invoke-ToolJson @('doctor', '--json')
    Assert-True ($doctor.success -and $doctor.offline -and $doctor.version -eq '1.2.0') 'doctor check failed.'

    $inspection = Invoke-ToolJson @('inspect', $inputPath, '--json')
    Assert-True ($inspection.success -and $inspection.results[0].inspection.sheetCount -eq 1) 'Synthetic BIFF inspection failed.'
    Assert-True ($inspection.results[0].inspection.activeProtectionRecordCount -eq 2) 'Protection records were not detected.'

    $unlocked = Invoke-ToolJson @('unlock', $inputPath, '--json')
    Assert-True ($unlocked.success -and $unlocked.succeeded -eq 1) 'Synthetic unlock failed.'
    $result = $unlocked.results[0]
    $outputPath = [string]$result.output
    Assert-True ([IO.File]::Exists($outputPath)) 'Unlock output was not created.'
    Assert-True ($result.changedByteCount -eq 6) 'Unexpected number of changed bytes.'
    Assert-True ($result.inspection.activeProtectionRecordCount -eq 0) 'Output still contains active protection.'
    Assert-True ((Get-FileHash -Algorithm SHA256 -LiteralPath $inputPath).Hash -eq $originalHash) 'Original input was modified.'
    Assert-True (@(Get-ChildItem -LiteralPath $testRoot -Filter '*.tmp-*' -Force).Count -eq 0) 'Temporary output was left behind.'
    Write-Output 'PASS: syntax, doctor, synthetic BIFF inspect, non-destructive unlock, and verification.'
} finally {
    if ([IO.Directory]::Exists($testRoot)) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
