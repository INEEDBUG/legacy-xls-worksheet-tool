Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

$Script:ToolVersion = '1.2.0'
$Script:OleSignature = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
$Script:EndOfChain = -2
$Script:ExitCode = 0

function Read-UInt16([byte[]]$Data, [int]$Offset) {
    return [BitConverter]::ToUInt16($Data, $Offset)
}

function Write-UInt16([byte[]]$Data, [int]$Offset, [UInt16]$Value) {
    $bytes = [BitConverter]::GetBytes($Value)
    [Buffer]::BlockCopy($bytes, 0, $Data, $Offset, 2)
}

function Read-Int32([byte[]]$Data, [int]$Offset) {
    return [BitConverter]::ToInt32($Data, $Offset)
}

function Read-Int64([byte[]]$Data, [int]$Offset) {
    return [BitConverter]::ToInt64($Data, $Offset)
}

function Get-SectorOffset($Info, [int]$SectorId) {
    if ($SectorId -lt 0) { throw 'OLE 扇区编号无效。' }
    $offset = ($SectorId + 1) * $Info.SectorSize
    if ($offset -lt 0 -or ($offset + $Info.SectorSize) -gt $Info.Bytes.Length) {
        throw 'OLE 扇区指向文件外部。'
    }
    return $offset
}

function Get-Chain([int]$StartSector, [int[]]$Table) {
    $result = New-Object 'System.Collections.Generic.List[int]'
    $seen = New-Object 'System.Collections.Generic.HashSet[int]'
    $sid = $StartSector
    while ($sid -ge 0 -and $sid -ne $Script:EndOfChain) {
        if ($sid -ge $Table.Length) { throw 'OLE FAT 链索引越界。' }
        if (-not $seen.Add($sid)) { throw 'OLE FAT 链包含循环。' }
        $result.Add($sid)
        $sid = $Table[$sid]
    }
    return ,([int[]]$result.ToArray())
}

function Read-RegularChain($Info, [int]$StartSector, [long]$RequestedBytes) {
    $chain = Get-Chain $StartSector $Info.Fat
    $maximum = [long]$chain.Count * $Info.SectorSize
    $length = [Math]::Min($RequestedBytes, $maximum)
    if ($length -gt [int]::MaxValue) { throw 'OLE 数据流过大。' }
    $output = New-Object byte[] ([int]$length)
    $written = 0
    foreach ($sid in $chain) {
        if ($written -ge $length) { break }
        $sourceOffset = Get-SectorOffset $Info $sid
        $take = [Math]::Min($Info.SectorSize, [int]$length - $written)
        [Buffer]::BlockCopy($Info.Bytes, $sourceOffset, $output, $written, $take)
        $written += $take
    }
    return ,$output
}

function Write-RegularChain([byte[]]$Output, $Info, [int]$StartSector, [byte[]]$Stream) {
    $chain = Get-Chain $StartSector $Info.Fat
    $read = 0
    foreach ($sid in $chain) {
        if ($read -ge $Stream.Length) { break }
        $destinationOffset = Get-SectorOffset $Info $sid
        $take = [Math]::Min($Info.SectorSize, $Stream.Length - $read)
        [Buffer]::BlockCopy($Stream, $read, $Output, $destinationOffset, $take)
        $read += $take
    }
    if ($read -ne $Stream.Length) { throw 'OLE FAT 链长度不足。' }
}

function Read-MiniChain($Info, [byte[]]$RootMiniStream, [int]$StartSector, [int]$RequestedBytes) {
    $chain = Get-Chain $StartSector $Info.MiniFat
    $output = New-Object byte[] $RequestedBytes
    $written = 0
    foreach ($sid in $chain) {
        if ($written -ge $RequestedBytes) { break }
        $sourceOffset = $sid * $Info.MiniSectorSize
        if ($sourceOffset -lt 0 -or ($sourceOffset + $Info.MiniSectorSize) -gt $RootMiniStream.Length) {
            throw 'OLE MiniFAT 指向文件外部。'
        }
        $take = [Math]::Min($Info.MiniSectorSize, $RequestedBytes - $written)
        [Buffer]::BlockCopy($RootMiniStream, $sourceOffset, $output, $written, $take)
        $written += $take
    }
    if ($written -ne $RequestedBytes) { throw 'OLE 迷你流长度不足。' }
    return ,$output
}

function Write-MiniChain($Info, [byte[]]$RootMiniStream, [int]$StartSector, [byte[]]$Stream) {
    $chain = Get-Chain $StartSector $Info.MiniFat
    $read = 0
    foreach ($sid in $chain) {
        if ($read -ge $Stream.Length) { break }
        $destinationOffset = $sid * $Info.MiniSectorSize
        $take = [Math]::Min($Info.MiniSectorSize, $Stream.Length - $read)
        [Buffer]::BlockCopy($Stream, $read, $RootMiniStream, $destinationOffset, $take)
        $read += $take
    }
    if ($read -ne $Stream.Length) { throw 'OLE MiniFAT 链长度不足。' }
}

function New-CfbInfo([byte[]]$Bytes) {
    if ($Bytes.Length -lt 512) { throw 'OLE 文件头过短。' }
    for ($i = 0; $i -lt 8; $i++) {
        if ($Bytes[$i] -ne $Script:OleSignature[$i]) { throw '文件不是有效的旧式 OLE .xls。' }
    }

    $sectorSize = 1 -shl (Read-UInt16 $Bytes 30)
    $miniSectorSize = 1 -shl (Read-UInt16 $Bytes 32)
    if ($sectorSize -ne 512 -and $sectorSize -ne 4096) { throw '不支持的 OLE 扇区大小。' }

    $info = [pscustomobject]@{
        Bytes = $Bytes
        SectorSize = $sectorSize
        MiniSectorSize = $miniSectorSize
        MiniCutoff = Read-Int32 $Bytes 56
        Fat = [int[]]@()
        MiniFat = [int[]]@()
        Entries = @()
        Root = $null
    }

    $firstDirectorySector = Read-Int32 $Bytes 48
    $firstMiniFatSector = Read-Int32 $Bytes 60
    $miniFatSectorCount = Read-Int32 $Bytes 64
    $firstDifatSector = Read-Int32 $Bytes 68
    $difatSectorCount = Read-Int32 $Bytes 72

    $difat = New-Object 'System.Collections.Generic.List[int]'
    for ($i = 0; $i -lt 109; $i++) {
        $sid = Read-Int32 $Bytes (76 + $i * 4)
        if ($sid -ge 0) { $difat.Add($sid) }
    }

    $difatSid = $firstDifatSector
    for ($n = 0; $n -lt $difatSectorCount -and $difatSid -ge 0; $n++) {
        $sectorOffset = Get-SectorOffset $info $difatSid
        for ($i = 0; $i -lt ($sectorSize / 4 - 1); $i++) {
            $sid = Read-Int32 $Bytes ($sectorOffset + $i * 4)
            if ($sid -ge 0) { $difat.Add($sid) }
        }
        $difatSid = Read-Int32 $Bytes ($sectorOffset + $sectorSize - 4)
    }

    $fat = New-Object 'System.Collections.Generic.List[int]'
    foreach ($sid in $difat) {
        $sectorOffset = Get-SectorOffset $info $sid
        for ($i = 0; $i -lt ($sectorSize / 4); $i++) {
            $fat.Add((Read-Int32 $Bytes ($sectorOffset + $i * 4)))
        }
    }
    $info.Fat = [int[]]$fat.ToArray()

    $directoryBytes = Read-RegularChain $info $firstDirectorySector ([long][int]::MaxValue)
    $entries = New-Object 'System.Collections.Generic.List[object]'
    for ($offset = 0; ($offset + 128) -le $directoryBytes.Length; $offset += 128) {
        $nameBytes = Read-UInt16 $directoryBytes ($offset + 64)
        $type = $directoryBytes[$offset + 66]
        if ($type -eq 0) { continue }
        if ($nameBytes -ge 2 -and $nameBytes -le 64) {
            $name = [Text.Encoding]::Unicode.GetString($directoryBytes, $offset, $nameBytes - 2)
        } elseif ($type -eq 5) {
            $name = '(Root)'
        } else {
            continue
        }
        $entry = [pscustomobject]@{
            Name = $name
            Type = [int]$type
            StartSector = Read-Int32 $directoryBytes ($offset + 116)
            Size = Read-Int64 $directoryBytes ($offset + 120)
        }
        $entries.Add($entry)
    }
    $info.Entries = [object[]]$entries.ToArray()
    $info.Root = @($entries | Where-Object { $_.Type -eq 5 } | Select-Object -First 1)[0]
    if ($null -eq $info.Root) { throw 'OLE 根目录缺失。' }

    if ($firstMiniFatSector -ge 0 -and $miniFatSectorCount -gt 0) {
        $miniFatBytes = Read-RegularChain $info $firstMiniFatSector ([long]$miniFatSectorCount * $sectorSize)
        $miniFat = New-Object int[] ($miniFatBytes.Length / 4)
        for ($i = 0; $i -lt $miniFat.Length; $i++) {
            $miniFat[$i] = Read-Int32 $miniFatBytes ($i * 4)
        }
        $info.MiniFat = $miniFat
    }
    return $info
}

function Find-CfbStream($Info, [string]$Name) {
    return @($Info.Entries | Where-Object { $_.Type -eq 2 -and $_.Name -ieq $Name } | Select-Object -First 1)[0]
}

function Read-CfbStream($Info, $Entry) {
    if ($Entry.Size -lt $Info.MiniCutoff) {
        $rootMiniStream = Read-RegularChain $Info $Info.Root.StartSector $Info.Root.Size
        return ,(Read-MiniChain $Info $rootMiniStream $Entry.StartSector ([int]$Entry.Size))
    }
    return ,(Read-RegularChain $Info $Entry.StartSector $Entry.Size)
}

function Write-CfbStreamCopy($Info, $Entry, [byte[]]$Stream) {
    if ($Stream.Length -ne $Entry.Size) { throw '仅支持等长修改，拒绝改变工作簿结构。' }
    $output = [byte[]]$Info.Bytes.Clone()
    if ($Entry.Size -lt $Info.MiniCutoff) {
        $rootMiniStream = Read-RegularChain $Info $Info.Root.StartSector $Info.Root.Size
        Write-MiniChain $Info $rootMiniStream $Entry.StartSector $Stream
        Write-RegularChain $output $Info $Info.Root.StartSector $rootMiniStream
    } else {
        Write-RegularChain $output $Info $Entry.StartSector $Stream
    }
    return ,$output
}

function Test-ProtectionFlag([UInt16]$Id) {
    return ($Id -eq 0x0012 -or $Id -eq 0x0019 -or $Id -eq 0x0063 -or $Id -eq 0x01BC)
}

function Test-PasswordRecord([UInt16]$Id) {
    return ($Id -eq 0x0013 -or $Id -eq 0x01BD)
}

function Read-BoundSheetName([byte[]]$Workbook, [int]$Offset, [int]$Length) {
    $charCount = [int]$Workbook[$Offset + 6]
    $isUnicode = (($Workbook[$Offset + 7] -band 1) -ne 0)
    $byteCount = $charCount
    if ($isUnicode) { $byteCount *= 2 }
    if ((8 + $byteCount) -gt $Length) { return '(名称损坏)' }
    if ($isUnicode) { return [Text.Encoding]::Unicode.GetString($Workbook, $Offset + 8, $byteCount) }
    return [Text.Encoding]::GetEncoding(28591).GetString($Workbook, $Offset + 8, $byteCount)
}

function Get-XlsInspection([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $info = New-CfbInfo $bytes
    if ((Find-CfbStream $info 'EncryptedPackage') -or (Find-CfbStream $info 'EncryptionInfo')) {
        return [pscustomobject][ordered]@{
            format = 'encrypted_ooxml_container'
            fileOpenEncrypted = $true
            sheetCount = 0
            activeProtectionRecordCount = 0
            passwordRecordCount = 0
            sheets = @()
            fileBytes = $bytes.Length
        }
    }

    $workbookEntry = Find-CfbStream $info 'Workbook'
    if ($null -eq $workbookEntry) { $workbookEntry = Find-CfbStream $info 'Book' }
    if ($null -eq $workbookEntry) { throw '未找到 Workbook 数据流；文件可能损坏或不是 BIFF .xls。' }
    $workbook = Read-CfbStream $info $workbookEntry

    $sheetNames = New-Object 'System.Collections.Generic.List[string]'
    $activeProtection = 0
    $passwordCount = 0
    $encrypted = $false
    $offset = 0
    while (($offset + 4) -le $workbook.Length) {
        $id = Read-UInt16 $workbook $offset
        $length = Read-UInt16 $workbook ($offset + 2)
        if (($offset + 4 + $length) -gt $workbook.Length) { throw "BIFF 记录在偏移 $offset 处损坏。" }
        $dataOffset = $offset + 4
        if ($id -eq 0x002F) {
            $encrypted = $true
        } elseif ($id -eq 0x0085 -and $length -ge 8) {
            $sheetNames.Add((Read-BoundSheetName $workbook $dataOffset $length))
        } elseif ((Test-ProtectionFlag $id) -and $length -ge 2 -and (Read-UInt16 $workbook $dataOffset) -ne 0) {
            $activeProtection++
        } elseif ((Test-PasswordRecord $id) -and $length -ge 2) {
            $passwordCount++
        }
        $offset += 4 + $length
    }
    if ($offset -ne $workbook.Length) { throw 'BIFF 数据末尾不完整。' }

    $sheets = @()
    for ($i = 0; $i -lt $sheetNames.Count; $i++) {
        $sheets += [pscustomobject][ordered]@{ index = $i + 1; name = $sheetNames[$i] }
    }
    return [pscustomobject][ordered]@{
        format = 'xls_biff_ole'
        fileOpenEncrypted = $encrypted
        sheetCount = $sheetNames.Count
        activeProtectionRecordCount = $activeProtection
        passwordRecordCount = $passwordCount
        sheets = $sheets
        fileBytes = $bytes.Length
    }
}

function Clear-XlsProtection([byte[]]$Workbook) {
    $modified = 0
    $offset = 0
    while (($offset + 4) -le $Workbook.Length) {
        $id = Read-UInt16 $Workbook $offset
        $length = Read-UInt16 $Workbook ($offset + 2)
        if (($offset + 4 + $length) -gt $Workbook.Length) { throw "BIFF 记录在偏移 $offset 处损坏。" }
        $dataOffset = $offset + 4
        if ($id -eq 0x002F) { throw '检测到打开文件加密，工具不会尝试绕过。' }
        if ($length -ge 2 -and ((Test-ProtectionFlag $id) -or (Test-PasswordRecord $id))) {
            if ((Read-UInt16 $Workbook $dataOffset) -ne 0) {
                Write-UInt16 $Workbook $dataOffset 0
                $modified++
            }
        }
        $offset += 4 + $length
    }
    return $modified
}

function Unlock-Xls([string]$InputPath, [string]$OutputPath) {
    $original = [IO.File]::ReadAllBytes($InputPath)
    $info = New-CfbInfo $original
    if ((Find-CfbStream $info 'EncryptedPackage') -or (Find-CfbStream $info 'EncryptionInfo')) {
        throw '检测到打开文件加密，无法在不知道密码的情况下安全处理。'
    }
    $workbookEntry = Find-CfbStream $info 'Workbook'
    if ($null -eq $workbookEntry) { $workbookEntry = Find-CfbStream $info 'Book' }
    if ($null -eq $workbookEntry) { throw '未找到 Workbook 数据流。' }

    $workbook = Read-CfbStream $info $workbookEntry
    $modified = Clear-XlsProtection $workbook
    if ($modified -eq 0) { throw '未发现需要解除的保护记录。' }
    $outputBytes = Write-CfbStreamCopy $info $workbookEntry $workbook

    $changed = 0
    for ($i = 0; $i -lt $original.Length; $i++) {
        if ($original[$i] -ne $outputBytes[$i]) { $changed++ }
    }

    $temp = $OutputPath + '.tmp-' + [Guid]::NewGuid().ToString('N')
    try {
        [IO.File]::WriteAllBytes($temp, $outputBytes)
        $verification = Get-XlsInspection $temp
        if ($verification.fileOpenEncrypted -or $verification.activeProtectionRecordCount -ne 0) {
            throw '输出复检失败：仍检测到保护或加密。'
        }
        [IO.File]::Move($temp, $OutputPath)
        return [pscustomobject][ordered]@{
            verification = $verification
            changedByteCount = $changed
            modifiedRecordCount = $modified
        }
    } finally {
        if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    }
}

function Get-OutputPath([string]$InputPath, [string]$OutputDirectory) {
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $directory = [IO.Path]::GetDirectoryName($InputPath)
    } else {
        $directory = [IO.Path]::GetFullPath($OutputDirectory)
    }
    if (-not [IO.Directory]::Exists($directory)) { [IO.Directory]::CreateDirectory($directory) | Out-Null }
    $baseName = [IO.Path]::GetFileNameWithoutExtension($InputPath) + '_已解除保护'
    $extension = [IO.Path]::GetExtension($InputPath)
    $candidate = [IO.Path]::Combine($directory, $baseName + $extension)
    $index = 2
    while ([IO.File]::Exists($candidate)) {
        $candidate = [IO.Path]::Combine($directory, "$baseName ($index)$extension")
        $index++
    }
    return $candidate
}

function Get-InputFiles([string[]]$Inputs, [bool]$Recursive) {
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($raw in $Inputs) {
        $clean = $raw.Trim().Trim('"')
        $path = [IO.Path]::GetFullPath($clean)
        if ([IO.File]::Exists($path)) {
            if ([IO.Path]::GetExtension($path) -ine '.xls') { throw "暂不支持该文件类型：$path" }
            $files.Add($path) | Out-Null
        } elseif ([IO.Directory]::Exists($path)) {
            $option = [IO.SearchOption]::TopDirectoryOnly
            if ($Recursive) { $option = [IO.SearchOption]::AllDirectories }
            foreach ($file in [IO.Directory]::EnumerateFiles($path, '*.xls', $option)) {
                if (-not [IO.Path]::GetFileNameWithoutExtension($file).EndsWith('_已解除保护', [StringComparison]::OrdinalIgnoreCase)) {
                    $files.Add([IO.Path]::GetFullPath($file)) | Out-Null
                }
            }
        } else {
            throw "路径不存在：$path"
        }
    }
    return @($files | Sort-Object)
}

function Show-Help {
    @'
ExcelUnlocker 1.2.0 - 旧式 Excel 工作表保护离线处理工具

用法：
  ExcelUnlocker.cmd doctor [--json]
  ExcelUnlocker.cmd inspect <文件或文件夹...> [--recursive] [--json]
  ExcelUnlocker.cmd unlock <文件或文件夹...> [--recursive] [--out <目录>] [--json]
  ExcelUnlocker.cmd <文件或文件夹...>               # 拖放/快捷解锁

命令：
  doctor   检查运行环境和工具能力，不修改文件
  inspect  识别保护类型、工作表数量及是否存在打开文件加密
  unlock   生成“_已解除保护.xls”副本，绝不覆盖原文件

选项：
  --recursive, -r   递归处理文件夹中的 .xls 文件
  --out <目录>      将输出副本集中保存到指定目录
  --json            输出稳定的 JSON，便于批处理
  --help, -h        显示帮助

限制：
  - 支持 BIFF8/OLE 格式的 .xls 工作表和工作簿保护。
  - 不破解“打开文件密码”（真正加密）；检测到后会安全跳过。
  - 不支持 .xlsx/.xlsm。
  - 工具仅应用于你有权处理的文件。
'@ | Write-Output
}

function New-Result([string]$Status, [string]$InputPath, $Inspection, [string]$OutputPath, [string]$Message, $UnlockResult) {
    $item = [ordered]@{
        status = $Status
        input = $InputPath
    }
    if ($OutputPath) { $item.output = $OutputPath }
    if ($Message) { $item.message = $Message }
    if ($Inspection) { $item.inspection = $Inspection }
    if ($UnlockResult) {
        $item.changedByteCount = $UnlockResult.changedByteCount
        $item.modifiedRecordCount = $UnlockResult.modifiedRecordCount
    }
    return [pscustomobject]$item
}

function Write-HumanResults($Results, [bool]$InspectOnly) {
    foreach ($result in $Results) {
        $marker = '[检查]'
        if ($result.status -eq 'success') { $marker = '[成功]' }
        elseif ($result.status -eq 'skipped') { $marker = '[跳过]' }
        elseif ($result.status -eq 'error') { $marker = '[错误]' }
        Write-Output "$marker $($result.input)"
        if ($result.PSObject.Properties.Name -contains 'inspection') {
            $i = $result.inspection
            $encryptedText = '否'
            if ($i.fileOpenEncrypted) { $encryptedText = '是' }
            Write-Output "       工作表：$($i.sheetCount)；启用保护记录：$($i.activeProtectionRecordCount)；打开加密：$encryptedText"
            if ($InspectOnly -and $i.sheets.Count -gt 0) {
                Write-Output ('       名称：' + (($i.sheets | ForEach-Object { $_.name }) -join '、'))
            }
        }
        if ($result.PSObject.Properties.Name -contains 'output') { Write-Output "       输出：$($result.output)" }
        if ($result.PSObject.Properties.Name -contains 'message') { Write-Output "       说明：$($result.message)" }
    }
    $successCount = @($Results | Where-Object { $_.status -eq 'success' -or $_.status -eq 'inspected' }).Count
    $skippedCount = @($Results | Where-Object { $_.status -eq 'skipped' }).Count
    $errorCount = @($Results | Where-Object { $_.status -eq 'error' }).Count
    Write-Output ''
    Write-Output "共 $($Results.Count) 个：成功 $successCount，跳过 $skippedCount，错误 $errorCount。"
}

function Invoke-Main([string[]]$Arguments) {
    $interactive = ($Arguments.Count -eq 0)
    if ($interactive) {
        Write-Output 'Excel 工作表离线解锁工具'
        Write-Output '支持旧式 .xls 工作表/工作簿保护；不会修改原文件。'
        Write-Output ''
        $inputPath = Read-Host '请拖入一个 .xls 文件或文件夹，然后按回车'
        if ([string]::IsNullOrWhiteSpace($inputPath)) { $Script:ExitCode = 0; return }
        $Arguments = @('unlock', $inputPath.Trim().Trim('"'))
        if ([IO.Directory]::Exists($inputPath.Trim().Trim('"'))) { $Arguments += '--recursive' }
    }

    $command = 'unlock'
    $index = 0
    if ($Arguments[0] -in @('doctor','inspect','unlock','help','--help','-h')) {
        $command = $Arguments[0].ToLowerInvariant()
        $index = 1
    }
    if ($command -in @('help','--help','-h')) {
        Show-Help
        $Script:ExitCode = 0
        return
    }

    $json = $false
    $recursive = $false
    $outputDirectory = ''
    $inputs = New-Object 'System.Collections.Generic.List[string]'
    while ($index -lt $Arguments.Count) {
        $arg = $Arguments[$index]
        if ($arg -eq '--json') {
            $json = $true
            $index++
        } elseif ($arg -eq '--recursive' -or $arg -eq '-r') {
            $recursive = $true
            $index++
        } elseif ($arg -eq '--out') {
            if (($index + 1) -ge $Arguments.Count) { throw '--out 后必须指定输出目录。' }
            $outputDirectory = $Arguments[$index + 1]
            $index += 2
        } elseif ($arg.StartsWith('-')) {
            throw "未知选项：$arg"
        } else {
            $inputs.Add($arg)
            $index++
        }
    }

    if ($command -eq 'doctor') {
        $doctor = [pscustomobject][ordered]@{
            success = $true
            command = 'doctor'
            version = $Script:ToolVersion
            offline = $true
            authenticationRequired = $false
            powershellVersion = $PSVersionTable.PSVersion.ToString()
            operatingSystem = [Environment]::OSVersion.VersionString
            supportedExtensions = @('.xls')
            capabilities = @(
                'inspect_legacy_xls_protection',
                'remove_worksheet_protection',
                'remove_workbook_structure_protection',
                'detect_file_open_encryption',
                'batch_folder_processing',
                'non_destructive_output'
            )
        }
        if ($json) { $doctor | ConvertTo-Json -Depth 8 }
        else {
            Write-Output "ExcelUnlocker $($Script:ToolVersion)"
            Write-Output '状态：可离线运行，无需登录或网络'
            Write-Output "PowerShell：$($PSVersionTable.PSVersion)"
            Write-Output '支持：.xls 工作表保护、工作簿结构保护、批量处理'
            Write-Output '安全：原文件不覆盖；真正的打开文件加密会被跳过'
        }
        $Script:ExitCode = 0
        return
    }

    if ($inputs.Count -eq 0) { throw '请至少提供一个 .xls 文件或文件夹。' }
    $files = Get-InputFiles $inputs.ToArray() $recursive
    $results = New-Object 'System.Collections.Generic.List[object]'
    foreach ($file in $files) {
        try {
            $inspection = Get-XlsInspection $file
            if ($command -eq 'inspect') {
                $results.Add((New-Result 'inspected' $file $inspection '' '' $null))
                continue
            }
            if ($inspection.fileOpenEncrypted) {
                $results.Add((New-Result 'skipped' $file $inspection '' '检测到真正的打开文件加密，工具不会尝试绕过。' $null))
                continue
            }
            if ($inspection.activeProtectionRecordCount -eq 0) {
                $results.Add((New-Result 'skipped' $file $inspection '' '未发现启用中的工作表或工作簿保护。' $null))
                continue
            }
            $outputPath = Get-OutputPath $file $outputDirectory
            $unlockResult = Unlock-Xls $file $outputPath
            $results.Add((New-Result 'success' $file $unlockResult.verification $outputPath '已生成副本并复检通过；原文件未修改。' $unlockResult))
        } catch {
            $results.Add((New-Result 'error' $file $null '' $_.Exception.Message $null))
        }
    }

    $successCount = @($results | Where-Object { $_.status -eq 'success' -or $_.status -eq 'inspected' }).Count
    $skippedCount = @($results | Where-Object { $_.status -eq 'skipped' }).Count
    $errorCount = @($results | Where-Object { $_.status -eq 'error' }).Count
    $envelope = [pscustomobject][ordered]@{
        success = ($errorCount -eq 0)
        command = $command
        processed = $results.Count
        succeeded = $successCount
        skipped = $skippedCount
        failed = $errorCount
        results = [object[]]$results.ToArray()
    }
    if ($json) { $envelope | ConvertTo-Json -Depth 12 }
    else { Write-HumanResults $results ($command -eq 'inspect') }

    if ($interactive) {
        Write-Output ''
        Read-Host '处理完成，按回车退出' | Out-Null
    }
    if ($errorCount -gt 0) { $Script:ExitCode = 1 }
    else { $Script:ExitCode = 0 }
}

try {
    Invoke-Main $args
    exit $Script:ExitCode
} catch {
    $jsonRequested = ($args -contains '--json')
    if ($jsonRequested) {
        [pscustomobject][ordered]@{
            success = $false
            error = [pscustomobject][ordered]@{
                code = 'runtime_error'
                message = $_.Exception.Message
            }
        } | ConvertTo-Json -Depth 6
    } else {
        [Console]::Error.WriteLine("错误：$($_.Exception.Message)")
        [Console]::Error.WriteLine('运行 ExcelUnlocker.cmd --help 查看帮助。')
    }
    exit 1
}
