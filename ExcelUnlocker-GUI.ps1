Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:Backend = Join-Path $script:BaseDir 'ExcelUnlocker.ps1'
$script:LastOutput = $null
$script:Busy = $false
$script:AutomationMode = (-not [string]::IsNullOrWhiteSpace($env:EXCEL_UNLOCKER_AUTOMATION))

function Quote-ProcessArgument([string]$Value) {
    if ($Value.Contains('"')) { throw '路径中不能包含双引号。' }
    return '"' + $Value + '"'
}

function Invoke-Backend([string]$Command, [string[]]$Paths) {
    if (-not (Test-Path -LiteralPath $script:Backend)) {
        throw '核心程序 ExcelUnlocker.ps1 不在当前目录。'
    }
    $arguments = @(
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Quote-ProcessArgument $script:Backend),
        $Command
    )
    foreach ($path in $Paths) {
        $arguments += Quote-ProcessArgument $path
    }
    $arguments += '--json'

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'powershell.exe'
    $startInfo.Arguments = $arguments -join ' '
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) { throw '无法启动核心处理程序。' }
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    if ([string]::IsNullOrWhiteSpace($stdout)) {
        if ([string]::IsNullOrWhiteSpace($stderr)) { $stderr = '核心处理程序没有返回结果。' }
        throw $stderr.Trim()
    }
    try {
        return $stdout | ConvertFrom-Json
    } catch {
        throw ('无法解析核心程序返回结果：' + $stdout.Trim())
    }
}

function New-Button([string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height) {
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 1
    $button.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $button
}

function Add-PressFeedback($Button, [Drawing.Color]$NormalColor, [Drawing.Color]$PressedColor) {
    $Button.Tag = @{
        NormalColor = $NormalColor
        PressedColor = $PressedColor
    }
    $Button.Add_MouseDown({
        if ($this.Enabled) { $this.BackColor = $this.Tag.PressedColor }
    })
    $Button.Add_MouseUp({
        if ($this.Enabled) { $this.BackColor = $this.Tag.NormalColor }
    })
    $Button.Add_MouseLeave({
        if ($this.Enabled) { $this.BackColor = $this.Tag.NormalColor }
    })
}

function Set-Busy([bool]$Busy, [string]$Text) {
    $script:Busy = $Busy
    $selectFilesButton.Enabled = -not $Busy
    $selectFolderButton.Enabled = -not $Busy
    $clearButton.Enabled = -not $Busy
    $unlockButton.Enabled = (-not $Busy -and $fileList.Items.Count -gt 0)
    if ($unlockButton.Enabled) {
        $unlockButton.BackColor = [Drawing.Color]::FromArgb(6, 118, 71)
        $unlockButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(6, 118, 71)
    } else {
        $unlockButton.BackColor = [Drawing.Color]::FromArgb(208, 213, 221)
        $unlockButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(208, 213, 221)
    }
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
    if ($Busy) {
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $progressBar.MarqueeAnimationSpeed = 28
    }
    $statusLabel.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

function Find-ListItem([string]$Path) {
    foreach ($item in $fileList.Items) {
        if ([string]::Equals([string]$item.Tag, $Path, [StringComparison]::OrdinalIgnoreCase)) {
            return $item
        }
    }
    return $null
}

function Add-FileRows([string[]]$Paths) {
    $newPaths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($path in $Paths) {
        $fullPath = [IO.Path]::GetFullPath($path)
        if ([IO.Path]::GetExtension($fullPath) -ine '.xls') { continue }
        if ($null -ne (Find-ListItem $fullPath)) { continue }

        $item = New-Object System.Windows.Forms.ListViewItem([IO.Path]::GetFileName($fullPath))
        $item.SubItems.Add([IO.Path]::GetDirectoryName($fullPath)) | Out-Null
        $item.SubItems.Add('等待检查') | Out-Null
        $item.SubItems.Add('') | Out-Null
        $item.Tag = $fullPath
        $fileList.Items.Add($item) | Out-Null
        $newPaths.Add($fullPath)
    }

    if ($newPaths.Count -eq 0) {
        $unlockButton.Enabled = ($fileList.Items.Count -gt 0)
        return
    }

    $emptyStatePanel.Visible = $false
    $fileList.Visible = $true
    $fileCountLabel.Text = "$($fileList.Items.Count) 个文件"
    Set-Busy $true "正在检查 $($newPaths.Count) 个文件……"
    try {
        $report = Invoke-Backend 'inspect' $newPaths.ToArray()
        foreach ($result in $report.results) {
            $item = Find-ListItem ([string]$result.input)
            if ($null -eq $item) { continue }
            if ($result.status -eq 'error') {
                $item.SubItems[2].Text = '检查失败'
                $item.ForeColor = [Drawing.Color]::Firebrick
                $item.ToolTipText = [string]$result.message
                continue
            }
            $inspection = $result.inspection
            if ($inspection.fileOpenEncrypted) {
                $item.SubItems[2].Text = '打开密码加密（不支持）'
                $item.ForeColor = [Drawing.Color]::DarkOrange
            } elseif ($inspection.activeProtectionRecordCount -gt 0) {
                $item.SubItems[2].Text = "可解除（$($inspection.activeProtectionRecordCount) 项保护）"
                $item.ForeColor = [Drawing.Color]::FromArgb(16, 101, 71)
            } else {
                $item.SubItems[2].Text = '无需处理'
                $item.ForeColor = [Drawing.Color]::DimGray
            }
        }
        $fileCountLabel.Text = "$($fileList.Items.Count) 个文件"
        $statusLabel.Text = "已添加 $($fileList.Items.Count) 个文件，检查完成。"
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            '检查失败',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $statusLabel.Text = '检查失败，请确认文件格式。'
    } finally {
        Set-Busy $false $statusLabel.Text
    }
}

function Select-ExcelFiles {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = '选择需要解除工作表保护的 Excel 文件'
    $dialog.Filter = '旧式 Excel 工作簿 (*.xls)|*.xls'
    $dialog.Multiselect = $true
    $dialog.CheckFileExists = $true
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-FileRows $dialog.FileNames
    }
}

function Select-ExcelFolder {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = '选择包含 .xls 文件的文件夹'
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $option = [IO.SearchOption]::TopDirectoryOnly
    if ($recursiveCheckBox.Checked) { $option = [IO.SearchOption]::AllDirectories }
    $files = [IO.Directory]::EnumerateFiles($dialog.SelectedPath, '*.xls', $option)
    Add-FileRows @($files)
}

function Unlock-SelectedFiles {
    if ($fileList.Items.Count -eq 0) { return }
    $paths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $fileList.Items) {
        $paths.Add([string]$item.Tag)
        $item.SubItems[2].Text = '处理中……'
        $item.ForeColor = [Drawing.Color]::Black
    }

    Set-Busy $true "正在处理 $($paths.Count) 个文件，请稍候……"
    try {
        $report = Invoke-Backend 'unlock' $paths.ToArray()
        $success = 0
        $skipped = 0
        $failed = 0
        foreach ($result in $report.results) {
            $item = Find-ListItem ([string]$result.input)
            if ($null -eq $item) { continue }
            if ($result.status -eq 'success') {
                $success++
                $item.SubItems[2].Text = '已解除保护'
                $item.ForeColor = [Drawing.Color]::FromArgb(16, 101, 71)
                $item.SubItems[3].Text = [string]$result.output
                $script:LastOutput = [string]$result.output
            } elseif ($result.status -eq 'skipped') {
                $skipped++
                if ($result.inspection.fileOpenEncrypted) {
                    $item.SubItems[2].Text = '打开密码加密（已跳过）'
                    $item.ForeColor = [Drawing.Color]::DarkOrange
                } else {
                    $item.SubItems[2].Text = '无需处理'
                    $item.ForeColor = [Drawing.Color]::DimGray
                }
                $item.ToolTipText = [string]$result.message
            } else {
                $failed++
                $item.SubItems[2].Text = '处理失败'
                $item.ForeColor = [Drawing.Color]::Firebrick
                $item.ToolTipText = [string]$result.message
            }
        }
        $openOutputButton.Enabled = ($null -ne $script:LastOutput)
        $statusLabel.Text = "完成：成功 $success，跳过 $skipped，失败 $failed。"
        $icon = [System.Windows.Forms.MessageBoxIcon]::Information
        if ($failed -gt 0) { $icon = [System.Windows.Forms.MessageBoxIcon]::Warning }
        if (-not $script:AutomationMode) {
            [System.Windows.Forms.MessageBox]::Show(
                "处理完成。`r`n`r`n成功：$success`r`n跳过：$skipped`r`n失败：$failed`r`n`r`n原文件没有被修改。",
                '处理结果',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                $icon
            ) | Out-Null
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            '处理失败',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        $statusLabel.Text = '处理失败。'
    } finally {
        Set-Busy $false $statusLabel.Text
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Unprotect · Excel 工作表解锁'
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Size = New-Object System.Drawing.Size(920, 650)
$form.MinimumSize = New-Object System.Drawing.Size(860, 600)
$form.BackColor = [Drawing.Color]::FromArgb(246, 247, 249)
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.AllowDrop = $true

$header = New-Object System.Windows.Forms.Panel
$header.Dock = [System.Windows.Forms.DockStyle]::Top
$header.Height = 76
$header.BackColor = [Drawing.Color]::White
$form.Controls.Add($header)

$brandMark = New-Object System.Windows.Forms.Label
$brandMark.Text = 'XL'
$brandMark.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$brandMark.ForeColor = [Drawing.Color]::White
$brandMark.BackColor = [Drawing.Color]::FromArgb(25, 36, 53)
$brandMark.Font = New-Object System.Drawing.Font('Segoe UI', 11, [Drawing.FontStyle]::Bold)
$brandMark.Location = New-Object System.Drawing.Point(24, 17)
$brandMark.Size = New-Object System.Drawing.Size(42, 42)
$header.Controls.Add($brandMark)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'Excel 工作表解锁'
$titleLabel.ForeColor = [Drawing.Color]::FromArgb(25, 36, 53)
$titleLabel.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 13.5, [Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(78, 14)
$header.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = '移除旧式 .xls 的工作表保护，并始终保留原文件'
$subtitleLabel.ForeColor = [Drawing.Color]::FromArgb(102, 112, 133)
$subtitleLabel.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$subtitleLabel.AutoSize = $true
$subtitleLabel.Location = New-Object System.Drawing.Point(80, 42)
$header.Controls.Add($subtitleLabel)

$offlineBadge = New-Object System.Windows.Forms.Label
$offlineBadge.Text = '●  离线模式'
$offlineBadge.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$offlineBadge.ForeColor = [Drawing.Color]::FromArgb(6, 118, 71)
$offlineBadge.BackColor = [Drawing.Color]::FromArgb(236, 253, 243)
$offlineBadge.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 8.5, [Drawing.FontStyle]::Bold)
$offlineBadge.Location = New-Object System.Drawing.Point(788, 24)
$offlineBadge.Size = New-Object System.Drawing.Size(100, 28)
$offlineBadge.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$header.Controls.Add($offlineBadge)

$headerDivider = New-Object System.Windows.Forms.Panel
$headerDivider.Dock = [System.Windows.Forms.DockStyle]::Bottom
$headerDivider.Height = 1
$headerDivider.BackColor = [Drawing.Color]::FromArgb(234, 236, 240)
$header.Controls.Add($headerDivider)

$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Location = New-Object System.Drawing.Point(24, 92)
$actionPanel.Size = New-Object System.Drawing.Size(864, 64)
$actionPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$actionPanel.BackColor = [Drawing.Color]::White
$actionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($actionPanel)

$selectFilesButton = New-Button '＋ 选择 .xls 文件' 14 12 158 38
$selectFilesButton.BackColor = [Drawing.Color]::FromArgb(25, 36, 53)
$selectFilesButton.ForeColor = [Drawing.Color]::White
$selectFilesButton.FlatAppearance.BorderColor = $selectFilesButton.BackColor
$actionPanel.Controls.Add($selectFilesButton)
Add-PressFeedback $selectFilesButton ([Drawing.Color]::FromArgb(25, 36, 53)) ([Drawing.Color]::FromArgb(12, 19, 30))

$selectFolderButton = New-Button '选择文件夹' 182 12 118 38
$selectFolderButton.BackColor = [Drawing.Color]::White
$selectFolderButton.ForeColor = [Drawing.Color]::FromArgb(52, 64, 84)
$selectFolderButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(208, 213, 221)
$actionPanel.Controls.Add($selectFolderButton)
Add-PressFeedback $selectFolderButton ([Drawing.Color]::White) ([Drawing.Color]::FromArgb(242, 244, 247))

$recursiveCheckBox = New-Object System.Windows.Forms.CheckBox
$recursiveCheckBox.Text = '包含子文件夹'
$recursiveCheckBox.AutoSize = $true
$recursiveCheckBox.Location = New-Object System.Drawing.Point(316, 23)
$recursiveCheckBox.ForeColor = [Drawing.Color]::FromArgb(71, 84, 103)
$actionPanel.Controls.Add($recursiveCheckBox)

$fileCountLabel = New-Object System.Windows.Forms.Label
$fileCountLabel.Text = '0 个文件'
$fileCountLabel.TextAlign = [Drawing.ContentAlignment]::MiddleRight
$fileCountLabel.ForeColor = [Drawing.Color]::FromArgb(102, 112, 133)
$fileCountLabel.Location = New-Object System.Drawing.Point(630, 17)
$fileCountLabel.Size = New-Object System.Drawing.Size(104, 28)
$fileCountLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$actionPanel.Controls.Add($fileCountLabel)

$clearButton = New-Button '清空' 744 12 104 38
$clearButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$clearButton.BackColor = [Drawing.Color]::White
$clearButton.ForeColor = [Drawing.Color]::FromArgb(71, 84, 103)
$clearButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(208, 213, 221)
$actionPanel.Controls.Add($clearButton)
Add-PressFeedback $clearButton ([Drawing.Color]::White) ([Drawing.Color]::FromArgb(242, 244, 247))

$fileList = New-Object System.Windows.Forms.ListView
$fileList.Location = New-Object System.Drawing.Point(24, 172)
$fileList.Size = New-Object System.Drawing.Size(864, 300)
$fileList.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$fileList.View = [System.Windows.Forms.View]::Details
$fileList.FullRowSelect = $true
$fileList.GridLines = $false
$fileList.HideSelection = $false
$fileList.ShowItemToolTips = $true
$fileList.BackColor = [Drawing.Color]::White
$fileList.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$fileList.ForeColor = [Drawing.Color]::FromArgb(52, 64, 84)
$fileList.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$fileList.Columns.Add('文件', 190) | Out-Null
$fileList.Columns.Add('位置', 310) | Out-Null
$fileList.Columns.Add('状态', 190) | Out-Null
$fileList.Columns.Add('输出副本', 360) | Out-Null
$fileList.Visible = $false
$form.Controls.Add($fileList)

$emptyStatePanel = New-Object System.Windows.Forms.Panel
$emptyStatePanel.Location = New-Object System.Drawing.Point(24, 172)
$emptyStatePanel.Size = New-Object System.Drawing.Size(864, 300)
$emptyStatePanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$emptyStatePanel.BackColor = [Drawing.Color]::White
$emptyStatePanel.AllowDrop = $true
$form.Controls.Add($emptyStatePanel)
$emptyStatePanel.BringToFront()
$emptyStatePanel.Add_Paint({
    $rectangle = $emptyStatePanel.ClientRectangle
    $rectangle.Inflate(-1, -1)
    $pen = New-Object Drawing.Pen([Drawing.Color]::FromArgb(208, 213, 221), 1)
    $pen.DashStyle = [Drawing.Drawing2D.DashStyle]::Dash
    $_.Graphics.DrawRectangle($pen, $rectangle)
    $pen.Dispose()
})

$emptyIcon = New-Object System.Windows.Forms.Label
$emptyIcon.Text = '＋'
$emptyIcon.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$emptyIcon.Font = New-Object System.Drawing.Font('Segoe UI', 21)
$emptyIcon.ForeColor = [Drawing.Color]::FromArgb(71, 84, 103)
$emptyIcon.BackColor = [Drawing.Color]::FromArgb(242, 244, 247)
$emptyIcon.Location = New-Object System.Drawing.Point(402, 44)
$emptyIcon.Size = New-Object System.Drawing.Size(60, 60)
$emptyIcon.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$emptyStatePanel.Controls.Add($emptyIcon)

$emptyTitle = New-Object System.Windows.Forms.Label
$emptyTitle.Text = '添加需要解除保护的 Excel 文件'
$emptyTitle.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$emptyTitle.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [Drawing.FontStyle]::Bold)
$emptyTitle.ForeColor = [Drawing.Color]::FromArgb(25, 36, 53)
$emptyTitle.Location = New-Object System.Drawing.Point(150, 116)
$emptyTitle.Size = New-Object System.Drawing.Size(564, 30)
$emptyTitle.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$emptyStatePanel.Controls.Add($emptyTitle)

$emptyDescription = New-Object System.Windows.Forms.Label
$emptyDescription.Text = '点击下方按钮选择文件，或直接拖放 .xls 到此区域'
$emptyDescription.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$emptyDescription.ForeColor = [Drawing.Color]::FromArgb(102, 112, 133)
$emptyDescription.Location = New-Object System.Drawing.Point(150, 148)
$emptyDescription.Size = New-Object System.Drawing.Size(564, 24)
$emptyDescription.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$emptyStatePanel.Controls.Add($emptyDescription)

$emptySelectButton = New-Button '选择 .xls 文件' 342 188 180 42
$emptySelectButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top
$emptySelectButton.BackColor = [Drawing.Color]::White
$emptySelectButton.ForeColor = [Drawing.Color]::FromArgb(25, 36, 53)
$emptySelectButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(208, 213, 221)
$emptyStatePanel.Controls.Add($emptySelectButton)
Add-PressFeedback $emptySelectButton ([Drawing.Color]::White) ([Drawing.Color]::FromArgb(242, 244, 247))

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = '输出副本保存在原文件旁；真正的“打开文件密码”会被安全跳过。'
$hintLabel.AutoSize = $true
$hintLabel.Location = New-Object System.Drawing.Point(25, 484)
$hintLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$hintLabel.ForeColor = [Drawing.Color]::FromArgb(102, 112, 133)
$form.Controls.Add($hintLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(24, 514)
$progressBar.Size = New-Object System.Drawing.Size(590, 6)
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = '准备就绪'
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(24, 534)
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$statusLabel.ForeColor = [Drawing.Color]::FromArgb(71, 84, 103)
$form.Controls.Add($statusLabel)

$openOutputButton = New-Button '打开输出位置' 628 506 118 44
$openOutputButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$openOutputButton.BackColor = [Drawing.Color]::White
$openOutputButton.ForeColor = [Drawing.Color]::FromArgb(52, 64, 84)
$openOutputButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(208, 213, 221)
$openOutputButton.Enabled = $false
$form.Controls.Add($openOutputButton)
Add-PressFeedback $openOutputButton ([Drawing.Color]::White) ([Drawing.Color]::FromArgb(242, 244, 247))

$unlockButton = New-Button '解除工作表保护' 756 506 132 44
$unlockButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$unlockButton.BackColor = [Drawing.Color]::FromArgb(6, 118, 71)
$unlockButton.ForeColor = [Drawing.Color]::White
$unlockButton.FlatAppearance.BorderColor = $unlockButton.BackColor
$unlockButton.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5, [Drawing.FontStyle]::Bold)
$unlockButton.Enabled = $false
$unlockButton.BackColor = [Drawing.Color]::FromArgb(208, 213, 221)
$unlockButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(208, 213, 221)
$form.Controls.Add($unlockButton)
Add-PressFeedback $unlockButton ([Drawing.Color]::FromArgb(6, 118, 71)) ([Drawing.Color]::FromArgb(4, 95, 57))

$selectFilesButton.Add_Click({ Select-ExcelFiles })
$emptySelectButton.Add_Click({ Select-ExcelFiles })
$selectFolderButton.Add_Click({ Select-ExcelFolder })
$clearButton.Add_Click({
    if ($script:Busy) { return }
    $fileList.Items.Clear()
    $fileList.Visible = $false
    $emptyStatePanel.Visible = $true
    $emptyStatePanel.BringToFront()
    $fileCountLabel.Text = '0 个文件'
    $script:LastOutput = $null
    $openOutputButton.Enabled = $false
    $unlockButton.Enabled = $false
    $statusLabel.Text = '准备就绪'
})
$unlockButton.Add_Click({ Unlock-SelectedFiles })
$openOutputButton.Add_Click({
    if ($null -eq $script:LastOutput) { return }
    if (Test-Path -LiteralPath $script:LastOutput) {
        Start-Process explorer.exe -ArgumentList ('/select,"' + $script:LastOutput + '"')
    }
})
$dragEnterHandler = {
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
}
$dragDropHandler = {
    $dropped = @($_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop))
    $files = New-Object 'System.Collections.Generic.List[string]'
    foreach ($entry in $dropped) {
        if ([IO.File]::Exists($entry) -and [IO.Path]::GetExtension($entry) -ieq '.xls') {
            $files.Add($entry)
        } elseif ([IO.Directory]::Exists($entry)) {
            $option = [IO.SearchOption]::TopDirectoryOnly
            if ($recursiveCheckBox.Checked) { $option = [IO.SearchOption]::AllDirectories }
            foreach ($file in [IO.Directory]::EnumerateFiles($entry, '*.xls', $option)) {
                $files.Add($file)
            }
        }
    }
    Add-FileRows $files.ToArray()
}
$form.Add_DragEnter($dragEnterHandler)
$form.Add_DragDrop($dragDropHandler)
$emptyStatePanel.Add_DragEnter($dragEnterHandler)
$emptyStatePanel.Add_DragDrop($dragDropHandler)
$fileList.AllowDrop = $true
$fileList.Add_DragEnter($dragEnterHandler)
$fileList.Add_DragDrop($dragDropHandler)

if (-not [string]::IsNullOrWhiteSpace($env:EXCEL_UNLOCKER_TEST_FILE)) {
    Add-FileRows @($env:EXCEL_UNLOCKER_TEST_FILE)
    if ($env:EXCEL_UNLOCKER_AUTOMATION -eq 'unlock') {
        Unlock-SelectedFiles
    }
}

if (-not [string]::IsNullOrWhiteSpace($env:EXCEL_UNLOCKER_SCREENSHOT)) {
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    $bitmap = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
    $form.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
    $bitmap.Save($env:EXCEL_UNLOCKER_SCREENSHOT, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
    $form.Close()
    exit 0
}

[System.Windows.Forms.Application]::Run($form)
