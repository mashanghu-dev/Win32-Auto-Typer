# QuickText - 快捷文本模板库工具 v2
# PowerShell + Windows Forms 实现
# 标签云流式布局 + 点击直接输入 + 拖拽排序

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# ===== Win32 API =====
$signature = @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    
    [DllImport("user32.dll")]
    public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);
    
    public const int SW_MINIMIZE = 6;
    public const int SW_RESTORE = 9;
    public const uint GW_HWNDNEXT = 2;
}
"@
Add-Type $signature -ReferencedAssemblies System.Drawing

# ===== 配置文件路径 =====
$script:ConfigFile = Join-Path $PSScriptRoot "quicktext_config.json"
$script:SettingsFile = Join-Path $PSScriptRoot "quicktext_settings.json"

# ===== 默认数据 =====
$script:DefaultData = @{
    tabs = @(
        @{
            name = "常用"
            phrases = @(
                "你好，很高兴认识你！",
                "谢谢你的帮助！",
                "请问有什么可以帮您？",
                "稍等一下，我帮您查询",
                "好的，收到",
                "不客气~"
            )
        }
    )
}

$script:DefaultSettings = @{
    opacity = 0.92
    topMost = $true
    autoMinimize = $true
    windowWidth = 420
    windowHeight = 560
}

# ===== 全局变量 =====
$script:autoMinimize = $true
$script:draggingItem = $null
$script:dragStartPos = $null
$script:lastActiveHwnd = [IntPtr]::Zero

# ===== 强制设置前台窗口（用 AttachThreadInput 绕过 Win10 限制）=====
function Force-SetForegroundWindow($hWnd) {
    if ($hWnd -eq [IntPtr]::Zero) { return $false }
    
    # 如果窗口最小化，先恢复
    if ([Win32]::IsIconic($hWnd)) {
        [Win32]::ShowWindow($hWnd, [Win32]::SW_RESTORE) | Out-Null
    }
    
    $foregroundHWnd = [Win32]::GetForegroundWindow()
    $foregroundThreadId = [Win32]::GetWindowThreadProcessId($foregroundHWnd, [ref]0)
    $targetThreadId = [Win32]::GetWindowThreadProcessId($hWnd, [ref]0)
    
    if ($foregroundThreadId -ne $targetThreadId) {
        # 附加线程输入
        [Win32]::AttachThreadInput($foregroundThreadId, $targetThreadId, $true) | Out-Null
        $result = [Win32]::SetForegroundWindow($hWnd)
        [Win32]::AttachThreadInput($foregroundThreadId, $targetThreadId, $false) | Out-Null
        return $result
    }
    
    return [Win32]::SetForegroundWindow($hWnd)
}

# ===== 加载/保存数据 =====
function Load-Data {
    if (Test-Path $script:ConfigFile) {
        try {
            $json = Get-Content $script:ConfigFile -Raw -Encoding UTF8
            return $json | ConvertFrom-Json
        } catch {
            return [pscustomobject]$script:DefaultData
        }
    }
    return [pscustomobject]$script:DefaultData
}

function Save-Data {
    $data = @{ tabs = @() }
    foreach ($tabPage in $script:tabControl.TabPages) {
        $tabData = @{
            name = $tabPage.Text
            phrases = @()
        }
        $flowPanel = $tabPage.Tag
        foreach ($ctrl in $flowPanel.Controls) {
            if ($ctrl.Tag -and $ctrl.Tag -ne "addBtn") {
                $tabData.phrases += $ctrl.Tag
            }
        }
        # 反转顺序
        [array]::Reverse($tabData.phrases)
        $data.tabs += $tabData
    }
    $data | ConvertTo-Json -Depth 10 | Set-Content $script:ConfigFile -Encoding UTF8
}

function Load-Settings {
    if (Test-Path $script:SettingsFile) {
        try {
            $json = Get-Content $script:SettingsFile -Raw -Encoding UTF8
            return $json | ConvertFrom-Json
        } catch {
            return [pscustomobject]$script:DefaultSettings
        }
    }
    return [pscustomobject]$script:DefaultSettings
}

function Save-Settings {
    $settings = @{
        opacity = $script:mainForm.Opacity
        topMost = $script:mainForm.TopMost
        autoMinimize = $script:autoMinimize
        windowWidth = $script:mainForm.Width
        windowHeight = $script:mainForm.Height
    }
    $settings | ConvertTo-Json -Depth 5 | Set-Content $script:SettingsFile -Encoding UTF8
}

# ===== 创建圆角标签按钮 =====
function New-PhraseLabel($text) {
    # 计算文本宽度
    $g = [System.Drawing.Graphics]::FromImage([System.Drawing.Bitmap]::new(1,1))
    $font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $sizeF = $g.MeasureString($text, $font)
    $g.Dispose()
    [int]$textWidth = $sizeF.Width + 24
    [int]$panelWidth = $textWidth + 60
    [int]$lblWidth = $textWidth - 2
    [int]$editX = $panelWidth - 50
    [int]$deleteX = $panelWidth - 28

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Height = 32
    $panel.Tag = $text
    $panel.Margin = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
    $panel.Cursor = "Hand"
    $panel.BackColor = [System.Drawing.Color]::Transparent
    $panel.Width = $panelWidth

    # 背景面板
    $bgPanel = New-Object System.Windows.Forms.Panel
    $bgPanel.Dock = "Fill"
    $bgPanel.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 245)
    $bgPanel.Cursor = "Hand"
    $bgPanel.Name = "bgPanel"

    # 拖拽手柄
    $dragHandle = New-Object System.Windows.Forms.Label
    $dragHandle.Text = "⋮⋮"
    $dragHandle.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 8)
    $dragHandle.ForeColor = [System.Drawing.Color]::FromArgb(150, 170, 170)
    $dragHandle.Location = New-Object System.Drawing.Point(6, 6)
    $dragHandle.Size = New-Object System.Drawing.Size(16, 20)
    $dragHandle.Cursor = "SizeAll"
    $dragHandle.TextAlign = "MiddleCenter"
    $dragHandle.BackColor = [System.Drawing.Color]::Transparent
    $dragHandle.Name = "dragHandle"

    # 文本标签
    $lblText = New-Object System.Windows.Forms.Label
    $lblText.Text = $text
    $lblText.Font = $font
    $lblText.ForeColor = [System.Drawing.Color]::FromArgb(40, 60, 60)
    $lblText.Location = New-Object System.Drawing.Point(26, 4)
    $lblText.Size = New-Object System.Drawing.Size($lblWidth, 24)
    $lblText.TextAlign = "MiddleLeft"
    $lblText.Cursor = "Hand"
    $lblText.BackColor = [System.Drawing.Color]::Transparent
    $lblText.AutoEllipsis = $true
    $lblText.Name = "lblText"

    # 编辑按钮
    $btnEdit = New-Object System.Windows.Forms.Label
    $btnEdit.Text = "✏"
    $btnEdit.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
    $btnEdit.ForeColor = [System.Drawing.Color]::FromArgb(100, 140, 140)
    $btnEdit.Size = New-Object System.Drawing.Size(20, 20)
    $btnEdit.Location = New-Object System.Drawing.Point($editX, 6)
    $btnEdit.TextAlign = "MiddleCenter"
    $btnEdit.Cursor = "Hand"
    $btnEdit.BackColor = [System.Drawing.Color]::Transparent
    $btnEdit.Name = "btnEdit"

    # 删除按钮
    $btnDelete = New-Object System.Windows.Forms.Label
    $btnDelete.Text = "✕"
    $btnDelete.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $btnDelete.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 100)
    $btnDelete.Size = New-Object System.Drawing.Size(20, 20)
    $btnDelete.Location = New-Object System.Drawing.Point($deleteX, 6)
    $btnDelete.TextAlign = "MiddleCenter"
    $btnDelete.Cursor = "Hand"
    $btnDelete.BackColor = [System.Drawing.Color]::Transparent
    $btnDelete.Name = "btnDelete"

    $bgPanel.Controls.Add($btnDelete)
    $bgPanel.Controls.Add($btnEdit)
    $bgPanel.Controls.Add($lblText)
    $bgPanel.Controls.Add($dragHandle)
    $panel.Controls.Add($bgPanel)

    # ===== 事件：悬停效果（通过 $this.Parent 获取 bgPanel）=====
    $enterHandler = {
        $this.Parent.BackColor = [System.Drawing.Color]::FromArgb(220, 240, 240)
    }.GetNewClosure()
    $leaveHandler = {
        $this.Parent.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 245)
    }.GetNewClosure()

    $bgPanel.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(220, 240, 240)
    })
    $bgPanel.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(235, 245, 245)
    })
    $lblText.Add_MouseEnter($enterHandler)
    $lblText.Add_MouseLeave($leaveHandler)
    $dragHandle.Add_MouseEnter($enterHandler)
    $dragHandle.Add_MouseLeave($leaveHandler)

    # ===== 事件：点击文本 → 发送到目标窗口 =====
    $clickHandler = {
        $outerPanel = $this.Parent.Parent
        Send-TextToActiveWindow $outerPanel.Tag
    }.GetNewClosure()
    $lblText.Add_Click($clickHandler)
    $bgPanel.Add_Click($clickHandler)

    # ===== 事件：编辑 =====
    $btnEdit.Add_Click({
        $outerPanel = $this.Parent.Parent
        Edit-Phrase $outerPanel
    }.GetNewClosure())
    $btnEdit.Add_MouseEnter({
        $this.ForeColor = [System.Drawing.Color]::FromArgb(30, 144, 255)
    })
    $btnEdit.Add_MouseLeave({
        $this.ForeColor = [System.Drawing.Color]::FromArgb(100, 140, 140)
    })

    # ===== 事件：删除 =====
    $btnDelete.Add_Click({
        $outerPanel = $this.Parent.Parent
        $result = [System.Windows.Forms.MessageBox]::Show(
            "确定要删除这条短语吗？",
            "删除确认",
            "YesNo",
            "Question"
        )
        if ($result -eq "Yes") {
            $flowPanel = $outerPanel.Parent
            $flowPanel.Controls.Remove($outerPanel)
            $outerPanel.Dispose()
            Save-Data
        }
    }.GetNewClosure())
    $btnDelete.Add_MouseEnter({
        $this.ForeColor = [System.Drawing.Color]::Red
    })
    $btnDelete.Add_MouseLeave({
        $this.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 100)
    })

    # ===== 事件：拖拽排序 =====
    $dragHandle.Add_MouseDown({
        param($sender, $e)
        if ($e.Button -eq "Left") {
            $outerPanel = $sender.Parent.Parent
            $script:draggingItem = $outerPanel
            $script:dragStartPos = $e.Location
            $outerPanel.BackColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
        }
    }.GetNewClosure())

    $dragHandle.Add_MouseMove({
        param($sender, $e)
        if ($script:draggingItem -and $e.Button -eq "Left") {
            $dist = [Math]::Abs($e.X - $script:dragStartPos.X) + [Math]::Abs($e.Y - $script:dragStartPos.Y)
            if ($dist -gt 5) {
                $outerPanel = $sender.Parent.Parent
                $outerPanel.DoDragDrop($outerPanel, "Move")
            }
        }
    }.GetNewClosure())

    $dragHandle.Add_MouseUp({
        param($sender, $e)
        $outerPanel = $sender.Parent.Parent
        $script:draggingItem = $null
        $outerPanel.BackColor = [System.Drawing.Color]::Transparent
    }.GetNewClosure())

    return $panel
}

# ===== 发送文本到当前活动窗口 =====
function Send-TextToActiveWindow($text) {
    $ourHwnd = [Win32]::GetForegroundWindow()
    
    # 复制到剪贴板
    [System.Windows.Forms.Clipboard]::SetText($text)
    
    # 找到目标窗口：优先用记录的上一个活动窗口，否则从 Z-order 找
    $targetHwnd = [IntPtr]::Zero
    
    if ($script:lastActiveHwnd -ne [IntPtr]::Zero -and 
        [Win32]::IsWindowVisible($script:lastActiveHwnd) -and 
        -not [Win32]::IsIconic($script:lastActiveHwnd)) {
        $targetHwnd = $script:lastActiveHwnd
    }
    
    if ($targetHwnd -eq [IntPtr]::Zero) {
        # 从 Z-order 往下找第一个可见且非最小化、有标题的窗口
        $hwnd = [Win32]::GetWindow($ourHwnd, [Win32]::GW_HWNDNEXT)
        while ($hwnd -ne [IntPtr]::Zero) {
            if ([Win32]::IsWindowVisible($hwnd) -and 
                -not [Win32]::IsIconic($hwnd) -and 
                [Win32]::GetWindowTextLength($hwnd) -gt 0) {
                $targetHwnd = $hwnd
                break
            }
            $hwnd = [Win32]::GetWindow($hwnd, [Win32]::GW_HWNDNEXT)
        }
    }
    
    if ($targetHwnd -eq [IntPtr]::Zero) {
        return  # 找不到目标窗口，只复制到剪贴板
    }
    
    # 先隐藏我们的窗口（让目标窗口自然成为前台）
    if ($script:autoMinimize) {
        $script:mainForm.Visible = $false
    }
    
    # 强制设置目标窗口为前台（AttachThreadInput 绕过 Win10 限制）
    Force-SetForegroundWindow $targetHwnd | Out-Null
    
    # 等待焦点切换稳定
    Start-Sleep -Milliseconds 300
    
    # 发送 Ctrl+V 粘贴
    try {
        [System.Windows.Forms.SendKeys]::SendWait("^v")
    } catch {
        # 失败的话内容还在剪贴板，用户可手动粘贴
    }
    
    # 如果不自动隐藏，确保我们的窗口还在 TopMost
    if (-not $script:autoMinimize) {
        $script:mainForm.TopMost = $true
    }
}

# ===== 编辑短语 =====
function Edit-Phrase($panel) {
    $oldText = $panel.Tag
    $newText = [Microsoft.VisualBasic.Interaction]::InputBox(
        "编辑短语内容：",
        "编辑快捷用语",
        $oldText
    )
    
    if ([string]::IsNullOrWhiteSpace($newText)) { return }
    
    $newText = $newText.Trim()
    $panel.Tag = $newText
    
    # 更新显示的文本（通过 Name 查找）
    $bgPanel = $panel.Controls["bgPanel"]
    $lblText = $bgPanel.Controls["lblText"]
    $btnEdit = $bgPanel.Controls["btnEdit"]
    $btnDelete = $bgPanel.Controls["btnDelete"]
    
    if ($lblText) {
        $lblText.Text = $newText
        # 重新计算宽度
        $g = [System.Drawing.Graphics]::FromImage([System.Drawing.Bitmap]::new(1,1))
        $sizeF = $g.MeasureString($newText, $lblText.Font)
        $g.Dispose()
        [int]$newTextWidth = $sizeF.Width + 24
        [int]$newPanelWidth = $newTextWidth + 60
        [int]$newLblWidth = $newTextWidth - 2
        [int]$newEditX = $newPanelWidth - 50
        [int]$newDeleteX = $newPanelWidth - 28
        
        $lblText.Size = New-Object System.Drawing.Size($newLblWidth, 24)
        $panel.Width = $newPanelWidth
        
        # 更新按钮位置
        if ($btnEdit) { $btnEdit.Location = New-Object System.Drawing.Point($newEditX, 6) }
        if ($btnDelete) { $btnDelete.Location = New-Object System.Drawing.Point($newDeleteX, 6) }
    }
    
    Save-Data
}

# ===== 创建 Tab 页 =====
function New-TabPage($tabName, $phrases) {
    $tabPage = New-Object System.Windows.Forms.TabPage
    $tabPage.Text = $tabName
    $tabPage.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 252)
    
    # 主容器
    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Dock = "Fill"
    $mainPanel.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 252)
    
    # 底部添加栏
    $bottomPanel = New-Object System.Windows.Forms.Panel
    $bottomPanel.Dock = "Bottom"
    $bottomPanel.Height = 50
    $bottomPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 245, 245)
    $bottomPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
    
    $txtAdd = New-Object System.Windows.Forms.TextBox
    $txtAdd.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $txtAdd.Location = New-Object System.Drawing.Point(10, 12)
    $txtAdd.Size = New-Object System.Drawing.Size(300, 28)
    $txtAdd.Anchor = "Left, Top, Right"
    $txtAdd.Text = ""
    
    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "＋ 添加"
    $btnAdd.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)
    $btnAdd.Location = New-Object System.Drawing.Point(316, 10)
    $btnAdd.Size = New-Object System.Drawing.Size(70, 30)
    $btnAdd.Anchor = "Top, Right"
    $btnAdd.BackColor = [System.Drawing.Color]::FromArgb(30, 144, 255)
    $btnAdd.ForeColor = [System.Drawing.Color]::White
    $btnAdd.FlatStyle = "Flat"
    $btnAdd.FlatAppearance.BorderSize = 0
    $btnAdd.Cursor = "Hand"
    
    $bottomPanel.Controls.Add($btnAdd)
    $bottomPanel.Controls.Add($txtAdd)
    
    # 流式布局面板（标签云）
    $flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowPanel.Dock = "Fill"
    $flowPanel.AutoScroll = $true
    $flowPanel.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 252)
    $flowPanel.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 8)
    $flowPanel.WrapContents = $true
    $flowPanel.FlowDirection = "LeftToRight"
    $flowPanel.AllowDrop = $true
    
    # 拖拽放置处理
    $flowPanel.Add_DragEnter({
        param($sender, $e)
        if ($script:draggingItem) {
            $e.Effect = "Move"
        }
    })
    
    $flowPanel.Add_DragDrop({
        param($sender, $e)
        $draggedItem = $script:draggingItem
        if ($draggedItem -eq $null) { return }
        
        # 找到放下位置的控件
        $point = $sender.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
        $targetItem = $null
        
        foreach ($ctrl in $sender.Controls) {
            if ($ctrl -ne $draggedItem -and $ctrl.Bounds.Contains($point)) {
                $targetItem = $ctrl
                break
            }
        }
        
        if ($targetItem) {
            $targetIndex = $sender.Controls.GetChildIndex($targetItem)
            $sender.Controls.SetChildIndex($draggedItem, $targetIndex)
        }
        
        $draggedItem.BackColor = [System.Drawing.Color]::Transparent
        $script:draggingItem = $null
        Save-Data
    })
    
    $tabPage.Tag = $flowPanel
    
    # 加载短语
    foreach ($phrase in $phrases) {
        $label = New-PhraseLabel $phrase
        $flowPanel.Controls.Add($label)
    }
    
    # 添加按钮事件
    $addAction = {
        $text = $txtAdd.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return }
        $label = New-PhraseLabel $text
        $flowPanel.Controls.Add($label)
        $txtAdd.Text = ""
        $txtAdd.Focus()
        Save-Data
    }.GetNewClosure()
    $btnAdd.Add_Click($addAction)
    $txtAdd.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq "Enter") {
            $e.SuppressKeyPress = $true
            $text = $sender.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($text)) { return }
            $label = New-PhraseLabel $text
            $flowPanel.Controls.Add($label)
            $sender.Text = ""
            $sender.Focus()
            Save-Data
        }
    }.GetNewClosure())
    
    $mainPanel.Controls.Add($flowPanel)
    $mainPanel.Controls.Add($bottomPanel)
    $tabPage.Controls.Add($mainPanel)
    
    return $tabPage
}

# ===== 主窗口 =====
$script:mainForm = New-Object System.Windows.Forms.Form
$script:mainForm.Text = "快捷文本模板库"
$script:mainForm.StartPosition = "CenterScreen"
$script:mainForm.FormBorderStyle = "Sizable"
$script:mainForm.MinimumSize = New-Object System.Drawing.Size(300, 360)
$script:mainForm.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 252)

# 加载设置
$settings = Load-Settings
$script:mainForm.Opacity = $settings.opacity
$script:mainForm.TopMost = $settings.topMost
$script:mainForm.Width = $settings.windowWidth
$script:mainForm.Height = $settings.windowHeight
$script:autoMinimize = $settings.autoMinimize

# ===== 顶部工具栏 =====
$toolStrip = New-Object System.Windows.Forms.ToolStrip
$toolStrip.GripStyle = "Hidden"
$toolStrip.BackColor = [System.Drawing.Color]::FromArgb(235, 242, 242)
$toolStrip.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

# 置顶按钮
$btnTopMost = New-Object System.Windows.Forms.ToolStripButton
$btnTopMost.Text = "📌 置顶"
$btnTopMost.CheckOnClick = $true
$btnTopMost.Checked = $script:mainForm.TopMost
$btnTopMost.Add_Click({
    $script:mainForm.TopMost = $btnTopMost.Checked
    Save-Settings
})

# 透明度下拉
$ddlOpacity = New-Object System.Windows.Forms.ToolStripDropDownButton
$ddlOpacity.Text = "💧 透明度"
for ($i = 100; $i -ge 30; $i -= 10) {
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = "$i%"
    $item.Tag = $i / 100.0
    $item.Add_Click({
        $script:mainForm.Opacity = $this.Tag
        Save-Settings
    })
    $ddlOpacity.DropDownItems.Add($item) | Out-Null
}

# 自动最小化按钮
$btnAutoMin = New-Object System.Windows.Forms.ToolStripButton
$btnAutoMin.Text = "🕐 自动发送"
$btnAutoMin.CheckOnClick = $true
$btnAutoMin.Checked = $script:autoMinimize
$btnAutoMin.ToolTipText = "点击后自动最小化并粘贴到目标窗口"
$btnAutoMin.Add_Click({
    $script:autoMinimize = $btnAutoMin.Checked
    Save-Settings
})

# 添加Tab按钮
$btnAddTab = New-Object System.Windows.Forms.ToolStripButton
$btnAddTab.Text = "➕ 新页签"
$btnAddTab.Add_Click({ Add-NewTab })

# 重命名Tab按钮
$btnRenameTab = New-Object System.Windows.Forms.ToolStripButton
$btnRenameTab.Text = "✏️ 重命名"
$btnRenameTab.Add_Click({ Rename-CurrentTab })

# 删除Tab按钮
$btnDeleteTab = New-Object System.Windows.Forms.ToolStripButton
$btnDeleteTab.Text = "🗑️ 删页签"
$btnDeleteTab.Add_Click({ Delete-CurrentTab })

# 导入按钮
$btnImport = New-Object System.Windows.Forms.ToolStripButton
$btnImport.Text = "📥 导入"
$btnImport.Add_Click({ Import-Config })

# 导出按钮
$btnExport = New-Object System.Windows.Forms.ToolStripButton
$btnExport.Text = "📤 导出"
$btnExport.Add_Click({ Export-Config })

$toolStrip.Items.AddRange(@(
    $btnTopMost, $ddlOpacity, $btnAutoMin, [System.Windows.Forms.ToolStripSeparator]::new(),
    $btnAddTab, $btnRenameTab, $btnDeleteTab, [System.Windows.Forms.ToolStripSeparator]::new(),
    $btnImport, $btnExport
))

# ===== Tab 控件 =====
$script:tabControl = New-Object System.Windows.Forms.TabControl
$script:tabControl.Dock = "Fill"
$script:tabControl.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)
$script:tabControl.DrawMode = "OwnerDrawFixed"
$script:tabControl.ItemSize = New-Object System.Drawing.Size(75, 26)

# 自绘Tab标签
$script:tabControl.Add_DrawItem({
    param($sender, $e)
    $tabPage = $sender.TabPages[$e.Index]
    $isSelected = $e.State -band [System.Windows.Forms.DrawItemState]::Selected
    
    $backColor = [System.Drawing.Color]::FromArgb(235, 242, 242)
    $textColor = [System.Drawing.Color]::FromArgb(80, 100, 100)
    
    if ($isSelected) {
        $backColor = [System.Drawing.Color]::FromArgb(250, 252, 252)
        $textColor = [System.Drawing.Color]::FromArgb(20, 120, 120)
    }
    
    $brush = New-Object System.Drawing.SolidBrush($backColor)
    $e.Graphics.FillRectangle($brush, $e.Bounds)
    $brush.Dispose()
    
    if ($isSelected) {
        $lineBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 160, 160))
        $e.Graphics.FillRectangle($lineBrush, $e.Bounds.X, $e.Bounds.Bottom - 2, $e.Bounds.Width, 2)
        $lineBrush.Dispose()
    }
    
    $stringFormat = New-Object System.Drawing.StringFormat
    $stringFormat.Alignment = "Center"
    $stringFormat.LineAlignment = "Center"
    $textBrush = New-Object System.Drawing.SolidBrush($textColor)
    $rectF = [System.Drawing.RectangleF]$e.Bounds
    $e.Graphics.DrawString($tabPage.Text, $e.Font, $textBrush, $rectF, $stringFormat)
    $textBrush.Dispose()
    $stringFormat.Dispose()
})

# ===== 添加新 Tab =====
function Add-NewTab {
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("请输入新页签名称：", "新建页签", "新页签")
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    
    $tabPage = New-TabPage $name.Trim() @()
    $script:tabControl.TabPages.Add($tabPage)
    $script:tabControl.SelectedTab = $tabPage
    Save-Data
}

# ===== 重命名当前 Tab =====
function Rename-CurrentTab {
    if ($script:tabControl.TabCount -eq 0) { return }
    $currentTab = $script:tabControl.SelectedTab
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("请输入新名称：", "重命名页签", $currentTab.Text)
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    $currentTab.Text = $name.Trim()
    Save-Data
}

# ===== 删除当前 Tab =====
function Delete-CurrentTab {
    if ($script:tabControl.TabCount -le 1) {
        [System.Windows.Forms.MessageBox]::Show("至少保留一个页签！", "提示", "OK", "Warning")
        return
    }
    $currentTab = $script:tabControl.SelectedTab
    $result = [System.Windows.Forms.MessageBox]::Show(
        "确定要删除页签「$($currentTab.Text)」吗？`n该页签下的所有快捷用语也会被删除。",
        "删除确认",
        "YesNo",
        "Question"
    )
    if ($result -eq "Yes") {
        $script:tabControl.TabPages.Remove($currentTab)
        $currentTab.Dispose()
        Save-Data
    }
}

# ===== 导入配置 =====
function Import-Config {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "JSON 配置文件 (*.json)|*.json|所有文件 (*.*)|*.*"
    $dialog.Title = "导入配置文件"
    
    if ($dialog.ShowDialog() -eq "OK") {
        try {
            $json = Get-Content $dialog.FileName -Raw -Encoding UTF8
            $data = $json | ConvertFrom-Json
            
            if (-not $data.tabs -or $data.tabs.Count -eq 0) {
                throw "配置文件格式不正确"
            }
            
            $result = [System.Windows.Forms.MessageBox]::Show(
                "导入将替换当前所有内容，确定继续吗？`n（将加载 $($data.tabs.Count) 个页签）",
                "导入确认",
                "YesNo",
                "Question"
            )
            
            if ($result -eq "Yes") {
                $script:tabControl.TabPages.Clear()
                
                foreach ($tab in $data.tabs) {
                    $phrases = @()
                    if ($tab.phrases) {
                        foreach ($phrase in $tab.phrases) {
                            $phrases += $phrase
                        }
                    }
                    $tabPage = New-TabPage $tab.name $phrases
                    $script:tabControl.TabPages.Add($tabPage)
                }
                
                Save-Data
                [System.Windows.Forms.MessageBox]::Show("导入成功！", "提示", "OK", "Information")
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("导入失败：$($_.Exception.Message)", "错误", "OK", "Error")
        }
    }
    $dialog.Dispose()
}

# ===== 导出配置 =====
function Export-Config {
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "JSON 配置文件 (*.json)|*.json"
    $dialog.Title = "导出配置文件"
    $dialog.FileName = "quicktext_export.json"
    
    if ($dialog.ShowDialog() -eq "OK") {
        try {
            Copy-Item $script:ConfigFile $dialog.FileName -Force
            [System.Windows.Forms.MessageBox]::Show("导出成功！", "提示", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("导出失败：$($_.Exception.Message)", "错误", "OK", "Error")
        }
    }
    $dialog.Dispose()
}

# ===== 初始化加载数据 =====
$data = Load-Data
$firstRun = -not (Test-Path $script:ConfigFile)
foreach ($tab in $data.tabs) {
    $phrases = @()
    if ($tab.phrases) {
        foreach ($phrase in $tab.phrases) {
            $phrases += $phrase
        }
    }
    $tabPage = New-TabPage $tab.name $phrases
    $script:tabControl.TabPages.Add($tabPage)
}

if ($firstRun) {
    Save-Data
}

# ===== 布局 =====
# 先加 TabControl（Fill），再加 ToolStrip（Top），这样 ToolStrip 会占顶部
$script:mainForm.Controls.Add($script:tabControl)
$script:mainForm.Controls.Add($toolStrip)

$script:mainForm.Add_FormClosing({
    Save-Settings
})

# ===== 定时器：跟踪上一个活动窗口 =====
$trackTimer = New-Object System.Windows.Forms.Timer
$trackTimer.Interval = 200
$trackTimer.Add_Tick({
    $currentHwnd = [Win32]::GetForegroundWindow()
    $ourHwnd = $script:mainForm.Handle
    if ($currentHwnd -ne [IntPtr]::Zero -and $currentHwnd -ne $ourHwnd) {
        # 只记录可见且非最小化的窗口
        if ([Win32]::IsWindowVisible($currentHwnd) -and -not [Win32]::IsIconic($currentHwnd)) {
            $script:lastActiveHwnd = $currentHwnd
        }
    }
})
$trackTimer.Start()

# ===== 启动 =====
$script:mainForm.ShowDialog() | Out-Null
