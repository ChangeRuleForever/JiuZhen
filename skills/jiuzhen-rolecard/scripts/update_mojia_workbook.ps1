$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Get-WorksheetByAnyName {
    param(
        [object]$Workbook,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        try {
            return $Workbook.Worksheets.Item($name)
        } catch {
        }
    }

    throw "Worksheet not found. Tried: $($Names -join ', ')"
}

function Set-MergedValue {
    param(
        [object]$Worksheet,
        [string]$Address,
        [object]$Value
    )

    $range = $Worksheet.Range($Address)
    if ($range.MergeCells) {
        $range.UnMerge() | Out-Null
    }
    $range.Merge() | Out-Null
    $range.Value2 = $Value
    Release-ComObject $range
}

function Set-CellValue {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$Column,
        [object]$Value
    )

    $cell = $Worksheet.Cells.Item($Row, $Column)
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [decimal] -or $Value -is [double] -or $Value -is [float]) {
        $cell.Value2 = [double]$Value
    } else {
        $cell.Value2 = [string]$Value
    }
    Release-ComObject $cell
}

function Write-TableRows {
    param(
        [object]$Worksheet,
        [int]$StartColumn,
        [int]$StartRow,
        [object[]]$Rows
    )

    $rowIndex = $StartRow
    foreach ($rowValues in $Rows) {
        if ($rowValues -is [string] -or $rowValues -isnot [System.Collections.IEnumerable]) {
            $cells = @($rowValues)
        } else {
            $cells = @($rowValues)
        }

        for ($offset = 0; $offset -lt $cells.Count; $offset++) {
            $cell = $Worksheet.Cells.Item([int]$rowIndex, [int]($StartColumn + $offset))
            $value = $cells[$offset]
            if ($value -is [int] -or $value -is [long] -or $value -is [decimal] -or $value -is [double] -or $value -is [float]) {
                $cell.Value2 = [double]$value
            } else {
                $cell.Value2 = [string]$value
            }
            Release-ComObject $cell
        }
        $rowIndex++
    }
}

function Get-VbaCodeBody {
    param([string]$Path)

    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $startIndex = $content.IndexOf("Option Explicit")
    if ($startIndex -lt 0) {
        throw "Option Explicit not found in $Path"
    }

    $body = $content.Substring($startIndex)
    $bodyLines = $body -split "`r?`n" | Where-Object { $_ -notmatch '^Attribute VB_' }
    return ($bodyLines -join "`r`n")
}

function Set-CodeModule {
    param(
        [object]$VBComponent,
        [string]$CodeBody
    )

    $codeModule = $VBComponent.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    $codeModule.AddFromString($CodeBody)
    Release-ComObject $codeModule
}

$skillRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\\..\\.."))
$workbookRoot = Join-Path $repoRoot "workbooks"
$templateRoot = Join-Path $workbookRoot "templates"
$workbookPath = Join-Path $templateRoot "新版玖镇角色卡_墨家.xlsm"
$workbookPath = [System.IO.Path]::GetFullPath($workbookPath)
$vbaRoot = Join-Path $skillRoot "references\\vba"
$vbaRoot = [System.IO.Path]::GetFullPath($vbaRoot)

$commonItems = @(
    @("等离子切割刀", 1, "通用墨器", "伤害1d5+1d3。"),
    @("单人飞行器", 3, "通用墨器", "如其名。"),
    @("手雷", 1, "通用墨器", "伤害2d10。"),
    @("无线通讯器", 2, "通用墨器", "可在无手机信号的时候沟通。"),
    @("简易呼吸器", 2, "通用墨器", "可以视为氧气管。"),
    @("抓钩", 3, "通用墨器", "如其名。"),
    @("生命探测仪", 5, "通用墨器", "如其名，50米。"),
    @("能量屏障", 2, "通用墨器", "20点一次性护盾。"),
    @("黑眼摄像头", 1, "通用墨器", "侦查。"),
    @("遥控小车", 1, "通用墨器", "侦查。")
)

$machineItems = @(
    @("高等合金小刀", 3, "机巧墨器", "伤害2d8。"),
    @("外骨骼紧身衣", 8, "机巧墨器", "使自身某两个属性+10。"),
    @("防身激光手枪", 4, "机巧墨器", "单发5点伤害，有幸运*100%的概率造成双倍伤害，20发弹匣，无备弹，不会损坏。"),
    @("佛怒唐莲", 1, "机巧墨器", "50无法闪避固定物理伤害，仅可以携带一个。"),
    @("臂装弹射器", 5, "机巧墨器", "可装入小型手雷，让手雷必中；发射后需一次次要行动装填。"),
    @("特种手雷", 3, "机巧墨器", "每枚3格；可选直伤、火伤、EMP、雷电任一或两种属性。"),
    @("电磁脉冲手枪", 3, "机巧墨器", "单发，1d8+2。"),
    @("等离子电磁狙击枪", 10, "机巧墨器", "单伤20；幸运*100%概率爆头双倍；五发子弹，五发备弹。"),
    @("脉冲步枪", 8, "机巧墨器", "单伤1d10，可三连发；3弹匣，30备弹。"),
    @("荆棘（手枪）", 6, "机巧墨器", "单伤2d8；附带三回合2d2毒伤。"),
    @("流明（手枪）", 6, "机巧墨器", "单伤2d6；命中可恢复自身或队友1d5生命。"),
    @("玫瑰（手枪）", 5, "机巧墨器", "单伤1d6+1d4；20%概率造成双倍伤害。"),
    @("AWM（静默弹药）", 12, "机巧墨器", "单伤2d15；击中附加1回合失明。"),
    @("晋景公榴弹发射器", 12, "机巧墨器", "单发装填，可发射药水弹药，直伤1d15。"),
    @("A137 追踪导弹单兵发射装置", 25, "机巧墨器", "单发装填；直伤1d30，爆炸4d10，带追踪效果。")
)

$deployItems = @(
    @("禁锢陷阱", 3, "布设墨器", "踩中的敌人承受20点禁锢。"),
    @("探测陷阱", 3, "布设墨器", "持续探测周围50米的各种物体。"),
    @("爆炸陷阱", 3, "布设墨器", "地雷，伤害4d10。"),
    @("震慑信标", 5, "布设墨器", "每次敌方攻击有一次30判定降低一个成功等级，4次后报废。"),
    @("加速信标", 5, "布设墨器", "己方敏捷闪避+5，可叠加。"),
    @("自动炮塔", 10, "布设墨器", "自动警戒发射子弹，伤害1d10，血量1。"),
    @("雷电信标", 5, "布设墨器", "接近敌人时造成伤害和麻痹/击晕/短路。"),
    @("EMP 信标", 5, "布设墨器", "瘫痪一定范围内的电子设备，持续1小时。"),
    @("治疗信标", 5, "布设墨器", "自动治疗范围内友方，3次治疗，1d5。"),
    @("治疗机器人", 10, "布设墨器", "消耗1小时回复1d10。"),
    @("特斯拉力场护盾", 3, "布设墨器", "10点生命，战后修理，可保护全队。"),
    @("等离子能量炮指示器", 50, "布设墨器", "仅可在墨家内部制作，不可移动，仅可使用一次。")
)

$specialItems = @(
    @("地脉测绘占位槽", 10, "职业特殊", "布设师25格特殊成长发动后，直到次日整备前固定占用10格工巧格。"),
    @("不稳定超级狙击枪", 25, "职业特殊", "机巧师40格究极技能；10+1d40真实伤害，无视护盾与护甲。"),
    @("不稳定超级狙击枪子弹", 2, "职业特殊", "每发占2格，供不稳定超级狙击枪装填使用。")
)

$pluginItems = @(
    @("属性插件", 5, "选一属性+20。"),
    @("技能插件", 5, "获得一个50的技能。"),
    @("战斗能力拓展插槽", 5, "获得一种新的战斗能力。"),
    @("火焰附魔插件", 5, "攻击附带1d8火焰伤害。"),
    @("雷电附魔插件", 5, "攻击附带1d8电流伤害。"),
    @("自动瞄准插件", 5, "魁人攻击无法闪避。"),
    @("飞行插件", 5, "魁人可飞行且移动+10。"),
    @("特种护甲插件", 5, "魁人护甲+10。"),
    @("机械嘲讽插件", 5, "敌人必须优先攻击该魁人。"),
    @("魁人外表自定义插件", 5, "自定义魁人外貌。"),
    @("网络插件", 5, "魁人可以联网。"),
    @("远程通讯插件", 5, "10公里以上也可命令魁人。"),
    @("槽位插件", 5, "槽位+3。")
)

$weaponItems = @(
    @("双持刀", 5, "2d8伤害，50%概率连打。"),
    @("长柄锤", 5, "1d20伤害，10%眩晕。"),
    @("双斧", 5, "2d6+5伤害。"),
    @("钩锁", 5, "1d6伤害，可将敌人拽到身前。"),
    @("偃月", 10, "2d15伤害，并击倒对方，50%概率连击。")
)

$themeReferenceRows = @(
    @("家族主题", "特点", "介绍", "代表颜色"),
    @("张家", "权威、强盛", "是玖镇最有话语权的家族，主要掌握道法、雷电", "雷霆紫"),
    @("岳家", "豪放、崇尚武力", "玖镇武功最强的家族，纯靠功夫打出一片天地，擅长岳家枪、棍、刀等，尤其喜欢美酒", "赤铜色"),
    @("林家", "神秘、隐世", "玖镇最看重风水的家族，与易家交好，擅长隐藏和出其不意的技巧", "青绿色"),
    @("易家", "睿智、避战", "能够在一定程度上预知未来，擅长预测并回避敌方攻击", "蓝色"),
    @("曹家", "中立、富有", "玖镇中的经商大家，擅长赚取经费并把资源转化为战斗力", "青玉色"),
    @("墨家", "神秘、技术强、不与外界沟通", "玖镇中较为神秘的家族，擅长各种机关道具，传言甚至在开发高达", "钢青灰"),
    @("药家", "崇尚自然、妙手回春", "制药家族，几乎所有草药和毒药均出自药家，擅长恢复和用毒", "绿色"),
    @("巫家", "全能、神秘、中立", "玖镇中最中立的家族，经营醉红尘，家族中几乎全部是女性", "粉色"),
    @("亓家", "诡异、强大、没有道德伦理", "被排挤出玖镇的家族，经常使用活人做实验，擅长召唤小鬼攻击", "血红色"),
    @("乐正家", "风雅、共鸣、控场支援", "擅长以乐器和灵魂音符进行控制、治疗、增速与灵魂打击", "雅乐金")
)

$weaponNames = $weaponItems | ForEach-Object { $_[0] }

$excel = $null
$workbook = $null
$wsSettings = $null
$wsMojia = $null
$wsKui = $null
$wsCatalog = $null
$vbProject = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($workbookPath)

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path ([System.IO.Path]::GetDirectoryName($workbookPath)) ("新版玖镇角色卡_墨家_backup_{0}.xlsm" -f $timestamp)
    $workbook.SaveCopyAs($backupPath)

    $wsSettings = Get-WorksheetByAnyName $workbook @("主题设置", "基本信息")
    $wsMojia = Get-WorksheetByAnyName $workbook @("墨家特性-背包槽位", "【墨家】工巧格与魁人")
    $wsKui = Get-WorksheetByAnyName $workbook @("魁人面板")
    $wsCatalog = Get-WorksheetByAnyName $workbook @("物品库", "【墨家】物品库")

    $legacyStrength = $(if (-not [string]::IsNullOrWhiteSpace([string]$wsMojia.Range("P10").Value2)) { $wsMojia.Range("P10").Value2 } else { $wsKui.Range("D4").Value2 })
    $legacyDexterity = $(if (-not [string]::IsNullOrWhiteSpace([string]$wsMojia.Range("P11").Value2)) { $wsMojia.Range("P11").Value2 } else { $wsKui.Range("B6").Value2 })
    $legacyConstitution = $(if (-not [string]::IsNullOrWhiteSpace([string]$wsMojia.Range("P12").Value2)) { $wsMojia.Range("P12").Value2 } else { $wsKui.Range("B4").Value2 })
    $legacyBattleAbility = $(if (-not [string]::IsNullOrWhiteSpace([string]$wsMojia.Range("P7").Value2)) { $wsMojia.Range("P7").Value2 } else { $wsKui.Range("B9").Value2 })
    $legacyDodge = $(if (-not [string]::IsNullOrWhiteSpace([string]$wsMojia.Range("P13").Value2)) { $wsMojia.Range("P13").Value2 } else { $wsKui.Range("B8").Value2 })
    $legacyPlugins = @()
    $currentPluginRange = @()
    for ($row = 5; $row -le 14; $row++) {
        $currentPluginRange += [string]$wsMojia.Cells.Item($row, 21).Value2
    }
    if (($currentPluginRange | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
        $legacyPlugins = $currentPluginRange
    } else {
        for ($row = 4; $row -le 13; $row++) {
            $legacyPlugins += [string]$wsKui.Cells.Item($row, 7).Value2
        }
    }

    $legacyWeapons = New-Object System.Collections.Generic.List[string]
    $currentWeapons = @()
    for ($row = 17; $row -le 19; $row++) {
        $currentWeapons += [string]$wsMojia.Cells.Item($row, 16).Value2
    }
    if (($currentWeapons | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
        foreach ($weapon in $currentWeapons) {
            if (-not [string]::IsNullOrWhiteSpace($weapon) -and -not $legacyWeapons.Contains($weapon)) {
                $legacyWeapons.Add($weapon)
            }
        }
    } else {
        for ($row = 31; $row -le 330; $row++) {
            $cell = $wsKui.Cells.Item($row, 2)
            try {
                $text = [string]$cell.Text
                if ([string]::IsNullOrWhiteSpace($text)) { continue }
                if ($cell.MergeCells -and $cell.MergeArea.Row -ne $row) { continue }
                if ($weaponNames -contains $text -and -not $legacyWeapons.Contains($text)) {
                    $legacyWeapons.Add($text)
                }
            } finally {
                Release-ComObject $cell
            }
        }
    }

    if ($wsSettings.Name -ne "基本信息") {
        $wsSettings.Name = "基本信息"
    }
    if ($wsMojia.Name -ne "【墨家】工巧格与魁人") {
        $wsMojia.Name = "【墨家】工巧格与魁人"
    }
    if ($wsCatalog.Name -ne "【墨家】物品库") {
        $wsCatalog.Name = "【墨家】物品库"
    }

    if ($wsSettings.Index -ne 1) {
        $wsSettings.Move($workbook.Worksheets.Item(1))
    }

    $wsSettings.Range("A1").Value2 = "基本信息"
    $wsSettings.Range("B2").Value2 = "墨家"
    Write-TableRows $wsSettings 1 5 $themeReferenceRows
    $themeValidation = $wsSettings.Range("B2").Validation
    $themeValidation.Delete()
    $themeValidation.Add(3, 1, 1, '=$A$6:$A$15')
    $themeValidation.IgnoreBlank = $true
    $themeValidation.InCellDropdown = $true
    Release-ComObject $themeValidation
    $wsSettings.Range("A16").Value2 = "说明"
    $wsSettings.Range("B16").Value2 = "切换当前主题后会自动刷新配色，并按需显示墨家专属页。"
    $wsSettings.Range("A17").Value2 = "墨家专属页"
    $wsSettings.Range("B17").Value2 = "仅当当前主题为""墨家""时显示【墨家】工巧格与魁人与【墨家】物品库。"

    $wsKui.Visible = 2

    $growthRows = @()
    $growthStartColumn = 11
    if ([string]$wsMojia.Range("L3").Text -eq "判定等级") {
        $growthStartColumn = 12
    }
    for ($row = 4; $row -le 13; $row++) {
        $growthRows += ,@(
            $wsMojia.Cells.Item($row, $growthStartColumn).Value2,
            $wsMojia.Cells.Item($row, $growthStartColumn + 1).Value2,
            $wsMojia.Cells.Item($row, $growthStartColumn + 2).Value2
        )
    }

    $catalogVisibleRange = $wsCatalog.Range("A1:AS80")
    $catalogVisibleRange.UnMerge() | Out-Null
    $catalogVisibleRange.ClearContents()
    Release-ComObject $catalogVisibleRange

    Set-MergedValue $wsCatalog "A1:AC1" "【墨家】物品库"
    Set-MergedValue $wsCatalog "A3:D3" "通用墨器"
    Set-MergedValue $wsCatalog "F3:I3" "机巧墨器"
    Set-MergedValue $wsCatalog "K3:N3" "布设墨器"
    Set-MergedValue $wsCatalog "P3:S3" "魁人模组"
    Set-MergedValue $wsCatalog "U3:X3" "魁人武器"
    Set-MergedValue $wsCatalog "Z3:AC3" "职业特殊"
    Set-MergedValue $wsCatalog "A22:AC22" "附录说明"
    Set-MergedValue $wsCatalog "A23:AC23" "1. 所有装备、插件、武器与挂载均占用工巧格；魁人本体不占格。"
    Set-MergedValue $wsCatalog "A24:AC24" "2. 每个魁人模组固定占用5格工巧格；当工巧格达到20/30/40/50时，模组上限各+1。"
    Set-MergedValue $wsCatalog "A25:AC25" "3. 枪械与弹药沿用原稿口径：每五发子弹视为1格；特种手雷为每枚3格。"
    Set-MergedValue $wsCatalog "A26:AC26" "4. 20格转职后，使用对应职业特性墨器时可获得额外加成。"

    Write-TableRows $wsCatalog 1 4 (, @("墨器", "占工巧格", "分类", "说明"))
    Write-TableRows $wsCatalog 6 4 (, @("墨器", "占工巧格", "分类", "说明"))
    Write-TableRows $wsCatalog 11 4 (, @("墨器", "占工巧格", "分类", "说明"))
    Write-TableRows $wsCatalog 16 4 (, @("模组", "占工巧格", "说明", "备注"))
    Write-TableRows $wsCatalog 21 4 (, @("武器", "占工巧格", "说明", "备注"))
    Write-TableRows $wsCatalog 26 4 (, @("墨器", "占工巧格", "适用职业", "说明"))

    $pluginVisibleRows = @()
    foreach ($item in $pluginItems) {
        $pluginVisibleRows += ,@($item[0], $item[1], $item[2], "每个模组固定占5工巧格")
    }

    $weaponVisibleRows = @()
    foreach ($item in $weaponItems) {
        $weaponVisibleRows += ,@($item[0], $item[1], $item[2], "")
    }

    Write-TableRows $wsCatalog 1 5 $commonItems
    Write-TableRows $wsCatalog 6 5 $machineItems
    Write-TableRows $wsCatalog 11 5 $deployItems
    Write-TableRows $wsCatalog 16 5 $pluginVisibleRows
    Write-TableRows $wsCatalog 21 5 $weaponVisibleRows
    Write-TableRows $wsCatalog 26 5 $specialItems

    $helperPersonal = @()
    $helperPersonal += ,@("空", 0, "移除选择", "用于清空本人背包配置。")
    $helperPersonal += $commonItems
    $helperPersonal += $machineItems
    $helperPersonal += $deployItems
    $helperPersonal += $specialItems

    $helperPlugin = @()
    $helperPlugin += ,@("空", 0, "用于取出当前魁人模组。")
    foreach ($item in $pluginItems) {
        $helperPlugin += ,@($item[0], $item[1], $item[2])
    }

    $helperWeapon = @()
    $helperWeapon += ,@("空", 0, "用于卸下当前魁人武器。")
    foreach ($item in $weaponItems) {
        $helperWeapon += ,@($item[0], $item[1], $item[2])
    }

    $helperRender = @()
    foreach ($item in $helperPersonal) {
        if ($item[0] -eq "空") { continue }
        $helperRender += ,@($item[0], $item[1], $item[2])
    }
    foreach ($item in $pluginItems) {
        $helperRender += ,@(("【魁人模组】" + $item[0]), $item[1], "魁人模组")
    }
    foreach ($item in $weaponItems) {
        $helperRender += ,@(("【魁人武器】" + $item[0]), $item[1], "魁人武器")
    }

    Write-TableRows $wsCatalog 31 4 (, @("本人墨器", "工巧格", "分类", "说明"))
    Write-TableRows $wsCatalog 31 5 $helperPersonal

    Write-TableRows $wsCatalog 36 4 (, @("魁人模组", "工巧格", "说明"))
    Write-TableRows $wsCatalog 36 5 $helperPlugin

    Write-TableRows $wsCatalog 40 4 (, @("魁人武器", "工巧格", "说明"))
    Write-TableRows $wsCatalog 40 5 $helperWeapon

    Write-TableRows $wsCatalog 44 4 (, @("渲染墨器", "工巧格", "分类"))
    Write-TableRows $wsCatalog 44 5 $helperRender

    Write-TableRows $wsCatalog 48 4 @(
        @("成长等级"),
        @("大成功"),
        @("极难"),
        @("困难"),
        @("普通"),
        @("失败")
    )

    Write-TableRows $wsCatalog 49 4 @(
        @("战斗能力"),
        @("近战"),
        @("远程"),
        @("投掷")
    )

    $wsCatalog.Columns("AD:AZ").Hidden = $true

    $mojiaHeaderRange = $wsMojia.Range("H1:W30")
    $mojiaHeaderRange.UnMerge() | Out-Null
    $mojiaHeaderRange.ClearContents()
    Release-ComObject $mojiaHeaderRange

    Set-MergedValue $wsMojia "A1:E1" "【墨家】工巧格与魁人"
    Set-MergedValue $wsMojia "H1:I1" "工巧格汇总"
    Set-MergedValue $wsMojia "K1:M1" "成长记录"
    Set-MergedValue $wsMojia "O1:W1" "魁人模组与挂载"
    Set-MergedValue $wsMojia "A3:E3" "本人背包配置"
    Set-MergedValue $wsMojia "O3:R3" "魁人基础"
    Set-MergedValue $wsMojia "T3:W3" "魁人模组"
    Set-MergedValue $wsMojia "O15:R15" "魁人武器"
    Set-MergedValue $wsMojia "O21:W21" "工巧阶段提要"
    Write-TableRows $wsMojia 8 3 @(
        @("基础工巧格", ""),
        @("成长额外工巧格", ""),
        @("当前总工巧格", ""),
        @("本人占用工巧格", ""),
        @("魁人占用工巧格", ""),
        @("当前剩余工巧格", "")
    )

    Write-TableRows $wsMojia 8 9 @(
        @("成长触发", "学习、拆解、修缮、调试、制造或实战验证，经KP认可后进行成长检定。"),
        @("大成功", "+5 格"),
        @("极难", "+3+1d2 格"),
        @("困难", "+2+1d3 格"),
        @("普通", "+1+1d4 格"),
        @("失败", "+1 格"),
        @("翻倍说明", "25格特殊成长先按成功等级结算，再将最终结果翻倍。"),
        @("布设代价", "地脉测绘发动后，直到次日整备前固定占用10格工巧格。")
    )

    Write-TableRows $wsMojia 11 3 (, @("判定等级", "实际增长", "备注"))
    Write-TableRows $wsMojia 11 4 $growthRows

    Write-TableRows $wsMojia 15 4 @(
        @("移动力", "同主人", "属性总和", ""),
        @("护甲", "", "建议总和", 200),
        @("生命值", 20, "模组上限", ""),
        @("战斗能力", "", "已选模组", ""),
        @("斗殴", 60, "剩余模组", ""),
        @("基础伤害", "2d6", "单模组占格", ""),
        @("力量", $legacyStrength, "模组占格", ""),
        @("敏捷", $legacyDexterity, "武器占格", ""),
        @("体质", $legacyConstitution, "魁人总占格", ""),
        @("基础闪避", $(if ([string]::IsNullOrWhiteSpace([string]$legacyDodge)) { "基础值" } else { $legacyDodge }), "", "")
    )

    Write-TableRows $wsMojia 20 4 (, @("序号", "模组", "占格", "说明"))
    for ($idx = 0; $idx -lt 10; $idx++) {
        $targetRow = 5 + $idx
        Set-CellValue $wsMojia $targetRow 20 ($idx + 1)
        if ($idx -lt $legacyPlugins.Count -and -not [string]::IsNullOrWhiteSpace($legacyPlugins[$idx])) {
            Set-CellValue $wsMojia $targetRow 21 $legacyPlugins[$idx]
        }
    }

    Write-TableRows $wsMojia 15 16 (, @("序号", "武器", "占格", "说明"))
    for ($idx = 0; $idx -lt 3; $idx++) {
        $targetRow = 17 + $idx
        Set-CellValue $wsMojia $targetRow 15 ($idx + 1)
        if ($idx -lt $legacyWeapons.Count) {
            Set-CellValue $wsMojia $targetRow 16 $legacyWeapons[$idx]
        }
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$legacyBattleAbility)) {
        $wsMojia.Range("P7").Value2 = $legacyBattleAbility
    }

    Set-MergedValue $wsMojia "Q13:R13" "每个魁人模组固定占5工巧格；若剩余模组为0或剩余工巧格不足5，则无法继续装载。"

    Write-TableRows $wsMojia 15 22 @(
        @("工巧格", "节点", "效果摘要", "", "", "", "", ""),
        @("15", "基础", "基础15工巧格，初始可携带1名魁人。", "", "", "", "", ""),
        @("20", "转职", "在机巧师、偃师、布设师三条路线中选择一条；使用对应特性墨器获得额外加成。", "", "", "", "", ""),
        @("25", "特殊成长", "每天1次声明职业特殊成长；机巧师掉上限、偃师停用魁人、布设师固定占10格，成长收益翻倍。", "", "", "", "", ""),
        @("30", "核心技能", "机巧师过载校爆；偃师同调代行；布设师交叉火网。", "", "", "", "", ""),
        @("35", "职业成熟", "机巧师百工并联；偃师入偃同躯；布设师全相布置。", "", "", "", "", ""),
        @("40", "究极技能", "机巧师不稳定超级狙击枪；偃师高达形态待补；布设师炮台合体。", "", "", "", "", "")
    )

    $wsMojia.Columns("J").ColumnWidth = 3
    $wsMojia.Columns("K").ColumnWidth = 12
    $wsMojia.Columns("L").ColumnWidth = 10
    $wsMojia.Columns("M").ColumnWidth = 18
    $wsMojia.Columns("O").ColumnWidth = 10
    $wsMojia.Columns("P").ColumnWidth = 14
    $wsMojia.Columns("Q").ColumnWidth = 10
    $wsMojia.Columns("R").ColumnWidth = 16
    $wsMojia.Columns("S").ColumnWidth = 3
    $wsMojia.Columns("T").ColumnWidth = 6
    $wsMojia.Columns("U").ColumnWidth = 16
    $wsMojia.Columns("V").ColumnWidth = 8
    $wsMojia.Columns("W").ColumnWidth = 26

    $wsCatalog.Columns("A").ColumnWidth = 18
    $wsCatalog.Columns("B").ColumnWidth = 10
    $wsCatalog.Columns("C").ColumnWidth = 12
    $wsCatalog.Columns("D").ColumnWidth = 30
    $wsCatalog.Columns("E").ColumnWidth = 3
    $wsCatalog.Columns("F").ColumnWidth = 18
    $wsCatalog.Columns("G").ColumnWidth = 10
    $wsCatalog.Columns("H").ColumnWidth = 12
    $wsCatalog.Columns("I").ColumnWidth = 32
    $wsCatalog.Columns("J").ColumnWidth = 3
    $wsCatalog.Columns("K").ColumnWidth = 18
    $wsCatalog.Columns("L").ColumnWidth = 10
    $wsCatalog.Columns("M").ColumnWidth = 12
    $wsCatalog.Columns("N").ColumnWidth = 32
    $wsCatalog.Columns("O").ColumnWidth = 3
    $wsCatalog.Columns("P").ColumnWidth = 18
    $wsCatalog.Columns("Q").ColumnWidth = 10
    $wsCatalog.Columns("R").ColumnWidth = 24
    $wsCatalog.Columns("S").ColumnWidth = 14
    $wsCatalog.Columns("T").ColumnWidth = 3
    $wsCatalog.Columns("U").ColumnWidth = 18
    $wsCatalog.Columns("V").ColumnWidth = 10
    $wsCatalog.Columns("W").ColumnWidth = 24
    $wsCatalog.Columns("X").ColumnWidth = 12
    $wsCatalog.Columns("Y").ColumnWidth = 3
    $wsCatalog.Columns("Z").ColumnWidth = 20
    $wsCatalog.Columns("AA").ColumnWidth = 10
    $wsCatalog.Columns("AB").ColumnWidth = 12
    $wsCatalog.Columns("AC").ColumnWidth = 34

    try {
        $refreshShape = $wsMojia.Shapes.Item("btnRefreshMojiaInventory")
        $refreshShape.Top = $wsMojia.Range("E3").Top
        $refreshShape.Left = $wsMojia.Range("E3").Left
        $refreshShape.Width = $wsMojia.Range("E3:F4").Width
        $refreshShape.Height = $wsMojia.Range("E3:F4").Height
        Release-ComObject $refreshShape
    } catch {
    }

    $vbProject = $workbook.VBProject
    Set-CodeModule $vbProject.VBComponents.Item("ThisWorkbook") (Get-VbaCodeBody (Join-Path $vbaRoot "ThisWorkbook.cls"))
    Set-CodeModule $vbProject.VBComponents.Item("Sheet3") (Get-VbaCodeBody (Join-Path $vbaRoot "Sheet3.cls"))
    Set-CodeModule $vbProject.VBComponents.Item("JiuzhenMojiaMacros") (Get-VbaCodeBody (Join-Path $vbaRoot "JiuzhenMojiaMacros.updated.bas"))

    $excel.Run("'" + $workbook.Name + "'!ApplyWorkbookTheme")
    $excel.CalculateFull()

    $workbook.Save()
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($true)
    }

    if ($null -ne $excel) {
        $excel.EnableEvents = $true
        $excel.ScreenUpdating = $true
        $excel.Quit()
    }

    Release-ComObject $vbProject
    Release-ComObject $wsCatalog
    Release-ComObject $wsKui
    Release-ComObject $wsMojia
    Release-ComObject $wsSettings
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
