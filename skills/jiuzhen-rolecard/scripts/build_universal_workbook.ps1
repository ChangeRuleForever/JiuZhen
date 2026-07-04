$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Get-WorksheetByName {
    param(
        [object]$Workbook,
        [string]$Name
    )

    try {
        return $Workbook.Worksheets.Item($Name)
    } catch {
        return $null
    }
}

function Remove-WorksheetIfExists {
    param(
        [object]$Workbook,
        [string]$Name
    )

    $ws = Get-WorksheetByName $Workbook $Name
    if ($null -ne $ws) {
        $ws.Delete()
        Release-ComObject $ws
    }
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
        $cells = @($rowValues)
        for ($offset = 0; $offset -lt $cells.Count; $offset++) {
            Set-CellValue $Worksheet $rowIndex ($StartColumn + $offset) $cells[$offset]
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

function Get-OrCreateStandardModule {
    param(
        [object]$VBProject,
        [string]$ModuleName
    )

    try {
        return $VBProject.VBComponents.Item($ModuleName)
    } catch {
        $module = $VBProject.VBComponents.Add(1)
        $module.Name = $ModuleName
        return $module
    }
}

function Parse-EntryText {
    param([string]$Text)

    $clean = [string]$Text
    $clean = $clean -replace '\s+', ' '
    $clean = $clean.Trim()

    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $null
    }

    $firstSpace = $clean.IndexOf(' ')
    if ($firstSpace -lt 0) {
        return [PSCustomObject]@{
            Name = $clean
            Detail = ""
        }
    }

    return [PSCustomObject]@{
        Name = $clean.Substring(0, $firstSpace).Trim()
        Detail = $clean.Substring($firstSpace + 1).Trim()
    }
}

function Add-HerbEntriesFromRange {
    param(
        [ref]$Rows,
        [object]$Worksheet,
        [string]$Stage,
        [string]$Category,
        [int]$Column,
        [int[]]$SourceRows
    )

    foreach ($sourceRow in $SourceRows) {
        $text = [string]$Worksheet.Cells.Item($sourceRow, $Column).Text
        $entry = Parse-EntryText $text
        if ($null -eq $entry) {
            continue
        }

        $Rows.Value += ,@(
            $Stage,
            $Category,
            [string]$entry.Name,
            [string]$entry.Detail,
            ""
        )
    }
}

function Add-EquipmentRowsFromColumnSet {
    param(
        [ref]$Rows,
        [object]$Worksheet,
        [string]$Category,
        [string]$SubCategory,
        [int]$StartRow,
        [int]$EndRow,
        [int]$NameColumn,
        [int]$EffectColumn,
        [int]$PriceColumn,
        [int]$ArmorColumn = 0,
        [int]$MagicResistColumn = 0,
        [int]$DurabilityColumn = 0,
        [int]$RequirementColumn = 0,
        [int]$DamageColumn = 0,
        [string]$Note = ""
    )

    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $name = ([string]$Worksheet.Cells.Item($row, $NameColumn).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($name -match '^(大印|胸甲（轻）|重|头盔|腿甲|饰品|长剑|阵旗|长枪|拳套|祭坛|手环|吊坠|丹炉|手杖)$') {
            continue
        }

        if ($name -match '大成功|极难|困难|普通|大失败') {
            continue
        }

        $Rows.Value += ,@(
            $Category,
            $SubCategory,
            $name,
            ([string]$Worksheet.Cells.Item($row, $EffectColumn).Text).Trim(),
            ([string]$Worksheet.Cells.Item($row, $PriceColumn).Text).Trim(),
            $(if ($ArmorColumn -gt 0) { ([string]$Worksheet.Cells.Item($row, $ArmorColumn).Text).Trim() } else { "" }),
            $(if ($MagicResistColumn -gt 0) { ([string]$Worksheet.Cells.Item($row, $MagicResistColumn).Text).Trim() } else { "" }),
            $(if ($DurabilityColumn -gt 0) { ([string]$Worksheet.Cells.Item($row, $DurabilityColumn).Text).Trim() } else { "" }),
            $(if ($RequirementColumn -gt 0) { ([string]$Worksheet.Cells.Item($row, $RequirementColumn).Text).Trim() } else { "" }),
            $(if ($DamageColumn -gt 0) { ([string]$Worksheet.Cells.Item($row, $DamageColumn).Text).Trim() } else { "" }),
            $Note
        )
    }
}

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

$skillRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\\..\\.."))
$sourceRoot = Join-Path $repoRoot "source"
$tableSourceRoot = Join-Path $sourceRoot "tables"
$workbookRoot = Join-Path $repoRoot "workbooks"
$templateRoot = Join-Path $workbookRoot "templates"
$sourcePath = Join-Path $templateRoot "新版玖镇角色卡_墨家.xlsm"
$targetPath = Join-Path $workbookRoot "玖镇新版角色卡.xlsm"
$dataPath = Join-Path $skillRoot "references\\data\\universal_rolecard_data.json"
$vbaRoot = Join-Path $skillRoot "references\\vba"
$herbSourcePath = Join-Path $tableSourceRoot "草药学列表.xlsx"
$equipmentSourcePath = Join-Path $tableSourceRoot "装备列表.xlsx"
$skipMacroRun = ($env:JIUZHEN_SKIP_MACRO_RUN -eq "1")

if (-not (Test-Path -LiteralPath $dataPath)) {
    throw "Data file not found: $dataPath"
}

Write-Host "[1/7] Loading universal data..."
$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 10

Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force

$excel = $null
$workbook = $null
$vbProject = $null
$wsSettings = $null
$wsBook = $null
$wsPersonal = $null
$wsAfterBook = $null
$wsTrait = $null
$wsSkill = $null
$wsEquipment = $null
$wsHerb = $null
$wsFollower = $null
$herbWorkbook = $null
$herbWs = $null
$equipmentWorkbook = $null
$equipmentWs = $null

try {
    Write-Host "[2/7] Opening target workbook in Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $workbook = $excel.Workbooks.Open($targetPath)

    $wsSettings = Get-WorksheetByName $workbook "基本信息"
    $wsBook = Get-WorksheetByName $workbook "书籍、特性与技能"
    $wsPersonal = Get-WorksheetByName $workbook "个人信息"

    if ($null -eq $wsSettings -or $null -eq $wsBook -or $null -eq $wsPersonal) {
        throw "Required worksheets missing in target workbook."
    }

    Write-TableRows $wsSettings 1 5 $themeReferenceRows
    $themeValidation = $wsSettings.Range("B2").Validation
    $themeValidation.Delete()
    $themeValidation.Add(3, 1, 1, '=$A$6:$A$15')
    $themeValidation.IgnoreBlank = $true
    $themeValidation.InCellDropdown = $true
    Release-ComObject $themeValidation
    $wsSettings.Range("A16").Value2 = "说明"
    $wsSettings.Range("B16").Value2 = "切换当前主题后会自动刷新配色，并按需显示墨家专属页。"
    $wsSettings.Range("A17").Value2 = "专属页说明"

    Remove-WorksheetIfExists $workbook "通用特性体系"
    Remove-WorksheetIfExists $workbook "技能列表"
    Remove-WorksheetIfExists $workbook "装备列表"
    Remove-WorksheetIfExists $workbook "【药家】草药学列表"
    Remove-WorksheetIfExists $workbook "随从列表"

    $wsAfterBook = $workbook.Worksheets.Item($wsBook.Index + 1)

    $wsTrait = $workbook.Worksheets.Add($wsAfterBook)
    $wsTrait.Name = "通用特性体系"

    $wsSkill = $workbook.Worksheets.Add($workbook.Worksheets.Item($wsTrait.Index + 1))
    $wsSkill.Name = "技能列表"

    $wsEquipment = $workbook.Worksheets.Add($workbook.Worksheets.Item($wsSkill.Index + 1))
    $wsEquipment.Name = "装备列表"

    $wsHerb = $workbook.Worksheets.Add($workbook.Worksheets.Item($wsEquipment.Index + 1))
    $wsHerb.Name = "【药家】草药学列表"

    $wsFollower = $workbook.Worksheets.Add($workbook.Worksheets.Item($wsHerb.Index + 1))
    $wsFollower.Name = "随从列表"

    $wsSettings.Range("B17").Value2 = "仅当当前主题为""墨家""时显示【墨家】工巧格与魁人与【墨家】物品库；仅当当前主题为""药家""时显示【药家】草药学列表。乐正家当前复用通用页，不新增专属工作表。"

    Write-Host "[3/7] Writing family overview, trait catalog, and skill catalog..."
    Set-MergedValue $wsSettings "A19:F19" "家族概览与升级方式"
    Write-TableRows $wsSettings 1 20 (, @("家族", "核心定位", "家族描述", "升级方式", "代表体系", "备注"))

    $familyRows = @()
    foreach ($family in $data.families) {
        $familyRows += ,@(
            [string]$family.family,
            [string]$family.positioning,
            [string]$family.description,
            [string]$family.upgrade,
            [string]$family.system,
            [string]$family.note
        )
    }
    Write-TableRows $wsSettings 1 21 $familyRows
    $wsSettings.Columns("A").ColumnWidth = 14
    $wsSettings.Columns("B").ColumnWidth = 22
    $wsSettings.Columns("C").ColumnWidth = 44
    $wsSettings.Columns("D").ColumnWidth = 82
    $wsSettings.Columns("E").ColumnWidth = 34
    $wsSettings.Columns("F").ColumnWidth = 42

    Set-MergedValue $wsTrait "A1:D1" "通用特性体系"
    Write-TableRows $wsTrait 1 2 (, @("家族", "分组", "特性名称", "特性介绍"))
    $traitRows = @()
    foreach ($trait in $data.traits) {
        $traitRows += ,@(
            [string]$trait.family,
            [string]$trait.group,
            [string]$trait.name,
            [string]$trait.description
        )
    }
    Write-TableRows $wsTrait 1 3 $traitRows
    $wsTrait.Columns("A").ColumnWidth = 10
    $wsTrait.Columns("B").ColumnWidth = 18
    $wsTrait.Columns("C").ColumnWidth = 18
    $wsTrait.Columns("D").ColumnWidth = 68
    $wsTrait.Columns("E:G").ColumnWidth = 11

    Set-MergedValue $wsSkill "A1:F1" "技能列表"
    Write-TableRows $wsSkill 1 2 (, @("家族", "分组", "技能名称", "技能介绍", "技能速度", "技能消耗"))
    $skillRows = @()
    foreach ($skill in $data.skills) {
        $skillRows += ,@(
            [string]$skill.family,
            [string]$skill.group,
            [string]$skill.name,
            [string]$skill.description,
            [string]$skill.speed,
            [string]$skill.cost
        )
    }
    Write-TableRows $wsSkill 1 3 $skillRows
    $wsSkill.Columns("A").ColumnWidth = 10
    $wsSkill.Columns("B").ColumnWidth = 22
    $wsSkill.Columns("C").ColumnWidth = 18
    $wsSkill.Columns("D").ColumnWidth = 74
    $wsSkill.Columns("E").ColumnWidth = 12
    $wsSkill.Columns("F").ColumnWidth = 18
    $wsSkill.Columns("G:I").ColumnWidth = 11

    Write-Host "[4/7] Building equipment catalog..."
    $equipmentWorkbook = $excel.Workbooks.Open($equipmentSourcePath)
    $equipmentWs = $equipmentWorkbook.Worksheets.Item(1)

    Set-MergedValue $wsEquipment "A1:K1" "装备列表"
    Set-MergedValue $wsEquipment "A2:K2" "可在下方直接新增自定义装备；请规范填写大类、子类与名称，人物/随从联动会自动刷新，表头下拉可按分类筛选。"
    Write-TableRows $wsEquipment 1 3 (, @("大类", "子类", "名称", "效果", "价格", "护甲", "魔抗", "耐久", "需求", "伤害", "备注"))

    $equipmentRows = @()
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "饰品" "大印" 3 43 1 2 3 0 0 0 0 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "防具" "胸甲（轻）" 2 21 4 5 6 7 8 9 10 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "防具" "胸甲（重）" 23 42 4 5 6 7 8 9 10 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "防具" "头盔" 44 63 4 5 6 7 8 9 10 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "防具" "腿甲" 65 82 4 5 6 7 8 9 10 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "饰品" "饰品" 2 42 13 14 15 0 0 0 16 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "武器" "长剑" 2 11 19 20 22 0 0 0 0 21 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "武器" "阵旗" 14 23 19 20 22 0 0 0 0 21 "阵旗没有价格的项目保留原表说明"
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "武器" "长枪" 2 31 23 24 26 0 0 0 0 25 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "武器" "拳套" 2 31 27 28 30 0 0 0 0 29 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "功能装备" "祭坛" 2 23 31 32 33 0 0 0 0 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "功能装备" "手环" 2 23 34 35 36 0 0 0 0 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "功能装备" "吊坠" 2 23 37 38 39 0 0 0 0 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "功能装备" "丹炉" 2 23 40 41 42 0 0 0 0 0 ""
    Add-EquipmentRowsFromColumnSet ([ref]$equipmentRows) $equipmentWs "功能装备" "手杖" 2 23 43 44 45 0 0 0 0 0 ""

    Write-TableRows $wsEquipment 1 4 $equipmentRows

    $equipmentLookupRows = @()
    foreach ($rowValues in $equipmentRows) {
        $equipmentLookupRows += ,@($rowValues[2], $rowValues[3], $rowValues[4], $rowValues[5], $rowValues[6], $rowValues[7], $rowValues[8], $rowValues[9])
    }
    Write-TableRows $wsEquipment 13 4 (, @("名称", "效果", "价格", "护甲", "魔抗", "耐久", "需求", "伤害"))
    Write-TableRows $wsEquipment 13 5 $equipmentLookupRows

    Write-TableRows $wsEquipment 22 4 (, @("头部装备", "胸部装备", "腿部装备", "饰品与大印", "随从武器"))
    Set-CellValue $wsEquipment 5 22 "空"
    Set-CellValue $wsEquipment 5 23 "空"
    Set-CellValue $wsEquipment 5 24 "空"
    Set-CellValue $wsEquipment 5 25 "空"
    Set-CellValue $wsEquipment 5 26 "空"

    $equipListRow = 6
    foreach ($rowValues in $equipmentRows) {
        switch ([string]$rowValues[1]) {
            "头盔" {
                $itemName = [string]$rowValues[2]
                Set-CellValue $wsEquipment $equipListRow 22 $itemName
                $equipListRow++
            }
        }
    }
    $chestRow = 6
    foreach ($rowValues in $equipmentRows) {
        if ([string]$rowValues[1] -eq "胸甲（轻）" -or [string]$rowValues[1] -eq "胸甲（重）") {
            $itemName = [string]$rowValues[2]
            Set-CellValue $wsEquipment $chestRow 23 $itemName
            $chestRow++
        }
    }
    $legRow = 6
    foreach ($rowValues in $equipmentRows) {
        if ([string]$rowValues[1] -eq "腿甲") {
            $itemName = [string]$rowValues[2]
            Set-CellValue $wsEquipment $legRow 24 $itemName
            $legRow++
        }
    }
    $accRow = 6
    foreach ($rowValues in $equipmentRows) {
        if ([string]$rowValues[0] -eq "饰品") {
            $itemName = [string]$rowValues[2]
            Set-CellValue $wsEquipment $accRow 25 $itemName
            $accRow++
        }
    }
    $weaponRow = 6
    foreach ($rowValues in $equipmentRows) {
        if ([string]$rowValues[0] -eq "武器") {
            $itemName = [string]$rowValues[2]
            Set-CellValue $wsEquipment $weaponRow 26 $itemName
            $weaponRow++
        }
    }

    $wsEquipment.Columns("A").ColumnWidth = 10
    $wsEquipment.Columns("B").ColumnWidth = 14
    $wsEquipment.Columns("C").ColumnWidth = 18
    $wsEquipment.Columns("D").ColumnWidth = 66
    $wsEquipment.Columns("E").ColumnWidth = 10
    $wsEquipment.Columns("F").ColumnWidth = 8
    $wsEquipment.Columns("G").ColumnWidth = 8
    $wsEquipment.Columns("H").ColumnWidth = 8
    $wsEquipment.Columns("I").ColumnWidth = 16
    $wsEquipment.Columns("J").ColumnWidth = 12
    $wsEquipment.Columns("K").ColumnWidth = 20
    $wsEquipment.Columns("L:Z").Hidden = $true
    $equipmentWorkbook.Close($false)
    Release-ComObject $equipmentWs
    Release-ComObject $equipmentWorkbook
    $equipmentWs = $null
    $equipmentWorkbook = $null

    Write-Host "[5/7] Building herb catalog..."
    $herbWorkbook = $excel.Workbooks.Open($herbSourcePath)
    $herbWs = $herbWorkbook.Worksheets.Item(1)

    Set-MergedValue $wsHerb "A1:E1" "【药家】草药学列表"
    Set-MergedValue $wsHerb "A2:E2" "按阶段整理草药、蛊术进阶、药方与特殊规则；当前主题切换为药家时显示。"
    Write-TableRows $wsHerb 1 4 (, @("阶段", "类别", "名称", "效果说明", "备注"))

    $herbRows = @()
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "入门" "草药" 1 (2..11)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "入门" "进阶（蛊术）" 2 (2..11)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "入门" "药方" 1 (14..19)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "中阶" "草药" 1 (22..31)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "中阶" "进阶" 2 (21..30)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "中阶" "药方" 1 (34..39)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "高阶" "草药" 1 (42..51)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "高阶" "进阶" 2 (41..50)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "高阶" "药方" 1 (54..58)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "特殊" "草药" 1 (61..66)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "特殊" "药方" 1 (69..74)
    Add-HerbEntriesFromRange ([ref]$herbRows) $herbWs "特殊" "药品" 1 (78..78)

    $herbRows += ,@(
        "通用规则",
        "购买与炼药",
        "草药获取与票券价格",
        "当每次到达新等级的时候，获得该等级每份草药 3 份。购买方面，低阶草药一票券可以购买 3 份，中阶一票券购买 2 份，高阶一票券购买 1 份。低阶药方和虫子售价每个 2 票券，中阶 4 票券，高阶 6 票券。每次炼药低阶半小时，中阶 1 小时，高阶 1 小时。",
        "源自原表第 55 行说明"
    )
    $herbRows += ,@(
        "通用规则",
        "经验成长",
        "药家经验体系",
        "每 20 经验一级，成长为大成功 +10、极难 +7+1d3、困难 +5+1d5、普通 +1+1d9、失败 +2。",
        "源自原表第 64 行说明"
    )

    Write-TableRows $wsHerb 1 5 $herbRows
    $wsHerb.Columns("A").ColumnWidth = 12
    $wsHerb.Columns("B").ColumnWidth = 18
    $wsHerb.Columns("C").ColumnWidth = 18
    $wsHerb.Columns("D").ColumnWidth = 78
    $wsHerb.Columns("E").ColumnWidth = 22
    $herbWorkbook.Close($false)
    Release-ComObject $herbWs
    Release-ComObject $herbWorkbook
    $herbWs = $null
    $herbWorkbook = $null

    Write-Host "[6/7] Building follower sheet and workbook helpers..."
    Set-MergedValue $wsFollower "A1:R1" "随从1"
    Set-MergedValue $wsFollower "A15:R15" "随从2"
    Set-MergedValue $wsFollower "A29:R29" "随从3"
    Set-MergedValue $wsFollower "A43:R43" "随从4"

    $blockStarts = @(1, 15, 29, 43)
    foreach ($startRow in $blockStarts) {
    Write-TableRows $wsFollower 1 ($startRow + 1) @(
            @("移动", "", "灵魂", "", "幸运", ""),
            @("体质", "", "力量", "", "", ""),
            @("体型", "", "魅力", "", "", ""),
            @("敏捷", "", "智力", "", "", "")
        )

        Set-CellValue $wsFollower ($startRow + 1) 7 "头部装备"
        Set-MergedValue $wsFollower ("H{0}:J{0}" -f ($startRow + 1)) "空"
        Set-MergedValue $wsFollower ("K{0}:R{0}" -f ($startRow + 1)) ""

        Set-CellValue $wsFollower ($startRow + 2) 7 "胸部装备"
        Set-MergedValue $wsFollower ("H{0}:J{0}" -f ($startRow + 2)) "空"
        Set-MergedValue $wsFollower ("K{0}:R{0}" -f ($startRow + 2)) ""

        Set-CellValue $wsFollower ($startRow + 3) 7 "腿部装备"
        Set-MergedValue $wsFollower ("H{0}:J{0}" -f ($startRow + 3)) "空"
        Set-MergedValue $wsFollower ("K{0}:R{0}" -f ($startRow + 3)) ""

        Set-CellValue $wsFollower ($startRow + 4) 7 "饰品"
        Set-MergedValue $wsFollower ("H{0}:J{0}" -f ($startRow + 4)) "空"
        Set-MergedValue $wsFollower ("K{0}:R{0}" -f ($startRow + 4)) ""

        Set-MergedValue $wsFollower ("A{0}:G{0}" -f ($startRow + 6)) "技能栏"
        Set-MergedValue $wsFollower ("H{0}:M{0}" -f ($startRow + 6)) "武器栏"
        Set-MergedValue $wsFollower ("N{0}:R{0}" -f ($startRow + 6)) "物品栏"

        Write-TableRows $wsFollower 1 ($startRow + 7) (, @("技能名称", "技能介绍", "", "", "", "", "", "武器名称", "武器伤害", "武器介绍", "", "", "", "物品名称", "物品介绍", "", "", ""))
        Set-MergedValue $wsFollower ("B{0}:G{0}" -f ($startRow + 7)) "技能介绍"
        Set-MergedValue $wsFollower ("J{0}:M{0}" -f ($startRow + 7)) "武器介绍"
        Set-MergedValue $wsFollower ("O{0}:R{0}" -f ($startRow + 7)) "物品介绍"

        for ($row = $startRow + 8; $row -le $startRow + 11; $row++) {
            Set-CellValue $wsFollower $row 1 "空"
            Set-CellValue $wsFollower $row 8 "空"
            Set-MergedValue $wsFollower ("B{0}:G{0}" -f $row) ""
            Set-MergedValue $wsFollower ("J{0}:M{0}" -f $row) ""
            Set-MergedValue $wsFollower ("O{0}:R{0}" -f $row) ""
        }
    }

    $wsFollower.Columns("A").ColumnWidth = 14
    $wsFollower.Columns("B").ColumnWidth = 10
    $wsFollower.Columns("C").ColumnWidth = 14
    $wsFollower.Columns("D").ColumnWidth = 10
    $wsFollower.Columns("E").ColumnWidth = 12
    $wsFollower.Columns("F").ColumnWidth = 10
    $wsFollower.Columns("G").ColumnWidth = 10
    $wsFollower.Columns("H").ColumnWidth = 14
    $wsFollower.Columns("I").ColumnWidth = 10
    $wsFollower.Columns("J").ColumnWidth = 10
    $wsFollower.Columns("K").ColumnWidth = 10
    $wsFollower.Columns("L").ColumnWidth = 10
    $wsFollower.Columns("M").ColumnWidth = 10
    $wsFollower.Columns("N").ColumnWidth = 14
    $wsFollower.Columns("O").ColumnWidth = 10
    $wsFollower.Columns("P").ColumnWidth = 10
    $wsFollower.Columns("Q").ColumnWidth = 10
    $wsFollower.Columns("R").ColumnWidth = 10

    $wsBook.Rows("1:1").Insert() | Out-Null
    Set-MergedValue $wsBook "A1:J1" "书籍配置"
    Set-MergedValue $wsBook "L1:U1" "特性配置"
    Set-MergedValue $wsBook "A19:U19" "技能配置"

    for ($row = 3; $row -le 18; $row++) {
        $wsBook.Range("L$row").Value2 = "空"
        $wsBook.Range("M$row").Value2 = ""
    }
    for ($row = 21; $row -le 36; $row++) {
        $wsBook.Range("A$row").Value2 = "空"
        $wsBook.Range("L$row").Value2 = "空"
    }
    $wsBook.Columns("W:AD").Hidden = $true

    $wsPersonal.Rows("23:28").Insert() | Out-Null

    Set-MergedValue $wsPersonal "A23:N23" "武器栏"
    Set-MergedValue $wsPersonal "A24:C24" "武器名称"
    Set-MergedValue $wsPersonal "D24:E24" "武器伤害"
    Set-MergedValue $wsPersonal "F24:N24" "武器介绍"

    foreach ($row in 25..28) {
        Set-MergedValue $wsPersonal ("A{0}:C{0}" -f $row) "空"
        Set-MergedValue $wsPersonal ("D{0}:E{0}" -f $row) ""
        Set-MergedValue $wsPersonal ("F{0}:N{0}" -f $row) ""
    }

    $wsPersonal.Range("P23:V36").ClearContents() | Out-Null
    Set-MergedValue $wsPersonal "P23:V23" "满足画风的细节"

    $detailLabels = @(
        "发色",
        "瞳色",
        "肤色",
        "理想",
        "羁绊",
        "私人爱好",
        "缺点",
        "个人信仰",
        "名言警句",
        "其他"
    )
    for ($index = 0; $index -lt $detailLabels.Count; $index++) {
        $labelRow = 24 + $index
        Set-MergedValue $wsPersonal ("P{0}:R{0}" -f $labelRow) $detailLabels[$index]
        Set-MergedValue $wsPersonal ("S{0}:V{0}" -f $labelRow) ""
    }

    Set-MergedValue $wsPersonal "P35:V35" "满足画风物品栏"

    foreach ($cellAddr in @("F13", "F15", "F17", "F19", "F21")) {
        $wsPersonal.Range($cellAddr).Value2 = "空"
    }

    Write-Host "[7/7] Updating VBA modules and finalizing workbook..."
    $vbProject = $workbook.VBProject
    Set-CodeModule $vbProject.VBComponents.Item("ThisWorkbook") (Get-VbaCodeBody (Join-Path $vbaRoot "ThisWorkbook.cls"))
    Set-CodeModule $vbProject.VBComponents.Item("Sheet3") (Get-VbaCodeBody (Join-Path $vbaRoot "Sheet3.cls"))
    Set-CodeModule $vbProject.VBComponents.Item("JiuzhenMojiaMacros") (Get-VbaCodeBody (Join-Path $vbaRoot "JiuzhenMojiaMacros.updated.bas"))
    Set-CodeModule (Get-OrCreateStandardModule $vbProject "JiuzhenUniversalMacros") (Get-VbaCodeBody (Join-Path $vbaRoot "JiuzhenUniversalMacros.bas"))

    if ($skipMacroRun) {
        Write-Host "Skipping ApplyWorkbookTheme because JIUZHEN_SKIP_MACRO_RUN=1."
    } else {
        Write-Host "Running ApplyWorkbookTheme..."
        $excel.Run("'" + $workbook.Name + "'!ApplyWorkbookTheme")
    }
    Write-Host "Calculating workbook..."
    $excel.CalculateFull()
    Write-Host "Saving workbook..."
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

    Release-ComObject $wsFollower
    Release-ComObject $wsHerb
    Release-ComObject $wsSkill
    Release-ComObject $wsTrait
    Release-ComObject $wsAfterBook
    Release-ComObject $wsPersonal
    Release-ComObject $wsBook
    Release-ComObject $wsSettings
    if ($null -ne $herbWorkbook) {
        try {
            $herbWorkbook.Close($false)
        } catch {
        }
    }
    Release-ComObject $herbWs
    Release-ComObject $herbWorkbook
    Release-ComObject $vbProject
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
