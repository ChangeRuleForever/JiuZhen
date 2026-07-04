# 通用层

## 先看这里

- 当前 skill 已经把原项目里的脚本、JSON 和 VBA 导出整合到 skill 内。
- 原始 Word / Excel 素材位于 `source/`，工作簿模板与产物位于 `workbooks/`，可维护文本资源统一放在本 skill 下。
- 常用入口：
  - `scripts/build_universal_rolecard_data.py`
  - `scripts/build_universal_workbook.ps1`
  - `scripts/update_mojia_workbook.ps1`
  - `references/data/universal_rolecard_data.json`
  - `references/data/family_sources.json`
  - `references/vba/*.bas|*.cls`

## 工作簿总览

| 工作表 | 当前状态 | 作用 | 常见修改点 |
| --- | --- | --- | --- |
| `基本信息` | 可见 | 主题切换、家族概览、说明区 | `B2` 当前主题；`A6:A15` 家族列表；`A19:F*` 家族概览 |
| `个人信息` | 可见 | 角色基础信息、属性、装备下拉 | `F13/F15/F17/F19/F21` 装备下拉；技能/属性公式区 |
| `书籍、特性与技能` | 可见 | 书籍、特性、双列技能配置 | `L2:L17` 特性下拉；`A20:A35` 与 `L20:L35` 技能下拉；`W:AD` 隐藏 helper |
| `通用特性体系` | 可见 | 全家族可检索特性目录 | `A:D` |
| `技能列表` | 可见 | 全家族可检索技能目录 | `A:F` |
| `装备列表` | 可见 | 装备总表与下拉 helper | `A:K` 可见目录；`M:Z` 隐藏 helper |
| `【药家】草药学列表` | 主题控制 | 药家草药/药方/蛊术目录 | `A:E`；仅药家主题显示 |
| `随从列表` | 可见 | 4 个随从块，技能/武器/物品下拉 | `A1:R56` |
| `【墨家】工巧格与魁人` | 主题控制 | 墨家背包、成长、魁人模组/武器 | `A:W`；按钮与验证都在此页 |
| `魁人面板` | `veryHidden` | 墨家旧/辅助手工面板 | 默认隐藏；仍挂着一个刷新按钮 |
| `【墨家】物品库` | 主题控制 | 墨家物品、插件、武器、helper | `A:AC` 可见；`AD:AZ` 隐藏 helper |
| `快速复制区` | 可见 | 面向在线文档的摘要输出 | `B:G` 全由公式拉取 |

## 主题与事件

- `基本信息!B2` 是全工作簿主题开关。
- `references/vba/ThisWorkbook.cls` 在 `Workbook_Open` 与 `Workbook_SheetChange` 中触发 `ApplyWorkbookTheme`。
- `references/vba/JiuzhenMojiaMacros.updated.bas`
  - 负责主题配色、墨家页显隐、背包刷新、魁人校验、保存前验证。
- `references/vba/JiuzhenUniversalMacros.bas`
  - 负责通用表配色、药家页显隐、通用特性/技能/装备/随从的下拉与公式回填。

## 通用下拉与 helper 关系

- `基本信息!B2`
  - 数据源：`基本信息!A6:A15`
- `书籍、特性与技能!L2:L17`
  - 数据源：同页 `W:W`
- `书籍、特性与技能!A20:A35` 与 `L20:L35`
  - 数据源：同页 `AA:AA`
- `个人信息!F13/F15/F17/F19/F21`
  - 数据源：`装备列表` 的隐藏 helper 列
- `随从列表`
  - 技能名称来自 `书籍、特性与技能!AA:AD`
  - 武器/装备来自 `装备列表` 的隐藏 helper 列

## 数据源优先级

1. `source/` 下的 Excel / Word 源文件。
2. 本 skill 的脚本与 JSON。
3. 本 skill 的 VBA 导出。
4. `workbooks/` 下的 `.xlsm`。

如果同一份信息已经在脚本或 JSON 中结构化，不要优先手改成品工作簿。

## 常见改动路径

### 改通用特性 / 技能 / 家族概览

- 先改 `references/data/universal_rolecard_data.json`，或改 `scripts/build_universal_rolecard_data.py` 的规则后重生成。
- 再运行 `scripts/build_universal_workbook.ps1`，让 `基本信息`、`通用特性体系`、`技能列表`、`随从列表` 与 `书籍、特性与技能` helper 一起更新。

### 改通用装备 / 武器下拉

- 改 `source/tables/装备列表.xlsx`。
- 再运行 `scripts/build_universal_workbook.ps1`，让 `装备列表` 和相关验证公式重建。

### 改主题颜色 / 新增家族主题

- 同时改：
  - `基本信息!A6:A15` 与家族概览文字
  - `references/vba/JiuzhenMojiaMacros.updated.bas` 的 `PrimaryColor` / `Palette`
  - `references/vba/JiuzhenUniversalMacros.bas` 的 `UniversalPrimaryColor` / `ThemeTone`
- 如果新增家族还需要专属页，再补显隐规则和脚本生成逻辑。

## 验证清单

- `基本信息!B2` 切换后能否刷新配色。
- `书籍、特性与技能` 的特性 / 技能下拉是否仍能拉到正确目录。
- `个人信息` 与 `随从列表` 的装备 / 技能描述回填公式是否仍正确。
- 专属页是否按主题正确显示或隐藏。
- 若改了 VBA，按钮名称与宏名是否仍一致。
