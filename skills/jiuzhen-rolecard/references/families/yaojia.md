# 药家

## 当前定位

- 当前通用 JSON 中 `药家` 特性数 `20`、技能数 `0`。
- 药家的“特性成长”走通用表，草药 / 药方 / 蛊术目录走专属隐藏表。

## 主要来源

- `source/families/药家.docx`
- `source/tables/草药学列表.xlsx`
- `references/data/universal_rolecard_data.json`
- `scripts/build_universal_workbook.ps1`
- `references/vba/JiuzhenUniversalMacros.bas`

## 需要特殊修改的表

- `【药家】草药学列表`
  - 仅在主题为 `药家` 时显示。
  - 当前是按 `草药学列表.xlsx` 重建出来的目录页，不建议手工长期维护。
  - 修改草药、药方、蛊术、阶段说明时，优先改 `source/tables/草药学列表.xlsx`，再运行 `scripts/build_universal_workbook.ps1`。
- `通用特性体系`
  - 所有药家等级特性仍通过这里进入下拉体系。
- `基本信息!A19:F*`
  - 药家家族概览与升级说明在这里维护。

## 注意点

- 药家的专属页显隐由 `JiuzhenUniversalMacros.bas` 的 `ApplyFamilySpecificSheetVisibility` 控制。
- 如果以后要把药家补成显式技能体系，不仅要增 `技能列表` 数据，还要检查 `书籍、特性与技能` helper 是否需要扩容。
