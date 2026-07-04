# 墨家

## 当前定位

- 当前通用 JSON 中 `墨家` 特性数 `0`、技能数 `0`。
- 墨家不依赖 `通用特性体系` / `技能列表` 承载主玩法，而是依赖专属工作表与 VBA。

## 主要来源

- `source/families/墨家_新.docx`
- `workbooks/templates/新版玖镇角色卡_墨家.xlsm`
- `scripts/update_mojia_workbook.ps1`
- `references/vba/JiuzhenMojiaMacros.updated.bas`
- `references/vba/ThisWorkbook.cls`
- `references/vba/Sheet3.cls`
- `references/previews/mojia/`
  - `墨家_成长重制预览稿.pdf`
  - `墨家_成长重制预览稿_含装备分类.pdf`
  - `墨家_成长重制预览稿_成长逻辑预览.pdf`
  - `墨家_附录预览.pdf`
  - 这些是视觉参考稿，不参与脚本生成，但在核对墨家规则编排、装备分类和成长说明时很有用。

## 专属工作表

- `【墨家】工巧格与魁人`
  - 主操作页。
  - `B5:B304` 本人背包物品下拉。
  - `K4:K13` 成长判定等级下拉。
  - `L4:L13` 成长增量下拉，固定值 `1,2,3,4,5`。
  - `P7` 魁人战斗能力下拉。
  - `U5:U14` 魁人模组下拉。
  - `P17:P19` 魁人武器下拉。
  - `I3:I8` 工巧格汇总公式。
  - `R6:R12` 模组上限、已选数、占格统计。
- `【墨家】物品库`
  - 可见目录在 `A:AC`。
  - 隐藏 helper 在 `AD:AZ`。
  - 关键 helper 组：
    - `AE:AF` 本人背包名称 / 占格
    - `AJ:AL` 模组名称 / 占格 / 说明
    - `AN:AP` 魁人武器名称 / 占格 / 说明
    - `AV:AW` 增长或战斗能力验证源
- `魁人面板`
  - 当前为 `veryHidden`，属于旧版或辅助面板。
  - 仍保留按钮和部分公式，不要随意删掉，除非同时清理 VBA 与按钮绑定。

## 控件与宏

- `btnRefreshMojiaInventory`
  - 位于 `【墨家】工巧格与魁人`
  - 绑定宏：`RefreshMojiaInventory`
- `btnRefreshKuiInventory`
  - 位于 `魁人面板`
  - 绑定宏：`RefreshKuiInventory`
- `Workbook_BeforeSave`
  - 通过 `CanSaveMojiaWorkbook` 拦截不合法的模组 / 工巧格状态。
- `Sheet3.Worksheet_Change`
  - 监听背包、成长、战斗能力、模组、武器相关区域，自动刷新背包并在非法选择时回滚。

## 需要特殊修改的表

- 新增 / 修改墨器、插件、魁人武器：
  - 先改 `【墨家】物品库`
  - 同时检查 helper 列是否仍覆盖全部条目
  - 如数量超出当前上限，补改 VBA 里对应的验证范围上限
- 修改工巧格规则、模组上限、占格逻辑：
  - 同时改工作表公式与 `JiuzhenMojiaMacros.updated.bas`
  - 重点看 `BASE_SLOT_COUNT`、`KUI_SINGLE_MODULE_SLOT_COST`、`ValidateKuiModuleState`
- 修改墨家主题配色或显隐：
  - 同时改 `ApplyMojiaSheetVisibility` 与 `Palette`
- 修改按钮或控件位置：
  - 同时检查 `update_mojia_workbook.ps1` 里对 shape 的定位逻辑

## 操作建议

- 涉及墨家结构性改动时，优先改 `scripts/update_mojia_workbook.ps1` 与 `references/vba/`，再回灌到 `.xlsm`。
- 只做单个数值或单条说明微调时，可以直接改工作表，但要回头确认验证、公式和刷新宏仍自洽。
