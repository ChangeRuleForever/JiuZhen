# 玖镇角色卡项目

这个仓库用于维护《玖镇》项目的角色卡工作簿、各家设定素材、通用/专属 VBA 宏，以及后续新增家族、技能、特性时所需的可重复构建脚本。

## 项目目标

- 把原本分散在 Excel、Word、VBA 中的内容整理成可版本管理的 Git 仓库。
- 把“可维护源文件”和“生成后的工作簿产物”拆开，减少直接手改 `.xlsm` 带来的混乱。
- 保留玖镇规则中“通用层 + 各家特性”的结构，方便后续继续加家族、改技能、调主题和修宏。

## 仓库结构

```text
JiuZhen/
├─ README.md
├─ .gitattributes
├─ .gitignore
├─ source/
│  ├─ families/              # 各家 Word 原始设定
│  └─ tables/                # 通用特性、装备、草药等 Excel 源表
├─ workbooks/
│  ├─ 玖镇新版角色卡.xlsm      # 当前生成/交付用工作簿
│  └─ templates/
│     └─ 新版玖镇角色卡_墨家.xlsm  # 构建用基础模板
└─ skills/
   └─ jiuzhen-rolecard/
      ├─ scripts/            # 构建与更新脚本
      ├─ references/
      │  ├─ data/            # 结构化 JSON 数据
      │  ├─ families/        # 各家维护说明
      │  ├─ vba/             # VBA 导出源码
      │  └─ previews/        # 预览/参考文件
      └─ SKILL.md
```

## 当前家族一览

| 家族 | 核心定位 | 代表机制 | 当前承载方式 |
| --- | --- | --- | --- |
| 张家 | 雷火道法正宗 | 雷法 / 火法 / 阵符剑术 | 通用表 |
| 岳家 | 纯武学前线 | 拳系 / 枪系 / 属性路线 | 通用表 |
| 林家 | 风水龙脉操盘 | 风水 / 龙脉 / 蒙天 | 通用表 |
| 易家 | 占卜与天谴 | 铜钱 / 占卜 / 术阵 | 通用表 |
| 曹家 | 商法与资源转换 | 商法 / 交涉 / 资源调度 | 通用表 |
| 墨家 | 机关与工巧格 | 工巧格 / 魁人 / 机巧师-偃师-布设师 | 墨家专属页 |
| 药家 | 丹蛊与持续经营 | 炼丹 / 蛊虫 / 药效叠加 | 通用表 + 药家专属页 |
| 巫家 | 附魔防护全能术 | 附魔 / 防护 / 能力 / 萨满 | 通用表 |
| 亓家 | 炼鬼与术阵异道 | 小鬼 / 炼鬼 / 术阵 | 通用表 |
| 乐正家 | 音律控场与共鸣支援 | 灵魂音符 / 打击乐 / 管弦乐 / 持续演奏 | 通用表 |

## 玖镇规则与维护约定

### 工作簿结构规则

- `基本信息!B2` 是全工作簿主题开关。
- `基本信息!A6:A15` 是当前全部家族主题列表。
- `书籍、特性与技能!W:AD`、`装备列表!M:Z`、`【墨家】物品库` 的 helper 区是隐藏且公式驱动的，不要把它们当普通展示区维护。
- `【墨家】工巧格与魁人`、`【墨家】物品库` 只在主题为 `墨家` 时显示。
- `【药家】草药学列表` 只在主题为 `药家` 时显示。
- 除墨家和药家外，其他家族目前都复用通用页，不单独新增可见专属工作表。

### 数据维护规则

- 优先改 `source/` 下的 Word / Excel 源文件，再改 skill 里的脚本、JSON 和 VBA。
- `skills/jiuzhen-rolecard/references/data/universal_rolecard_data.json` 是通用家族信息、特性和技能的结构化结果。
- `skills/jiuzhen-rolecard/references/vba/` 是 VBA 的可维护源码，不要只在 Excel 里改完就结束。
- 墨家的专属玩法主要依赖专属页和 VBA，不依赖 `通用特性体系` / `技能列表` 承载主体。

### 乐正家特别规则

- 乐正家文档是按“阶段 -> 技能 -> 特性”交替写的。
- 维护时必须严格分开：
  - `特性` 只能进入 `通用特性体系`
  - `技能` 只能进入 `技能列表`
- 不允许再把乐正家的特性错误写进技能 helper。

### 源文件优先级

1. `source/` 下的 Word / Excel 原始素材
2. `skills/jiuzhen-rolecard/scripts/` 与 `references/data/`
3. `skills/jiuzhen-rolecard/references/vba/`
4. `workbooks/` 下的 `.xlsm`

## 常用工作流

### 1. 重新提取家族来源文本

```powershell
python .\skills\jiuzhen-rolecard\scripts\extract_family_sources.py > .\skills\jiuzhen-rolecard\references\data\family_sources.json
```

### 2. 重建通用家族 JSON

```powershell
python .\skills\jiuzhen-rolecard\scripts\build_universal_rolecard_data.py
```

### 3. 更新墨家基础模板

```powershell
powershell -ExecutionPolicy Bypass -File .\skills\jiuzhen-rolecard\scripts\update_mojia_workbook.ps1
```

### 4. 基于模板重建最终工作簿

```powershell
powershell -ExecutionPolicy Bypass -File .\skills\jiuzhen-rolecard\scripts\build_universal_workbook.ps1
```

## 运行前提

- Windows 环境
- 本机已安装桌面版 Excel
- Excel 已允许访问 VBA Project Object Model
- Python 环境可用 `openpyxl`、`python-docx`

## Git 约定

- `.xlsm`、`.xlsx`、`.docx` 视为二进制文件，优先通过配套脚本和文档改动来表达结构变化。
- 自动生成的备份工作簿、Office 临时锁文件、Python 缓存已在 `.gitignore` 中排除。
- 如果只是局部微调，仍建议把最终逻辑回写到 `scripts/`、`references/data/` 或 `references/vba/`，避免仓库只留下结果、不留下原因。

## 深入说明

- 通用维护说明：`skills/jiuzhen-rolecard/references/common.md`
- 各家维护说明：`skills/jiuzhen-rolecard/references/families/*.md`
- 主 skill 入口：`skills/jiuzhen-rolecard/SKILL.md`
