---
name: visio-diagram
description: 在 Windows 上用 Visio 生成 .vsdx 图，支持 ER 图、流程图、时序图，也能分析并注册用户提供的新模板。用户提到 Visio、.vsdx、ER 图/实体关系图、流程图/业务流程/审批流程、时序图/序列图/调用链路/接口交互，或要把文档里的流程画成图时，使用本技能。用户只说“画图”“生成图”“做个图”没指明类型时也要先用本技能确认类型。要分析或导入新的 Visio 模板时同样使用本技能。
compatibility: Windows；需已安装 Microsoft Visio；Python 需 pythoncom、win32com（pywin32）
---

# Visio 图绘制

依据用户给的资料，用内置模板生成新的 `.vsdx`。已有图不是输入：不读、不解压、不比对。

## 1. 确定画什么
按用户的说法选一份 reference，避免无关规则占上下文：

| 用户说的 | reference |
|---|---|
| ER 图、实体关系图、E-R 图、实体属性图 | `references/er.md` |
| 流程图、业务流程、操作流程、审批流程 | `references/flowchart.md` |
| 时序图、序列图、交互时序、调用链路、接口交互 | `references/sequence.md` |
| 添加、导入、注册、扩展新模板 | `references/template-intake.md` |

用户只说“画图”“绘制 Visio”“生成 vsdx”时，先问「要画 ER 图、流程图还是时序图？」，再读对应文件。

论文里往往有多类图，先对照能力边界再回答“画哪张”：

| 图类型 | 能画 | 不能画 |
|---|---|---|
| ER | 单个实体的属性图（1–10 个属性，含一个主键） | 系统 E-R 图、类图、用例图 |
| 流程图 | 模板槽位内的业务流程（3–8 个节点） | 泳道图、跨部门流程图 |
| 时序图 | 模板固定的 4–6 条泳道、14/18/22 条消息 | 其他泳道数或消息数 |

能力边界之外的图没有模板，直接说明不能画。“范围”以用户原话为准：只处理他点名的那几张，候选只从能画的对象里选，图号按论文实际顺序核对；数不出来就问用户要图号，不要凑一个看似合理的选项。选项文字只复述论文原文，不替用户加取舍标准（如“相关性高”“比较重要”）。

## 2. 确定依据
输入只有用户资料（论文文字、SQL、口头描述）和本 skill 的模板说明。

1. 在项目根下找 `.docx` / `.pdf` / `.md` 候选，只找论文，不顺带扫描旧图目录。
2. 用 `AskUserQuestion` 问「本次图依据哪份论文？」：最多 3 个候选加「还没有论文」，其他路径走 Other。
3. 有论文时先读目标图对应的正文章节，再做选模板、报图号之类的判断。只看目录、标题或图号就下结论，等于拿猜测当证据。
4. 只读**文字**：描述图的段落、字段表、流程/调用说明。论文里嵌的图不解析——那些正是要重画的对象。
5. 没有论文时：ER 图用描述或 SQL 生成；流程图和时序图没有流程/调用依据，告诉用户需要先提供论文或对应章节。

## 3. 生成
三种图统一两条：JSON 放 `json/<类型>/`，输出放 `draw/<类型>/`。`project_root` 用户没指定就用当前工作目录。

流程图和时序图的 JSON 顶层只有 `title`、`template` 和 `nodes`（流程图）或 `messages`（时序图）。多包一层 `user`/`type`/`data` 会让 `check` 直接失败。

时序图先按 `references/sequence.md` 让用户选模板，并填 `role`（第一泳道角色）和 `frontend_lane`（前端泳道对应的具体页面）。

1. 在 `<项目根>/json/<类型>/` 写 JSON。
2. `check` 一次，`ok=true` 再继续。
3. `draw` 生成 `.vsdx`。
4. `verify` 一次，从文件读回文字核对。
5. 只对报错的下标用 `set-text`，改完再 `verify`。重画整张图要再跑一遍 Visio 并覆盖整个文件，改一个形状不值得付这个代价。

```bash
# 在 skill 目录执行；<类型> 换成 flowchart 或 sequence
python -m scripts.<类型>.tools check --json <项目根>/json/<类型>/<图>.json
python -m scripts.<类型>.tools draw --json <项目根>/json/<类型>/<图>.json [--project-root <项目根>] [--overwrite]
python -m scripts.<类型>.tools verify --json <项目根>/json/<类型>/<图>.json [--project-root <项目根>]
python -m scripts.<类型>.tools set-text [--project-root <项目根>] --title <title> --template <模板> --index <下标> --text <新文字>
```

字数用 `len()` 判定（流程图 ≤4、时序图 ≤6），中英混排时目测不可靠。`draw` 成功只说明文件写入，所以还要 `verify` 读回确认。

ER 图只有 `er_draw` 一个入口：JSON `entity` + `attributes` 放 `json/er/<实体>.json`，输出 `draw/er/<实体>.vsdx`，没有覆盖保护。完整规则见 `references/er.md`。

三种图都适用：

- `project_root` 传给脚本时用绝对路径。Visio 收到相对路径会报“文件名无效”。
- 只调 `scripts/**/tools.py` 的现成命令。临时脚本和 `python -c` / `node -e` 在引号、转义上容易出错，结果也难核对。
- 不解压、不读取任何 `.vsdx`。模板结构已写在各 reference，重复解压既费上下文，结论还可能不一致。
- 不改 `template/` 下的原件：脚本复制后再编辑，改原件会让之后每次绘图都继承改动。
- 校验脚本是唯一判据，按报错修。
- 目标已存在时 `draw` 会拒绝；先说明路径和覆盖范围，确认后加 `--overwrite`。
- 结束前删除自己写的临时 JSON 和目录。

## 资源
- ER：`scripts/er/tools.py`、`scripts/er/er_contract.py`、`template/er/*.vsdx`
- 流程图：`scripts/flowchart/tools.py`、`template/flowchart/template_1..4.vsdx`
- 时序图：`scripts/sequence/tools.py`、`template/sequence/`（template_4/5 常用，1/2 六层，6/7 通用四泳道）
- 模板入库：`scripts/analyze_visio_template.py`、`scripts/register_visio_template.py`、`template/registry.json`
