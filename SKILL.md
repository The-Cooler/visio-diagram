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

ER 图的 JSON 是 `{"entity": "...", "attributes": [...]}`，放 `json/er/<实体>.json`，输出 `draw/er/<实体>.vsdx`，没有覆盖保护。细节见 `references/er.md`。

时序图先按 `references/sequence.md` 让用户选模板，并填 `role`（第一泳道角色）和 `frontend_lane`（前端泳道对应的具体页面）。

1. 在 `<项目根>/json/<类型>/` 写 JSON。
2. `check` 一次，`ok=true` 再继续。
3. `draw` 生成 `.vsdx`。
4. `verify` 一次，从文件读回文字核对。
5. `verify` 报错时按下面的方式修正，然后重新 `verify`。重画整张图要再跑一遍 Visio 并覆盖整个文件，改一个形状不值得付这个代价。

**check 报错后怎么改 JSON**

`check` 的错误信息会点名 `nodes[i]` 或 `messages[i]`。用 Edit 工具改那个键的值，一次一处：

- 只超字数 → 只改那一条的 `"text"`，例如 `"生成令牌"` → `"发令牌"`。
- 只流向不对 → 只改那一条的 `"from"` / `"to"`。
- 多条报错 → 逐条 Edit，一条一次。

不要用 Write 重新生成整份 JSON，也不要用脚本或命令行去改写 JSON。重新生成会让没有报错的条目一起被替换，多出的改动无法核对；Edit 每次只动一个键，改了什么一目了然。

**verify 报错后怎么改文件**

`verify` 读回的是生成好的 `.vsdx`，这时不用再碰 JSON，直接用 `set-text` 改对应下标；改完再 `verify` 一次即可。

```bash
# 在 skill 目录执行；<类型> 换成 flowchart 或 sequence
python -m scripts.<类型>.tools check --json <项目根>/json/<类型>/<图>.json
python -m scripts.<类型>.tools draw --json <项目根>/json/<类型>/<图>.json [--project-root <项目根>] [--overwrite]
python -m scripts.<类型>.tools verify --json <项目根>/json/<类型>/<图>.json [--project-root <项目根>]
python -m scripts.<类型>.tools set-text [--project-root <项目根>] --title <title> --template <模板> --index <下标> --text <新文字>
```

字数用 `len()` 判定（流程图 ≤4、时序图 ≤6），中英混排时目测不可靠。`draw` 成功只说明文件写入，所以还要 `verify` 读回确认。

## 接口清单
下面就是全部对外接口和返回结构。**不要为了确认签名去读 `.py` 源码**；这里没有的才需要读。

导包方式：`from scripts.flowchart.tools import ...`、`from scripts.sequence.slots import ...`、`from scripts.er.tools import ...`。

### CLI（在 skill 目录执行）
| 命令 | 作用 |
|---|---|
| `python -m scripts.flowchart.tools {check,draw,verify,set-text}` | 流程图四个步骤 |
| `python -m scripts.sequence.tools {templates,check,draw,verify,set-text}` | 时序图五个步骤，`templates` 列模板 |
| `python -m scripts.er <json> [--project-root] [--user]` | ER 校验并出图 |
| `python scripts/analyze_visio_template.py <vsdx> [--diagram-type]` | 分析新模板 |
| `python scripts/register_visio_template.py <vsdx> --diagram-type X --confirmed` | 注册新模板 |
| `python scripts/flowchart/flowchart_contract.py --json '<json>'` | 独立契约校验 |
| `python scripts/sequence/sequence_contract.py --json '<json>'` | 独立契约校验 |

`--project-root` 省略时用当前工作目录；`--overwrite` 只在确认覆盖后加。

### Python 函数
可直接 `from scripts.<类型>.tools import ...` 调用，参数与 CLI 一致。

**流程图**（`scripts.flowchart`）
| 函数 | 参数 | 返回 |
|---|---|---|
| `flowchart_check` | `(template, nodes)` | `{ok, errors, max_length, lengths, template}` |
| `flowchart_draw` | `(project_root, user, title, template, nodes, overwrite=False)` | `{success, message}` |
| `flowchart_verify` | `(project_root, title, template)` | `{ok, errors, max_length, shapes:[{index,text,length}]}` |
| `flowchart_set_node` | `(project_root, title, template, index, text)` | `{success, message}` |
| `roles_for_template` | `(template)` | 槽位 role 元组 |
| `validate_flowchart_content` | `({template, nodes})` | `[错误字符串]` |

常量：`FLOWCHART_TEXT_MAX_LEN = 4`，`ROLE_SHAPE_HINT`（role → 形状说明）。

**时序图**（`scripts.sequence`）
| 函数 | 参数 | 返回 |
|---|---|---|
| `sequence_check` | `(template, messages, frontend_lane=None, role=None)` | `{ok, errors, max_length, lengths, frontend_lane, role}` |
| `sequence_draw` | `(project_root, user, title, messages, template, frontend_lane=None, role=None, overwrite=False)` | `{success, message}` |
| `sequence_verify` | `(project_root, title, template, frontend_lane=None, role=None)` | `{ok, errors, max_length, messages:[{index,text,length}], frontend_lane, role}` |
| `sequence_set_message` | `(project_root, title, template, index, text)` | `{success, message}` |
| `template_options` | `()` | `[{template,name,lanes,messages,desc}]` |
| `slots_for_template` | `(template)` | `((from,to), ...)` |
| `first_lane_name` | `(template)` | 第一泳道默认名 |

常量：`SEQUENCE_TEXT_MAX_LEN = 6`，`FRONTEND_LANE_MAX_LEN = 8`。

**ER**（`scripts.er`）
| 函数 | 参数 | 返回 |
|---|---|---|
| `er_draw` | `(project_root, json_path, user=None)` | `{success, message}` |
| `validate_er_payload` | `(payload)` | `[错误字符串]` |
| `normalize_attrs` | `(attributes)` | 主键排首位的属性列表 |
| `choose_template_name` | `(attr_count)` | `"ER_odd.vsdx"` 或 `"ER_even.vsdx"` |

常量：`ER_TEMPLATE_ATTR_SLOTS = 10`。

**模板入库**
| 函数 | 参数 | 返回 |
|---|---|---|
| `analyze_vsdx` | `(path, diagram_type=None)` | 页面、形状、连接、流向、复杂度、推断类型 |
| `register_template` | `(source, diagram_type, template_id, replace, confirmed)` | 写入后的元数据 |

三种图都适用：

- `project_root` 传给脚本时用绝对路径。Visio 收到相对路径会报“文件名无效”。
- 只调 `scripts/**/tools.py` 的现成命令。临时脚本和 `python -c` / `node -e` 在引号、转义上容易出错，结果也难核对。
- 写 JSON 用 Write 一次成稿；之后改 JSON 一律用 Edit 改对应键，不用 Write 整份重写，也不写脚本改 JSON。
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
