# visio-diagram

用 Microsoft Visio 生成 `.vsdx` 图的 Claude Skill，支持 ER 图、流程图、时序图，以及分析/注册用户提供的新模板。

> **迁移说明**：本项目原为 Agent 项目（入口为 `agents/openai.yaml`），现已改为 Claude Skill。现在的入口是 `SKILL.md` 的 frontmatter（`name` + `description`），由 Claude Code 自动发现和触发；`agents/openai.yaml` 仅作为旧入口保留，不再维护。

## 目录结构

```
visio-diagram/
├── SKILL.md                  # Skill 入口：触发描述、类型路由、论文证据、生成流程
├── README.md
├── references/               # 按图类型拆分的规则，按需加载
│   ├── er.md
│   ├── flowchart.md
│   ├── sequence.md
│   └── template-intake.md    # 新模板分析与入库
├── scripts/
│   ├── er/                   # ER：tools.py（er_draw）、er_contract.py、__main__.py
│   ├── flowchart/            # 流程图：tools.py、spec.py、flowchart_contract.py
│   ├── sequence/             # 时序图：tools.py、slots.py、sequence_contract.py
│   ├── analyze_visio_template.py
│   └── register_visio_template.py
├── template/
│   ├── er/                   # ER_odd.vsdx / ER_even.vsdx
│   ├── flowchart/            # template_1.vsdx … template_4.vsdx
│   ├── sequence/             # template_1/2/4/5/6/7.vsdx
│   └── registry.json
└── agents/openai.yaml        # 旧 Agent 入口（保留，不再使用）
```

## 环境要求

- Windows，已安装 Microsoft Visio
- Python 3.9+
- `pywin32`（提供 `pythoncom`、`win32com.client`）

## 安装

把整个目录放到 Claude 的 skills 目录下：

```
~/.claude/skills/visio-diagram/
```

或使用符号链接指向本仓库。重启会话后，Claude 会自动读取 `SKILL.md` 的 frontmatter。

## 使用方式

向 Claude 提出绘图需求即可，例如：

- “把论文第三章的用户实体属性图画成 Visio”
- “根据这份论文画登录流程的时序图”
- “帮我注册这个新的 vsdx 模板”

Claude 会按 `SKILL.md` 的流程执行：

1. **确认类型** —— 读对应的 `references/*.md`
2. **对照能力边界** —— 确认要画的图在本技能范围内
3. **确认依据** —— 有论文就只读文字，不解析论文里嵌的图；没有论文时 ER 可用描述/SQL，流程图和时序图无法生成
4. **生成** —— 写 JSON → `check` → `draw` → `verify`，必要时 `set-text`

## 命令速查

所有命令都在本 skill 目录下执行。

### 流程图

```bash
python -m scripts.flowchart.tools check  --json <项目根>/json/flowchart/<图>.json
python -m scripts.flowchart.tools draw   --json <项目根>/json/flowchart/<图>.json [--project-root <项目根>] [--overwrite]
python -m scripts.flowchart.tools verify --json <项目根>/json/flowchart/<图>.json [--project-root <项目根>]
python -m scripts.flowchart.tools set-text [--project-root <项目根>] --title <title> --template <1-4> --index <下标> --text <新文字>
```

### 时序图

```bash
python -m scripts.sequence.tools templates        # 列出模板、泳道、消息数和说明
python -m scripts.sequence.tools check  --json <项目根>/json/sequence/<图>.json
python -m scripts.sequence.tools draw   --json <项目根>/json/sequence/<图>.json [--project-root <项目根>] [--overwrite]
python -m scripts.sequence.tools verify --json <项目根>/json/sequence/<图>.json [--project-root <项目根>]
python -m scripts.sequence.tools set-text [--project-root <项目根>] --title <title> --template <template_N> --index <下标> --text <新文字>
```

### ER 图

```bash
python -m scripts.er <项目根>/json/er/<实体>.json [--project-root <项目根>]
```

### 模板入库

```bash
python scripts/analyze_visio_template.py "<新模板.vsdx>" [--diagram-type flowchart]
python scripts/register_visio_template.py "<新模板.vsdx>" --diagram-type flowchart --template-id custom_xxx --confirmed
```

`--confirmed` 只有在用户明确同意入库后才能加；`--replace` 仅在用户明确要求替换旧模板时使用。

## 路径约定

| 内容 | 位置 |
|---|---|
| 输入 JSON | `<project_root>/json/<类型>/<名称>.json` |
| 输出 vsdx | `<project_root>/draw/<类型>/<名称>.vsdx` |
| 未指定 `project_root` | 使用当前工作目录 |

三种图统一：ER 不再使用 `draw/<user>/er/` 这一层。

## 模板

### ER

按属性数量奇偶选择：奇数 `ER_odd.vsdx`，偶数 `ER_even.vsdx`。模板提供 10 个属性槽，属性不足时多余形状会被删除。

### 流程图

| 模板 | 节点数 | 槽位顺序 |
|---|---|---|
| template_1 | 8 | start, process, process, decision, process, process, process, end |
| template_2 | 7 | start, input, process, decision, process, process, end |
| template_3 | 7 | start, process, process, decision, process, process, end |
| template_4 | 7 | start, process, process, decision, process, process, end |

### 时序图

| 模板 | 消息数 | 泳道 | 说明 |
|---|---|---|---|
| template_1 | 22 | 管理员/前端/控制器/服务层/Mapper/数据库 | Spring Boot 六层 |
| template_2 | 22 | 用户/前端/路由层/服务层/仓储层/存储层 | Python 六层 |
| template_4 | 18 | 管理员/前端/控制器/服务层/数据库 | Spring Boot 五层（常用） |
| template_5 | 18 | 管理员/前端/路由层/服务层/数据库 | Python 五层（常用） |
| template_6 | 14 | 管理员/前端/服务层/数据库 | 通用四泳道（服务层） |
| template_7 | 14 | 管理员/前端/控制器/数据库 | 通用四泳道（控制器，推荐） |

选模板前必须让用户确认，选项里要写出泳道。

## 内容约束

| 项目 | 限制 | 判定方式 |
|---|---|---|
| 流程图每个形状 | ≤ 4 个字 | 代码 `len()` |
| 时序图每条消息 | ≤ 6 个字 | 代码 `len()` |
| 时序图 `frontend_lane` | ≤ 8 个字，不能仍是“前端” | 代码 `len()` |
| ER 属性数 | 1–10 条，最多一个 `pk` | `er_contract` |
| 流程图文字字符集 | 中文/英文/数字 | 正则 |

时序图的两条泳道标题会按业务语境替换：`role` 换第一泳道（管理员/用户），`frontend_lane` 换前端泳道。替换只改文字，不动字号、位置、连接线和其他泳道。

## 设计原则

- **文字优先**：绘图依据是论文的文字描述、字段表和流程说明；论文里嵌的图不解析，因为那正是要重画的对象。
- **校验用代码**：字数、槽位数、流向都由脚本判定，不依赖目测。
- **读回确认**：`draw` 成功后用 `verify` 从生成文件读回文字复核。
- **单点修改**：字数不对只用 `set-text` 改对应形状，不重画整张图。
- **不覆盖模板**：脚本复制模板后再编辑，`template/` 下的原件始终不变。
- **不静默覆盖输出**：目标文件已存在时 `draw` 会拒绝，确认后加 `--overwrite`。

## 已知情况

- `template/sequence/template_3.vsdx.tmp` 未在 `registry.json` 中登记，也不在任何模板表中，不参与绘图。
- `scripts/er/er_contract.py` 中的 `er_draw` 是早期实现，输出路径与当前约定不同；实际使用 `scripts/er/tools.py` 的 `er_draw`，入口为 `python -m scripts.er`。
