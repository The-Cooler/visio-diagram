# Template Intake

## 适用范围
只用于用户提供新 `.vsdx` 并要求分析/入库的情况。绘图任务中不要用这里的脚本或任何解压/XML 方式去探查内置模板；内置模板结构见 `flowchart.md` / `sequence.md` 的表格。

## 流程
1. 确认用户给了可读的 `.vsdx` 路径。
2. 先分析，不复制文件。
3. 用中文汇总：推断的图类型与置信度、节点/连接线数量、流向或消息顺序、复杂度、以及缺失文本槽位之类的警告。
4. 类型不确定时问用户属于 `er` / `flowchart` / `sequence` 哪一类。
5. 问用户是否加入模板库，得到明确同意（是/同意/加入/确认）后再入库。

## 分析命令
在 skill 目录执行，或传绝对路径：

```powershell
python scripts/analyze_visio_template.py "<模板路径.vsdx>"
python scripts/analyze_visio_template.py "<模板路径.vsdx>" --diagram-type flowchart
```

## 入库命令
得到同意后，加 `--confirmed`：

```powershell
python scripts/register_visio_template.py "<模板路径.vsdx>" --diagram-type flowchart --template-id custom_order_flow --confirmed
```

- `--diagram-type` 只能是 `er`、`flowchart`、`sequence`。
- `--template-id` 用小写字母、数字、下划线；省略时由源文件名生成 `custom_<source-name>`。
- 已有模板不会被覆盖，除非用户明确要求替换并加 `--replace`。

## 存储位置
- 模板：`template/<diagram-type>/<template-id>.vsdx`
- 元数据：`template/<diagram-type>/<template-id>.json`
- 注册表：`template/registry.json`

## 入库后的影响
新模板可以被绘图使用。如果它的槽位数或流向不在现有校验脚本覆盖范围内，要先更新对应的校验脚本，再拿它绘图；否则 `check` 会按旧模板的槽位报错。
