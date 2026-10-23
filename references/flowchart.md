# Flowchart Diagram

通用规则见 `SKILL.md` 的「1. 确定画什么」「2. 确定依据」「3. 生成」。本文件只写流程图特有的内容。

## 依据论文
按 `SKILL.md` 的「2. 确定依据」先确定论文。从文字里提取节点（业务流程段落、步骤说明），论文里嵌的流程图不解析——那张图正是要重画的对象。论文没写到的分支不要自行补。

## JSON
```json
{
  "title": "登录流程",
  "template": 1,
  "nodes": [
    {"role": "start", "text": "开始"},
    {"role": "process", "text": "输入账号"},
    {"role": "process", "text": "校验账号"},
    {"role": "decision", "text": "是否通过"},
    {"role": "process", "text": "生成令牌"},
    {"role": "process", "text": "提示失败"},
    {"role": "process", "text": "登录完成"},
    {"role": "end", "text": "结束"}
  ]
}
```
`nodes[i].role` 要与所选模板第 i 个槽位一致，否则文字会落到类型不对的形状上。

## 模板槽位
| Template | Nodes | Role Sequence |
|---|---|---|
| 1 | 8 | start, process, process, decision, process, process, process, end |
| 2 | 7 | start, input, process, decision, process, process, end |
| 3 | 7 | start, process, process, decision, process, process, end |
| 4 | 7 | start, process, process, decision, process, process, end |

模板 1 的 8 个槽位依次是：开始 → 处理 → 处理 → 判断 → 是分支 → 否分支 → 汇合 → 结束。模板 3 与 4 结构相同、版面不同；模板 2 多一个平行四边形输入框。

选模板时优先更复杂的 `1 > 2 > 3 > 4`；同一轮多张流程图不要每张都用同一个模板。

## 内容约束
- 每个形状不超过 4 个字（脚本用 `len()` 判定）。超出会挤压箭头和相邻形状。
- 只允许中文、英文、数字；符号和空白会让文字换行或溢出。
- `decision` 槽位以 `是否` 开头（如 `是否登录`），符合菱形判断框的语义。
- `input` 槽位含动作动词（编辑/写入/输入/填写/录入/添加/新建/创建），符合平行四边形输入框的语义。

## 输出
JSON 放 `json/flowchart/<title>.json`，输出 `draw/flowchart/<title>.vsdx`。
