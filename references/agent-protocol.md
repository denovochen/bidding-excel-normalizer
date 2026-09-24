# 平台协议 v1

平台使用 `scripts/excel_agent.py`；其内部复用既有解析引擎，没有第二套 Excel 解析实现。默认响应预算6000字符，`run --response-chars` 可在4000至12000间配置。应在部署平台实测工具截断阈值后选更小预算。正常交接退出0，异常退出2；stdout 是单个 JSON，宿主不得将命令状态文本拼接到该 JSON 文件。

## 调用

```bash
python scripts/excel_agent.py run <输入> --output <本次任务的空结果目录>
python scripts/excel_agent.py status --state <返回的session.json>
python scripts/excel_agent.py question --state <session.json> --id <question_id>
python scripts/excel_agent.py question --state <session.json> --id <question_id> --section columns --offset 0 --limit 8
python scripts/excel_agent.py inspect --state <session.json> --region <region_id> --start 10 --count 20 --columns A:H
python scripts/excel_agent.py resolve --state <session.json> --answers <返回的answer_file>
```

`question` 无 section 时返回证据字段索引；数组/对象分页，复杂对象可用 `answer_template.columns`、`source_block_examples.0` 等子路径。每次保留 `next_command`，超大条目返回子字段索引，不截断 JSON。`inspect` 自动计算闭区间，并按响应预算缩短返回行数；继续命令从实际下一行开始，不漏行。原有单元格显示长度限制通过 `values_may_be_truncated` 明示。

会话保存在输出目录旁的独立 `.excel-agent-*` 中；底层结构状态仍由原引擎维护。平台响应按大小选择当前批次问题，并生成新的答案文件：

```json
{
  "batch_id": "脚本生成",
  "state_version": 1,
  "answers": {"脚本给出的question_id": "填写本批答案"}
}
```

`answers` 内的结构答案遵循 semantic-review；项目关系和中标决定遵循相应参考。可以提交当前问题的非空子集。不同批次/陈旧版本/历史问题被拒绝；相同批次与完全相同答案的重试返回当前快照，不重复处理，冲突答案拒绝覆盖。提交期间加排他锁；若进程异常中断，报告中断并保留状态供核对，不盲目重试底层变更。跨进程崩溃的自动事务恢复尚未实现。

## 宿主职责

程序负责解析、分块、校验与发布；宿主按 `decision_owner` 调用模型或真实用户。语义判断执行环境仅需 question/inspect/resolve 的受限工具，不能因安装了 Skill 就认为任意 shell、文件工具已经被禁用。限制工具、裁剪模型历史、可信用户回执与预算中断需宿主实现，本仓库没有假装提供这些权限能力。

`decision_owner=user` 时，问题文字和候选值原样展示。无唯一推荐的候选以“不确定”为首项；支持空选择的宿主应不预选，不支持的宿主仍须等待用户实际提交。宿主应记录 question_id、候选值、文件指纹、actor 和回执 ID；模型写入的本地 JSON 不等于可信回执。

交付使用 `delivery` 的程序摘要。`review_reasons` 按 issue 代码汇总，其中受影响记录可能重叠，不能逐项相加得到队列总数。辅助审计提示独立统计；`unique_company_count` 是文本名称去重数量，不是已确认工商实体数。

底层 `excel_ledger.py` 的0/2/3/4/5退出码、plan兼容入口继续保留。平台会话和引擎 state 不混用。旧产物可只读 validate；1.8.0未完成会话应使用原版本完成，不能跨版本恢复。
