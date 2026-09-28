# 平台协议 v2 / Skill 1.8.2

平台使用 `scripts/excel_agent.py`；其内部复用既有解析引擎，没有第二套 Excel 解析实现。默认响应预算6000字符，`run --response-chars` 可在4000至12000间配置。应在部署平台实测工具截断阈值后选更小预算。正常交接退出0，异常退出2；stdout 是单个 JSON，宿主不得将命令状态文本拼接到该 JSON 文件。

## 调用

```bash
python scripts/excel_agent.py run <输入> --output <本次任务的空结果目录>
python scripts/excel_agent.py status --state <返回的session.json>
python scripts/excel_agent.py question --state <session.json> --id <question_id>
python scripts/excel_agent.py question --state <session.json> --id <question_id> --section columns --offset 0 --limit 8
python scripts/excel_agent.py inspect --state <session.json> --region <region_id> --start 10 --count 20 --columns A:H
# 复制返回的 answer_command，再追加 --id 和 --json；批次字段不自行编造
python scripts/excel_agent.py answer --state <session.json> --batch-id <batch_id> --state-version <版本> --id <question_id> --json '{"columns":{"C":"evidence"},"basis":"C列为辅助说明，其他列采用已核对的模板角色"}'
```

`run` 默认先返回 `task_type=scope`，目标为投标企业参与明细及关联汇总；可用 `--scope '用户指定业务范围'` 覆盖。范围按实际资料语义选择，不按位置/表名过滤。原引擎 `excel_ledger.py run` 未传 --scope 时仍走全量流程。1.8.2 的包版本由 doctor.skill_version 标识；parser_version 保持1.8.1，读取协议v1旧会话时不补插范围阶段或重跑。新会话写 protocol_version=2，防止旧入口误读新范围状态；新会话请用1.8.2继续。

`answer --id` 接受单个问题的答案对象/字符串，不带 answers 外壳。省略 --id 时 --json 为当前问题ID到答案的映射，适用于同批多个已确定答案，统一校验后一次提交，不增加逐题轮次。结构题在模板上合并 columns 和其他指定字段，未修改的有依据字段保留；basis 必须自行填写。关系/中标选择支持 --json '"unresolved"'。范围题的 exclude 以 sheet_id 为键，未排除表保留；草稿中用 null 撤销某项排除。详情见 scope-selection.md。

单次 --json 最多2000字符，basis最多256字。普通问题一次提交；宽表按一组字段用 --draft 分段保存，最后一次去掉 --draft 提交。answer_saved 不推进批次；status --resume 核验后可继续该草稿。不调用 edit_file/write_file，不读取答案文件来猜缩进。answer 自动提交，不再额外执行同批 resolve。语义错误仍由现有引擎处理；未解释列、未知字段、缺失排除理由先明确报错，不消耗结构判断次数。非空理由的业务正确性不能仅靠格式校验保证。

v2交接使用 answer_command；不再额外返回旧 resolve 的 next_command，以免重复提交并占用响应预算。问题摘要的 answer_defaults 不重复列映射，列建议见 columns[].suggested_role；完整模板仍可用 question --section answer_template 读取。

`question` 无 section 时返回证据字段索引；数组/对象分页，复杂对象可用 `answer_template.columns`、`source_block_examples.0` 等子路径。每次保留 `next_command`，超大条目返回子字段索引，不截断 JSON。`inspect` 自动计算闭区间，并按响应预算缩短返回行数；继续命令从实际下一行开始，不漏行。原有单元格显示长度限制通过 `values_may_be_truncated` 明示。

会话保存在输出目录旁的独立 `.excel-agent-*` 中；底层结构状态仍由原引擎维护。平台响应按大小选择当前批次问题，并生成新的答案文件：

```json
{
  "batch_id": "脚本生成",
  "state_version": 1,
  "answers": {"脚本给出的question_id": "填写本批答案"}
}
```

这些文件由程序管理，旧 `resolve --answers` 仅供兼容与测试。结构答案遵循 semantic-review，项目关系和中标决定遵循相应参考。不同批次/陈旧版本/历史问题被拒绝；同一批次已提交字段的相同重试返回当前快照，不重复处理，冲突答案拒绝覆盖。提交期间加排他锁；若进程异常中断，报告中断并保留状态供核对，不盲目重试底层变更。跨进程崩溃的自动事务恢复尚未实现。

## 宿主职责

`workflow` 是程序生成的固定待办投影，名称和顺序始终为读取并检查Excel、清洗并整理数据、生成并校验结果；范围、结构、表关系、项目关系及企业确认只改变第二项内部 phase。sync_tool=write_todos、sync_before_next_command=true 表示先原样同步 workflow.todos，再读取问题/继续命令；不从模型说明推断进度。这是工具调用约定，不是宿主强制拦截器。

`status` 只读核验当前会话、输入指纹、引擎状态、答案批次和已提交回执，不重新解析Excel或重写答案；已发布状态另核验产物。默认返回暂停状态；明确恢复时使用 `status --resume`，核验通过后把当前阶段投影为进行中，再继续返回的同一批次。旧1.8.1会话保持兼容。中断或遗留锁不能直接重新 run/resolve，必须先核对提交结果。

暂停核验和错误响应不会把当前项标为 in_progress。模型流式中断发生在Skill未运行时，Skill无法主动修改宿主状态；宿主须在失败/取消/中断回调中停止显示运行中，再在恢复核验完成后同步程序状态。只在提示词中要求调用 write_todos，不能提供宿主级强制保证。

程序负责解析、分块、校验与发布；宿主按 decision_owner 调用模型或真实用户。正常语义阶段使用 question/inspect/answer，不需要手工文件编辑。安装 Skill 不会禁用任意 shell、文件工具。限制工具、裁剪模型历史、可信用户回执与预算中断需宿主实现，本仓库没有假装提供这些权限能力。

`decision_owner=user` 时，问题文字和候选值原样展示。无唯一推荐的候选以“不确定”为首项；支持空选择的宿主应不预选，不支持的宿主仍须等待用户实际提交。宿主应记录 question_id、候选值、文件指纹、actor 和回执 ID；模型写入的本地 JSON 不等于可信回执。

交付只使用 files 中的 final.csv、review_queue.csv；internal_files 中的 ledger.json 仍在目录内生成和校验，不能附给用户。无需询问交付格式。delivery 为程序摘要，excluded_sheet_count 表示业务排除数量；具体目标与理由留在审计。review_reasons 按 issue 汇总，受影响记录可能重叠，不能逐项相加。辅助审计提示独立统计；unique_company_count 是名称文本去重数量，不是已确认工商实体数。

平台入口验收须确认实际运行快照已加载 bidding-excel-normalizer。当前目标平台的预加载取已启用 Skills 的交集，因此专用智能体应同时启用并预加载；仅安装包不等于实际加载。本仓库不修改平台配置、服务或超时。

底层 `excel_ledger.py` 的0/2/3/4/5退出码、plan兼容入口继续保留。平台会话和引擎 state 不混用。旧产物可只读 validate；1.8.0未完成会话应使用原版本完成，不能跨版本恢复。
