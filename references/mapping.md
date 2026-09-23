# 字段映射 v1

每次 `run` 都先返回 `mapping_required`，供 Agent 轻量审阅每份文件和各结构区域。沿用 `inspection.suggested_plan` 的文件名、哈希和 Sheet 顺序，在工作区保存 plan.json 后执行 `run --plan`。模型只理解结构并修订声明式计划，不逐行处理企业、不判断名称对应、不修改解析源码。

每个 Sheet 使用 `parse`、`review` 或有 reason 的 `skip`。`skip` 仅适合空表或已明确非业务内容；非空 skip 仍会生成复核项。无法可靠解释的业务区域使用 `review_regions`，不能跳过后宣称全部完成。parse 示例：

```json
{
  "name": "原Sheet名",
  "action": "parse",
  "ignored_rows": [{"start": 1, "end": 1, "reason": "标题"}],
  "review_regions": [],
  "tables": [{
    "header_rows": [2],
    "data_start_row": 3,
    "data_end_row": 100,
    "columns": {"project_name": "A", "bidder_name": "C", "award_name": "D"},
    "column_dispositions": [
      {"column": "B", "disposition": "group_context", "mode": "repeated", "header": "批次", "reason": "批次构成独立招标组边界"},
      {"column": "E", "disposition": "evidence", "header": "其他说明", "reason": "仅保留来源证据，不进入16列业务字段"}
    ],
    "project_mode": "blocks",
    "group_mode": "project",
    "bidder_separator": "single",
    "award_completeness": {"status": "complete", "basis_type": "structural", "basis": "该区域完整列出投标企业并记录最终中标结果"},
    "award_mode": "auto",
    "structure_warnings": [],
    "summary_markers": ["项目汇总", "合计", "小计", "总计"],
    "non_tender_markers": ["未招投标", "未招标"]
  }]
}
```

XLSX 有效内容由 OOXML 稀疏预扫描确定，空白样式不会产生业务列；0、False、公式、错误及隐藏单元格中的内容不能丢弃。保留原坐标和相关合并，表头仍通过合并锚点继承。超过读取范围的真实内容或必须保留的合并按整份文件失败记录在 `inspection.source_failures`，不能在 plan 中静默忽略超限列；其余文件继续结构审阅。

所有有效列必须有明确去向。`columns` 保存业务字段和匹配辅助字段；其余列逐项写入 `column_dispositions`：

| disposition | 含义 |
|---|---|
| context | 项目或区域上下文，保存在审计中 |
| evidence | 辅助来源证据，不进入最终业务列 |
| group_context | 批次、轮次等内部组边界；`mode` 为 repeated 或 blocks |
| ignore | 明确无关列，必须写具体原因 |
| unrecognized | 尚未识别；存在时禁止执行并继续返回 mapping_required |

`columns` 可用角色：project_name/code/serial/year/owner、agent、lot_name/code、bidder_name/count/price/legal_person、award_name/status/price/legal_person、rank、notes。至少识别项目、标段或企业中的一项；不同角色不能映射同列。代理、实施主体和中标候选人不能当成投标企业。存在投标列时必须映射 bidder_name；只列中标企业时才仅映射 award_name。

所有非空行必须由表头、数据范围、`ignored_rows` 或 `review_regions` 覆盖，数据区间不能重叠。同 Sheet 混合结构拆为多个 tables。`structure_warnings` 中的新表头或结构变化必须通过拆区或复核区域消解后清空。补映射最多尝试两次；仍无可靠映射时将最小范围写入 `review_regions`，保留其他已确定结果。

| 分组选项 | 含义 |
|---|---|
| project_mode=none | 无项目名称/编号，project_id=null |
| merged | 按真实项目合并锚点继承 |
| blocks | 无合并、仅块首填写项目，块内继承项目上下文 |
| repeated | 平表按项目编号、名称、年度、实施主体组合 |
| group_mode=source | 纯名册按来源区域去重，不表示共同投标 |
| row | 每行一个组，单元格内企业名单展开 |
| anchor | 多行一组，group_start_field 必须是真实稀疏或合并锚点 |
| lot | 按项目及明确标段分组 |
| project | 已确认每项目只有一个招标组 |

投标数量每行重复时不能作为 anchor。相同项目和标段存在不同批次、轮次或独立块时，优先把对应列设为 group_context；也可拆为多个 table。组 ID 包含区域边界，Python 只在组内匹配和去重。

`bidder_separator` 可为 single、delimited 或 lines。名单顺序不是排名；联合体无法确认成员关系时进入复核。`award_mode=auto` 检查是否可安全把同行企业作为人工确认推荐项；`name_match` 明确只按组内名称推荐；`row_aligned` 表示区域已确认逐行对应，但非精确名称仍需用户选择。三种模式都只自动确认精确名称，并执行合并、名单、重复和跨行冲突检查。

`award_completeness.status` 为 complete、partial 或 unknown；`basis_type` 为 explicit、structural 或 none，basis 必须写区域级依据。不能仅因存在中标列就设 complete，也不能因中标单元格稀疏就设 partial。Python 会逐组核验边界、读取覆盖和名称对应；只有 complete 且核验通过的组才把未匹配企业标“否”。

缺失业务值留空，不补造编号、排名或结果。输出保留投标名称；plan 中不写公司别名或逐对名称决定。非精确中标名称在执行阶段按 [中标企业人工确认](award-review.md) 处理。需要补证据时使用 `inspect --sheet <名称> --rows <起始:结束>`，一次最多100行，不读取完整 ledger 或全部企业名单。
