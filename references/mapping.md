# 字段映射 v1

仅在 `mapping_required` 时读取。沿用 inspection.suggested_plan 中的文件名、哈希、Sheet 顺序；Agent 内部补全 plan.json，再执行 `run --plan`。缺失业务值留空，疑点进入复核，不向用户询问或生成临时清洗代码。

每个 Sheet 必须是 `parse` 或有 reason 的 `skip`。不要跳过无法识别的业务数据；隐藏 Sheet 也需覆盖。parse 的最小结构：

```json
{
  "name": "原Sheet名",
  "action": "parse",
  "ignored_rows": [{"start": 1, "end": 1, "reason": "标题"}],
  "tables": [{
    "header_rows": [2],
    "data_start_row": 3,
    "data_end_row": 100,
    "columns": {"bidder_name": "B"},
    "project_mode": "none",
    "group_mode": "source",
    "bidder_separator": "single",
    "award_list_complete": false,
    "award_mode": "auto",
    "summary_markers": ["项目汇总", "合计", "小计", "总计"],
    "non_tender_markers": ["未招投标", "未招标"]
  }]
}
```

范围和列标用实际值。所有非空行必须被表头、数据范围或有理由的 ignored_rows 覆盖；数据区间不能重叠。同 Sheet 混合结构拆成多个 tables，各区可复用同一 header_rows，分别设置 group_mode、award_mode 和 bidder_separator；同一合并项目仍保留统一 ID。仅保存 plan 对象，外层保持 `{"schema_version":1,"sources":[{"file_name":…,"sha256":…,"sheets":[…]}]}`。

`columns` 可用角色：project_name/code/serial/year/owner、agent、lot_name/code、bidder_name/count/price/legal_person、award_name/status/price/legal_person、rank、notes（斜线表示共用前缀，例如 project_code）。至少识别项目、标段或企业中的一项；不同角色不可映射同列。代理/实施主体不是投标企业，中标候选人不是中标企业。只列中标企业时仅映射 award_name；原“是否中标”映射 award_status。

| 选项 | 含义 |
|---|---|
| project_mode=none | 无项目名称/编号，project_id=null |
| merged | 按真实项目合并锚点继承，同名不同锚点不合并 |
| blocks | 无合并、仅块首有项目值，块内只继承项目上下文 |
| repeated | 平表按项目编号、名称、年度、实施主体组合 |
| group_mode=source | 纯名册按来源表区去重，不表示共同投标 |
| row | 每行一个组，单元格内企业名单展开 |
| anchor | 多行一组，另设 group_start_field 为已映射的组首字段，如 bidder_count |
| lot | 按明确标段名/编号及项目分组 |
| project | 已确认每项目只有一个招标组 |
| bidder_separator=single | 每格一家企业 |
| delimited | 括号外顿号/逗号/分号分隔；多行均有企业后缀则拆分，否则按折行处理 |
| lines | 明确每行一家企业，按换行及名单分隔符拆分 |
| award_mode=auto | 默认，由脚本检查单元格及组内证据选择行对应或名单匹配 |
| name_match | 明确采用组内名单匹配，适用一格多家、合并、集中或重复展示 |
| row_aligned | 已确认单企业行与中标明细逐行对应；脚本仍检查合并、名单、重复与跨行冲突 |

不全表填充企业、报价、排名；名单顺序不是排名。代理可能覆盖多个组，不能默认按代理分组；同名项目应按实际块区分。缺少组边界的行隔离，勿猜跨文件归属。联合体保留复核，不能拆成独立投标人。

按真实标段/数量锚点分组，不能跨组继承。只填首行不证明该行企业中标，不能根据非空比例指定 row_aligned；优先 auto。一格多家、跨行合并、重复展示或整组行都有值时按组内名单匹配。row_presence 已停用。项目汇总只作审计，不增加一组或企业记录。

完整中标结果设 award_list_complete=true；部分结果或无结果设 false。缺少结果不等于全部未中标。中标对应由脚本完成，输出保留投标名称；不调用模型判断名称，不在 plan 中写公司名映射。summary_markers 完整匹配项目/企业/代理字段；non_tender_markers 仅保存原采购方式，不覆盖明确中标事实，不独立触发复核。

结构信息不足时仅局部 inspect；无需读取全表或源码。核对产物规则时才读 [输出契约](output-contract.md)。
