# 输出契约 v1

最终目录恰好包含 final.csv、review_queue.csv、ledger.json。两个 CSV 均为 UTF-8 BOM、标准引号转义、CRLF，分别从1连续编号：

```text
序号,项目名称,项目编号,标段名称,标段编号,公司名称,中标与否,投标排名,文件类别,依据文件路径,来源页码,提取方式,证据文本,置信度,复核状态,解析结果生成日期时间
```

业务值有则填、无则空，不补造官方编号、标段、排名或中标结果。类别为 excel_ledger，提取方式为“台账整理”，来源为原文件名，页码为空，时间为带时区 ISO 8601。

置信度不是正确概率。脚本按提取来源、分组依据和中标对应方式计算规则可靠程度，逐条组成、依据和问题代码保存在 `records.confidence`；独立复核占位行记0。待复核不以提高分数或清零为目标。

存在投标列时，公司名称只能来自投标字段。仅清理全半角、空白、折行和完整企业后缀后的“投标”附注；不纠正错字、重复地名或异常后缀，不用中标名称覆盖。只有中标列的表按 award_company 保留原记录，不虚构投标来源。

先按项目块、区域、标段和 group_context 分组。项目仅块首填写时按 blocks 继承；投标数量每行重复不能作为组起点；相同项目和标段的不同批次、轮次或表区必须保持不同组。匹配和去重只能发生在组内。

每条 occurrence 保存稳定的 `source_block_id`。去重键至少包含来源投标块和企业规范名，因此同一企业参加不同标段或独立块时必须保留多条参与；`summary.cross_block_deduplication_count` 必须为0。映射了项目角色的 bidder roster 不允许发布 projectless 企业，`summary.bidder_roster_projectless_record_count` 必须为0。

1.8 的 `source_block_id` 根据原文件、Sheet、物理块起始行和投标列计算，不复制 plan 的 group ID。`source_blocks` 保留来源行与边界证据；同一表头拆成多个 table 仍共同检查原始边界，不能以拆表绕过候选完整性核验。标段无法对应到多个独立来源块时，参与记录全部保留，但停止中标推荐。

`source_coverage` 核验已映射企业来源行的提取/隔离守恒；非空原企业来源必须对应完整 occurrences 或独立复核。未解析区域另由 SOURCE_REGION_UNRESOLVED 等问题保存来源范围。`checks.artifact_consistency`、`checks.source_accounting` 与 `checks.review_queue` 分别报告文件一致性、来源核验和队列状态；validated 不表示所有业务问题已确认。旧产物没有来源清单时返回 source_accounting=not_available_legacy。

组内仅自动确认规范化后精确同名的投标企业，以及原表明确提供的中标状态。非精确名称不因 `row_aligned`、法人、金额、简称或相似度自动确认。

非精确名称生成一个人工确认任务。安全的单企业同行结构优先推荐同行投标企业；否则按组内名称接近程度推荐。推荐只影响提问工具的第一选项，不改变中标状态。法人、报价、中标金额及单位继续保留在来源审计中，但不参与自动匹配或推荐排序，缺失时不产生匹配问题。

用户可选择推荐的 `record_id`、通过 Other 输入企业名称或选择“不确定”。Other 文本必须在当前招标组内唯一对应已有投标名称；不能新增企业。人工选择后由 Python 重建结果，匹配依据记为 `user_selection`；选择不确定时保留原问题和空白中标状态。

中标问题展示项目、标段及投标/中标来源位置，避免同名项目的多个组混淆。只接受当前批次的回答。`confirmation_provenance=local_answer_file_not_host_verified` 表示回答来自本地答案文件，未由宿主可信回执证明；不能把该字段误称为已认证的用户身份。

明确对应填“是”。区域声明 complete，且本组边界、读取覆盖、中标名单和对应关系均核验通过时，其他企业才填“否”；partial 或 unknown 只保留能够确认的“是”，其他企业留空。没有任何中标结果时不得将全组标“否”。完整性声明及逐组核验原因保存在 `groups.award_completeness`。

原“是否中标”可映射；未知值留空，和名单结果冲突时留空复核。跨行合并的中标状态或排名不能扩散给多家企业。“未招投标”备注只保存采购上下文，不覆盖明确结果。

同组重复记录保留全部 occurrences。一致非空的排名和中标状态可以补齐；冲突字段留空并生成记录级问题，结果不依赖来源行顺序。字段级问题 ID 包含字段和单元格；名称对应问题包含中标单元格/原名称；组级问题按代码和明确作用域归并。

复核交互完成前，草稿和 state 仅保存在内部运行目录，不作为最终文件展示。完成后 `review_queue.csv` 复制仍未解决的 final 行，仅重编号并在证据前添加原因；用户已确认的问题从队列消失，选择不确定的问题继续保留。无法形成 final 的文件级或区域级问题作为独立复核行，不能进入 records、record_count、unique_company_count 或参与记录数。一个输入文件不可解析时，其他正常文件仍可交付；权限、输出损坏和内部不变量失败仍终止并返回 error。

ledger 保存 sources、mapping、matching_policy、projects/groups、records、issues、resolutions、row_audit、summary、unique_companies 和 CSV 哈希。启用跨表关系时还保存 relationships 与 relationship_resolutions。`groups.award_matches` 保存原中标名、单元格、推荐候选、人工选择、依据与状态；`occurrences.column_evidence` 保存上下文/辅助列。内部 ID 不是官方编号。

辅助列逐行证据通过 `reason_ref` 指向 mapping 中对应 table_id 的 column_dispositions，完整判断依据只保存一次，避免长说明随单元格重复膨胀。旧产物的内联 reason 继续可读。局部辅助字段错误只影响对应来源记录，不能把一个单元格的问题扩散到全组；项目/标段/名单完整性问题仍约束整组候选。

价格和法人这四个仅审计角色的未计算公式/错误值保存到 `audit_warnings`，保留字段、来源单元格及原值，不计入业务复核问题或修改企业复核状态。涉及同组重复记录字段冲突时仍进入业务复核，不能把可能的分组异常降为辅助提示。

新产物的覆盖摘要还包含 `cross_block_duplicate_participation_count`、`cross_block_deduplication_count`、`incomplete_bidder_group_count` 和 `bidder_roster_projectless_record_count`。候选范围不完整时 `award_matches.status=blocked`，不返回截断候选。旧 parser_version 的 plan 和产物继续按其原 summary 只读校验。

`relationships` 必须包含可追溯的来源项目快照、目标项目、候选和组级链接。项目关系确定但标段范围不可靠时，相关记录只保留项目字段，标段字段为空，并由 `LOT_SCOPE_UNRESOLVED` 关联到 `review_queue.csv`；不能把项目级“否”解释为任何具体标段的未中标结论。

真实单表范围或必须保留的结构超限属于文件级可恢复失败：整份文件记录为 `sources.status=unreadable`，保留文件名、哈希和包含 Sheet/触发位置/阈值的 error，并产生 `SOURCE_UNREADABLE` 独立复核行；其他来源继续处理，不交付该文件的截断数据。文件大小、解压规模、Sheet 数、有效单元格总数及结构展开规模限制仍为批次级错误。全部输入均可恢复失败时，允许发布仅有失败审计的三个产物，final.csv 只有表头，业务记录与企业数量为0；退出码0表示产物发布成功，不代表输入解析成功。

unique_companies 从 final 非空公司名按 NFKC、去空白、casefold 去重，保持首次顺序并包含待复核名称；它不是已确认工商实体数。CSV 公式形文本加单引号并保存可恢复记录。发布前校验名称来源、组归属、置信度、表头、统计、哈希和复核关联。

CSV 证据字段允许合法换行。记录数以标准 CSV parser 读取结果或 `ledger.summary.record_count/review_record_count` 为准，不能用物理文本行数减表头计算。

1.7 增加跨表项目关系审计。关系只在 plan 显式声明时启用；旧 plan 继续使用原单表逻辑。项目名称不精确时只保存脚本提供的有界候选和结构化决定；不能通过中标企业反向选择项目。旧版结果仍可只读 validate，新结果不得覆盖原目录。
