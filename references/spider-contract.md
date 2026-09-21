# 采集名单契约

仅在维护下游接入时读取；普通 Excel 整理不需要。Skill 只生成三个文件，不读取凭据或实际调用服务。

现有 Gitee pipeline.py 消费 document_completed.companies（:341），按 NFKC、去空白、casefold 去重（:357），再提交（:377）：

```text
POST /api/v1/company-crawl/jobs/{id}/companies
{"document_id":"来源标识","companies":["企业甲"],"event_seq":事件序号}
```

Gateway 的 store.py:377 按 (job_id, normalized_name) 去重并追加来源；crawler.py:216 每家调用一次：

```text
POST /spider/crawl
{"companyNames":["企业甲"],"fetchDeepInfo":false,"fetchBiddingDetail":false,"relationExpansionDepth":1}
```

随后保存 runId 并轮询。企业采集不要求项目/标段；不同作业无全局去重保证。database_only 不提交；database_then_crawl 仅在启用客户端预查时过滤已有企业。

Excel 的 ledger.unique_companies 提供相同规范化规则的全批次非空名单；后续编排按 Gateway 每批最多 1000 名提交并携带来源。需要取名单时调用 `python scripts/excel_ledger.py companies <产物目录>`，仅向调用程序输出 JSON，避免让模型展开全名单。它不会新增文件或发请求。

final 仍保留跨项目/组参与记录，缺失业务值允许空；纯名单不猜测项目归属。Report API 目前仍要求 mineru_job_id，格式兼容不表示已接通 Excel 报告。上述行号对应本次核实的客户端/Runtime 快照，不是线上实时状态。
