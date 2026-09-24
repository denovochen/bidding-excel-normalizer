# 中标企业人工确认

`decision_owner=user` 表示必须使用真实用户回复。脚本拒绝当前批次之外的答案，但本地 JSON 不能证明是谁作出的选择；`confirmation_provenance=local_answer_file_not_host_verified` 明确记录这一限制。只有宿主提供可信提问回执才能进一步验证来源，不能把模型自行选择描述为用户确认。宿主没有提问能力时保存 unresolved，继续交付带复核项的结果。

仅在脚本返回 `kind=award_review_required` 时读取。该结果中的 `questions` 已符合内置“向用户提问”工具参数，不要补充候选、法人、金额、完整投标名单或相似度，也不要改写推荐项。

直接调用：

```text
ask_user_question({"questions": result.questions})
```

有唯一推荐时保留以下交互路径：

1. Python 推荐的投标企业，标签以 `(Recommended)` 结尾，值为稳定 `record_id`。
2. “不确定”，值为 `unresolved`。
3. 工具的 Other 输入框，允许用户输入企业名称。

名称评分并列时不设置推荐项，问题保留少量有界候选和“不确定”，不确定排在首项；候选排序不代表证据更强。支持不预选的宿主应不预选，始终等待真实用户提交。超预算问题标记 `details_required` 时，先按字段分页取得原始 question/options，再展示，不能让用户回答仅含索引的问题。

每批最多5个问题，平台响应可能按预算进一步缩小批次。工具返回后，将答案对象保存到返回的 `answer_file.answers`，保留 batch_id/state_version；不要改写公司别名或替用户选择。执行返回的 `next_command`。不累计历史答案。

- 再次返回 `award_review_required`：继续调用提问工具。`validation_errors` 非空时，简短说明对应输入未唯一匹配当前组，然后重问返回的问题。
- 返回 `result`：只交付最终三个文件。
- 返回 `error`：如实说明失败，不修改 state 或已发布目录。

Other 文本仅能唯一匹配当前招标组已有投标名称。找不到或同时命中多家时保持待确认，绝不新增企业。用户选择“不确定”后不再重问，该组继续保留在最终 `review_queue.csv`。

草稿产物和 state 是内部运行数据，不向用户展示，不读入模型上下文。人工选择逐批持久化；最终文件由 Python 根据 `record_id` 重建，模型不直接编辑 CSV 或 ledger。
