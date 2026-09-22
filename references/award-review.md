# 中标企业人工确认

仅在脚本返回 `kind=award_review_required` 时读取。该结果中的 `questions` 已符合内置“向用户提问”工具参数，不要补充候选、法人、金额、完整投标名单或相似度，也不要改写推荐项。

直接调用：

```text
ask_user_question({"questions": result.questions})
```

每个问题只提供三条交互路径：

1. Python 推荐的投标企业，标签以 `(Recommended)` 结尾，值为稳定 `record_id`。
2. “不确定”，值为 `unresolved`。
3. 工具的 Other 输入框，允许用户输入企业名称。

脚本每批最多返回5个问题。工具返回后，将 `answer` 对象原样保存为 state 同目录的 `answers.json`；不要把答案改写成公司别名或自行替用户选择。执行：

```bash
python scripts/excel_ledger.py resolve --state <state.json完整路径> --answers <answers.json完整路径>
```

- 再次返回 `award_review_required`：继续调用提问工具。`validation_errors` 非空时，简短说明对应输入未唯一匹配当前组，然后重问返回的问题。
- 返回 `result`：只交付最终三个文件。
- 返回 `error`：如实说明失败，不修改 state 或已发布目录。

Other 文本仅能唯一匹配当前招标组已有投标名称。找不到或同时命中多家时保持待确认，绝不新增企业。用户选择“不确定”后不再重问，该组继续保留在最终 `review_queue.csv`。

草稿产物和 state 是内部运行数据，不向用户展示，不读入模型上下文。人工选择逐批持久化；最终文件由 Python 根据 `record_id` 重建，模型不直接编辑 CSV 或 ledger。
