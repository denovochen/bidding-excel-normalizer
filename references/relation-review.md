# 项目关系复核 v1

仅在脚本返回 `kind=relationship_review_required` 时读取。

脚本每批最多返回 5 个项目问题。候选已按年度和名称相似度有界生成；模型只能从 `options` 选择候选项目或“不确定”，不能新增项目、使用中标企业反推项目，或把相似度当成自动确认阈值。

判断时比较项目身份语义，包括年份、行政区域、乡镇/街道、村或片区、建设/改造类型及资金来源。普通格式差异、`年/年度`、标点或明显单字排版错误可以支持同一项目；乡镇、村片、项目类型或资金来源冲突时选择“不确定”。

这是内部结构复核，不默认向用户展示。将模型选择保存为 state 同目录的 answers.json：

```json
{
  "relation_review_xxx": "project_xxx"
}
```

不能可靠确认时：

```json
{
  "relation_review_xxx": "unresolved"
}
```

然后执行：

```bash
python scripts/excel_ledger.py resolve --state <state.json完整路径> --answers <answers.json完整路径>
```

- 再次返回 `relationship_review_required`：继续处理下一批。
- 返回 `award_review_required`：转入 [中标企业确认](award-review.md)。
- 返回 `result`：交付最终三个文件。
- 返回 `error`：如实说明，不修改 state、原 Excel 或已发布目录。

关系决定由 Python 校验并写入 `ledger.relationship_resolutions`；模型不直接修改项目、企业或中标状态。
