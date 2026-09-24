# 项目关系复核 v1

仅在脚本返回 `kind=relationship_review_required` 时读取。

脚本每批最多返回5个项目问题，平台入口还会按大小缩小批次。候选已按显式冲突、年度和名称相似度排序；只能从 `options` 选择有依据的项目或“不确定”，不能新增项目、使用中标企业反推项目，或把相似度当成自动确认阈值。

判断时比较项目身份语义，包括年份、行政区域、乡镇/街道、村或片区、建设/改造类型及资金来源。普通格式差异、`年/年度`、标点或明显单字排版错误可以支持同一项目；乡镇、村片、项目类型或资金来源冲突时选择“不确定”。

主工程与增做/追加工程不能直接等同。同名同年度的不同来源块，必须有唯一项目编号或实施主体等区分依据；不能选第一个。脚本标记 selection_blockers 的候选不可提交，全部候选均受阻时直接保留复核，不强迫模型在错误候选中选择。

这是内部结构复核，不默认向用户展示。将选择保存到返回的 answer_file.answers，保留批次信息；下例仅为 answers 成员：

```json
{
  "relation_review_xxx": {
    "value": "project_xxx",
    "basis": "说明项目身份一致且不存在范围冲突的依据",
    "evidence": ["返回的汇总来源", "返回的目标名册来源"]
  }
}
```

不能可靠确认时：

```json
{
  "relation_review_xxx": "unresolved"
}
```

选择项目时须填写非空 basis，并原样引用该候选的 `evidence_refs`，不能用其他项目单元格替代。引用存在只证明来源可追溯，不自动证明语义正确。然后执行返回的 `next_command`。

- 再次返回 `relationship_review_required`：继续处理下一批。
- 返回 `award_review_required`：转入 [中标企业确认](award-review.md)。
- 返回 `result`：交付最终三个文件。
- 返回 `error`：如实说明，不修改 state、原 Excel 或已发布目录。

关系决定由 Python 校验并写入 `ledger.relationship_resolutions`，新模型决定保存 actor=model、basis=model_selection 及候选证据；不会标为用户确认。模型不直接修改项目、企业或中标状态。
