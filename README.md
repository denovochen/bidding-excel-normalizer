# bidding-excel-normalizer

独立的招投标 Excel 整理 Skill，当前版本 **1.4.0**。输入 XLS/XLSX，交付 final.csv、review_queue.csv、ledger.json。

脚本读取、分组、清洗、对应中标记录并校验；模型仅在陌生表头需要内部补映射时参与。取消逐对名称判断、决定文件、固定进度计划和具体公司名映射。存在投标列时，最终公司名称始终来自投标列；中标栏仅用于对应结果与审计，不替换、纠正投标全称。

## 文件位置

| 路径 | 用途 |
|---|---|
| SKILL.md | 精简执行入口及交付要求 |
| scripts/excel_ledger.py | 运行、局部检查、产物校验及读取去重名单 |
| scripts/ledger_core/ | 结构读取、确定性中标匹配、审计和发布 |
| references/mapping.md | 陌生表头和混合区域的内部映射 |
| references/output-contract.md | 三个产物、匹配依据及异常规则 |
| tests/ | 合成表格回归及外部真实样本验收 |

## 开发验证

```bash
python -B -m unittest discover -s tests
python scripts/excel_ledger.py run <Excel路径> --output <工作区内的结果目录>
```

已识别结构直接 run 完成，退出码0；有疑点也交付复核文件，不调用模型补名称。退出码3仅表示需要内部结构映射，退出码2为输入/产物错误。可选 --progress 是脚本日志，不要求宿主创建任务计划。

auto 按组检查单企业行、真实合并、名单和重复展示；只有同行关系得到证据支持才使用同行对应。其他情况在本组做精确名称、唯一法人加金额或高相似度且明显领先的匹配。row_aligned 可用于明确的行对应区域，仍受结构和冲突检查约束。冲突或无法唯一对应继续复核。规则、阈值、单位假设和选中的记录写入 ledger，不表示完成工商身份核验。

投标栏自身的错字、重复地名或异常公司后缀会保留；只清理格式和完整企业后缀后的“投标”附注。纯中标名单继续按 award_company 角色保留。unique_companies 是去重文本名称集合，不能称为已确认的企业实体数量。

1.4 移除 resolve/resume 和模型判断状态。1.3未完成任务使用旧版继续，或新版 run 到新目录；旧版本产物仍可 validate。既有文件和版本历史不删除。本仓库尚未接入 Spider 或报告服务，修改代码不等于已部署到 QM。

原始业务 Excel、处理结果、缓存和凭据不提交。tests/verify_samples.py 从仓库外读取真实样本，仓库不附带业务数据。
