# bidding-excel-normalizer

独立的招投标 Excel 整理 Skill，当前版本 **1.6.0**。输入 XLS/XLSX，必要时通过内置提问工具确认非精确中标名称，最终交付 final.csv、review_queue.csv、ledger.json。

脚本全量扫描、分组、清洗、对应中标记录并校验；模型轻量审阅每份文件的结构区域，只修订 plan.json，不逐行判断企业名称。存在投标列时，最终公司名称始终来自投标列；中标栏仅用于对应结果与审计，不替换、纠正投标全称。

## 文件位置

| 路径 | 用途 |
|---|---|
| SKILL.md | 精简执行入口及交付要求 |
| scripts/excel_ledger.py | 运行、局部检查、产物校验及读取去重名单 |
| scripts/ledger_core/ | 结构读取、确定性中标匹配、审计和发布 |
| references/mapping.md | 陌生表头和混合区域的内部映射 |
| references/award-review.md | 非精确中标名称的用户确认流程 |
| references/output-contract.md | 三个产物、匹配依据及异常规则 |
| tests/ | 合成表格回归及外部真实样本验收 |

## 开发验证

推荐 Python 3.12；本次实际验证版本为 Python 3.12.14。实际依赖为 `openpyxl`（XLSX）和 `xlrd`（XLS），安装方式：

```bash
python -m pip install -r requirements.txt
```

```bash
python -B -m unittest discover -s tests
python scripts/excel_ledger.py run <Excel路径> --output <工作区内的结果目录>
python scripts/excel_ledger.py resolve --state <state.json> --answers <answers.json>
```

首次 run 对每份文件返回结构清单和建议 plan，退出码3；Agent 审阅结构后用 `run --plan` 继续。非精确中标名称返回最多5个内置提问问题和退出码4；将回答保存为 JSON 后用 `resolve --state ... --answers ...` 继续，全部问题处理后才发布三个产物。退出码0表示产物已发布，退出码2表示不可恢复错误。

自动确认限于组内精确名称和原表明确中标状态。安全同行优先作为推荐项，否则按组内名称接近程度推荐；法人和金额只作来源审计。用户可选择推荐项、Other 输入当前组投标名称或“不确定”，最终文件始终由 Python 重建。

投标栏自身的错字、重复地名或异常公司后缀会保留；只清理格式和完整企业后缀后的“投标”附注。纯中标名单继续按 award_company 角色保留。unique_companies 是去重文本名称集合，不能称为已确认的企业实体数量。

1.6 增加组级人工确认、内置提问批次、持久化 state 和确定性重新发布；删除法人/金额自动匹配。旧版本产物仍可 validate。本仓库尚未接入 Spider 或报告服务，修改代码不等于已部署到 QM/Qwen。

原始业务 Excel、处理结果、缓存和凭据不提交。tests/verify_samples.py 从仓库外读取真实样本，仓库不附带业务数据。
