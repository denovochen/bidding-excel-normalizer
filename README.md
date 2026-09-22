# bidding-excel-normalizer

独立的招投标 Excel 整理 Skill，当前版本 **1.5.0**。输入 XLS/XLSX，交付 final.csv、review_queue.csv、ledger.json。

脚本全量扫描、分组、清洗、对应中标记录并校验；模型轻量审阅每份文件的结构区域，只修订 plan.json，不逐行判断企业名称。存在投标列时，最终公司名称始终来自投标列；中标栏仅用于对应结果与审计，不替换、纠正投标全称。

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

推荐 Python 3.12；本次实际验证版本为 Python 3.12.14。实际依赖为 `openpyxl`（XLSX）和 `xlrd`（XLS），安装方式：

```bash
python -m pip install -r requirements.txt
```

```bash
python -B -m unittest discover -s tests
python scripts/excel_ledger.py run <Excel路径> --output <工作区内的结果目录>
```

首次 run 对每份文件返回结构清单和建议 plan，退出码3；Agent 审阅完整表头、列统计、代表区域及异常片段后用 `run --plan` 继续。可恢复结构问题最多局部补映射两次，仍无法确认时写入 review_regions 并交付其他结果。退出码0表示三个产物已发布，退出码2表示不可恢复的输入、权限、输出或内部校验错误。

auto 按组检查单企业行、真实合并、名单和重复展示。自动确认限于精确名称、无需单侧单位假设的唯一法人加金额，或已经声明并通过结构校验的 row_aligned。名称相似度和简称只排序候选；单侧金额单位沿用只作审计证据。冲突或无法唯一对应继续复核。

投标栏自身的错字、重复地名或异常公司后缀会保留；只清理格式和完整企业后缀后的“投标”附注。纯中标名单继续按 award_company 角色保留。unique_companies 是去重文本名称集合，不能称为已确认的企业实体数量。

1.5 增加全区域结构审阅、有效列去向、批次/轮次组边界、中标完整性依据、规则置信度和文件/区域级交付退路。策略收紧会增加复核量；不为保持零复核恢复纯相似度自动确认。旧版本产物仍可 validate。本仓库尚未接入 Spider 或报告服务，修改代码不等于已部署到 QM/Qwen。

原始业务 Excel、处理结果、缓存和凭据不提交。tests/verify_samples.py 从仓库外读取真实样本，仓库不附带业务数据。
