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

推荐 Python 3.12；本次实际验证版本为 Python 3.12.14。实际依赖为 `openpyxl`（XLSX）、`xlrd`（XLS）和 `defusedxml`（有界 XML 预扫描，禁用 DTD/实体），安装方式：

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

## 读取边界与失败隔离

XLSX 先流式扫描 OOXML 中实际序列化的单元格，再使用 openpyxl 只读模式在经过校验的有效范围内读取。声明 dimension、远端空白样式以及筛选/打印命名范围不决定业务边界，不按文件名或 Sheet 名裁列，也不遍历到 XFD 的空白矩形。非空值、0、False、公式（包括共享公式与空缓存）、Excel 错误及隐藏内容都参与检查，坐标保持原样。

有内容锚点的合并完整保留；与内容及有值合并范围相交的空白合并保留，远离该范围的纯空白合并忽略。相交判断使用固定边界，空白合并不会通过链式扩张吸收其他远端区域。保留的合并在构造 Sheet 之前检查范围，纵向项目继承和横向多列表头沿用原逻辑。批注、超链接和表定义也检查范围，不能作为纯格式残留忽略；无值超链接保持普通读取模式的 target/location 填值行为。外部工作簿链接不访问、不刷新，公式保留为待核对的来源内容。

若合并非锚点仍序列化真实内容，整份文件进入可恢复失败，避免合并继承掩盖该内容。

原安全阈值保持：文件 32 MiB、ZIP 最多 4096 项且解压总量 128 MiB、最多 64 个 Sheet、全工作簿最多 500,000 个有效单元格；单表有效范围最多 100,000 行、256 列且矩形面积最多 500,000。合并索引和超链接展开另受 500,000 的规模限制，防止结构重复造成无界展开。

真实内容、公式范围或必须保留的结构超过单表范围时，整份文件按可恢复失败隔离，错误保存 Sheet、触发坐标/区域及阈值；不会截取允许列后冒充完整解析。同批正常文件继续结构审阅及必要的中标确认。失败来源进入 inspection.source_failures、ledger.sources 和 SOURCE_UNREADABLE 独立复核项，不计入企业或业务记录数。run --plan 和 resolve 重读输入时遵循相同规则。

文件大小、ZIP 解压规模、Sheet 数、有效单元格总数和结构展开规模超限仍终止批次。权限、缺失依赖、内部不变量或产物校验错误也不降级为文件级失败。

全部输入均为可恢复失败时沿用现有行为：inspect 返回空 sources 和失败清单，退出码0；run 生成只有失败审计的三个产物，final.csv 无业务行、record_count 为0，退出码0仅表示审计产物已发布，不表示任何输入解析成功。
