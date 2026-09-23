# bidding-excel-normalizer

独立的招投标 Excel 整理 Skill，当前版本 **1.7.1**。输入 XLS/XLSX，支持陌生结构映射、跨 Sheet 项目关系、项目候选复核和非精确中标名称确认，最终交付 final.csv、review_queue.csv、ledger.json。

脚本全量扫描、分组、清洗、关联项目、对应中标记录并校验；模型轻量审阅结构并只编写 plan patch 或选择有界项目候选，不逐行判断企业名称。存在投标列时，最终公司名称始终来自投标列；中标栏仅用于对应结果与审计，不替换、纠正投标全称。

## 文件位置

| 路径 | 用途 |
|---|---|
| SKILL.md | 精简执行入口及交付要求 |
| scripts/excel_ledger.py | 运行、局部检查、产物校验及读取去重名单 |
| scripts/ledger_core/ | 结构读取、确定性中标匹配、审计和发布 |
| references/mapping.md | 陌生表头和混合区域的内部映射 |
| references/cross-sheet-relations.md | 汇总表与投标明细表的通用跨表关系 |
| references/relation-review.md | 非精确项目名称候选的内部复核 |
| references/award-review.md | 非精确中标名称的用户确认流程 |
| references/output-contract.md | 三个产物、匹配依据及异常规则 |
| references/real-sample-acceptance.md | 三份仓库外真实样本的统计、哈希和未验证项 |
| tests/ | 合成表格回归及外部真实样本验收 |

## 开发验证

推荐 Python 3.12；本次实际验证版本为 Python 3.12.14。运行依赖为 `openpyxl`（XLSX）和 `xlrd`（XLS）。XML 预扫描在解析前拒绝 DTD/实体，不需要额外 XML 依赖。安装方式：

```bash
python -m pip install -r requirements.txt
```

```bash
python -B -m unittest discover -s tests
python scripts/excel_ledger.py doctor
python scripts/excel_ledger.py run <Excel路径> --output <工作区内的结果目录>
python scripts/excel_ledger.py resolve --state <state.json> --answers <answers.json>
```

首次 run 将完整 inspection 保存到内部 `.excel-ledger-work-*` 目录，stdout 只返回有界摘要和 `inspection_path`，退出码3。Agent 以 plan patch 修订表结构和跨表关系，再用 `plan` 命令生成完整 plan。非精确项目关系返回退出码5；非精确中标名称返回退出码4。两者都通过同一 `resolve --state ... --answers ...` 状态机逐批处理。退出码0表示产物已发布，退出码2表示不可恢复错误。

自动确认限于组内精确名称和原表明确中标状态。安全同行优先作为推荐项，否则按组内名称接近程度推荐；法人和金额只作来源审计。用户可选择推荐项、Other 输入当前组投标名称或“不确定”，最终文件始终由 Python 重建。

投标栏自身的错字、重复地名或异常公司后缀会保留；只清理格式和完整企业后缀后的“投标”附注。纯中标名单继续按 award_company 角色保留。unique_companies 是去重文本名称集合，不能称为已确认的企业实体数量。

1.7 增加标准表类型、项目名称主导的跨 Sheet 关系、标段可选细分、项目关系候选复核、紧凑 inspection、plan patch 和依赖预检。旧 plan 未声明 relationships 时继续执行 1.6 单表逻辑；旧版本产物仍可 validate。本仓库尚未接入 Spider 或报告服务，修改代码不等于已部署到 QM/Qwen。

1.7.1 增加 bidder roster 项目覆盖不变量、部分合并/块首项目继承、`bidder_serial`、来源投标块审计及跨块去重保护。映射了项目角色的投标企业若仍无法归属项目，将返回 `mapping_required`，不会进入跨表关系、中标候选或最终发布。候选范围不完整时关系和中标推荐均停止，award-only 表的投标数量只保留为证据。CSV 记录数始终由 CSV 解析器或 `ledger.summary` 返回，不按物理文本行统计。

原始业务 Excel、处理结果、缓存和凭据不提交。tests/verify_samples.py 从仓库外读取真实样本，仓库不附带业务数据。

三份真实样本的基线统计和限制见 [仓库外真实样本验收](references/real-sample-acceptance.md)。验收脚本只模拟选择 Python 提供的有界候选；业务真值仍需外部确认。

## 读取边界与失败隔离

XLSX 先流式扫描 OOXML 中实际序列化的单元格，再使用 openpyxl 只读模式在经过校验的有效范围内读取。声明 dimension、远端空白样式以及筛选/打印命名范围不决定业务边界，不按文件名或 Sheet 名裁列，也不遍历到 XFD 的空白矩形。非空值、0、False、公式（包括共享公式与空缓存）、Excel 错误及隐藏内容都参与检查，坐标保持原样。

有内容锚点的合并完整保留；与内容及有值合并范围相交的空白合并保留，远离该范围的纯空白合并忽略。相交判断使用固定边界，空白合并不会通过链式扩张吸收其他远端区域。保留的合并在构造 Sheet 之前检查范围，纵向项目继承和横向多列表头沿用原逻辑。批注、超链接和表定义也检查范围，不能作为纯格式残留忽略；无值超链接保持普通读取模式的 target/location 填值行为。外部工作簿链接不访问、不刷新，公式保留为待核对的来源内容。

若合并非锚点仍序列化真实内容，整份文件进入可恢复失败，避免合并继承掩盖该内容。

原安全阈值保持：文件 32 MiB、ZIP 最多 4096 项且解压总量 128 MiB、最多 64 个 Sheet、全工作簿最多 500,000 个有效单元格；单表有效范围最多 100,000 行、256 列且矩形面积最多 500,000。合并索引和超链接展开另受 500,000 的规模限制，防止结构重复造成无界展开。

真实内容、公式范围或必须保留的结构超过单表范围时，整份文件按可恢复失败隔离，错误保存 Sheet、触发坐标/区域及阈值；不会截取允许列后冒充完整解析。同批正常文件继续结构审阅及必要的中标确认。失败来源进入 inspection.source_failures、ledger.sources 和 SOURCE_UNREADABLE 独立复核项，不计入企业或业务记录数。run --plan 和 resolve 重读输入时遵循相同规则。

文件大小、ZIP 解压规模、Sheet 数、有效单元格总数和结构展开规模超限仍终止批次。权限、缺失依赖、内部不变量或产物校验错误也不降级为文件级失败。

全部输入均为可恢复失败时沿用现有行为：inspect 返回空 sources 和失败清单，退出码0；run 生成只有失败审计的三个产物，final.csv 无业务行、record_count 为0，退出码0仅表示审计产物已发布，不表示任何输入解析成功。
