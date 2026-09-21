# bidding-excel-normalizer

独立的招投标 Excel 整理 Skill，当前版本为 **1.3.0**。输入 `.xls` / `.xlsx`，输出 `final.csv`、`review_queue.csv`、`ledger.json`。

脚本读取、分组、清洗和校验全量数据；宿主模型只在未知结构或少量企业名称差异时参与。按招标组匹配中标名单，不用“中标格有值即本行企业中标”，不维护具体公司名映射。缺失业务字段留空，无法确定的记录保留复核，处理中不要求用户逐项确认。

## 文件位置

| 路径 | 用途 |
|---|---|
| `SKILL.md` | 模型执行入口与结果交付规则 |
| `scripts/excel_ledger.py` | 运行、恢复、校验及读取企业名单 |
| `scripts/ledger_core/` | Excel 读取、分组、名称判断交接和产物校验 |
| `references/` | 按需读取的结构映射、名称判断、输出及下游名单规则 |
| `tests/` | 行为与契约测试，使用合成数据 |

## 开发验证

```bash
python -B -m unittest discover -s tests
```

由宿主按 `SKILL.md` 执行：

```bash
python scripts/excel_ledger.py run <Excel路径> --output <结果目录>
```

退出码3表示宿主需要内部补充结构映射，4表示需要内部判断名称对；继续执行至 `kind=result` 才完成。任务状态可通过 `resume --state <状态路径>` 恢复。名称判断规则见 [name-review.md](references/name-review.md)。Skill 不规定宿主进度计划；`--progress` 仅为可选的脚本阶段日志。

名称判断交接返回 `decisions_path`，位于 `state.json` 同目录；模型写文件与 `resolve --decisions` 使用同一完整路径。输出目录须位于宿主允许写入的工作区内。已有1.3任务可通过 resume 获取路径，无需重新处理。

输出格式见 [output-contract.md](references/output-contract.md)。两个 CSV 保持固定16列、UTF-8 BOM；ledger 保留来源与清洗依据。去重企业名单是文本名称集合，不等于已确认工商实体数。

此仓库维护独立整理能力，尚未接通 Spider 或风险报告服务。模型对边界名称可能作出不同决定；现有任务保存并复用已接受决定。当前代码通过78项测试，并验证可恢复1.3.0的两份真实样本任务；QM 本地模型的完整运行仍需实际验收。

原始业务 Excel、处理结果、临时任务状态和凭据由 `.gitignore` 排除，不提交到仓库。`tests/verify_samples.py` 是维护者使用外部样本的验收工具，仓库不附带业务样本。
