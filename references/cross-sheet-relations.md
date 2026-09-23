# 跨表项目关系 v1

仅在 inspection 返回 `relationship_candidates`，或已确认工作簿把中标结果与投标企业拆在不同区域时读取。

## 标准角色

`project_name`、`project_code`、`project_year`、`lot_name`、`bidder_name` 和 `award_name` 是内部角色，不是原表字段名称。原表可以使用任意字段名、列顺序、多行或合并表头；先按 [字段映射](mapping.md) 将区域映射到标准角色，再判断关系。

| table_kind | 角色要求 | 用途 |
|---|---|---|
| award_summary | 项目角色 + award_name | 提供项目或标段的最终中标结果 |
| bidder_roster | 项目角色 + bidder_name | 提供项目下的投标企业名单 |
| complete_results | 项目角色 + bidder_name + award_name | 区域内部已含完整关系，通常无需跨表 |
| other | 不符合以上组合 | 不自动参加跨表关系 |

不按文件名、Sheet 名、年度 Sheet 命名或固定列号判断表类型。只有已映射角色和区域语义能决定 table_kind。

## 关系声明

确认候选关系后，在 plan patch 中写入：

```json
{
  "relationships": [{
    "id": "construction_awards",
    "kind": "award_to_bidder_roster",
    "award_tables": ["table_summary"],
    "bidder_tables": ["table_2019", "table_2020"],
    "project_keys": ["project_code", "project_name"]
  }]
}
```

`project_code` 存在时优先精确关联；否则使用规范化后的 `project_name`。规范化只统一 Unicode、空白、标点及“年/年度”形式，不删除资金来源、地区、片区或项目类型等可能改变项目身份的内容。

## 关联约束

1. 项目名称是跨表关联主字段，年度只作候选约束。
2. 标段不是项目关联的必要条件；两侧标段文本不一致不能阻止项目级关联。
3. 两侧标段均可明确规范为同一编号时，Python 可拆分标段。
4. 多标段无法唯一拆分时，保留项目级中标结果并生成 `LOT_SCOPE_UNRESOLVED`；不得声称其他企业属于某一具体标段的未中标企业。
5. 中标企业只能在项目关系建立后与目标投标名单匹配，不能反向用于选择项目。
6. 同一企业参加多个项目或多个标段不构成项目关系证据。
7. 精确且唯一的项目键由 Python 自动关联；非精确或不唯一的名称进入 `relationship_review_required`。
8. 无候选项目时生成独立复核项，不丢弃投标名单，也不编造关系。

模型不得读取全部投标企业来推断项目，不得编写临时 Python，不得直接编辑 final.csv 或 ledger.json。
