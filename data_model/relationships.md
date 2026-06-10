# Data Model — Relationships

## Current (prototype)

Single fact table, no joins needed.

```
fact_ar_performance
```

---

## Phase 2 — Star Schema

### Tables

| Table | Type | Grain |
|---|---|---|
| `fact_ar_performance` | Fact | 1 row per customer per month |
| `dim_date` | Dimension | 1 row per month |
| `dim_customer` | Dimension | 1 row per customer |
| `dim_country` | Dimension | 1 row per country |

### Relationships (all single-direction, many-to-one)

```
fact_ar_performance[month_date]  →  dim_date[date]         (active)
fact_ar_performance[customer]    →  dim_customer[customer] (active)
fact_ar_performance[country]     →  dim_country[country]   (active)
```

### How to create in Power BI Desktop

1. Open **Model view** (left rail icon)
2. Drag `fact_ar_performance[month_date]` → `dim_date[date]`
3. Drag `fact_ar_performance[customer]` → `dim_customer[customer]`
4. Drag `fact_ar_performance[country]` → `dim_country[country]`
5. For each relationship set:
   - Cardinality: **Many to one (*:1)**
   - Cross filter direction: **Single**
   - Make this relationship active: **Yes**
