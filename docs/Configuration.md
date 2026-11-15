# Configuration

Every adjustable constant is at the top of `src/CrossSheetAggregator.bas`.
**Source code is the single source of truth** — numbers here reference
names only.

## Two Modules

| Module | Entry | What it does | Input sheets | Output sheet |
|---|---|---|---|---|
| **Link Collector** | `RunModule1()` | Merge rows from Source A + Source B by entity link, 3-tier dedup | SourceA, SourceB | MergedLinks |
| **Identifier Collector** | `RunModule2()` | Match Source A + Source C by entity ID extracted from links | SourceA, SourceC | MergedCodes |
| **Run both** | `RunAllAggregations()` | Run Link Collector, then Identifier Collector | All three | Both output sheets |

---

## Sheet Name Constants

| Constant | Default | When to change |
|---|---|---|
| `SHEET_SOURCE_A` | `"SourceA"` | Your primary data sheet has a different name |
| `SHEET_SOURCE_B` | `"SourceB"` | Your secondary / backup data sheet has a different name |
| `SHEET_SOURCE_C` | `"SourceC"` | Your reference sheet for Module 2 has a different name |
| `SHEET_MERGED_LINKS` | `"MergedLinks"` | Where Module 1 writes output |
| `SHEET_MERGED_CODES` | `"MergedCodes"` | Where Module 2 writes output |

Output sheets are **recreated** on every run — changing the name is fine
but make sure the new name doesn't collide with anything you want to keep.

## Column Letter Constants

| Constant | Default | What it is | Which module uses it |
|---|---|---|---|
| `COL_LINK` | `"I"` | Column that holds the entity / product link | Both |
| `COL_ID_FIELD` | `"B"` | Source C column holding the identifier / code | Module 2 |
| `COL_ID_LINK` | `"G"` | Source C column holding the link (to extract entity ID for matching) | Module 2 |

Source A / B column layout is read by letter index (column C, D, E, J, R, T, Y, AB
etc. in the Module 1 helpers). These are hardcoded — if your layout differs,
update the `WriteDestRow` and `BuildNameDictionary` calls with the correct
column letters. They're clearly marked with `' col X, adjust ` comments.

## Regex Patterns

| Constant | Default | What it matches | When to change |
|---|---|---|---|
| `BATCH_ID_REGEX` | `"[A-Z]{4}\d{7}"` | 11-char batch IDs (4 uppercase letters + 7 digits) | Your batch IDs follow a different pattern |
| `ENTITY_ID_REGEX` | `"[A-Z][0-9A-Z]{9}"` | 10-char entity identifiers (1 uppercase + 9 alnum) | Your marketplace / platform uses a different format |
| `TARGET_DOMAIN` | `"amazon.example-domain"` | Domain suffix to rebuild sanitized links | Point at your target domain |

## Formatting Constants

| Constant | Default | What it is |
|---|---|---|
| `CLR_HEADER_BG` | `RGB(31, 78, 121)` | Dark blue header background |
| `CLR_HEADER_FG` | `RGB(255, 255, 255)` | White header font |

Change freely to match your company template.

---

## Workbook Layout Assumptions

| Sheet | Required | Structure |
|---|---|---|
| **Source A** (main) | Required for both modules | Row 1–3 = headers; data starts row 4. Link column present. |
| **Source B** | Required for Module 1 | Similar structure to Source A. |
| **Source C** | Required for Module 2 | Has both an identifier/code column AND a link column (to extract entity IDs). |
| **MergedLinks** (output) | Auto-created each Module 1 run | 12 columns: Batch ID, Name (CN), Name (EN), Identifier, Percent, Description, Price, Weight, Link, Material, Source, Price-per-kg |
| **MergedCodes** (output) | Auto-created each Module 2 run | 10 columns: Name (CN), Name (EN), Identifier, Percent, Description, Weight, Price, Link, Material, Source |

## How 3-Tier Dedup Works

| Rule | Condition | Action |
|---|---|---|
| **Rule 1 — Full duplicate** | Same link + CN name + identifier + percent + price + weight all match | Keep first occurrence, skip current |
| **Rule 2 — Same product, diff name** | Same link + identifier + percent + price + weight match, but CN name differs | Merge names with "/" separator, skip current |
| **Rule 3 — Diff product, same link** | Only link matches, everything else differs | Keep both rows |

Rule 1 → 2 → 3 priority. Code checks in that order, exits on first match.

## Second Domain Example — Address Pool Consolidation

To show generality, here's how you'd point this tool at a completely
different use case:

| Original scenario | Second scenario |
|---|---|
| Source A + Source B product link merge | **Address Pool Consolidation** — merge shipping addresses from ERP + manual Excel |
| Entity ID from marketplace links | **Postal code + last 4 of phone** as your entity-key extractor |
| Source C = reference codes table | **Clean address master** — zip + standard city name |
| Batch ID = container tracking number | **District code / routing code** — any 11-char or fixed-format ID |

### What you actually change

1. **Sheet names** — change `SHEET_SOURCE_A`, `SHEET_SOURCE_B`, `SHEET_SOURCE_C`.
2. **Column letters** — change `COL_LINK`, `COL_ID_FIELD`, `COL_ID_LINK`.
3. **Regex patterns** — point `ENTITY_ID_REGEX` at whatever makes sense for your
   "entity ID", and `BATCH_ID_REGEX` at your batch/batch-number format.
4. **Nothing else.** Same dedup rules, same output formatting, same hyperlinks.

Same code, same algorithm, different domain.
