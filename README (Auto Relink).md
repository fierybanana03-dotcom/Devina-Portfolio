# `auto_relink.py` (v8) — Auto-Relink & Sheet-Name Preservation for 质量日报

A robust Python utility designed to automatically update external XLOOKUP and formula links in Excel workbooks for the **质量日报 (Quality Daily Report)**. 

It dynamically maps external workbook references to new source files based on semantic concept matching, updates relationship paths, safely patches sheet formulas without corrupting sheet names, and recalculates formula value caches via COM automation (WPS/Excel).

---

## 🌟 Key Features

* **Targeted Worksheet Patching:** Exclusively patches external formula references in `xl/worksheets/sheet5.xml` (`透视资料 Hasil Pivot`). Leaves `综合 Keseluruhan` and other sheets completely untouched.
* **Semantic Concept Classification:** Uses predefined keyword rules and exclusion lists (`CONCEPTS`) to automatically match legacy external target paths to modern reference files.
* **Safe XML Formula Patching:** Employs regex capturing `([^'!]+)` to extract exact sheet names and rewrite formulas to explicit current file paths without breaking workbook XML integrity.
* **COM Background Auto-Recalculation:** Spawns a background WPS Office (`ET.Application`) or Microsoft Excel (`Excel.Application`) instance to recalculate updated formulas and save the cached values.
* **Resilient Error Handling:** Gracefully catches COM external source security locks (e.g., error `-2147352567`) without failing the file generation process.

---

## 📐 Supported Data Concepts (`CONCEPTS`)

The script classifies external source files into 15 specific concepts:

| Concept Name | Keyword Triggers | Excluded / Forbidden Terms | Multi-Slot Allowed? |
| :--- | :--- | :--- | :---: |
| **SLA** | `时效签收及时率`, `sla tepat waktu`, `waktu rencana ttd semula` | `投诉报表`, `komplain reguler` | ✅ (Multiple slots) |
| **Komplain** | `投诉报表`, `komplain reguler` | `时效签收及时率`, `sla tepat waktu`, `waktu rencana ttd semula` | ❌ |
| **Lostscan** | `漏扫`, `lostscan` | *None* | ❌ |
| **PBTM** | `分批配载率明细`, `pbtm`, `rumus` | *None* | ❌ |
| **ResiKembali** | `回单返回及时率监控`, `resi kembali` | *None* | ❌ |
| **Transit** | `中转及时率报表` | *None* | ❌ |
| **Keberangkatan** | `干线发车准点率`, `keberangkatan` | *None* | ❌ |
| **Perjalanan** | `干线运行合格率`, `perjalanan` | *None* | ❌ |
| **Arbitrase** | `仲裁破损遗失报告`, `arbitrase` | *None* | ❌ |
| **CollectionPoint** | `进港分拨`, `collection point` | *None* | ❌ |
| **ReturCodTT** | `retur cod tt`, `退件看板` | *None* | ❌ |
| **ReturLazada** | `retur lazada`, `退件看板` | *None* | ❌ |
| **PickUp** | `订单揽收及时率`, `pick up` | *None* | ❌ |
| **BukaWaybill** | `开单未走货率`, `buka waybill` | *None* | ❌ |
| **Gateway** | `中心妥投率`, `gateway` | *None* | ❌ |

---

## 🚀 Installation & Requirements

### Prerequisites
* **Python:** 3.8 or higher
* **OS:** Windows (required for COM auto-recalculation via WPS/Excel)

### Python Dependencies
Optional but recommended for auto-recalculation:
```bash
pip install pywin32
```

---

## 💻 Usage

### Basic Command Syntax

```bash
python auto_relink.py <path_to_workbook.xlsx> --refs-folder <path_to_reference_folder> --output <path_to_output.xlsx>
```

### Options & Arguments

* `workbook` *(Position 1)*: Path to the target template Excel workbook.
* `--refs-folder` *(Required)*: Directory containing the updated daily source Excel files (`.xlsx`, `.xls`, `.xlsm`).
* `--output` *(Required)*: Path where the relinked workbook will be saved.
* `--dry-run`: Pre-analyzes external links and prints the classification matrix without modifying files.
* `--no-recalc`: Skips spawning WPS/Excel background instances for formula value baking.

---

## 🛠️ Usage Examples

### 1. Execute Complete Relinking with Value Baking
```bash
python auto_relink.py "template_质量日报.xlsx" --refs-folder "./daily_sources_2026_08_31" --output "质量日报_2026_08_31.xlsx"
```

### 2. Perform a Dry-Run Inspection
```bash
python auto_relink.py "template_质量日报.xlsx" --refs-folder "./daily_sources" --output "out.xlsx" --dry-run
```

### 3. Fast Execution (Skip Background Calculation)
```bash
python auto_relink.py "template_质量日报.xlsx" --refs-folder "./daily_sources" --output "out.xlsx" --no-recalc
```

---

## ⚙️ Architecture & Execution Flow

```
+-------------------------------------------------------------------+
| 1. Zip Analysis & Index Discovery                                |
|    - Inspect xl/externalLinks/_rels/externalLink*.xml.rels       |
|    - Parse active formula link tokens [N] in sheet5.xml           |
+-------------------------------------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------+
| 2. Semantic Classification & Matching                             |
|    - Classify targets using CONCEPTS keyword matrix                |
|    - Match best reference file in --refs-folder by modification time |
+-------------------------------------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------+
| 3. Zip Stream Patching & Rebuilding                               |
|    - Rewrite Target attributes in externalLink relationships      |
|    - Patch sheet5.xml formulas with explicit current file paths    |
+-------------------------------------------------------------------+
                                  |
                                  v
+-------------------------------------------------------------------+
| 4. Value Baking (Optional COM Trigger)                            |
|    - Spawn ET.Application / Excel.Application                    |
|    - Update links & Execute CalculateFull()                      |
+-------------------------------------------------------------------+
```

---

## 🔄 Version 8 Changelog

* **Syntax Fixes:** Standardized raw string formatting across docstrings to prevent `SyntaxError` on Windows backslashes.
* **COM Resilience:** Wrapped `wb.UpdateLink()` and `CalculateFull()` in guarded try-except blocks to catch OS protected external source locks (`-2147352567`) gracefully.
* **Enhanced Safety:** Preserved exact XML sheet-name extraction regex (`([^'!]+)`) introduced in v7.

---

## 🛡️ License & Maintenance

Designed for internal automation of **质量日报** operations. Maintain concept keyword mappings inside `CONCEPTS` if source report naming conventions change.
