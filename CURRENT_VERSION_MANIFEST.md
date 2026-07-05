# CURRENT VERSION MANIFEST

Dokumen ini menjelaskan file yang dianggap **current** dalam repository Project Hades ini.

---

## Controller / 5 Shortcut

```text
src/vba/00_controller/HADES_5_SHORTCUTS_QC_MENU_V2_LOCKED.bas
```

Menggantikan:

```text
archive/old_modules_replaced/HADES_5_SHORTCUTS_PHASE5D_CURRENT.bas
```

Shortcut aktif:

```text
HADES_PRECHECK_MASTER
HADES_PREPARE_MASTER
HADES_EXECUTE_LAYOUT
HADES_QC_FINAL
HADES_FINALIZE_CONVERT
```

---

## Core Foundation

```text
src/vba/01_core_foundation/HADES_CORE_REPORT_PHASE2.bas
src/vba/01_core_foundation/HADES_CORE_IO_PATHS_PHASE5.bas
src/vba/01_core_foundation/HADES_CORE_TEXT_NORMALIZE_PHASE5.bas
src/vba/01_core_foundation/HADES_CORE_ORDER_DB_PHASE5.bas
src/vba/01_core_foundation/HADES_CORE_GEOMETRY_SELECTION_PHASE5.bas
src/vba/01_core_foundation/HADES_CORE_SELF_TEST_PHASE5.bas
```

Jangan hapus selama Phase Report / Finalize masih memanggil core.

---

## Phase Report / Lock

```text
src/vba/02_phase_report_lock/HADES_QC_FINAL_REPORT_PHASE3C.bas
src/vba/02_phase_report_lock/HADES_QC_FINAL_REPORT_PHASE4.bas
src/vba/02_phase_report_lock/HADES_REPORT_VIEWER_PHASE5D_FIX3.bas
src/vba/02_phase_report_lock/HADES_FINALIZE_CONVERT_PHASE3C.bas
src/vba/02_phase_report_lock/HADES_FINALIZE_CONVERT_PHASE4.bas
```

Berfungsi sebagai report chain dan convert lock.

---

## QC Engines

```text
src/vba/03_engines/qc/QC_TRANSPARENCY_POWERCLIP_CHECK_V1_0_USER_SUPPLIED.bas
src/vba/03_engines/qc/VBA IDPO CHECK.bas
src/vba/03_engines/qc/VBA_QC_SIZE_CHECK_V8_4.bas
src/vba/03_engines/qc/VBA_QC_TYPO_CHECK_V13_2_GREEN_MARKER.bas
src/vba/03_engines/qc/GROUP_STRUCTURE_CHECK_V1_CLEANED.bas
```

QC Typo current:

```text
VBA_QC_TYPO_CHECK_V13_2_GREEN_MARKER.bas
```

Entry point tetap:

```text
QC_TYPO_CHECK
HADES_QC_TYPO_REPORT
```

---

## Layout Engines

```text
src/vba/03_engines/layout/VBA_AUTO_DUPLICATE_V2_3_1_ADAPTIVE_GRID_FIXED.bas
src/vba/03_engines/layout/VBA_AUTO_RENAME_V5_3_VISUAL_LIGATURE_BREAKER.bas
src/vba/03_engines/layout/VBA_AUTO_MASS_NESTING_V3_5_ROW_MAJOR_6X2.bas
src/vba/03_engines/layout/HADES_RECORD_PATTERN_CATALOG_V3_2_AUTOONLY.bas
src/vba/03_engines/layout/HADES_AUTO_NESTING_TEMPLATE_V1_3_SIZE_BLOCK_LOCK.bas
```

Catatan:

```text
Auto Duplicate, Auto Nesting LRP, dan Auto Mass Nesting adalah tiga jalur layout berbeda.
Jangan dijalankan otomatis berurutan tanpa pilihan operator.
```

---

## Prepare Master Engines

```text
src/vba/03_engines/prepare_master/BUILD_TYPO_TEMPLATE_V2_CLEANED.bas
src/vba/03_engines/prepare_master/VBA AI TEXT.bas
src/vba/03_engines/prepare_master/VBA AUTO ARRANGE.bas
src/vba/03_engines/prepare_master/VBA AUTO RE-CONTOUR.bas
```

---

## Database Mining Tools

```text
src/vba/03_engines/database_mining/MINE_SIZE_DATABASE_SINGLE.bas
src/vba/03_engines/database_mining/MINE_SPLIT_FRONT_DATABASE.bas
```

Tool ini bukan workflow layout harian.

---

## Python Tools

```text
src/python/order_extractor/hades_order_extractor_v4_5_metadata_reference.py
src/python/pattern_fetcher/hades_pattern_fetcher_v2_2.py
src/python/launcher/HADES_START.bat
```

---

## Future / Not Yet Active

```text
VBA Black CMYK Detector
Python Pattern Router berbasis database kode pola RND
HADES Python Command Center final
Core SizeDB Registry Phase6
```
