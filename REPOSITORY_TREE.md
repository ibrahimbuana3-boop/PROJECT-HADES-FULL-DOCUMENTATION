# Repository Tree

```text
PROJECT_HADES_FULL_REPOSITORY_CURRENT/
├─ .gitignore
├─ ACTIVE_INSTALL_PACKAGE.md
├─ CURRENT_VERSION_MANIFEST.md
├─ LICENSE
├─ README.md
├─ archive/
│  ├─ old_modules_replaced/
│  │  ├─ HADES_5_SHORTCUTS_PHASE5D_CURRENT.bas
│  │  ├─ HADES_AUTO_MASS_NESTING_V3_4_LIGATURE_BREAKER.bas
│  │  ├─ VBA AUTO DUPLICATE.bas
│  │  ├─ VBA AUTO MASS NESTING.bas
│  │  ├─ VBA AUTO RENAME.bas
│  │  ├─ VBA RECORD PATTERN CATALOG.bas
│  │  └─ VBA_QC_TYPO_CHECK_V12_2_CORE_BRIDGE.bas
│  └─ original_upload_snapshot/
│     ├─ ACTIVE_INSTALL_PACKAGE.md
│     ├─ BUILD_TYPO_TEMPLATE_V2_CLEANED.bas
│     ├─ GROUP_STRUCTURE_CHECK_V1_CLEANED.bas
│     ├─ HADES_5_SHORTCUTS_PHASE5D_CURRENT.bas
│     ├─ HADES_AUTO_MASS_NESTING_V3_4_LIGATURE_BREAKER.bas
│     ├─ HADES_AUTO_NESTING_TEMPLATE_V1_3_SIZE_BLOCK_LOCK.bas
│     ├─ HADES_CORE_GEOMETRY_SELECTION_PHASE5.bas
│     ├─ HADES_CORE_IO_PATHS_PHASE5.bas
│     ├─ HADES_CORE_ORDER_DB_PHASE5.bas
│     ├─ HADES_CORE_REPORT_PHASE2.bas
│     ├─ HADES_CORE_SELF_TEST_PHASE5.bas
│     ├─ HADES_CORE_TEXT_NORMALIZE_PHASE5.bas
│     ├─ HADES_FINALIZE_CONVERT_PHASE3C.bas
│     ├─ HADES_FINALIZE_CONVERT_PHASE4.bas
│     ├─ HADES_QC_FINAL_REPORT_PHASE3C.bas
│     ├─ HADES_QC_FINAL_REPORT_PHASE4.bas
│     ├─ HADES_RECORD_PATTERN_CATALOG_V3_2_AUTOONLY.bas
│     ├─ HADES_REPORT_VIEWER_PHASE5D_FIX3.bas
│     ├─ LICENSE
│     ├─ PHASE_REDUCTION_MAP.md
│     ├─ QC_TRANSPARENCY_POWERCLIP_CHECK_V1_0_USER_SUPPLIED.bas
│     ├─ README.md
│     ├─ VBA AI TEXT.bas
│     ├─ VBA AUTO ARRANGE.bas
│     ├─ VBA AUTO DUPLICATE.bas
│     ├─ VBA AUTO MASS NESTING.bas
│     ├─ VBA AUTO RE-CONTOUR.bas
│     ├─ VBA AUTO RENAME.bas
│     ├─ VBA IDPO CHECK.bas
│     ├─ VBA RECORD PATTERN CATALOG.bas
│     ├─ VBA_QC_SIZE_CHECK_V8_4.bas
│     └─ VBA_QC_TYPO_CHECK_V12_2_CORE_BRIDGE.bas
├─ data/
│  └─ templates/
│     ├─ HADES_PATTERN_CATALOG_CURRENT.example.txt
│     ├─ Order.example.txt
│     ├─ SizeDB_Jersey.example.txt
│     ├─ SizeDB_SplitFront.example.txt
│     └─ TypoTemplate_Current.example.txt
├─ docs/
│  ├─ ARCHITECTURE.md
│  ├─ FUTURE_CMYK_BLACK_DETECTOR.md
│  ├─ GLOSSARY.md
│  ├─ IMPORT_GUIDE_CORELDRAW_VBA.md
│  ├─ MINING_DATABASE.md
│  ├─ PATTERN_ROUTER_FUTURE.md
│  ├─ PROJECT_STATUS.md
│  ├─ PYTHON_TOOLS.md
│  ├─ QC_FINAL_MENU_V2_LOCKED.md
│  ├─ QC_TYPO_GREEN_MARKER.md
│  ├─ SIZEDB_GUIDE.md
│  └─ WORKFLOW_5_SHORTCUTS.md
├─ release_notes/
│  └─ CHANGELOG.md
└─ src/
   ├─ python/
   │  ├─ launcher/
   │  │  └─ HADES_START.bat
   │  ├─ order_extractor/
   │  │  ├─ HADES_ORDER_EXTRACTOR.bat
   │  │  ├─ hades_order_extractor_v4_5_metadata_reference.py
   │  │  └─ requirements.txt
   │  └─ pattern_fetcher/
   │     ├─ CARI_POLA.bat
   │     └─ hades_pattern_fetcher_v2_2.py
   └─ vba/
      ├─ 00_controller/
      │  └─ HADES_5_SHORTCUTS_QC_MENU_V2_LOCKED.bas
      ├─ 01_core_foundation/
      │  ├─ HADES_CORE_GEOMETRY_SELECTION_PHASE5.bas
      │  ├─ HADES_CORE_IO_PATHS_PHASE5.bas
      │  ├─ HADES_CORE_ORDER_DB_PHASE5.bas
      │  ├─ HADES_CORE_REPORT_PHASE2.bas
      │  ├─ HADES_CORE_SELF_TEST_PHASE5.bas
      │  └─ HADES_CORE_TEXT_NORMALIZE_PHASE5.bas
      ├─ 02_phase_report_lock/
      │  ├─ HADES_FINALIZE_CONVERT_PHASE3C.bas
      │  ├─ HADES_FINALIZE_CONVERT_PHASE4.bas
      │  ├─ HADES_QC_FINAL_REPORT_PHASE3C.bas
      │  ├─ HADES_QC_FINAL_REPORT_PHASE4.bas
      │  └─ HADES_REPORT_VIEWER_PHASE5D_FIX3.bas
      ├─ 03_engines/
      │  ├─ database_mining/
      │  │  ├─ MINE_SIZE_DATABASE_SINGLE.bas
      │  │  └─ MINE_SPLIT_FRONT_DATABASE.bas
      │  ├─ layout/
      │  │  ├─ HADES_AUTO_NESTING_TEMPLATE_V1_3_SIZE_BLOCK_LOCK.bas
      │  │  ├─ HADES_RECORD_PATTERN_CATALOG_V3_2_AUTOONLY.bas
      │  │  ├─ VBA_AUTO_DUPLICATE_V2_3_1_ADAPTIVE_GRID_FIXED.bas
      │  │  ├─ VBA_AUTO_MASS_NESTING_V3_5_ROW_MAJOR_6X2.bas
      │  │  └─ VBA_AUTO_RENAME_V5_3_VISUAL_LIGATURE_BREAKER.bas
      │  ├─ prepare_master/
      │  │  ├─ BUILD_TYPO_TEMPLATE_V2_CLEANED.bas
      │  │  ├─ VBA AI TEXT.bas
      │  │  ├─ VBA AUTO ARRANGE.bas
      │  │  └─ VBA AUTO RE-CONTOUR.bas
      │  └─ qc/
      │     ├─ GROUP_STRUCTURE_CHECK_V1_CLEANED.bas
      │     ├─ QC_TRANSPARENCY_POWERCLIP_CHECK_V1_0_USER_SUPPLIED.bas
      │     ├─ VBA IDPO CHECK.bas
      │     ├─ VBA_QC_SIZE_CHECK_V8_4.bas
      │     └─ VBA_QC_TYPO_CHECK_V13_2_GREEN_MARKER.bas
      └─ 04_experimental_or_future/
```
