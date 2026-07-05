# Architecture

Project Hades memakai arsitektur hybrid:

```text
PDF PO / ERP / NAS
↓
Python tools
↓
Documents runtime files
↓
CorelDRAW VBA standalone engines
↓
Core / PHASE foundation
↓
Controller 5 shortcut
↓
QC final lock
```

---

## Lapisan 1 — Input Tools

```text
src/python/order_extractor
src/python/pattern_fetcher
```

- Order Extractor membuat `Order.txt`.
- Pattern Fetcher mengambil file pola exact dari NAS ke folder lokal.

---

## Lapisan 2 — Runtime Database

```text
Documents\Order.txt
Documents\SizeDB_*.txt
Documents\TypoTemplate_Current.txt
Documents\HADES_PATTERN_CATALOG_CURRENT.txt
```

File ini adalah jembatan antara Python dan CorelDRAW VBA.

---

## Lapisan 3 — VBA Engines

Standalone engine adalah mesin kerja asli:

```text
QC Size
QC Typo
IDPO Check
PowerClip & Transparency Check
Group Structure Check
Auto Rename
Auto Duplicate
Record Pattern Catalog
Auto Mass Nesting
Auto Nesting Template / LRP
Mining Database
```

---

## Lapisan 4 — Core Foundation

Core foundation berisi helper umum:

```text
path Documents
baca Order.txt
baca SizeDB
normalisasi size
normalisasi teks
helper geometry
helper report
self-test
```

Core tidak menggantikan engine. Core hanya mengurangi pengulangan fungsi dasar.

---

## Lapisan 5 — 5 Shortcut Controller

```text
HADES_PRECHECK_MASTER
HADES_PREPARE_MASTER
HADES_EXECUTE_LAYOUT
HADES_QC_FINAL
HADES_FINALIZE_CONVERT
```

Controller adalah pintu operator.

---

## Lapisan 6 — QC Lock

Finalize Convert hanya boleh jalan setelah Global Final Report PASS.

```text
QC Final FAIL
↓
Finalize Convert BLOCKED
```
