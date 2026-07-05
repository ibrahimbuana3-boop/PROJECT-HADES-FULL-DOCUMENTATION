# Project Hades

**Project Hades** adalah sistem otomasi dan Quality Control hybrid untuk workflow layout produksi garmen/sportswear sublimasi berbasis **CorelDRAW VBA + Python + file database ringan**.

Project ini dibuat untuk memangkas pekerjaan manual layouter, mengurangi risiko salah size, salah nama/nomor, salah pola, dan mencegah file produksi masuk tahap convert/press sebelum lolos QC.

> Status repositori ini: **current working repository snapshot**.  
> Beberapa macro sudah working di workflow produksi, beberapa modul bersifat experimental/future spec.  
> Selalu import module VBA dengan cara **replace**, bukan berdampingan, untuk menghindari `Duplicate Procedure`.

---

## Ringkasan Sistem

Project Hades bukan lagi sekadar kumpulan macro. Struktur sistemnya terdiri dari:

```text
Python Input Tools
↓
Database / SizeDB / Template
↓
CorelDRAW VBA Engines
↓
Core / PHASE Foundation
↓
5 Shortcut Controller
↓
QC Final Report + Convert Lock
```

5 shortcut utama:

```text
1. HADES_PRECHECK_MASTER
2. HADES_PREPARE_MASTER
3. HADES_EXECUTE_LAYOUT
4. HADES_QC_FINAL
5. HADES_FINALIZE_CONVERT
```

---

## Fitur Utama

### 1. Python Input Tools

```text
src/python/order_extractor
src/python/pattern_fetcher
src/python/launcher
```

Fungsi:

- Membaca PDF PO dan membuat `Documents\Order.txt`.
- Mengambil file pola dari NAS berdasarkan kode pola.
- Mengurangi kerja manual search pola di database.
- Menyiapkan data untuk CorelDRAW workflow.

### 2. SizeDB dan Mining Tools

```text
src/vba/03_engines/database_mining
```

Fungsi:

- Menambang dimensi pola jersey dari outline merah.
- Membuat `SizeDB_*.txt`.
- Mendukung jersey body depan/belakang dan split-front seperti jaket/kemeja/rompi.

### 3. CorelDRAW VBA Workflow

```text
src/vba/00_controller
src/vba/01_core_foundation
src/vba/02_phase_report_lock
src/vba/03_engines
```

Fungsi:

- Auto Duplicate
- Auto Rename
- Auto Mass Nesting
- Auto Nesting Template / LRP
- Record Pattern Catalog
- QC Size
- QC Typo
- IDPO Check
- Transparency & PowerClip Check
- Group Structure Check
- Finalize Convert Lock

### 4. QC Final Menu V2 Locked

Shortcut `HADES_QC_FINAL` sekarang menjadi menu:

```text
1 = Jalankan PowerClip & Transparency
2 = Jalankan IDPO Check
3 = Jalankan Size Check
4 = Jalankan Typo Check
5 = Jalankan Group Structure Check
6 = Jalankan semuanya secara berurutan
```

Aturan lock:

```text
Pilihan 1-5:
- menjalankan QC mandiri
- Finalize Convert di-invalidasi

Pilihan 6:
- menjalankan Global Final Report
- jika semua PASS → Finalize Convert boleh
- jika ada satu FAIL → Finalize Convert diblokir
```

### 5. QC Typo Green Marker

QC Typo terbaru menambahkan fitur:

```text
Jika TYPO FAIL:
- panel body depan/belakang dalam group bermasalah diberi outline hijau
- shortcut lama tetap sama:
  QC_TYPO_CHECK
  HADES_QC_TYPO_REPORT
```

---

## Struktur Repository

```text
PROJECT_HADES/
├─ README.md
├─ LICENSE
├─ ACTIVE_INSTALL_PACKAGE.md
├─ CURRENT_VERSION_MANIFEST.md
├─ .gitignore
├─ docs/
├─ data/
│  └─ templates/
├─ src/
│  ├─ python/
│  │  ├─ order_extractor/
│  │  ├─ pattern_fetcher/
│  │  └─ launcher/
│  └─ vba/
│     ├─ 00_controller/
│     ├─ 01_core_foundation/
│     ├─ 02_phase_report_lock/
│     ├─ 03_engines/
│     │  ├─ qc/
│     │  ├─ layout/
│     │  ├─ prepare_master/
│     │  └─ database_mining/
│     └─ 04_experimental_or_future/
├─ archive/
│  ├─ original_upload_snapshot/
│  └─ old_modules_replaced/
└─ release_notes/
```

---

## Import Cepat ke CorelDRAW

Import module dengan urutan:

```text
1. src/vba/01_core_foundation
2. src/vba/02_phase_report_lock
3. src/vba/03_engines/qc
4. src/vba/03_engines/prepare_master
5. src/vba/03_engines/layout
6. src/vba/03_engines/database_mining
7. src/vba/00_controller
```

Penting:

```text
Jangan import module lama berdampingan dengan module baru yang punya nama Sub sama.
Hapus/replace module lama terlebih dahulu.
```

Controller aktif:

```text
src/vba/00_controller/HADES_5_SHORTCUTS_QC_MENU_V2_LOCKED.bas
```

Module lama yang diganti:

```text
HADES_5_SHORTCUTS_PHASE5D_CURRENT.bas
```

---

## File Database Runtime

Project Hades memakai file runtime di folder `Documents` Windows:

```text
Documents\Order.txt
Documents\SizeDB_*.txt
Documents\TypoTemplate_Current.txt
Documents\HADES_PATTERN_CATALOG_CURRENT.txt
Documents\HADES_PATTERN_INBOX\
```

---

## Format Order.txt

```text
@JENIS_PESANAN=JERSEY
@JENIS_POLA=JERSEY REGULER
@MODEL_JAHIT=DEWASA PRIA
@SIZEDB=SizeDB_Pria.txt
@IDPO=355863

M|Meilan||
XL|Miftah||
XL|Waode Rahayu||
XL|Alfiah||
```

Format row:

```text
SIZE|NAMA|NOMOR|NICKNAME
```

---

## Format SizeDB Jersey

```text
SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
```

Contoh:

```text
M|54.600|75.104|74.937
L|56.600|77.104|76.937
XL|58.600|79.104|78.937
```

## Format SizeDB Split Front

```text
SIZE|LEBAR_BELAKANG|LEBAR_DEPAN|TINGGI_DEPAN|TINGGI_BELAKANG
```

Contoh:

```text
M|56.000|29.500|74.000|75.000
```

---

## Catatan Status

### Aktif / Current

```text
QC Final Menu V2 Locked
QC Typo V13.2G Green Marker
Auto Rename V5.3 Visual Ligature Breaker
Auto Duplicate V2.3.1 Adaptive Grid Fixed
Record Pattern Catalog V3.2 AutoOnly
Auto Mass Nesting V3.5 Row-Major 6x2
```

### Wacana / Future

```text
VBA Black CMYK Detector
Python Pattern Router otomatis berbasis database kode pola RND
Python Command Center terpadu
Core SizeDB Registry Phase6
```

---

## Disclaimer Teknis

Repo ini adalah dokumentasi dan snapshot implementasi Project Hades.  
Macro VBA harus diuji ulang di CorelDRAW 2021 sebelum dijadikan release produksi penuh, karena environment CorelDRAW, versi VBA, font, path Windows, dan struktur file pabrik dapat memengaruhi hasil runtime.
