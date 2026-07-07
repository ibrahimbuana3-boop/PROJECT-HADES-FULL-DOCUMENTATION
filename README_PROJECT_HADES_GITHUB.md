# PROJECT H.A.D.E.S.

**Hybrid Automation & Data Execution Systems**  
*Sistem Otomasi Hibrida & Eksekusi Data untuk workflow layout, nesting, QC, dan finalisasi produksi sublimasi sportswear berbasis CorelDRAW VBA + Python.*

---

## 1. Ringkasan

**Project H.A.D.E.S.** adalah sistem otomasi internal untuk membantu proses produksi layout di CorelDRAW, terutama pada workflow sportswear/sublimasi yang melibatkan data order, master desain, master pola, size database, auto duplicate, auto rename, nesting, quality control, dan final convert.

H.A.D.E.S. bukan hanya satu macro VBA. Ia adalah kumpulan modul yang disusun sebagai pipeline produksi:

```text
PRECHECK MASTER
↓
PREPARE MASTER
↓
EXECUTE LAYOUT
↓
QC FINAL
↓
FINALIZE CONVERT
```

Fokus utama project ini adalah mengurangi pekerjaan manual berulang dan menekan human error seperti:

```text
- salah jumlah duplicate per size
- salah size
- salah nama atlet
- salah nomor atlet
- pasangan nama-nomor tertukar
- IDPO belum diganti
- font Jepang/CJK tidak terbaca
- typo akibat copy-paste manual
- struktur group layout tidak rapi
- panel hasil mass nesting tidak bisa dilacak identitasnya
- file difinalisasi sebelum QC benar-benar PASS
```

Project ini dibuat untuk membantu operator/layouter bekerja lebih aman, lebih cepat, dan lebih konsisten tanpa harus mengandalkan ingatan manual di setiap tahap.

---

## 2. Filosofi Project

H.A.D.E.S. mengambil nama dari konsep “dunia bawah” bukan sebagai simbol gelap, tetapi sebagai metafora sistem yang bekerja diam-diam di bawah permukaan workflow produksi.

Operator melihat CorelDRAW, layout, pola, nama, nomor, dan hasil print. Namun di balik itu, H.A.D.E.S. menjaga data, struktur, validasi, metadata, dan alur eksekusi.

Secara filosofi:

```text
Manusia tetap memegang kontrol visual dan keputusan produksi.
H.A.D.E.S. mengambil alih bagian repetitif, rawan slip, dan berbasis aturan.
```

Tujuannya bukan menggantikan kreativitas layouter, tetapi menggantikan bagian mekanis yang sering menjadi sumber reject.

---

## 3. Stack Teknologi

Project H.A.D.E.S. terdiri dari dua sisi utama.

### 3.1 CorelDRAW VBA

Digunakan untuk:

```text
- membaca selection di CorelDRAW
- membaca red outline pola
- membaca text aktif
- membaca dan menulis metadata Shape.Name
- auto duplicate
- auto rename
- auto nesting template LRP
- auto mass nesting
- QC size
- QC typo
- QC IDPO
- report final
- convert/finalize
```

Target utama saat ini:

```text
CorelDRAW 2021 VBA
```

### 3.2 Python

Digunakan untuk:

```text
- membaca PDF Purchase Order
- mengekstrak size, nama, nomor, nickname
- membuat Order.txt
- menambahkan metadata order seperti @SIZEDB dan @IDPO
- pattern fetcher / pencarian pola dari folder database/NAS
- launcher/command center sederhana
```

Dependency Python utama:

```bash
pip install pdfplumber watchdog
```

Dependency ini juga tercantum di `requirements.txt` pada dokumentasi project.

---

## 4. File Data Utama

### 4.1 Order.txt

`Order.txt` adalah sumber data order yang dibaca oleh VBA.

Lokasi default:

```text
C:\Users\<USER>\Documents\Order.txt
```

Format utama:

```text
SIZE|NAMA|NOMOR|NICKNAME
```

Contoh:

```text
M|RINA|7|
M|SUSI|11|
XL|TUTI|13|
```

Field boleh kosong, tetapi separator `|` tetap dipertahankan:

```text
M||10|
M|DONI||DN
XL|||
```

Duplicate row harus tetap ditulis literal karena mewakili quantity.

### 4.2 Metadata Order

`Order.txt` juga mendukung metadata dengan prefix `@`.

Contoh:

```text
@JENIS_PESANAN=JERSEY
@JENIS_POLA=JERSEY REGULER
@MODEL_JAHIT=DEWASA PRIA
@SIZEDB=SizeDB_Pria.txt
@IDPO=365455
M|RINA|7|
M|SUSI|11|
XL|TUTI|13|
```

Metadata penting:

```text
@SIZEDB  → mengunci database size yang harus dipakai
@IDPO    → nomor PO/ID produksi yang harus divalidasi
```

Jika `@SIZEDB` ada tetapi file database tidak ditemukan, sistem sebaiknya FAIL/warning, bukan diam-diam memilih database lain.

---

## 5. Size Database

SizeDB disimpan di folder Documents.

Contoh:

```text
SizeDB_Pria.txt
SizeDB_Wanita.txt
SizeDB_Anak.txt
SizeDB_Jaket.txt
SizeDB_JaketAnak.txt
SizeDB_CelanaPria.txt
SizeDB_CelanaWanita.txt
SizeDB_CelanaAnak.txt
SizeDB_PriaSlimFit.txt
SizeDB_WanitaSlimFit.txt
```

### 5.1 Format Jersey Regular

```text
SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
```

Contoh:

```text
M|54.600|75.104|74.937
L|56.600|77.104|76.937
XL|58.600|79.104|78.937
```

### 5.2 Format Jaket / Split Front

```text
SIZE|LEBAR_BELAKANG|LEBAR_DEPAN|TINGGI_DEPAN|TINGGI_BELAKANG
```

### 5.3 Format Celana

```text
SIZE|L_DEPAN|L_BELAKANG
```

SizeDB digunakan oleh banyak modul:

```text
- QC Size
- QC Typo
- QC Typo Nest
- Auto Duplicate
- Record Pattern Catalog
- Auto Mass Nesting
- panel/size detection
```

---

## 6. Struktur 5 Shortcut Utama

Project H.A.D.E.S. dirancang agar operator tidak perlu menjalankan puluhan macro secara acak. Struktur utamanya diringkas menjadi 5 shortcut besar.

```text
1. HADES_PRECHECK_MASTER
2. HADES_PREPARE_MASTER
3. HADES_EXECUTE_LAYOUT
4. HADES_QC_FINAL
5. HADES_FINALIZE_CONVERT
```

### 6.1 HADES_PRECHECK_MASTER

Untuk mengecek risiko pada master sebelum layout dijalankan.

Contoh modul:

```text
- Transparency / PowerClip Check
- Visual risk check
- master object preflight
```

Tujuan:

```text
mencegah masalah visual/struktur masuk ke tahap layout.
```

### 6.2 HADES_PREPARE_MASTER

Untuk menyiapkan data/template sebelum eksekusi layout.

Contoh modul:

```text
- Auto Arrange Master Pola
- Auto Re-Contour Placeholder
- Build Typo Template
- Build Nesting Template LRP
- Record Pattern Catalog untuk Mass Nesting
```

### 6.3 HADES_EXECUTE_LAYOUT

Untuk menjalankan proses layout berdasarkan jalur workflow.

Dua jalur utama:

```text
Jalur LRP / Template Layout
Jalur Mass Nesting / Panel Layout
```

### 6.4 HADES_QC_FINAL

Untuk menjalankan QC produksi sebelum finalisasi.

Contoh modul:

```text
- QC Size
- QC Typo Auto Dispatcher
- QC Typo Normal
- QC Typo Nest
- IDPO Check
- Group Structure Check
- PowerClip/Transparency Check
- Global Final Report
```

### 6.5 HADES_FINALIZE_CONVERT

Tahap akhir. Idealnya hanya boleh berjalan jika QC Final PASS.

Fungsi:

```text
- cek QC lock
- cek selection fingerprint
- cek report PASS terbaru
- convert/finalize
- cleanup metadata bila diperlukan
```

---

## 7. Dua Jalur Layout Resmi

H.A.D.E.S. mengenali dua jenis workflow layout yang berbeda.

---

### 7.1 Jalur LRP / Template-Based Layout

LRP bukan Mass Nesting. LRP adalah workflow berbasis template.

Alur ideal:

```text
Build Typo Template
↓
Build Nesting Template LRP
↓
Auto Nesting Template LRP
↓
Auto Duplicate Adaptive Grid
↓
Auto Rename
↓
QC Size
↓
QC Typo Normal
↓
QC Final
```

Ciri utama:

```text
- hasil layout masih berupa jersey/set relatif utuh
- 1 group biasanya mewakili 1 jersey / 1 order row
- cocok untuk QC Typo Normal
- tidak memakai metadata HADES_AMN
```

### 7.2 Jalur Mass Nesting / Panel-Based Layout

Mass Nesting berbeda dari LRP. Ia tidak memakai template LRP dan tidak memakai Auto Duplicate Adaptive Grid, karena duplicate dan rename sudah menjadi bagian internal dari Auto Mass Nesting.

Alur ideal:

```text
Build Typo Template
↓
Record Pattern Catalog
↓
Auto Mass Nesting
↓
Auto Rename internal
↓
QC Size
↓
QC Typo Nest
↓
QC Final
```

Ciri utama:

```text
- satu jersey/order row bisa terpecah menjadi banyak panel
- body depan, body belakang, sleeve, collar bisa berada di box berbeda
- membutuhkan metadata HADES_AMN|ROW=...
- cocok untuk QC Typo Nest
```

---

## 8. QC Typo Auto Dispatcher

Salah satu upgrade terbaru Project H.A.D.E.S. adalah penyatuan pintu QC Typo.

Sebelumnya ada dua engine:

```text
QC Typo Normal → untuk LRP / Auto Duplicate
QC Typo Nest   → untuk Auto Mass Nesting
```

Sekarang keduanya dapat disatukan melalui dispatcher:

```text
QC_TYPO_CHECK
↓
scan selection
↓
jika ditemukan HADES_AMN|ROW=
    jalankan QC Typo Nest
else
    jalankan QC Typo Normal
```

Artinya operator cukup menjalankan:

```text
QC Typo Check
```

H.A.D.E.S. yang menentukan mode yang cocok.

### 8.1 QC Typo Normal

Dipakai untuk layout LRP.

Validasi:

```text
- nama atlet
- nomor atlet
- nickname
- pasangan nama-nomor
- expected count dari TypoTemplate
- indikasi nama/nomor tertukar
- hard exact validation untuk PASS
- fuzzy hanya untuk memilih kandidat report saat FAIL
```

### 8.2 QC Typo Nest

Dipakai untuk hasil Auto Mass Nesting.

Validasi:

```text
- membaca metadata HADES_AMN|ROW=
- mengumpulkan panel yang tersebar berdasarkan row order
- membaca Order.txt
- membaca TypoTemplate_Current.txt
- membaca SizeDB
- validasi nama/nomor/nickname per row
- kompatibel dengan Visual Ligature Breaker
- targeted green marker pada panel tempat typo terdeteksi
- marker dapat hilang dengan sekali Undo
```

QC Typo Nest V1.3 sudah berhasil diuji dan bekerja pada workflow Mass Nesting.

---

## 9. Metadata Internal

H.A.D.E.S. menggunakan metadata internal melalui `Shape.Name` untuk melacak struktur dan asal objek.

### 9.1 Metadata Pattern Catalog

```text
HADES_PC_UID|...
```

Digunakan oleh Record Pattern Catalog untuk menandai source panel.

Setelah Auto Mass Nesting selesai, metadata source ini boleh dibersihkan.

### 9.2 Metadata Auto Mass Nesting

```text
HADES_AMN|ROW=7|SIZE=M|PANEL=BODY_FRONT|UID=...
```

Digunakan oleh QC Typo Nest untuk mengetahui panel hasil nesting milik order row berapa.

Metadata ini sebaiknya dipertahankan minimal sampai QC Final selesai.

### 9.3 Metadata LRP

Jika nanti diperlukan, LRP sebaiknya memakai metadata berbeda, misalnya:

```text
HADES_LRP|ROW=...|SIZE=...|COPY=...
HADES_LRP_TEMPLATE|SLOT=...
```

Jangan memakai `HADES_AMN` untuk LRP, agar QC Typo Auto Dispatcher tidak salah mengira layout LRP sebagai Mass Nesting.

---

## 10. Modul Utama VBA

Dokumentasi lengkap dan file `.bas` berada di ZIP full documentation. README ini hanya menjelaskan peta sistem dan cara membacanya.

### 10.1 Controller

```text
HADES_5_SHORTCUTS_QC_MENU_V2_LOCKED.bas
```

Pintu utama 5 shortcut.

### 10.2 Core Foundation

```text
HADES_CORE_IO_PATHS_PHASE5.bas
HADES_CORE_ORDER_DB_PHASE5.bas
HADES_CORE_TEXT_NORMALIZE_PHASE5.bas
HADES_CORE_GEOMETRY_SELECTION_PHASE5.bas
HADES_CORE_REPORT_PHASE2.bas
HADES_CORE_SELF_TEST_PHASE5.bas
```

Fungsi:

```text
- path helper
- UTF-8 read/write
- Order.txt parser
- SizeDB loader
- text normalize
- red/green detection
- selection fingerprint
- report helper
```

### 10.3 QC Engines

```text
QC_TRANSPARENCY_POWERCLIP_CHECK_V1_0_USER_SUPPLIED.bas
VBA IDPO CHECK.bas
VBA_QC_SIZE_CHECK_V8_4.bas
VBA_QC_TYPO_CHECK_V13_2_GREEN_MARKER.bas
VBA_QC_TYPO_NEST_V1_3_TARGETED_GREEN_MARKER.bas
VBA_QC_TYPO_AUTO_DISPATCHER_V1_0.bas
GROUP_STRUCTURE_CHECK_V1_CLEANED.bas
```

Catatan: nama file bisa berbeda jika repository sudah diupdate, tetapi fungsi besarnya tetap sama.

### 10.4 Layout Engines

```text
HADES_AUTO_NESTING_TEMPLATE_V1_3_SIZE_BLOCK_LOCK.bas
HADES_RECORD_PATTERN_CATALOG_V3_3_AUTOONLY_UID_MARKER.bas
VBA_AUTO_DUPLICATE_V2_3_1_ADAPTIVE_GRID_FIXED.bas
VBA_AUTO_MASS_NESTING_V3_4_UID_METADATA.bas
VBA_AUTO_RENAME_V5_3_VISUAL_LIGATURE_BREAKER.bas
```

### 10.5 Prepare Master Engines

```text
BUILD_TYPO_TEMPLATE_V2_CLEANED.bas
VBA AI TEXT.bas
VBA AUTO ARRANGE.bas
VBA AUTO RE-CONTOUR.bas
```

### 10.6 Database Mining

```text
MINE_SIZE_DATABASE_SINGLE.bas
MINE_SPLIT_FRONT_DATABASE.bas
MINE_PANTS.bas
```

Digunakan untuk membangun SizeDB dari pola master.

---

## 11. Python Tools

Folder Python di dokumentasi berisi beberapa tool pendukung.

### 11.1 Order Extractor

Fungsi:

```text
PDF PO → Order.txt
```

Menggunakan:

```bash
pip install pdfplumber watchdog
```

Output:

```text
Documents\Order.txt
```

### 11.2 Pattern Fetcher

Fungsi:

```text
mencari/copy pattern file dari database/NAS berdasarkan kode pola
```

### 11.3 Launcher

Rencana/arah pengembangan:

```text
HADES Python Command Center
```

Tujuan:

```text
menggabungkan Order Extractor, Pattern Fetcher, watcher, dan menu Python dalam satu launcher.
```

---

## 12. Urutan Import VBA

Urutan import penting agar module yang saling memanggil tidak bentrok.

Urutan umum:

```text
1. Core Foundation
2. Report Lock / Finalize Gate
3. QC Engines
4. Prepare Master Engines
5. Layout Engines
6. Database Mining Tools
7. Controller terakhir
```

Controller sebaiknya di-import terakhir karena memanggil module lain.

Jika memakai patch terbaru, jangan import versi lama berdampingan dengan versi baru yang memiliki public Sub sama.

Contoh konflik yang harus dihindari:

```text
QC_TYPO_CHECK lama
vs
QC_TYPO_AUTO_DISPATCHER baru
```

Jika dispatcher sudah dipakai, maka public entry `QC_TYPO_CHECK` sebaiknya dimiliki oleh dispatcher, bukan engine normal lama.

---

## 13. Workflow Produksi Ringkas

### 13.1 Workflow LRP

```text
1. Python membuat Order.txt dari PDF PO
2. Operator menyiapkan master desain dan master pola
3. PREPARE MASTER:
   - Build Typo Template
   - Build Nesting Template LRP
4. EXECUTE LAYOUT:
   - Auto Nesting Template LRP
   - Auto Duplicate Adaptive Grid
   - Auto Rename
5. QC FINAL:
   - QC Size
   - QC Typo Auto → memilih QC Typo Normal
   - IDPO Check
   - Group Structure Check
6. FINALIZE CONVERT jika QC PASS
```

### 13.2 Workflow Mass Nesting

```text
1. Python membuat Order.txt dari PDF PO
2. Operator menyiapkan panel source
3. PREPARE MASTER:
   - Build Typo Template
   - Record Pattern Catalog
4. EXECUTE LAYOUT:
   - Auto Mass Nesting
   - Auto Rename internal
   - metadata HADES_AMN dibuat
5. QC FINAL:
   - QC Size
   - QC Typo Auto → memilih QC Typo Nest
   - IDPO Check
   - Group Structure Check
6. FINALIZE CONVERT jika QC PASS
```

---

## 14. Status Modul Saat Ini

Status ringkas:

```text
QC Size Check                 : stabil / produksi-ready
Auto Rename                   : kuat / sangat berguna
Auto Duplicate Adaptive Grid  : stabil untuk jalur LRP
QC Typo Normal                : matang untuk LRP
QC Typo Nest                  : berhasil work sampai V1.3
QC Typo Auto Dispatcher       : arsitektur benar, perlu integrasi controller/report
Record Pattern Catalog UID    : fondasi Mass Nesting
Auto Mass Nesting UID         : memakai engine yang hasil nestingnya paling cocok + metadata
IDPO Check                    : layak jadi QC wajib
Group Structure Check         : berguna sebagai struktur guard
Build Typo Template           : fondasi Auto Rename/QC Typo
Build Nesting Template LRP    : fondasi LRP template-based layout
Finalize Convert Lock         : next milestone penting
Auto Layout sejati            : eksperimental, belum menjadi fokus utama
```

---

## 15. Roadmap Terdekat

Prioritas utama setelah integrasi QC Typo Nest dan dispatcher:

```text
1. Integrasikan QC Typo Auto Dispatcher ke HADES_QC_FINAL.
2. Pastikan QC Final Report memanggil dispatcher, bukan engine lama langsung.
3. Buat/rapikan PASS Lock agar Finalize Convert hanya bisa jalan jika QC Final PASS.
4. Pastikan selection fingerprint/report lock tidak basi.
5. Rapikan PREPARE_MASTER agar Build Typo, Build Nesting LRP, dan Record Pattern Catalog mudah dipilih.
6. Stabilkan dua jalur EXECUTE_LAYOUT: LRP dan Mass Nesting.
```

Auto Layout sejati tetap menjadi visi jangka panjang, tetapi belum menjadi fokus karena membutuhkan deteksi desain/pola yang jauh lebih kompleks.

---

## 16. Catatan Penting

Project H.A.D.E.S. dibuat untuk workflow spesifik CorelDRAW + produksi sportswear/sublimasi. Banyak modul bergantung pada konvensi internal seperti:

```text
- red outline pola RGB 255,0,0
- Order.txt di Documents
- SizeDB di Documents
- layout digroup dengan struktur tertentu
- text masih aktif sebelum convert
- metadata Shape.Name tidak dibersihkan sebelum QC selesai
```

Sebelum digunakan pada produksi nyata:

```text
- backup file CorelDRAW
- test di copy file
- jalankan QC sebelum convert
- jangan convert text sebelum QC Typo dan IDPO Check selesai
- jangan hapus metadata HADES_AMN sebelum QC Typo Nest selesai
```

---

## 17. Prinsip Pengembangan

Prinsip utama Project H.A.D.E.S.:

```text
Jangan membuat operator memilih hal teknis jika sistem bisa mendeteksinya.
Jangan menggabungkan semua engine menjadi satu kode raksasa.
Gunakan dispatcher/controller untuk memilih engine yang benar.
Pertahankan backward compatibility selama mungkin.
Pisahkan jalur LRP dan Mass Nesting secara jelas.
Buat QC sebagai gerbang, bukan sekadar laporan.
```

Dengan prinsip ini, H.A.D.E.S. berkembang dari kumpulan macro menjadi sistem produksi bertahap:

```text
Data → Layout → Metadata → QC → Lock → Finalize
```

---

## 18. Penutup

Project H.A.D.E.S. adalah sistem otomasi produksi yang lahir dari kebutuhan nyata operator layout: mengurangi repetisi, mengurangi slip, dan membuat pekerjaan CorelDRAW yang sebelumnya sangat manual menjadi lebih terstruktur.

Nilai utamanya bukan hanya kecepatan, tetapi **kontrol kesalahan**.

```text
H.A.D.E.S. tidak hanya membantu layout.
H.A.D.E.S. menjaga agar layout yang salah tidak mudah lolos ke tahap final.
```

