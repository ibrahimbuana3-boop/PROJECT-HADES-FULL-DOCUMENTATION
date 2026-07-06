# Project H.A.D.E.S.

**H.A.D.E.S.** = **H**ybrid **A**utomation & **D**ata **E**xecution **S**ystem  
**Sistem Otomasi Hibrida & Eksekusi Data**

Project H.A.D.E.S. adalah sistem otomasi dan QC internal untuk workflow layout sportswear sublimasi berbasis **Python + CorelDRAW VBA**. Sistem ini dirancang untuk menghubungkan data Purchase Order dari ERP/PDF dengan proses eksekusi layout, auto rename, duplicate, nesting, dan quality control di CorelDRAW.

H.A.D.E.S. bukan sekadar kumpulan macro. Ia adalah sistem produksi yang mencoba menjawab masalah nyata operator: pekerjaan repetitif, risiko salah size, salah jumlah, salah nama, salah nomor, lupa IDPO, font tidak mendukung karakter tertentu, dan kesalahan lain yang dapat menyebabkan reject produksi.

---

## 1. Makna Nama H.A.D.E.S.

### Hybrid

Project ini bekerja secara hybrid karena menggabungkan dua dunia:

- **Python** sebagai data layer.
- **CorelDRAW VBA** sebagai execution layer.

Python membaca dan membersihkan data dari PDF PO/ERP, lalu menghasilkan `Order.txt` sebagai jembatan data. CorelDRAW VBA kemudian memakai data tersebut untuk melakukan otomasi dan QC pada layout nyata di CorelDRAW.

### Automation

H.A.D.E.S. memangkas proses yang repetitif, seperti:

- menyalin data PO ke format kerja,
- mengganti nama/nomor/nickname secara manual,
- menghitung quantity size,
- menggandakan layout berdasarkan order,
- mengecek size satu per satu,
- mengecek IDPO secara manual,
- menyusun panel layout tertentu secara massal.

### Data

Data utama H.A.D.E.S. berasal dari PO dan disimpan dalam format runtime sederhana:

```txt
Documents\Order.txt
```

Format stabil:

```txt
@JENIS_PESANAN=JERSEY
@JENIS_POLA=JERSEY REGULER
@MODEL_JAHIT=DEWASA PRIA
@SIZEDB=SizeDB_Pria.txt
@IDPO=365455
M|RINA|7|
M|SUSI|11|
XL|TUTI|13|TUTI
```

Baris metadata diawali `@`. Baris order memakai format:

```txt
SIZE|NAMA|NOMOR|NICKNAME
```

Duplicate row harus dipertahankan karena berarti quantity.

### Execution

Data dari PO tidak hanya disimpan sebagai teks. Data tersebut dieksekusi menjadi aksi nyata:

- jika PO meminta size M sebanyak 3 pcs, layout harus mencerminkan M sebanyak 3 pcs;
- jika PO berisi nama RINA nomor 7, Auto Rename dan QC Typo harus mengikuti data itu;
- jika metadata berisi `@SIZEDB=SizeDB_Pria.txt`, QC Size harus memakai database ukuran tersebut;
- jika metadata berisi `@IDPO=365455`, IDPO Check harus memvalidasi angka tersebut di layout.

### System

H.A.D.E.S. adalah sistem, bukan satu macro tunggal. Ia terdiri dari:

- Python Auto-PO,
- `Order.txt`,
- `SizeDB_*.txt`,
- Core Loader,
- Auto Rename,
- Auto Duplicate,
- Auto Mass Nesting,
- QC Size,
- QC Typo,
- IDPO Check,
- Group Structure Check,
- Final QC Report,
- Finalize Convert Gate.

---

## 2. Filosofi Project H.A.D.E.S.

H.A.D.E.S. lahir dari kebutuhan operator, bukan dari konsep software formal di awal. Sistem ini dibuat untuk mendukung workflow produksi yang sudah berjalan, terutama di titik-titik yang rawan human error.

Project ini tidak dimaksudkan untuk melanggar SOP. Sebaliknya, H.A.D.E.S. berperan sebagai lapisan pendukung di bawah SOP:

> SOP mengatur alur resmi. H.A.D.E.S. mencoba menutup celah-celah kecil yang sering menyebabkan reject produksi.

Contoh celah yang ingin ditutup:

- operator lupa mengganti IDPO,
- IDPO lama terbawa ke order baru,
- nama atlet typo,
- nomor atlet salah,
- size salah jumlah,
- database ukuran salah dipilih,
- font tidak mendukung karakter Unicode/CJK,
- ligature membuat karakter berubah atau hilang,
- layout sudah convert sebelum QC selesai.

Dengan kata lain, H.A.D.E.S. adalah sistem **anti-reject** dan **workflow guard**.

---

## 3. Mitologi dan Simbolisme Hades

Nama Hades dipilih bukan untuk menggambarkan sesuatu yang negatif, tetapi karena proyek ini memiliki karakter “bekerja dari bawah layar”.

Hades dalam mitologi sering diasosiasikan dengan dunia bawah: wilayah yang tidak selalu terlihat dari permukaan, tetapi memiliki struktur, hukum, dan peran penting dalam menjaga keseimbangan.

Project H.A.D.E.S. juga bekerja seperti itu:

- banyak prosesnya terjadi di balik layar,
- tidak selalu memiliki tampilan interaktif,
- tidak menampilkan animasi proses yang mencolok,
- hanya ada jeda beberapa detik, lalu hasil layout/QC muncul,
- banyak validasinya bersifat invisible sampai terjadi error.

Namun justru di situlah perannya: menjaga produksi dari bawah, sebelum kesalahan muncul di permukaan.

Kalimat simboliknya:

> Project H.A.D.E.S. adalah sistem bawah layar yang lahir dari kebutuhan operator, bekerja di luar sorotan visual, tetapi bertujuan menjaga produksi agar tidak jatuh ke kesalahan yang sama.

---

## 4. Posisi H.A.D.E.S. terhadap Tool Nesting Lain

Beberapa plugin nesting seperti Aimari Nest berfokus pada efisiensi material: bagaimana panel disusun agar kain atau area print lebih hemat.

H.A.D.E.S. memiliki fokus yang berbeda.

| Aspek | Tool Nesting Umum | Project H.A.D.E.S. |
|---|---|---|
| Fokus utama | Efisiensi material/nesting | Anti-reject, QC, workflow automation |
| Input konteks | Sering melalui klik manual panel/size | `Order.txt`, `@SIZEDB`, SizeDB, Pattern Catalog |
| Visual process | Biasanya lebih interaktif | Banyak bekerja sebagai batch process |
| Auto Rename | Belum tentu menjadi fokus | Fitur penting |
| Unicode fallback | Belum tentu ada | Dipertahankan sebagai fitur produksi |
| QC Size/Typo/IDPO | Belum tentu ada | Bagian inti sistem |
| Filosofi | Material optimization | Production correctness |

H.A.D.E.S. tidak harus menjadi plugin nesting paling hemat kain. Auto Mass Nesting dalam H.A.D.E.S. bertujuan mengurangi pekerjaan susun manual, menjaga struktur layout tetap terkendali, dan tetap kompatibel dengan QC.

---

## 5. Arsitektur Umum

```txt
PDF PO / ERP
    ↓
Python Robot Auto-PO
    ↓
Documents\Order.txt
    ↓
CorelDRAW VBA Core Loader
    ↓
Auto Duplicate / Auto Rename / Pattern Catalog / Auto Mass Nesting
    ↓
QC Size / QC Typo / IDPO Check / Group Structure Check
    ↓
Final QC Report
    ↓
Finalize Convert Gate
```

### Python Layer

Python bertugas:

- memantau folder Downloads,
- membaca PDF PO,
- membersihkan data order,
- mempertahankan duplicate row,
- mempertahankan empty field,
- menulis metadata `@`,
- menentukan `@SIZEDB`,
- mengambil `@IDPO`,
- menulis `Documents\Order.txt` dengan format CRLF Windows.

### Data Bridge

File penting:

```txt
Documents\Order.txt
Documents\SizeDB_*.txt
Documents\HADES_PATTERN_CATALOG_CURRENT.txt
```

`Order.txt` adalah kontrak data runtime.  
`SizeDB_*.txt` adalah kontrak geometri.  
Pattern Catalog adalah peta panel master untuk Auto Mass Nesting.

### VBA Execution Layer

CorelDRAW VBA bertugas:

- membaca Order.txt,
- membaca metadata,
- memilih SizeDB yang tepat,
- mengeksekusi duplicate/rename/nesting,
- membaca geometri red outline,
- melakukan QC,
- menulis report,
- mengunci finalisasi jika QC belum PASS.

---

## 6. Workflow Produksi

Workflow utama:

```txt
1. Operator menerima PDF PO dari ERP.
2. Python Auto-PO membaca PDF.
3. Python membuat Documents\Order.txt.
4. Operator membuka master desain dan pola di CorelDRAW.
5. Operator membuat layout.
6. Auto Duplicate / Auto Rename membantu proses repetitif.
7. QC Size memvalidasi jumlah dan size berdasarkan geometri.
8. QC Typo memvalidasi nama, nomor, dan nickname.
9. IDPO Check memvalidasi IDPO kecil di layout.
10. Final QC Report membuat rekap PASS/FAIL.
11. Finalize Convert hanya boleh lanjut jika QC final PASS.
```

Workflow alternatif Auto Mass Nesting:

```txt
1. Master layout masih group per size.
2. Run HADES_RECORD_PATTERN_CATALOG.
3. Pattern Catalog mendeteksi size group dari SizeDB.
4. Setiap panel child mendapat UID/marker sementara.
5. Operator ungroup parent size sekali.
6. Select panel source.
7. Run HADES_AUTO_MASS_NESTING.
8. Sistem duplicate, auto rename, dan nesting berdasarkan Order.txt.
9. Hasil diletakkan ke box produksi.
10. Marker Pattern Catalog dibersihkan setelah sukses.
11. Hasil tetap harus melewati QC final.
```

---

## 7. Core Data: Order.txt

Format stabil:

```txt
SIZE|NAMA|NOMOR|NICKNAME
```

Contoh:

```txt
M|RINA|7|
M|SUSI|11|
XL|TUTI|13|TUTI
```

Empty field boleh, separator tetap wajib:

```txt
M|||
M|DONI||
M||10|
```

Metadata diawali `@`:

```txt
@JENIS_PESANAN=JERSEY
@JENIS_POLA=JERSEY REGULER
@MODEL_JAHIT=DEWASA PRIA
@SIZEDB=SizeDB_Pria.txt
@IDPO=365455
```

Macro yang membaca order rows wajib skip metadata `@`.

---

## 8. Core Data: SizeDB

Lokasi umum:

```txt
Documents\SizeDB_*.txt
```

Format jersey regular:

```txt
SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
```

Format split front:

```txt
SIZE|LEBAR_BELAKANG|LEBAR_DEPAN|TINGGI_DEPAN|TINGGI_BELAKANG
```

Format celana:

```txt
SIZE|L_DEPAN|L_BELAKANG
```

Prinsip penting:

- `@SIZEDB` di Order.txt adalah context lock.
- Jika `@SIZEDB` ada, VBA harus memakai file itu.
- Popup manual hanya fallback jika metadata tidak ada.
- Jika `@SIZEDB` ada tetapi file tidak ditemukan, sistem harus warning/FAIL.
- Jangan menebak database dari dimensi global karena ukuran antar produk bisa overlap.

---

## 9. Pattern Catalog

Pattern Catalog digunakan terutama oleh Auto Mass Nesting.

Tujuannya:

- membaca master layout yang masih group per size,
- mendeteksi size group dari body panel dan SizeDB,
- mencatat semua child panel dalam group tersebut,
- memberi konteks size ke panel yang tidak punya SizeDB sendiri seperti lengan, kerah, dan tulangan,
- memberi UID/marker sementara agar panel source dapat dilacak setelah ungroup.

Contoh master:

```txt
Group M
├─ Baju depan
├─ Baju belakang
├─ Lengan kanan
├─ Lengan kiri
├─ Kerah
└─ Tulangan
```

Setelah catalog:

```txt
M|BODY_FRONT|BODY|...
M|BODY_BACK|BODY|...
M|SLEEVE|SLEEVE|...
M|SLEEVE|SLEEVE|...
M|COLLAR|SMALL|...
M|TULANGAN|SMALL|...
```

Lengan, kerah, dan tulangan tidak perlu SizeDB sendiri karena mereka mewarisi size dari parent group.

---

## 10. Auto Mass Nesting

Auto Mass Nesting adalah engine layout alternatif/advanced. Untuk saat ini, modul ini belum dimasukkan ke 5 shortcut utama PHASE.

Prinsip baru Auto Mass Nesting:

```txt
Bukan 6 panel x 2 baris.
Tetapi maksimal 6 box produksi ke samping.
```

Setiap box memiliki batas fisik:

```txt
178 cm x 255 cm
```

Setiap box berisi gugusan panel yang diprioritaskan homogen:

- size sama,
- bucket sama,
- panel type sama,
- body depan dan belakang tidak dicampur jika memungkinkan,
- lengan dikumpulkan sesama lengan,
- kerah/tulangan masuk small panel box.

Jika box ke-7 dibuat, maka harus turun ke baris berikutnya, bukan terus memanjang ke kanan memenuhi desktop CorelDRAW.

Auto Rename internal dalam Auto Mass Nesting tetap dipertahankan karena duplicate dan rename adalah satu paket dalam workflow mass nesting.

---

## 11. Quality Control Modules

### QC Size Check

Fungsi:

- membaca Order.txt,
- membaca expected quantity per size,
- membaca SizeDB sesuai `@SIZEDB`,
- scan red outline panel dari selection,
- cocokkan dimensi dengan SizeDB,
- bandingkan jumlah layout dengan order,
- PASS/FAIL berdasarkan size dan quantity.

QC Size adalah geometric validator. Ia tidak membaca label visual size.

### QC Typo Check

Fungsi:

- membaca Order.txt,
- scan teks aktif sebelum convert,
- validasi nama, nomor, nickname,
- menghindari false match substring,
- mengabaikan IDPO 6 digit jika konteksnya bukan nomor atlet,
- mendukung nested group scan.

### IDPO Check

Fungsi:

- membaca `@IDPO` dari Order.txt,
- mencari teks kecil IDPO sekitar 0.3–0.6 cm,
- mendeteksi placeholder `IDPO`,
- mendeteksi IDPO lama atau angka 6 digit yang tidak sesuai.

### Final QC Report

Fungsi:

- menjalankan/mengumpulkan hasil QC,
- menulis report,
- membuat final QC lock,
- menjadi gate sebelum Finalize Convert.

---

## 12. Prinsip Desain

1. Jangan rewrite dari nol jika engine lama masih bisa dipertahankan.
2. Jangan mengubah format Order.txt sembarangan.
3. Jangan mengubah format SizeDB sembarangan.
4. Pertahankan backward compatibility.
5. Gunakan `@SIZEDB` sebagai context lock.
6. Macro harus robust terhadap nested group.
7. Macro tidak boleh bergantung pada white page.
8. Macro harus bisa bekerja di desktop CorelDRAW.
9. Macro yang memberi marker visual harus bisa di-undo atau cleanup.
10. Jika ada error, report harus jelas.

---

## 13. Kekuatan dan Kelemahan

### Kekuatan

- Dekat dengan masalah nyata operator.
- Mengurangi kerja repetitif.
- Mencegah human error sebelum produksi.
- Menggunakan data PO sebagai sumber kebenaran.
- Memakai SizeDB sebagai validator geometri.
- Mendukung Unicode fallback dan ligature handling.
- Bisa membuat QC final sebelum convert.

### Kelemahan saat ini

- Masih abstrak.
- Belum memiliki UI interaktif seperti plugin komersial.
- Banyak proses berjalan di balik layar.
- Output sering berupa popup/report TXT.
- Membutuhkan disiplin struktur file dan group layout.

Kelemahan ini bukan kelemahan logika, melainkan kelemahan presentasi. Roadmap jangka menengah dapat mencakup Command Center, progress status, box guide, preview/dry run, dan report viewer yang lebih visual.

---

## 14. Identity Statement

Project H.A.D.E.S. adalah sistem otomasi hibrida dan QC produksi sportswear sublimasi yang menghubungkan data PO dari ERP dengan eksekusi layout di CorelDRAW. Sistem ini bekerja sebagai lapisan bawah layar untuk mempercepat proses operator, menjaga konsistensi data, dan mencegah reject produksi sebelum file masuk tahap final.

Versi pendek:

> H.A.D.E.S. mengubah data PO menjadi aksi layout dan QC di CorelDRAW.

Versi teknis:

> H.A.D.E.S. adalah rule-based production intelligence system berbasis Python + CorelDRAW VBA untuk otomasi layout, validasi order, dan pencegahan reject produksi.

Versi filosofis:

> H.A.D.E.S. bekerja dari bawah layar: tidak selalu terlihat, tetapi menjaga produksi dari kesalahan yang sering luput dari mata manusia.

---

## 15. Status Saat Ini

Project H.A.D.E.S. sudah memiliki beberapa modul aktif:

- Python Robot Auto-PO,
- Order.txt generator,
- Auto Rename,
- Auto Duplicate,
- QC Size Check,
- QC Typo Check,
- IDPO Check,
- Final QC Report,
- Finalize Convert Gate,
- Pattern Catalog,
- Auto Mass Nesting dengan Auto Rename internal.

Auto Mass Nesting masih diposisikan sebagai engine alternatif/advanced dan belum dimasukkan ke shortcut PHASE utama. Fokus utama H.A.D.E.S. tetap: **anti-reject, data-driven layout, dan workflow automation**.
