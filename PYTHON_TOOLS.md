# Python Tools

Project Hades memakai Python sebagai alat pra-produksi.

---

## Order Extractor

Folder:

```text
src/python/order_extractor
```

Fungsi:

```text
PDF PO
↓
Order.txt
```

Output:

```text
Documents\Order.txt
```

Isi Order.txt:

```text
metadata @...
baris SIZE|NAMA|NOMOR|NICKNAME
```

---

## Pattern Fetcher

Folder:

```text
src/python/pattern_fetcher
```

Fungsi:

```text
Kode pola dari ERP
↓
Cari file di NAS
↓
Copy exact file pola ke local
```

Output:

```text
Documents\HADES_PATTERN_INBOX\FETCH_YYYYMMDD_HHMMSS
```

---

## Masalah Yang Diselesaikan Pattern Fetcher

Workflow manual lama:

```text
copy kode pola
paste ke search database pola
hapus 1 digit terakhir agar semua size muncul
pilih/drag manual
```

Risiko:

```text
PO minta M,L,XL
operator bisa salah ambil S,M,L
```

Pattern Fetcher mengurangi risiko dengan mencari kode exact yang dipaste.

---

## Future: Command Center

Kondisi sekarang:

```text
Order Extractor dan CARI_POLA.bat masih dua alat/window.
```

Future:

```text
HADES Python Command Center
= Order watcher + Pattern Fetcher dalam satu launcher.
```
