# SizeDB Guide

SizeDB adalah database dimensi pola resmi dalam centimeter.

---

## Jersey Regular / Body Depan-Belakang

Format:

```text
SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
```

Contoh:

```text
XS|50.600|71.104|70.937
S|52.600|73.104|72.937
M|54.600|75.104|74.937
L|56.600|77.104|76.937
XL|58.600|79.104|78.937
```

Digunakan oleh:

```text
QC_SIZE_CHECK
QC_TYPO_CHECK
AUTO_DUPLICATE
RECORD_PATTERN
AUTO_MASS_NESTING
```

---

## Split Front

Format:

```text
SIZE|LEBAR_BELAKANG|LEBAR_DEPAN|TINGGI_DEPAN|TINGGI_BELAKANG
```

Contoh:

```text
M|56.000|29.500|74.000|75.000
```

Digunakan untuk:

```text
Jaket
Jaket Anak
Kemeja
Rompi
```

---

## Penambahan SizeDB Baru

Jika format masih sama seperti jersey biasa:

```text
SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
```

maka cukup:

```text
1. Taruh file SizeDB baru di Documents.
2. Pastikan Order.txt menulis @SIZEDB=nama_file_baru.txt.
3. Engine membaca file itu.
```

Contoh:

```text
@SIZEDB=SizeDB_JunkiesWanita.txt
```

Jika ingin muncul di popup fallback manual, baru registry/popup perlu direvisi.
