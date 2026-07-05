# Future — Python Pattern Router

Pattern Fetcher saat ini masih menerima kode pola dari manusia.

Target masa depan:

```text
PDF PO
↓
Python baca spesifikasi jersey
↓
Python baca size yang diminta
↓
Python cocokkan ke database kode pola RND
↓
Python copy file pola exact dari NAS
```

---

## Kenapa Butuh Database RND

Mapping size/kode pola tidak universal.

Contoh dummy:

```text
Jersey(43)
Pria(1)
Lengan Pendek(1)
Tanpa Lis(5)
Kerah Oblong(12)
Size S(80)
→ 431151280
```

Tetapi size XS tidak selalu digit terakhir tertentu. Di spesifikasi lain, XS bisa 79, S 80, M 81, dst.

Karena itu sistem tidak boleh hanya mengandalkan rumus digit global.

---

## Source of Truth

Yang dibutuhkan:

```text
Database kode pola lengkap dari RND.
```

Minimal kolom:

```text
JENIS_PESANAN
JENIS_POLA
MODEL_JAHIT
GENDER
BENTUK_LENGAN
JENIS_LIS_LENGAN
MANSET
BENTUK_KRAH
KANTONG
SIZE
KODE_POLA
NAMA_FILE_POLA
PATH_NAS
STATUS_AKTIF
```

---

## Prinsip

```text
Jangan menebak.
Gunakan database resmi.
Jika ambiguous, STOP dan minta manusia memilih.
Mapping hasil koreksi bisa disimpan untuk job berikutnya.
```
