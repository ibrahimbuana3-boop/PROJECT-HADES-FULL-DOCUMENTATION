# Database Mining Tools

Mining tools bukan workflow layout harian. Ini alat admin/database.

---

## MINE_SIZE_DATABASE

File:

```text
src/vba/03_engines/database_mining/MINE_SIZE_DATABASE_SINGLE.bas
```

Fungsi:

```text
Menambang body depan + belakang dari size terkecil sampai terbesar.
```

Output:

```text
SIZE|LEBAR|TINGGI_DEPAN|TINGGI_BELAKANG
```

Cara pakai:

```text
1. Select semua body depan + belakang.
2. Pastikan outline merah.
3. Jalankan MINE_SIZE_DATABASE.
4. Masukkan size terkecil.
5. Simpan sebagai SizeDB_*.txt.
```

---

## MINE_SPLIT_FRONT

File:

```text
src/vba/03_engines/database_mining/MINE_SPLIT_FRONT_DATABASE.bas
```

Fungsi:

```text
Menambang 1 panel belakang + 1 panel depan.
```

Output:

```text
SIZE|LEBAR_BELAKANG|LEBAR_DEPAN|TINGGI_DEPAN|TINGGI_BELAKANG
```

Cara pakai:

```text
1. Select tepat 2 panel.
2. Jalankan MINE_SPLIT_FRONT.
3. Input size.
4. Input nama database.
```
