# QC Final Menu V2 Locked

Module:

```text
src/vba/00_controller/HADES_5_SHORTCUTS_QC_MENU_V2_LOCKED.bas
```

Shortcut tetap:

```text
HADES_QC_FINAL
```

Menu:

```text
1. Jalankan PowerClip & Transparency
2. Jalankan IDPO Check
3. Jalankan Size Check
4. Jalankan Typo Check
5. Jalankan Group Structure Check
6. Jalankan semuanya secara berurutan
```

---

## Logika Lock

### Pilihan 1–5

Dipakai untuk cek mandiri.

```text
Run QC mandiri
↓
Finalize Convert invalidated
```

Operator wajib menjalankan pilihan 6 untuk membuka convert.

### Pilihan 6

Dipakai untuk final gate.

```text
Run Global Final Report
↓
Jika semua PASS → Finalize Convert ALLOWED
Jika satu saja FAIL → Finalize Convert BLOCKED
```

---

## Alasan Desain

Cek mandiri berguna untuk debugging, tetapi tidak boleh dianggap sebagai final pass produksi. Final pass resmi hanya dari Global Final Report.
