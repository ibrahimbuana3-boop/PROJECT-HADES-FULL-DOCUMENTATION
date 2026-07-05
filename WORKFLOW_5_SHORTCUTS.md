# Workflow 5 Shortcut

## 1. HADES_PRECHECK_MASTER

Tujuan:

```text
Cek file master sebelum diproses lebih jauh.
```

Isi saat ini:

```text
PowerClip & Transparency Check
```

Future:

```text
CMYK Black Detector
```

---

## 2. HADES_PREPARE_MASTER

Tujuan:

```text
Menyiapkan master sebelum layout.
```

Isi:

```text
Auto Arrange
Auto Re-Contour Placeholder
Build Typo Template
```

Catatan:

```text
Record Pattern Catalog V3.2 AutoOnly masih standalone, tetapi bisa diposisikan sebagai bagian prepare untuk jalur Auto Mass Nesting.
```

---

## 3. HADES_EXECUTE_LAYOUT

Tujuan:

```text
Menjalankan layout.
```

Jalur layout tidak boleh otomatis dipaksa berurutan karena berbeda mode:

```text
1. Auto Duplicate biasa
2. Auto Nesting Template / LRP
3. Auto Mass Nesting
```

---

## 4. HADES_QC_FINAL

Shortcut ini memakai QC Final Menu V2 Locked.

Pilihan:

```text
1 = PowerClip & Transparency
2 = IDPO Check
3 = Size Check
4 = Typo Check
5 = Group Structure Check
6 = Jalankan semuanya berurutan
```

Aturan:

```text
Pilihan 1-5 → Finalize Convert di-invalidasi.
Pilihan 6 → jika PASS semua, Finalize Convert boleh.
```

---

## 5. HADES_FINALIZE_CONVERT

Tujuan:

```text
Convert / finalisasi hanya setelah QC final PASS.
```

Jika QC final belum PASS atau ada FAIL:

```text
Finalize Convert BLOCKED
```
