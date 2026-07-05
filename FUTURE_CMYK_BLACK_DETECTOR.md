# Future — VBA Black CMYK Detector

Status:

```text
WACANA / BELUM AKTIF
```

Tujuan:

```text
Mendeteksi warna hitam CMYK solid yang berisiko keluar abu-abu saat print/press sublimasi.
```

---

## Target Dicek

```text
Uniform fill vector
Uniform outline vector
Text fill uniform
Object dalam group / nested group
```

## Diabaikan

```text
Bitmap
Gradient / fountain fill
Pattern fill
Texture fill
Isi PowerClip
```

Alasan isi PowerClip diabaikan:

```text
PowerClip sudah harus FAIL oleh PowerClip & Transparency Check sebelum CMYK Black Detector dijalankan.
```

---

## Rencana Integrasi

Saat dibuat, module baru dapat masuk ke:

```text
HADES_PRECHECK_MASTER
HADES_QC_FINAL_MENU_V3_LOCKED
```

QC Final Menu V2 belum berubah karena detector belum aktif.
