# QC Typo Green Marker

Current file:

```text
src/vba/03_engines/qc/VBA_QC_TYPO_CHECK_V13_2_GREEN_MARKER.bas
```

Entry point tetap:

```text
QC_TYPO_CHECK
HADES_QC_TYPO_REPORT
HADES5C_QC_TYPO_CORE_SMOKE_TEST
```

---

## Tujuan Upgrade

Saat Typo Check mendeteksi FAIL, layouter perlu tahu panel mana yang bermasalah tanpa membaca report panjang.

Fitur baru:

```text
Typo FAIL
↓
Cari panel body dalam group bermasalah
↓
Cocokkan panel dengan SizeDB
↓
Panel diberi outline hijau
```

---

## Prinsip

- Tidak mengubah shortcut.
- Tidak mengubah format Order.txt.
- Tidak mengubah SizeDB.
- Tidak mengubah QC Final Menu V2.
- Hanya menambahkan penanda visual panel bermasalah.

---

## Catatan Warna

Hijau dalam konteks ini berarti:

```text
Panel ditandai oleh sistem QC.
```

Bukan berarti panel tersebut PASS.
