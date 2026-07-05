# Import Guide CorelDRAW VBA

## Prinsip Penting

```text
Replace module lama.
Jangan import berdampingan jika nama Sub sama.
```

Jika terjadi duplicate procedure:

```text
Compile error:
Ambiguous name detected
```

---

## Urutan Import

```text
1. Core foundation
2. Phase report/lock
3. QC engines
4. Prepare master engines
5. Layout engines
6. Database mining tools
7. Controller terakhir
```

Controller harus terakhir agar 5 shortcut terbaru yang aktif.

---

## Module Yang Tidak Boleh Dobel

```text
HADES_5_SHORTCUTS_PHASE5D_CURRENT.bas
HADES_5_SHORTCUTS_QC_MENU_V2_LOCKED.bas
```

Pilih hanya V2 Locked.

```text
VBA_QC_TYPO_CHECK_V12_2_CORE_BRIDGE.bas
VBA_QC_TYPO_CHECK_V13_2_GREEN_MARKER.bas
```

Pilih hanya V13.2 Green Marker.

---

## Setelah Import

Jalankan compile/debug di VBA Editor CorelDRAW sebelum produksi.
