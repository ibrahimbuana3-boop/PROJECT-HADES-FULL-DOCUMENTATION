"""
PROJECT HADES — ORDER EXTRACTOR V4.5 METADATA REFERENCE

Purpose:
- Membaca PDF PO produksi.
- Menghasilkan Documents\Order.txt.
- Menulis metadata Hades:
  @JENIS_PESANAN
  @JENIS_POLA
  @MODEL_JAHIT
  @SIZEDB
  @IDPO
- Menulis row order:
  SIZE|NAMA|NOMOR|NICKNAME

Catatan:
- Ini reference implementation repository.
- Sesuaikan regex/kolom dengan format PDF PO aktif di pabrik.
"""

from __future__ import annotations

import os
import re
import time
from pathlib import Path

try:
    import pdfplumber
except ImportError as exc:
    raise SystemExit("Install dependency dulu: pip install pdfplumber") from exc

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    Observer = None
    FileSystemEventHandler = object


DOCUMENTS = Path.home() / "Documents"
DOWNLOADS = Path.home() / "Downloads"
ORDER_PATH = DOCUMENTS / "Order.txt"

WATCH_DOWNLOADS = True
DELETE_PDF_AFTER_SUCCESS = False
UPPERCASE_OUTPUT = False

SIZE_WHITELIST = {
    "XXS", "XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL",
}


def clean_text(text: str) -> str:
    text = text or ""
    text = text.replace("\ufeff", "").replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalize_size(raw: str) -> str:
    raw = clean_text(raw).upper()
    raw = re.sub(r"\s+\d+\s*-\s*\d+\s*(TH|TAHUN)?", "", raw).strip()

    alias = {
        "XXL": "2XL",
        "XXXL": "3XL",
        "XXXXL": "4XL",
        "XXXXXL": "5XL",
        "XXXXXXL": "6XL",
    }

    raw = alias.get(raw, raw)

    return raw if raw in SIZE_WHITELIST else ""


def extract_idpo(text: str) -> str:
    # Biasanya berada pada bagian Kode Produk / Jumlah Kode Produk.
    m = re.search(r"(?<!\d)(\d{6})(?!\d)\s*\(Jumlah\s+Kode\s+Produk", text, re.I)

    if m:
        return m.group(1)

    # Fallback: ambil 6 digit pertama setelah label Kode Produk.
    m = re.search(r"Kode\s+Produk\s+.*?(?<!\d)(\d{6})(?!\d)", text, re.I | re.S)

    if m:
        return m.group(1)

    return ""


def infer_sizedb(meta: dict[str, str]) -> str:
    jenis = meta.get("JENIS_PESANAN", "").upper()
    pola = meta.get("JENIS_POLA", "").upper()
    model = meta.get("MODEL_JAHIT", "").upper()

    joined = f"{jenis} {pola} {model}"

    if "CELANA" in joined:
        if "ANAK" in joined:
            return "SizeDB_CelanaAnak.txt"
        if "WANITA" in joined:
            return "SizeDB_CelanaWanita.txt"
        return "SizeDB_CelanaPria.txt"

    if "JAKET" in joined:
        if "ANAK" in joined:
            return "SizeDB_JaketAnak.txt"
        return "SizeDB_Jaket.txt"

    if "SLIM" in joined:
        if "WANITA" in joined:
            return "SizeDB_WanitaSlimFit.txt"
        return "SizeDB_PriaSlimFit.txt"

    if "ANAK" in joined:
        return "SizeDB_Anak.txt"

    if "WANITA" in joined:
        return "SizeDB_Wanita.txt"

    return "SizeDB_Pria.txt"


def extract_metadata(text: str) -> dict[str, str]:
    meta: dict[str, str] = {}

    patterns = {
        "JENIS_PESANAN": r"Jenis\s+Pesanan\s+(.+)",
        "JENIS_POLA": r"Jenis\s+Pola\s+(.+)",
        "MODEL_JAHIT": r"Model\s+Jahit\s+(.+)",
    }

    for key, pattern in patterns.items():
        m = re.search(pattern, text, re.I)

        if m:
            line = clean_text(m.group(1))
            line = re.split(r"\s{2,}|Kualitas\s+Produk|Bentuk\s+Lengan|Warna", line)[0].strip()
            meta[key] = line

    meta["IDPO"] = extract_idpo(text)
    meta["SIZEDB"] = infer_sizedb(meta)

    return meta


def parse_order_rows(text: str) -> list[tuple[str, str, str, str]]:
    """
    Parser sederhana untuk blok:
    SIZE NO NAMA NICKNAME BONUS
    M Meilan N
    XL Miftah N

    Untuk format PDF lain, perkuat parser sesuai struktur tabel asli.
    """
    rows: list[tuple[str, str, str, str]] = []

    lines = [clean_text(x) for x in text.splitlines()]
    start = -1

    for i, line in enumerate(lines):
        if re.search(r"\bSIZE\b", line, re.I) and re.search(r"\bNAMA\b", line, re.I):
            start = i + 1
            break

    if start < 0:
        return rows

    for line in lines[start:]:
        if not line:
            continue

        parts = line.split()
        if not parts:
            continue

        size = normalize_size(parts[0])
        if not size:
            continue

        tail = parts[1:]

        # Jika kolom NO berisi angka jersey, simpan sebagai nomor.
        nomor = ""
        name_parts = []
        nickname = ""

        if tail and re.fullmatch(r"\d{1,3}", tail[0]):
            nomor = tail[0]
            tail = tail[1:]

        # Bonus sering N di ujung; abaikan.
        if tail and tail[-1].upper() in {"N", "Y"}:
            tail = tail[:-1]

        # Nickname bisa ditingkatkan bila PDF punya kolom eksplisit.
        name_parts = tail

        nama = " ".join(name_parts).strip()

        if UPPERCASE_OUTPUT:
            nama = nama.upper()
            nickname = nickname.upper()

        rows.append((size, nama, nomor, nickname))

    return rows


def pdf_to_text(path: Path) -> str:
    chunks = []

    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")

    return "\n".join(chunks)


def write_order_txt(meta: dict[str, str], rows: list[tuple[str, str, str, str]]) -> None:
    lines = []

    for key in ["JENIS_PESANAN", "JENIS_POLA", "MODEL_JAHIT", "SIZEDB", "IDPO"]:
        val = meta.get(key, "")

        if val:
            lines.append(f"@{key}={val}")

    lines.append("")

    for size, nama, nomor, nickname in rows:
        lines.append(f"{size}|{nama}|{nomor}|{nickname}")

    ORDER_PATH.write_text("\r\n".join(lines).strip() + "\r\n", encoding="utf-8")


def process_pdf(path: Path) -> bool:
    text = pdf_to_text(path)
    meta = extract_metadata(text)
    rows = parse_order_rows(text)

    if not rows:
        print(f"[SKIP] Tidak ada row order valid: {path.name}")
        return False

    write_order_txt(meta, rows)
    print(f"[OK] Order.txt dibuat: {ORDER_PATH}")

    return True


if Observer is not None:

    class PdfHandler(FileSystemEventHandler):
        def on_created(self, event):
            if event.is_directory:
                return

            path = Path(event.src_path)

            if path.suffix.lower() != ".pdf":
                return

            time.sleep(1.5)

            try:
                ok = process_pdf(path)

                if ok and DELETE_PDF_AFTER_SUCCESS:
                    path.unlink(missing_ok=True)

            except Exception as exc:
                print(f"[ERROR] {path.name}: {exc}")


def main():
    print("PROJECT HADES — ORDER EXTRACTOR V4.5 METADATA REFERENCE")
    print(f"Output: {ORDER_PATH}")

    if not WATCH_DOWNLOADS or Observer is None:
        pdf = input("Path PDF: ").strip().strip('"')

        if pdf:
            process_pdf(Path(pdf))

        return

    observer = Observer()
    observer.schedule(PdfHandler(), str(DOWNLOADS), recursive=False)
    observer.start()

    print(f"Watching: {DOWNLOADS}")
    print("Tekan Ctrl+C untuk keluar.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()


if __name__ == "__main__":
    main()
