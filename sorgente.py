"""
Lettura della sbobina originale: testo e immagini, in ordine di documento.

L'immagine viene sostituita nel testo da un segnaposto `[[IMG:n]]`, e il file
finisce in una cartella accanto all'output. Il segnaposto è testo a tutti gli
effetti: attraversa il chunking, il prompt e il Markdown senza codice speciale,
e viene riconvertito in immagine solo alla fine, dentro il Word.
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

IMG_MARKER = "[[IMG:{n}]]"
IMG_MARKER_RE = re.compile(r"\[\[IMG:(\d+)\]\]")

CARTELLA_IMMAGINI = "immagini"
MIN_LATO_PX = 120          # sotto questa soglia sono icone, loghi, pallini: si scartano
MIN_PESO_BYTE = 6000


def cartella_immagini(output: str | Path) -> Path:
    return Path(output).resolve().parent / CARTELLA_IMMAGINI


def percorso_immagine(cartella: Path, n: int) -> Path | None:
    """Ritrova l'immagine n qualunque sia l'estensione con cui è stata salvata."""
    trovate = sorted(cartella.glob(f"img_{n:03d}.*")) if cartella.is_dir() else []
    return trovate[0] if trovate else None


class _Raccolta:
    """Salva le immagini scartando doppioni (loghi ripetuti) e miniature."""

    def __init__(self, cartella: Path):
        self.cartella = cartella
        self.cartella.mkdir(parents=True, exist_ok=True)
        self.viste: dict[str, int] = {}
        self.immagini: dict[int, Path] = {}
        self.n = 0

    def aggiungi(self, blob: bytes, ext: str, lato_min: int | None = None) -> int | None:
        if len(blob) < MIN_PESO_BYTE:
            return None
        if lato_min is not None and lato_min < MIN_LATO_PX:
            return None

        impronta = hashlib.md5(blob).hexdigest()
        if impronta in self.viste:
            return None                       # gia' incontrata: e' un elemento grafico ripetuto

        self.n += 1
        percorso = self.cartella / f"img_{self.n:03d}.{ext.lstrip('.')}"
        percorso.write_bytes(blob)
        self.viste[impronta] = self.n
        self.immagini[self.n] = percorso
        return self.n


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------


def _da_docx(percorso: Path, raccolta: _Raccolta) -> str:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(percorso))
    parti: list[str] = []

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            para = Paragraph(child, doc)
            pezzi = [para.text.strip()]

            for blip in child.findall(f".//{qn('a:blip')}"):
                rid = blip.get(qn("r:embed"))
                if not rid:
                    continue
                try:
                    parte = doc.part.related_parts[rid]
                except KeyError:
                    continue
                ext = Path(str(parte.partname)).suffix or ".png"
                n = raccolta.aggiungi(parte.blob, ext)
                if n:
                    pezzi.append(IMG_MARKER.format(n=n))

            riga = " ".join(p for p in pezzi if p)
            if riga:
                parti.append(riga)

        elif child.tag == qn("w:tbl"):
            righe = []
            for row in Table(child, doc).rows:
                celle = [c.text.strip().replace("\n", " ") for c in row.cells]
                if any(celle):
                    righe.append(" | ".join(celle))
            if righe:
                parti.append("\n".join(righe))

    return "\n\n".join(parti)


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------


def _da_pdf(percorso: Path, raccolta: _Raccolta) -> str:
    import pymupdf

    parti: list[str] = []
    with pymupdf.open(str(percorso)) as doc:
        for pagina in doc:
            blocchi = pagina.get_text("dict").get("blocks", [])
            # in ordine di lettura: dall'alto al basso, poi da sinistra a destra
            blocchi.sort(key=lambda b: (round(b["bbox"][1], 1), round(b["bbox"][0], 1)))

            for blocco in blocchi:
                if blocco.get("type") == 0:                       # testo
                    testo = "".join(
                        span.get("text", "")
                        for riga in blocco.get("lines", [])
                        for span in riga.get("spans", [])
                    ).strip()
                    if testo:
                        parti.append(testo)
                elif blocco.get("image"):                          # immagine
                    lato = min(blocco.get("width", 0), blocco.get("height", 0))
                    n = raccolta.aggiungi(blocco["image"], blocco.get("ext", "png"), lato)
                    if n:
                        parti.append(IMG_MARKER.format(n=n))

    return "\n\n".join(parti)


# ---------------------------------------------------------------------------


def estrai(percorso: str | Path, cartella: Path) -> tuple[str, dict[int, Path]]:
    """Restituisce (testo con segnaposto, {numero: file immagine})."""
    percorso = Path(percorso)
    suffisso = percorso.suffix.lower()

    if suffisso not in (".docx", ".pdf"):                 # testo semplice: niente immagini
        return percorso.read_text(encoding="utf-8", errors="replace").strip(), {}

    raccolta = _Raccolta(cartella)
    testo = _da_docx(percorso, raccolta) if suffisso == ".docx" else _da_pdf(percorso, raccolta)
    return testo.strip(), raccolta.immagini
