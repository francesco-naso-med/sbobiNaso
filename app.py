"""
sbobiNaso — interfaccia locale per trasformare una sbobina in dispensa.

    streamlit run app.py

Tutto il lavoro vero sta in dispensa.py: qui c'è solo l'interfaccia. La chiave API
resta sul computer di chi la usa, non viene salvata su disco né inviata altrove
che a Google.
"""

from __future__ import annotations

import io
import re
import tempfile
import time
from pathlib import Path

import streamlit as st

import dispensa as D
import sorgente

st.set_page_config(page_title="sbobiNaso", page_icon="👃", layout="centered")

TEMPLATE_DI_SERIE = Path(__file__).parent / "template.docx"

ANIMAZIONE = """
<div class="sn-box">
  <div class="sn-naso">👃</div>
  <div class="sn-onde"><span></span><span></span><span></span></div>
</div>
<style>
.sn-box { display:flex; align-items:center; gap:.9rem; padding:.6rem 0 .2rem; }
.sn-naso { font-size:2.1rem; line-height:1; animation:sn-ondeggia 1.6s ease-in-out infinite; }
@keyframes sn-ondeggia {
  0%,100% { transform:translateX(0) rotate(0deg); }
  30%     { transform:translateX(6px) rotate(7deg); }
  60%     { transform:translateX(-3px) rotate(-4deg); }
}
.sn-onde { display:flex; gap:.35rem; }
.sn-onde span {
  width:.5rem; height:.5rem; border-radius:50%; background:currentColor; opacity:.25;
  animation:sn-pulsa 1.2s ease-in-out infinite;
}
.sn-onde span:nth-child(2) { animation-delay:.2s; }
.sn-onde span:nth-child(3) { animation-delay:.4s; }
@keyframes sn-pulsa { 0%,100% { opacity:.2; transform:scale(.8); } 50% { opacity:.9; transform:scale(1.15); } }
</style>
"""

# Ruotano durante l'elaborazione: ironiche ma non invadenti.
MESSAGGI = [
    "Leggo tutto 🤓",
    "Certo che 'sta sbobina l'hanno scritta coi piedi",
    "Tolgo gli «ehm» e le gasteme",
    "Complimenti agli sbobinatori",
    "Mo te l'aggiusto io",
    "Rileggo un po', famm sta' tranquill",
    "Metto i puntini sulle i",
    "Ma davvero è in italiano?",
]

MODELLI = [
    ("gemini-3.5-flash-lite", "Flash Lite — 500 dispense/giorno gratis, consigliato"),
    ("gemini-3.1-flash-lite", "Flash Lite (generazione precedente) — 500/giorno"),
    ("gemini-3.5-flash", "Flash pieno — più accurato, solo 20 richieste/giorno"),
    ("gemini-3.6-flash", "Flash pieno (più recente) — 20 richieste/giorno"),
]


# ---------------------------------------------------------------------------
# Elaborazione
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner="Apro la sbobina…")
def leggi_sorgente(nome: str, dati: bytes) -> tuple[str, dict[int, str], str]:
    """Estrae testo e immagini dal file caricato. In cache: non si rifà a ogni clic."""
    cartella = Path(tempfile.mkdtemp())
    percorso = cartella / nome
    percorso.write_bytes(dati)
    testo, immagini = sorgente.estrai(percorso, cartella / sorgente.CARTELLA_IMMAGINI)
    return testo, {n: str(f) for n, f in immagini.items()}, str(cartella / sorgente.CARTELLA_IMMAGINI)


def costruisci_docx(markdown: str, template: str | None, opzioni: dict, worker) -> tuple[bytes, int]:
    """Markdown → Word in memoria. Restituisce i byte e il numero di capitoli."""
    capitoli: dict[int, str] = {}
    if opzioni["capitoli"]:
        titoli = re.findall(r"^##\s+(.+)$", markdown, re.M)
        if titoli:
            capitoli = worker.raggruppa_capitoli(titoli)

    builder = D.DocxBuilder(template, Path(opzioni["cartella_immagini"]))
    if opzioni["indice"]:
        builder.add_toc()
    builder.add_markdown(markdown, chapters=capitoli)
    builder.add_footer(opzioni["firma"], numeri=opzioni["numeri"])

    buffer = io.BytesIO()
    builder.doc.save(buffer)
    return buffer.getvalue(), len(capitoli)


def elabora(chunks: list[str], immagini: dict, opzioni: dict, barra, stato) -> str | None:
    """Elabora i blocchi mancanti. Ritorna il messaggio di stop, o None se finisce."""
    worker = D.GeminiWorker(model=opzioni["modello"], fallbacks=list(D.FALLBACK_MODELS),
                            api_key=opzioni["api_key"])
    st.session_state.worker = worker
    fatti: list[str] = st.session_state.blocchi

    for i in range(len(fatti), len(chunks)):
        coda = fatti[-1][-D.CONTEXT_TAIL_CHARS:] if fatti else ""
        stato.markdown(f"**{MESSAGGI[i % len(MESSAGGI)]}**  \n"
                       f"Blocco {i + 1} di {len(chunks)}")
        try:
            fatti.append(worker.elaborate(chunks[i], i + 1, len(chunks),
                                          context_tail=coda, immagini=immagini))
        except D.QuotaExhaustedError as err:
            return str(err)
        except Exception as err:  # noqa: BLE001 - va mostrato all'utente, non nascosto
            return f"Errore sul blocco {i + 1}: {err}"
        barra.progress((i + 1) / len(chunks), text=f"{i + 1} / {len(chunks)} blocchi")
        if opzioni["pausa"] and i + 1 < len(chunks):
            time.sleep(opzioni["pausa"])
    return None


# ---------------------------------------------------------------------------
# Interfaccia
# ---------------------------------------------------------------------------

st.title("sbobiNaso 👃")
st.caption("Tu dammi la sbobina, io ti do la dispensa")

for chiave, default in (("blocchi", []), ("docx", None), ("worker", None), ("impronta", None)):
    st.session_state.setdefault(chiave, default)

with st.sidebar:
    st.subheader("Chiave API")
    api_key = st.text_input(
        "Chiave Gemini", type="password", placeholder="Spara…",
        help="Gratuita su aistudio.google.com/apikey. Resta sul tuo computer.",
    )
    st.caption("[Prendi la tua chiave](https://aistudio.google.com/apikey) — è gratis.")

    st.subheader("Modello")
    modello = st.selectbox("Modello", [m for m, _ in MODELLI],
                           format_func=lambda m: dict(MODELLI)[m], label_visibility="collapsed")

    st.subheader("Documento")
    indice = st.checkbox("Indice iniziale", value=True)
    capitoli = st.checkbox("Divisione in capitoli", value=True,
                           help="Una chiamata API in più: raggruppa le sezioni per argomento.")
    numeri = st.checkbox("Numeri di pagina", value=True)
    st.caption(f"Piè di pagina: _{D.FIRMA}_")

    with st.expander("Impostazioni avanzate"):
        template = st.file_uploader(
            "Template Word personale", type=["docx"],
            help="Da qui vengono font, margini e stili. Senza, si usa il template di sbobiNaso.",
        )
        chunk = st.slider("Caratteri per blocco", 1500, 8000, D.CHUNK_TARGET_CHARS, step=250,
                          help="Blocchi più corti = più fedeltà ai dettagli, più chiamate.")
        pausa = st.slider("Pausa fra le chiamate (secondi)", 0.0, 10.0, 4.0, step=0.5,
                          help="Serve a non superare il limite di richieste al minuto.")

sbobina = st.file_uploader("La sbobina grezza", type=["txt", "md", "docx", "pdf"],
                           help="Da Word e PDF vengono estratte anche le immagini, "
                                "che finiscono nella dispensa al punto giusto. Almeno spero")

if not sbobina:
    st.info("Carica la sbobina per iniziare: testo, Word o PDF, anche disordinato.")
    st.stop()

testo, immagini_str, cartella_immagini = leggi_sorgente(sbobina.name, sbobina.getvalue())
immagini = {int(n): Path(f) for n, f in immagini_str.items()}
chunks = D.chunk_text(testo, target=chunk)

impronta = (sbobina.name, len(testo), chunk)
if st.session_state.impronta != impronta:          # sbobina o taglio cambiati: si riparte
    st.session_state.update(impronta=impronta, blocchi=[], docx=None)

col1, col2, col3, col4 = st.columns(4)
col1.metric("Caratteri", f"{len(testo):,}".replace(",", "."))
col2.metric("Blocchi", len(chunks))
col3.metric("Immagini", len(immagini))
col4.metric("Tempo stimato", f"~{round(len(chunks) * (pausa + 5) / 60)} min")

fatti = len(st.session_state.blocchi)
if fatti:
    st.progress(fatti / len(chunks), text=f"{fatti} / {len(chunks)} blocchi già elaborati")

etichetta = "Crea la dispensa" if not fatti else "Riprendi da dove eravamo"
if st.button(etichetta, type="primary", use_container_width=True, disabled=not api_key):
    if not api_key:
        st.error("Serve la chiave API: la trovi nella barra laterale.")
        st.stop()

    if template:                                       # override dalle avanzate
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp.write(template.getvalue())
        tmp.close()
        percorso_template = tmp.name
    else:
        percorso_template = str(TEMPLATE_DI_SERIE) if TEMPLATE_DI_SERIE.is_file() else None

    opzioni = {"modello": modello, "indice": indice, "capitoli": capitoli,
               "numeri": numeri, "firma": D.FIRMA, "pausa": pausa,
               "cartella_immagini": cartella_immagini, "api_key": api_key}

    st.markdown(ANIMAZIONE, unsafe_allow_html=True)   # scritta una volta: non riparte a ogni blocco
    barra = st.progress(fatti / len(chunks), text=f"{fatti} / {len(chunks)} blocchi")
    stato = st.empty()
    problema = elabora(chunks, immagini, opzioni, barra, stato)
    stato.empty()

    if st.session_state.blocchi:
        markdown = "\n\n".join(st.session_state.blocchi)
        with st.spinner("Impagino il Word…"):
            try:
                documento, n_capitoli = costruisci_docx(
                    markdown, percorso_template, opzioni, st.session_state.worker
                )
                st.session_state.docx = documento
                st.session_state.capitoli = n_capitoli
            except Exception as err:  # noqa: BLE001
                st.error(f"Il testo è salvo, ma l'impaginazione è fallita: {err}")

    if problema:
        st.warning(f"Fermati a {len(st.session_state.blocchi)} blocchi su {len(chunks)}. "
                   f"{problema}\n\nQuello che c'è è scaricabile qui sotto, e il pulsante "
                   "riprende da questo punto senza rifare nulla.")
    else:
        st.success("Dispensa pronta.")

if st.session_state.docx:
    nome = Path(sbobina.name).stem
    st.download_button("⬇️  Scarica la dispensa (.docx)", st.session_state.docx,
                       file_name=f"{nome}_dispensa.docx", use_container_width=True,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.download_button("⬇️  Scarica il testo grezzo (.md)",
                       "\n\n".join(st.session_state.blocchi).encode("utf-8"),
                       file_name=f"{nome}_dispensa.md", use_container_width=True)
    if st.session_state.get("capitoli"):
        st.caption(f"{st.session_state.capitoli} capitoli. "
                   "In Word, se l'indice è vuoto: clic destro sull'indice → «Aggiorna campo».")

st.divider()
st.caption("sbobiNaso")
