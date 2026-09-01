# sbobiNaso 👃

La sbobina entra grezza, esce dispensa. **Senza riassumere niente.**

Prende la trascrizione di una lezione — anche disordinata, piena di "ehm" e di frasi
a metà — e la trasforma in una dispensa in Word: italiano scritto, titoli, elenchi,
termini chiave in grassetto, indice, capitoli e numeri di pagina. Tutti i dettagli
della lezione restano dov'erano: dosaggi, numeri, nomi, esempi, eccezioni.

Legge `.txt`, `.md`, `.docx` e `.pdf`. Da Word e PDF estrae anche **le immagini**, le
fa guardare al modello e le rimette nella dispensa al punto giusto.

## Cosa ti serve

1. **Python 3.9 o superiore.** Su Mac di solito c'è già: apri il Terminale e scrivi
   `python3 --version`. Se manca, si scarica da [python.org](https://www.python.org/downloads/).
2. **Una chiave API di Google AI Studio**, gratuita: la generi in trenta secondi su
   [aistudio.google.com/apikey](https://aistudio.google.com/apikey). Inizia con `AIza`.

## Installazione (una volta sola)

Apri il Terminale nella cartella di sbobiNaso e lancia:

```bash
pip3 install -r requirements.txt
```

## Come si usa

```bash
streamlit run app.py
```

Si apre una pagina nel browser. Incolla la chiave nella barra laterale, carica la
sbobina, premi **Crea la dispensa** e aspetta. Alla fine scarichi il Word.

Qualche indicazione:

- Il modello preimpostato (**Flash Lite**) è gratuito fino a 500 elaborazioni al
  giorno. I modelli "Flash pieni" sono più accurati ma si fermano a 20 al giorno.
- Se la quota finisce a metà lavoro, non perdi niente: il pulsante diventa
  *Riprendi da dove eravamo* e continua da lì.
- La chiave resta sul tuo computer. Non viene salvata su disco e non va da nessuna
  parte se non a Google, per elaborare la tua sbobina.
- L'indice, in Word, si compila all'apertura. Se resta vuoto: clic destro
  sull'indice → «Aggiorna campo».
- Il documento eredita font, margini e stili da `template.docx`. Se ne vuoi uno tuo,
  caricalo in *Impostazioni avanzate*.

## Da terminale, senza interfaccia

```bash
export GEMINI_API_KEY="la-tua-chiave"
python3 dispensa.py -i sbobina.docx
```

`--limit 3` per provare su pochi blocchi, `--resume` per riprendere, `--only-docx`
per rigenerare il Word senza rielaborare, `--help` per tutto il resto.

## Licenza

[MIT](LICENSE). Usalo, modificalo, distribuiscilo, anche in progetti tuoi: basta
tenere la nota di copyright.

---

Le dispense che produce portano la firma di sbobiNaso in fondo
a ogni pagina: è il modo in cui lo strumento resta gratis.
