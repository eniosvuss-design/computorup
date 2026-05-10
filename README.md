# ComputoRUP v0.2

Un solo file. Nessun modulo da importare. Nessuna cartella complicata.

---

## Cosa ti serve (una volta sola)

Apri il **Terminale** e lancia questi comandi **uno alla volta**:

```bash
brew install poppler
```
```bash
pip3 install streamlit pandas openpyxl rapidfuzz opencv-python-headless
```

---

## Avvio

Ogni volta che vuoi usarlo:

```bash
cd ~/Desktop/ComputoRUP2
streamlit run computorup.py
```

Il browser si apre da solo su `http://localhost:8501`

---

## Se non hai Homebrew

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Poi `brew install poppler` e poi le librerie Python.

---

## Cosa fa

- **Prezziario**: carica il PDF del Prezzario FVG 2025 → estrae ~6.500 voci in 6 secondi
- **Computo**: crea il computo, aggiungi voci dal prezziario o manualmente
- **Video sopralluogo**: carica un video del cantiere, annota gli interventi frame per frame
- **Preventivo**: carica il PDF del preventivo ditta
- **Confronto**: abbina le voci del preventivo al prezziario, calcola scostamenti
- **Esporta**: genera Excel professionale con 9 fogli

## File generati

- `computorup.db` – database locale (backup = copia questo file)
- `exports/` – file Excel generati
