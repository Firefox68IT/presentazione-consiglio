# Travel Council Presentation Skill

Skill Codex per generare presentazioni PowerPoint di viaggio a partire da cartelle tappa con foto, video e `Testo.txt`.

## Quick Start

1. Copia la skill in:
   `C:\Users\<utente>\.codex\skills\travel-council-presentation`
2. Riavvia Codex.
3. Apri la cartella del progetto viaggio.
4. Esegui lo script oppure chiedi a Codex di usare la skill.

## Cosa fa

- crea una copertina e una slide introduttiva per ogni tappa
- usa `Testo.txt` come testo della slide sezione quando presente
- ordina foto e video in ordine cronologico
- genera clip video da 15 secondi
- inserisce i video nella sezione corretta della presentazione
- imposta i video con autoplay dopo 3 secondi e playback fullscreen
- genera un file `presentazione_generata.txt` di riepilogo

## Struttura attesa del progetto

```text
project-root/
  MATERIALE...pdf
  20260206 New Delhi Ashalayam tempio/
    Testo.txt
    IMG....jpg
    VID....mp4
  20260207 Another place/
    ...
```

Le cartelle tappa vengono rilevate da nomi che iniziano con una data nel formato `YYYYMMDD`.

## Contenuto del repository

```text
skills/travel-council-presentation/
  SKILL.md
  agents/openai.yaml
  scripts/generate_presentation.ps1

docs/INSTALL_SKILL.md
README.md
```

## Requisiti

- Windows
- Microsoft PowerPoint installato
- `ffmpeg` disponibile nel `PATH`
  oppure in `C:\Users\<utente>\tools\ffmpeg\bin\ffmpeg.exe`

## Uso diretto dello script

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\skills\travel-council-presentation\scripts\generate_presentation.ps1
```

Con percorso personalizzato:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\skills\travel-council-presentation\scripts\generate_presentation.ps1 -ProjectRoot "C:\Percorso\Progetto" -OutputName "Presentazione finale.pptx"
```

## Installazione su altri computer

La guida completa è in:

[docs/INSTALL_SKILL.md](docs/INSTALL_SKILL.md)

## Output generati

Nel `ProjectRoot` vengono creati:

- il file PowerPoint finale
- `presentazione_generata.txt`
- la cartella `clip_video_15s` con le clip da 15 secondi
