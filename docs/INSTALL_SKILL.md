# Installare la skill su altri computer

## 1. Copiare la cartella della skill

Copia questa cartella del repository:

`skills\travel-council-presentation`

nella cartella skill di Codex del computer di destinazione.

Percorso tipico su Windows:

`C:\Users\<nome-utente>\.codex\skills\travel-council-presentation`

## 2. Verificare i prerequisiti

Sul computer di destinazione servono:

- Windows
- Microsoft PowerPoint installato
- `ffmpeg` disponibile nel `PATH`
  oppure in:
  `C:\Users\<nome-utente>\tools\ffmpeg\bin\ffmpeg.exe`

## 3. Riavviare Codex

Dopo aver copiato la cartella, riavvia Codex oppure apri una nuova sessione, così la skill viene rilevata.

## 4. Come usarla

Apri la cartella del progetto viaggio e chiedi qualcosa come:

- `Genera la presentazione del viaggio usando la skill travel-council-presentation`
- `Rigenera il PowerPoint dalle cartelle tappa con Testo.txt, foto e video`

Oppure esegui direttamente lo script:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\skills\travel-council-presentation\scripts\generate_presentation.ps1
```

Con parametri personalizzati:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\skills\travel-council-presentation\scripts\generate_presentation.ps1 -ProjectRoot "C:\Percorso\Progetto" -OutputName "Presentazione finale.pptx"
```

## 5. Output generati

La skill crea nel `ProjectRoot`:

- il file PowerPoint finale
- `presentazione_generata.txt`
- la cartella `clip_video_15s` con le clip da 15 secondi

## 6. Versionare nel repository

Per portarla facilmente su altri computer via git, committa almeno questi file:

- `skills/travel-council-presentation/SKILL.md`
- `skills/travel-council-presentation/agents/openai.yaml`
- `skills/travel-council-presentation/scripts/generate_presentation.ps1`
- `docs/INSTALL_SKILL.md`
