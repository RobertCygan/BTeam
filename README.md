# Bolig Team Steigen

Nettapp for kommunale boliger i Steigen: hver bolig er et kort med leietaker, datoer og frister.
**Alt ligger i én fil: `index.html`.** Åpne den i nettleseren og logg inn — ingen installasjon,
ingen server og ingen Excel-fil å velge.

## Innlogging og sikkerhet

- Appen åpner med en innloggingsside (bruker og passord oppgis separat — de står ikke i denne filen).
- Dataene i `index.html` er **kryptert** (AES-256-GCM, nøkkel fra PBKDF2-SHA256 med 600 000 iterasjoner).
  Uten riktig bruker og passord kan dataene ikke leses — heller ikke i kildekoden.
- Innloggingen huskes i fanen til den lukkes. **⋯ → Logg ut** logger ut med en gang.
- Endringer som lagres i nettleseren, er også kryptert.

## Funksjoner

- **Kort** med adresse (eller initialer, navn, Kundenr./Abb.nr. — velg under «Vis på kortene»).
- **Tre datoer på hvert kort**: Inn (innflytting), Ut (kontraktslutt) og Reg. (neste KPI-regulering).
- **Fargekoder for kontraktslutt** (hele kortet): 🟥 under 1 måned, 🟧 1–2 måneder, 🟨 2–3 måneder,
  skravert med mørkerød kant = utløpt. KPI-reguleringen har de samme fristene som merkelapp ved datoen.
- **Frister** øverst, med antall i parentes, f.eks. «under 1 md. (ant. 1)». Klikk for å filtrere.
- **Alle boliger / Etter gate** — i gatevisningen er det ett kort per gate; klikk for å se boligene der.
- **Redigering** med blyanten på hvert kort: alle felt kan endres. **Marker som ledig** tømmer
  leietakerens data med ett klikk. **Ny bolig** oppretter én eller flere boliger på samme adresse.
  Alle endringer kan angres.
- **Kjønn** (♀/♂) — fra kolonnen «Kjønn», ellers antatt ut fra fornavnet; kan endres med blyanten.
- **Søk** i alle felt samtidig.
- **Lyst / mørkt tema** — knappen ☀/☾ øverst til høyre (også på innloggingssiden). Uten valg følger
  appen innstillingen på datamaskinen.

## Sende appen til andre

`index.html` er helt selvstendig: den trenger ingen andre filer og ingen internettforbindelse.
Filen kan kopieres eller sendes og åpnes i Chrome, Edge, Firefox eller Safari på en annen datamaskin.

- Lagre filen og åpne den i nettleseren. Forhåndsvisning i e-post, Teams eller OneDrive kjører ikke appen.
- Noen e-postsystemer blokkerer `.html`-vedlegg — pakk filen i en `.zip` ved behov.
- Den andre datamaskinen starter med dataene som er bygget inn i filen. Endringer gjort i nettleseren
  her følger ikke med — flytt dem med «Eksporter til Excel» og «⋯ → Last inn Excel-fil».

## Hvor lagres endringene?

- Endringer lagres automatisk (kryptert) **i nettleseren** på denne datamaskinen.
- **Eksporter til Excel** laster ned alle data som `.xlsx` (samme kolonner som originalen + «Kjønn» og
  «Merknad»). Bruk den som sikkerhetskopi og når dataene skal deles. En oransje prikk på knappen betyr at
  det finnes endringer som ikke er eksportert.
- **⋯ → Last inn Excel-fil** erstatter dataene i appen med en Excel-fil.
- **⋯ → Tilbakestill til innebygde data** forkaster endringene i nettleseren.

## For utviklere: bygge `index.html`

Kildekoden ligger i `src/`. `tools/build.py` setter alt sammen til én fil, krypterer dataene fra
Excel-filen og skriver `index.html`:

```bash
pip install openpyxl cryptography
python3 tools/build.py --data Data.xlsx          # spør etter passordet
```

| Fil | Innhold |
| --- | --- |
| `src/template.html` | sidens oppbygning (innlogging, verktøylinje, panel, redigering) |
| `src/style.css` | utseende (lyst og mørkt tema, mobil) |
| `src/app.js` | kort, frister, søk, redigering, gater, lagring, eksport |
| `src/login.js` | innlogging og kryptering |
| `src/navn-kjonn.js` | fornavn → antatt kjønn |
| `src/xlsx-reader.js`, `src/xlsx-writer.js` | lese og skrive `.xlsx` uten eksterne biblioteker |
| `tools/build.py` | bygger `index.html` |

Salt og iterasjoner gjenbrukes fra forrige `index.html`, slik at endringer som er lagret i nettleseren,
fortsatt kan åpnes etter en ny bygging (så lenge bruker og passord er de samme).

## Personopplysninger

`index.html` inneholder bare krypterte data. `Data.xlsx` inneholder derimot personopplysninger i klartekst,
og dette repositoriet er **offentlig** — fjern filen fra repositoriet og git-historikken, eller gjør
repositoriet privat.
