# Leietakere — kort over kommunale boliger

En enkel nettapp (bare HTML/CSS/JS, ingen installasjon og ingen eksterne biblioteker) som viser
leietakerne og boligene som kort. **Åpne `index.html` i nettleseren** — dataene er bygget inn, så det
er ikke nødvendig å velge noen Excel-fil.

## Funksjoner

- **Kort** — hver bolig er ett kort. Som standard viser kortet **adressen**.
  Under «Vis på kortene» kan du i tillegg velge initialer, fullt navn og Kundenr./Abb.nr.
- **Tre datoer på hvert kort**: **Inn** (innflytting), **Ut** (kontraktslutt) og **Reg.** (neste KPI-regulering).
- **Fargekoder for kontraktslutt** (hele kortet): 🟥 under 1 måned, 🟧 1–2 måneder, 🟨 2–3 måneder,
  skravert med mørkerød kant = kontrakten har utløpt.
  KPI-reguleringen har de samme fristene som en merkelapp ved datoen («forfalt» når datoen er passert).
- **Oversikt over frister** øverst — klikk på en kategori for å vise bare de kortene.
- **Kjønn** (♀/♂) — fra kolonnen «Kjønn», ellers antatt ut fra fornavnet. Kan endres med blyanten.
- **Søk** i alle felt samtidig (store/små bokstaver, æ/ø/å og mellomrom i numre spiller ingen rolle).

### Redigering ✎

- Hvert kort har en **blyant**. Der kan alt endres: navn, initialer (fylles ut automatisk), kjønn,
  numre, telefon, adresse, bolig, datoer, beløp og merknad.
- **Marker som ledig** — ett klikk tømmer leietakerens data (navn, numre, datoer, telefon, merknader).
  Adressen, boligen og ARTnr beholdes, og leien flyttes til «Beløp 1» slik som for de andre ledige boligene.
- **Ny bolig** — knappen over kortene. Med «Antall boliger» kan flere boliger på samme adresse opprettes
  på én gang (f.eks. 4 nye leiligheter i et nytt hus).
- **Slett** — fjerner et kort.
- Alle endringer kan **angres** fra meldingen nederst på skjermen.

### Etter gate

Velg **Etter gate** for å se ett kort per gate (f.eks. «Eldridhågen» med alle 19 boligene), med antall
leietakere, ledige boliger, husnumre og frister. Klikk på gaten for å se bare de boligene.
«Ny bolig i denne gaten» fyller ut gatenavnet automatisk.

## Hvor lagres endringene?

- Endringer lagres **automatisk i nettleseren** på denne datamaskinen (localStorage) og er der neste gang
  `index.html` åpnes — i den samme nettleseren.
- **Eksporter til Excel** laster ned alle data som en `.xlsx`-fil (samme kolonner som originalen, pluss
  «Kjønn» og «Merknad»). En oransje prikk på knappen betyr at det finnes endringer som ikke er eksportert.
  Bruk eksporten som sikkerhetskopi og når dataene skal deles med andre.
- **⋯ → Last inn Excel-fil** erstatter dataene i appen med innholdet i en Excel-fil (f.eks. en eksport).
- **⋯ → Tilbakestill til innebygde data** forkaster endringene i nettleseren.

Hvis nettleserdata slettes (eller en annen nettleser/datamaskin brukes), vises de innebygde dataene igjen —
eksporter derfor jevnlig.

## Filer

| Fil | Innhold |
| --- | --- |
| `index.html` | siden |
| `style.css` | utseende (lyst og mørkt tema, mobilvisning) |
| `app.js` | kort, redigering, frister, søk, gruppering etter gate, lagring |
| `data.js` | innebygde data, generert fra `Data.xlsx` |
| `navn-kjonn.js` | ordliste over fornavn for å anta kjønn |
| `xlsx-reader.js` / `xlsx-writer.js` | lesing og skriving av `.xlsx`-filer uten eksterne biblioteker |

## Personopplysninger

`Data.xlsx` og `data.js` inneholder personopplysninger (navn, adresser, telefonnumre). Dette repositoriet er
**offentlig** — vurder å gjøre det privat, eller å fjerne filene fra repositoriet og git-historikken.
