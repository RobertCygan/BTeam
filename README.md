# Leietakere — kort fra Excel-filen

En enkel nettapp (bare HTML/CSS/JS, ingen installasjon og ingen eksterne biblioteker) som viser
leietakerne fra `Data.xlsx` som kort.

## Funksjoner

- **Kort** — hver leietaker i arket blir ett kort. Som standard viser kortet bare **adressen** (kolonne K).
- **Vis på kortene:** — velg hva som skal stå på kortene (flere felt samtidig er mulig):
  initialer (kolonne H), fullt navn (G), adresse (K), Kundenr. og Abb.nr. (C og B).
- **Tre datoer på hvert kort, alltid** — uansett hva som er valgt:
  - **Inn** – dato innflytting (L),
  - **Ut** – dato utflytting / kontraktslutt (M),
  - **Reg.** – neste regulering etter KPI (N).
- **Fargekoder for kontraktslutt** — hele kortet farges når kontrakten snart utløper:

  | Farge | Betydning |
  | --- | --- |
  | 🟥 knallrød | under 1 måned igjen |
  | 🟧 oransje | 1–2 måneder igjen |
  | 🟨 gul | 2–3 måneder igjen |
  | skravert med mørkerød kant | kontrakten har allerede utløpt |

- **KPI-regulering** — samme frister vises som en farget merkelapp ved «Reg.»-datoen
  (rød / oransje / gul, og «forfalt» når datoen er passert uten å ha blitt oppdatert).
- **Oversikt over frister** øverst: antall kort i hver kategori for kontraktslutt og KPI-regulering.
  Klikk på en kategori for å vise bare de kortene.
- **Kjønn** — hvert kort har symbolet ♀ eller ♂. Kjønnet hentes fra kolonnen «Kjønn» (M/K) hvis den finnes i
  Excel-filen, ellers antas det ut fra fornavnet (ordlisten ligger i `navn-kjonn.js`).
  Står det **?** på kortet, er kjønnet ukjent — legg da til kolonnen «Kjønn» i Excel.
  Tjenester (f.eks. «Helse og omsorg», «Nav») får et eget symbol.
- **Utvidet kort** — klikk på et kort for å se alle data fra raden: varsler om frister, numre med
  kopieringsknapp, beløp, telefon som lenke og kommentarer fra cellene i Excel. `Esc` eller `×` lukker kortet.
- **Søk** i alle kolonner samtidig: navn, initialer, adresse, Kundenr., Abb.nr., telefon, kommentarer …
  Store/små bokstaver, æ/ø/å og mellomrom i numre spiller ingen rolle (`molnmoa` finner «Mølnmoa»,
  `45254373` finner «452 54 373»). Snarvei: `/`.
- **Sortering**: som i filen, A–Å, etter kontraktslutt eller etter KPI-regulering (snarest først).
- **Ledige boliger** (rader uten leietaker) er skjult som standard og kan slås på.

Datoene sammenlignes med dagens dato i nettleseren, så fargene oppdaterer seg selv fra dag til dag.

## Slik starter du

Dataene leses rett fra Excel-filen, så når `Data.xlsx` byttes ut, viser appen de nye dataene med en gang.

**Måte 1 — åpne `index.html` i nettleseren.**
Nettleseren lar ikke en side som er åpnet fra disken lese andre filer på egen hånd, så appen ber deg velge
Excel-filen (knapp, eller dra filen inn på siden). Filen leses bare lokalt i nettleseren — ingenting sendes noe sted.

**Måte 2 — lokal server** (da lastes `Data.xlsx` fra mappen automatisk):

```bash
python -m http.server 8000
```

og åpne <http://localhost:8000>.

Du kan når som helst laste inn en annen fil med knappen **Last inn Excel-fil**.

## Filer

| Fil | Innhold |
| --- | --- |
| `index.html` | siden |
| `style.css` | utseende (lyst og mørkt tema, mobilvisning) |
| `app.js` | kort, valg av felt, frister, søk, utvidet kort |
| `navn-kjonn.js` | ordliste over fornavn for å anta kjønn |
| `xlsx-reader.js` | enkel leser for `.xlsx`-filer, skrevet uten eksterne biblioteker |

Kolonnene finnes ut fra overskriftene i raden med «Navn» og «Adresse»; hvis en overskrift endres, brukes
kolonnebokstaven (A–P). Nye rader og små endringer i arket krever derfor ingen endringer i koden.

## Personopplysninger

`Data.xlsx` inneholder personopplysninger (navn, adresser, telefonnumre). Dette repositoriet er
**offentlig** — vurder å fjerne filen fra repositoriet (og fra git-historikken) eller å gjøre repositoriet
privat. Appen fungerer også uten filen i repositoriet: last den inn fra disken.
