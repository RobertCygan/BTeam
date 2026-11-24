# Klienci — kafelki z pliku Excel

Prosta aplikacja webowa (sam HTML/CSS/JS, bez instalacji i bez zewnętrznych bibliotek),
która pokazuje klientów z pliku `Data.xlsx` jako kafelki.

## Co potrafi

- **Kafelki** — każdy klient z arkusza to jeden kafelek. Domyślnie na kafelku widać tylko **adres** (kolumna K).
- **Na kafelkach:** — lista, z której wybierasz, co ma być na kafelkach (można zaznaczyć kilka pól naraz):
  - Inicjały (kolumna H),
  - Imię i nazwisko (kolumna G),
  - Adres (kolumna K),
  - Kundenr. i Abb.nr. (kolumny C i B).
- **Rozwijanie** — kliknięcie kafelka płynnie rozwija go na całą szerokość i pokazuje wszystkie dane z wiersza
  (numery z przyciskiem „kopiuj”, daty z informacją „za 2 miesiące / rok temu”, kwoty, telefon jako link,
  komentarze dodane do komórek w Excelu). `Esc` albo `×` zwija kafelek.
- **Wyszukiwanie** po wszystkich kolumnach naraz: nazwisko, inicjały, adres, Kundenr., Abb.nr., telefon,
  komentarze… Bez znaczenia są wielkie litery, polskie/norweskie znaki (`molnmoa` znajdzie „Mølnmoa”)
  i spacje w numerach (`45254373` znajdzie „452 54 373”). Jeśli trafienie jest w polu, którego nie widać na
  kafelku, kafelek pokazuje to pole z podświetleniem. Skrót: `/` przechodzi do wyszukiwarki.
- **Kolejność**: jak w pliku albo alfabetycznie według pola na kafelku.
- **Wolne lokale** (wiersze bez klienta) — ukryte domyślnie, można je włączyć przełącznikiem.

## Jak uruchomić

Dane są czytane bezpośrednio z pliku Excel, więc po podmianie `Data.xlsx` aplikacja od razu pokazuje nowe dane.

**Sposób 1 — otwórz `index.html` w przeglądarce.**
Przeglądarka nie pozwala stronie otwartej z dysku samodzielnie czytać innych plików, więc aplikacja poprosi
o wybranie pliku Excel (przycisk albo przeciągnięcie pliku na stronę). Plik jest odczytywany wyłącznie
lokalnie, w tej przeglądarce — nic nie jest nigdzie wysyłane.

**Sposób 2 — lokalny serwer** (wtedy `Data.xlsx` z tego folderu wczyta się automatycznie):

```bash
python -m http.server 8000
```

i otwórz <http://localhost:8000>.

W każdej chwili można wczytać inny plik przyciskiem **Wczytaj plik Excel**.

## Pliki

| Plik | Co zawiera |
| --- | --- |
| `index.html` | strona |
| `style.css` | wygląd (jasny i ciemny motyw, wersja na telefon) |
| `app.js` | kafelki, wybór pól, wyszukiwanie, rozwijanie |
| `xlsx-reader.js` | minimalny czytnik plików `.xlsx` napisany bez zewnętrznych bibliotek |

Kolumny są rozpoznawane po nagłówkach w wierszu z „Navn” i „Adresse”; jeśli nagłówek zostanie zmieniony,
aplikacja używa litery kolumny (A–P), więc nowe wiersze i drobne zmiany w arkuszu nie wymagają zmian w kodzie.

## Uwaga o danych osobowych

`Data.xlsx` zawiera dane osobowe (imiona i nazwiska, adresy, numery telefonów). To repozytorium jest
**publiczne** — rozważ usunięcie pliku z repozytorium (i z historii git) albo zmianę repozytorium na prywatne.
Aplikacja działa także bez pliku w repozytorium: wystarczy wczytać go z dysku.
