# Agrimatie Visserij Scraper

Haalt automatisch alle grafiekdata op van agrimatie.nl/visserij en exporteert
ze naar één Excel-bestand met een tabblad per pagina.

## Installatie (éénmalig)

### Stap 1: Python installeren
Zorg dat Python 3.8+ geïnstalleerd is: https://www.python.org/downloads/

### Stap 2: Bibliotheken installeren
Open een terminal/opdrachtprompt en voer uit:

```
pip install playwright openpyxl
python -m playwright install chromium
```

## Gebruik

```
python scrape_agrimatie.py
```

Het script opent automatisch een onzichtbare browser, laadt alle vier de pagina's
en exporteert de grafiekdata naar:

    agrimatie_visserij_JJJJMMDD.xlsx

## Uitvoer (Excel-bestand)

Het bestand bevat de volgende tabbladen:

| Tabblad                   | Pagina                                                      |
|---------------------------|-------------------------------------------------------------|
| README                    | Uitleg en overzicht                                         |
| Kottervisserij            | aanvoer-en-besomming-in-de-kottervisserij                  |
| Mosselcultuur             | aanvoer-en-besomming-in-de-mosselcultuur                   |
| Overige kleine zeevisserij| aanvoer-en-besomming-in-de-overige-kleine-zeevisserij      |
| Grote zeevisserij         | sectoren/grote-zeevisserij                                  |

Per figuur staan de data in kolommen: eerste kolom = jaar/categorie,
daarna één kolom per dataserie (bijv. soort vis of vloottype).

## Meer pagina's toevoegen

Open `scrape_agrimatie.py` en voeg een item toe aan de `PAGES`-lijst:

```python
PAGES = [
    ...
    {
        "url": "https://agrimatie.nl/visserij/atotz/mijn-nieuwe-pagina/",
        "tab": "Mijn pagina",
    },
]
```

## Probleemoplossing

**"⚠ Geen grafiekdata gevonden"**
De website gebruikt mogelijk een ander charting-framework dan Highcharts.
Probeer:
1. Script opnieuw draaien (soms time-out bij eerste poging)
2. Headless mode uitzetten: zoek `headless=True` en verander in `headless=False`
   — dan opent een zichtbaar browservenster zodat je kunt zien wat er laadt.
3. Als de website een CSV-downloadknop heeft (onder de grafiek), is dat de
   betrouwbaarste route voor die specifieke figuur.

**Playwright installatiefout**
Zorg dat je als beheerder/administrator het `playwright install`-commando uitvoert.

## Technische werking

1. Playwright start een headless Chromium-browser
2. Per pagina wordt gewacht tot alle netwerkverzoeken voltooid zijn (networkidle)
3. Vervolgens wordt gescrolld zodat lazy-loaded grafieken renderen
4. Het script zoekt naar `window.Highcharts.charts` — het grafiekobject dat
   Highcharts (het charting-framework van agrimatie.nl) in de browser aanmaakt
5. Per grafiek worden titel, categorieën (x-as) en alle series (y-waarden)
   uitgelezen via JavaScript
6. De data wordt weggeschreven naar Excel met openpyxl

## Vereisten

- Python 3.8+
- playwright
- openpyxl
- Internetverbinding
