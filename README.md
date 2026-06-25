# PlantWijzer – tweede interface

Deze versie gebruikt dezelfde `data/typologies/` en `data/layers/` mappen als de bestaande Planten Databank.

## Deployen
Kopieer `index.html`, `assets/portal.css`, `assets/portal.js` en `data/config.json` naar een aparte GitHub Pages repository of branch. Kopieer daarna ook je bestaande `data/typologies/` en `data/layers/` mappen.

## Postcodevelden in de Ecoflora-Excelbestanden
De interface herkent deze optionele kolommen:
- `postcode` of `postcodes`: bijvoorbeeld `3600|3601|3690`
- wildcards: bijvoorbeeld `36*`
- `postcode_van` en `postcode_tot`: bijvoorbeeld 3500 en 3999

Zodra minstens één plantenrij postcodegegevens bevat, wordt de selectie strikt op postcode gefilterd. Zonder postcodevelden toont de app de algemene selectie.

## Biodiversiteitsranking
De app telt numerieke waarden op uit kolommen waarvan de naam woorden bevat zoals:
`vogel`, `vlinder`, `mot`, `bij`, `gastheer`, `biodiversiteit`, `keystone`, `insect`.

De score telt daar bovenop:
- AMBER-match: +30
- streek-eigen/regionale match: +25

Pas labels, gewichten en links aan onder `portal` in `data/config.json`.

## Fytoremediatie
Samenvatting:
`data/layers/fytoremediatie/Fytoremediatie_<type>.xlsx`

Detail:
`data/layers/fytoremediatie/detail/Fytoremediatie_<type>_<medium>_<pollutant>_detail.xlsx`
