# LEG1 Dashboard – Struktur

Diese Dokumentation beschreibt die dynamische, Name-Manager-gesteuerte Struktur des LEG1-Dashboards.

## Grundprinzip

Die Bereichsgrenzen werden über benannte Marker bestimmt. Ein `_NOK`-Marker ist gleichzeitig Endmarker des unmittelbar davorliegenden Bereichs und Ergebnisfeld für die Prüfung dieses Bereichs.

## Besondere Marker

- `LEG1_basisdaten` – Endmarker des Basisdatenbereichs
- `_NOK`-Marker – dynamische Endmarker und Ergebnisfelder der jeweiligen Bereiche

## Zeilenmarker

- `Dashboard_1_summen`
- `Dashboard_2_schwelle_1`
- `Dashboard_3_schwelle_2`
- `Dashboard_4_bezüge`
- `Dashboard_spieler_1`

## Ziel

Neue Bereiche sollen durch Hinzufügen entsprechender Name-Manager-Marker integriert werden können, ohne die zentrale Bereichslogik im VBA-Code anpassen zu müssen.
