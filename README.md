# ZeitApp V7.2

V7.2 basiert auf der Optik von V7 und enthält die wichtigen Fixes:

- Überstundenanzeige korrigiert
  - negative Überstunden werden jetzt auch mobil korrekt mit Minus angezeigt
- Google Sheets Export angepasst
  - `fetch` nutzt jetzt `text/plain;charset=utf-8`
  - das reduziert `Failed to fetch` bei Google Apps Script deutlich
- `doGet()` im Apps Script enthalten
  - Browser-Test der Webhook-URL zeigt jetzt sauber `Webhook läuft`

Wichtig:
- In Google Apps Script die neue Datei `google-apps-script.gs` komplett ersetzen
- Danach die Web-App neu bereitstellen
- In der App die neue `/exec`-URL eintragen und speichern
