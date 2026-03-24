# ZeitApp V8

V8 bringt automatischen Sync und Wiederherstellung:

- neuer Eintrag wird lokal gespeichert und direkt automatisch nach Google Sheets synchronisiert
- offene Einträge können weiter manuell gesammelt synchronisiert werden
- Rücklesen aus Google Sheets möglich
- Google Sheet dient damit als Sicherung und Wiederherstellungsquelle
- Korrigierte Einträge behalten ihre ID und werden im Sheet aktualisiert

Wichtig:
- `google-apps-script.gs` komplett in Google Apps Script ersetzen
- Web-App neu bereitstellen
- neue `/exec`-URL in der App speichern

Hinweis:
- App bleibt lokal schnell und offlinefähig
- wenn beim Speichern kein Netz da ist, bleibt der Eintrag lokal offen und kann später synchronisiert werden
