# Google-Sheets-Zugriff absichern (ZeitApp)

Die Google-Sheets-Anbindung uebertraegt personenbezogene Arbeits- und Fahrdaten. Die oeffentliche Repository-Datei `google-apps-script.gs` enthaelt **keine** Zugangsdaten. Die bereitgestellte Google-Apps-Script-Web-App muss separat aktualisiert werden.

## Einmalige Umstellung (vor Merge/Live-Schaltung)

1. Die vorhandenen Daten in Google Sheets sichern. Die bisherige funktionierende App und Bereitstellung zunaechst beibehalten.
2. In [Google Apps Script](https://script.google.com/) das zugehoerige Projekt oeffnen und den Code aus `google-apps-script.gs` uebernehmen.
3. Unter **Projekteinstellungen -> Skripteigenschaften** eine Eigenschaft mit dem Namen `ZEITAPP_ACCESS_KEY` anlegen. Als Wert einen **neuen zufaelligen geheimen Schluessel mit mindestens 32 Zeichen** eintragen. Beispiel zum *lokalen* Erzeugen in PowerShell:

   ```powershell
   $bytes = [byte[]]::new(32)
   [System.Security.Cryptography.RandomNumberGenerator]::Fill($bytes)
   [Convert]::ToHexString($bytes)
   ```

   Den erzeugten Wert niemals in GitHub, in Issues, in Screenshots oder in Chat-Nachrichten veroeffentlichen.
4. Unter **Bereitstellen -> Bereitstellungen verwalten** die verwendete Web-App-Bereitstellung aktualisieren. Die Berechtigungen so eng wie moeglich vergeben und die bisherige Web-App-URL auf Bedarf pruefen. Bei einer Web-App fuer `Jeder` bleibt die URL zwar erreichbar, aber Lese- und Schreiboperationen werden zusaetzlich durch den geheimen Schluessel geschuetzt. Dieser clientseitige Schluessel ist **kein Ersatz** fuer eine echte Nutzeranmeldung oder einen serverseitigen Identitaetsdienst.
5. Im neuen ZeitApp-Stand unter **Google-Sheets-Zugriffsschluessel** denselben Schluessel eintragen und **Einstellungen speichern**. Der Wert wird im Browserprofil lokal gespeichert. **Auf gemeinsam genutzten Geraeten ist das keine sichere Geheimnisablage.**
6. Ein Testeintrag und das Ruecklesen pruefen. Ein fehlender oder falscher Zugriffsschluessel muss mit `Zugriff verweigert` abgelehnt werden. Erst anschliessend die Codeaenderung in `main` uebernehmen.

## Sicherheitsgrenzen

- Vorhandene Bereitstellungen aendern sich durch ein GitHub-Commit **nicht**. Fuer bereits veroeffentlichte, ungeschuetzte Web-App-URLs ist die Aktualisierung in Google Apps Script erforderlich.
- Es gibt keine Geheimnisse im GitHub-Quelltext. Skripteigenschaften liegen im Apps-Script-Projekt; Zugriff auf dessen Bearbeitung muss eingeschraenkt bleiben.
- Der Schluessel wird bei Verwendung ueber HTTPS an den Dienst gesendet und liegt auf dem verwendeten Endgeraet im IndexedDB-Speicher. XSS, kompromittierte Endgeraete oder fremde Benutzer desselben Browserprofils koennen ihn gefaehrden.
- Ein gemeinsamer Schluessel erlaubt keine personengenauen Berechtigungen. Fuer mehrere Mitarbeiter oder besonders sensible Daten empfiehlt sich eine serverseitige Anmeldung mit individuellen Rechten.
- Ein vorhandenes Datenleck kann durch diese Aenderung nicht rueckwirkend beseitigt werden.
