# Definities – installatiehandleiding voor IT

Deze handleiding beschrijft hoe de Word-add-in *Definities* op een eigen interne webserver wordt gezet en via Microsoft 365 wordt uitgerold. Wat de add-in voor gebruikers doet, staat in de [README](../README.md).

## In het kort

| | |
|---|---|
| **Soort** | Office-add-in (taakvenster) voor Word, met een XML-manifest |
| **Onderdelen** | Statische bestanden (HTML, JavaScript, CSS, PNG) en een manifest. Geen database, geen server-side code. |
| **Hosting** | Elke webserver die statische bestanden via HTTPS serveert: IIS, Apache, nginx of een intranetomgeving |
| **Uitrol** | Microsoft 365-beheercentrum → *Integrated apps* |
| **Rechten in Word** | `ReadWriteDocument` (zie [Beveiliging en privacy](#beveiliging-en-privacy)) |
| **Gegevens** | De documenttekst blijft op de werkplek. Er wordt niets verstuurd, opgeslagen of gelogd. |

## Eisen

### Werkplekken

- **Word-versies:**
  - Microsoft 365-apps of Office 2021 en nieuwer, op Windows of Mac;
  - Word op het web.
- **Windows:** de add-in heeft de moderne webweergave van Office nodig (Microsoft Edge WebView2). Bij actuele Microsoft 365-apps op Windows 10 en 11 is die standaard aanwezig. Oudere eenmalig gekochte versies, zoals Office 2016 en 2019, gebruiken een oudere webweergave waarin de add-in niet werkt.
- **Netwerk:** de werkplekken moeten de interne server bereiken, ook bij thuiswerken (via VPN of een andere route). Ook moeten ze `appsforoffice.microsoft.com` bereiken voor de standaard Office-bibliotheek van Microsoft. Bij Microsoft 365 is dat normaal al toegestaan.

### Webserver

- **HTTPS** met een certificaat dat de werkplekken vertrouwen, bijvoorbeeld van de interne CA. Office weigert add-ins over HTTP of met een onvertrouwd certificaat.
- **Inbedden toegestaan:** Word op het web toont het paneel in een iframe. De server mag voor deze bestanden dus geen `X-Frame-Options: DENY` of `SAMEORIGIN` sturen, en ook geen `Content-Security-Policy` met een `frame-ancestors` die Office uitsluit.
- **Geen bijzondere configuratie.** De gebruikelijke MIME-types voor `.html`, `.js`, `.css`, `.png` en `.txt` volstaan.
- **Caching (aanbevolen):** stel voor `.html` en `.js` een korte cachetijd in, of `Cache-Control: no-cache`. Dan komen updates sneller bij gebruikers aan.

## Installatie

### Stap 1 – Pakket maken

Het manifest bevat een plaatshouder-adres (`https://addins.example.com/definities/`). Dat moet worden vervangen door het adres waar de add-in komt te staan. Er zijn twee manieren.

**Met Node.js (aanbevolen, Node.js 18 of nieuwer):**

```
npm run build -- https://intranet.kantoor.nl/addins/definities/
```

Dit maakt de map `dist/`:

- `dist/web/`: de bestanden voor de webserver;
- `dist/manifest.xml`: het manifest met het ingevulde adres.

Het script weigert adressen zonder `https://` en de plaatshouder zelf. Er zijn geen npm-pakketten nodig, `npm install` is dus niet nodig.

**Zonder Node.js:**

1. Kopieer de map `web/`.
2. Open `manifest.xml` in een teksteditor.
3. Vervang overal `https://addins.example.com/definities/` door het eigen adres, met een `/` aan het eind. Het komt 10 keer voor.

### Stap 2 – Bestanden op de server zetten

Kopieer de inhoud van `dist/web/` (of `web/`) naar de gekozen locatie op de server.

**Controle:** open `https://<adres>/taskpane.html` in een browser op een werkplek. Je hoort dan een voorbeeldcontract met het zijpaneel te zien. Klik in de tekst op *Completion Date*: het paneel toont de definitie.

### Stap 3 – Uitrollen via Microsoft 365

1. Ga in het **Microsoft 365-beheercentrum** naar *Settings → Integrated apps* en kies *Upload custom apps*.
2. Kies als type *Office Add-in* en upload `dist/manifest.xml`.
3. Wijs de add-in toe aan een testgroep. Na de test wijs je hem toe aan de juiste gebruikers of groepen.

Het kan tot 24 uur duren, soms langer, voordat in Word op het tabblad **Start** de knop **Definities** verschijnt. Word opnieuw starten helpt.

### Snel testen zonder centrale uitrol

Als het beleid het toestaat, kan een testgebruiker de add-in zelf laden in **Word op het web**: *Invoegen → Invoegtoepassingen → Mijn invoegtoepassingen → Mijn invoegtoepassing uploaden*, en dan `dist/manifest.xml` kiezen.

## Updates en verwijderen

- **Nieuwe versie van de add-in:** vervang de bestanden op de server. Gebruikers krijgen de nieuwe versie de volgende keer dat ze het paneel openen. Soms moet Word eerst helemaal afgesloten worden, omdat Office add-ins in een cache bewaart.
- **Wijziging in het manifest** (bijvoorbeeld een ander serveradres): verhoog `<Version>` in `manifest.xml` en maak het pakket opnieuw. Werk daarna de add-in in *Integrated apps* bij met het nieuwe manifest.
- **Verwijderen:** verwijder de add-in in *Integrated apps*. De bestanden op de server kunnen daarna weg.

## Beveiliging en privacy

- **Verwerking op de werkplek:** het document wordt gelezen met de Office JavaScript API, in Word op de werkplek. Het herkennen van definities gebeurt daar ook.
- **Netwerkverkeer:** de add-in maakt alleen verbinding met de eigen server (om de eigen bestanden op te halen) en met `appsforoffice.microsoft.com` (de Office-bibliotheek). Er zijn geen andere netwerkverzoeken, geen analytics, geen cookies en geen opslag in de browser. De server ziet alleen gewone verzoeken om statische bestanden.
- **Waarom `ReadWriteDocument`:** de knoppen *Ga naar definitie* en *Terug naar tekst* verplaatsen de cursor in het document. Die rechten zijn aangevraagd zodat dat zeker werkt. De add-in bevat geen code die tekst of opmaak wijzigt.
- **Externe code:** behalve de Office-bibliotheek van Microsoft is dat alleen `web/vendor/mammoth.browser.min.js` (mammoth.js 1.8.0, BSD-2-licentie). Die bibliotheek wordt alleen geladen in de demo-modus buiten Word, als iemand daar een `.docx` opent. In Word wordt hij nooit geladen. In het bestand zijn losse stuurtekens in strings vervangen door de gelijkwaardige `\xNN`-notatie; verder is het ongewijzigd.
- **Controleerbaar:** de eigen code is klein en leesbaar (`web/*.js`, samen ongeveer 900 regels). De herkenning van definities wordt automatisch getest (`npm test`).

## Problemen oplossen

| Probleem | Mogelijke oorzaak en oplossing |
|---|---|
| De knop **Definities** verschijnt niet | De uitrol kan tot 24 uur duren. Controleer of de gebruiker aan de add-in is toegewezen en start Word opnieuw. |
| Het paneel blijft leeg of geeft een foutmelding | Is de server bereikbaar vanaf de werkplek (VPN)? Wordt het certificaat vertrouwd? Opent `https://<adres>/taskpane.html` in de browser? Is `appsforoffice.microsoft.com` bereikbaar? In Word op het web: stuurt de server een `X-Frame-Options`-header? |
| Na een update blijft de oude versie zichtbaar | Sluit Word helemaal af. Helpt dat niet, leeg dan de Office-cache: in Windows de inhoud van `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, terwijl Word gesloten is. |
| "Geen definities gevonden" | Het document heeft een opmaak die nog niet herkend wordt. Geef een (geanonimiseerd) voorbeeld door, dan kan de herkenning worden uitgebreid. |
| Een term ontbreekt of de verkeerde definitie verschijnt | Lees het document opnieuw in met **↻** in het paneel. Blijft het probleem, geef dan een voorbeeld door. |

## Inhoud van het pakket

| Bestand | Doel |
|---|---|
| `taskpane.html`, `taskpane.css` | Het zijpaneel |
| `taskpane.js` | Weergave en koppeling met Word |
| `definitions.js` | Herkenning van definities en termen |
| `demo.js` | Demo-modus als de pagina buiten Word wordt geopend |
| `assets/icon-*.png` | Pictogrammen voor het lint in Word |
| `vendor/` | mammoth.js en de licentie daarvan (alleen voor de demo-modus) |
