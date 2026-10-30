# Definities – Word-add-in voor gedefinieerde termen

Klik in een contract op een gedefinieerde term (bijvoorbeeld *Completion Date*). Het zijpaneel toont direct de betekenis. Staat in die definitie weer een gedefinieerde term, dan klik je daarop door. Bovenin het paneel zie je het pad, bijvoorbeeld **Completion Date › Business Day**, zodat je altijd terug kunt.

## Wat het doet

- **Leest de definities uit het document**:
  - definitietabellen met de term links en de betekenis rechts (een kopregel als *Term / Meaning* of *Begrip / Betekenis* wordt overgeslagen);
  - definitiebepalingen, zoals `"Agreement" means …`, `"Encumbrance": …`, `"Koper" betekent …` en `Onder "Aandelen" wordt verstaan: …`;
  - definities in de lopende tekst, zoals `(the "Seller")`, `(hierna: "Koper")` en `each a "Party" and together the "Parties"`.
- **Herkent termen in de tekst**, ook in het meervoud (*Business Days*) of de bezittelijke vorm (*Seller's*). Langere termen gaan voor kortere: *Completion Date* wordt niet als *Completion* gezien.
- **Uitvouwen in de tekst**: termen binnen een definitie zijn klikbaar. Een klik vouwt de betekenis van die term tussen [ ] op die plek in de tekst uit, zodat je de hele definitie in één keer leest. Nogmaals klikken klapt hem weer in. **Alles uitvouwen** vouwt in één keer de hele keten uit. Een kringverwijzing wordt niet opnieuw uitgevouwen.
- **Losse term openen**: met ↗ in een uitgevouwen stuk open je die term als hoofdterm. Via het pad bovenin ga je terug.
- **Ga naar definitie / Terug naar tekst**: springt naar de definitie in de schedule en daarna weer terug naar waar je was.
- **Zoeken**: het zoekveld filtert de lijst met alle termen in het document.

**Vertrouwelijkheid:** alles gebeurt lokaal in Word. Er wordt geen contracttekst naar een server gestuurd. De server levert alleen de programmacode van het paneel. De add-in vraagt in het manifest lees- en schrijfrechten aan, omdat Word dat nodig heeft om de cursor naar de definitie te verplaatsen. De add-in wijzigt het document nooit.

**Waarom klikken en geen hover?** Word geeft add-ins geen signaal als de muis ergens boven hangt, alleen als de cursor verplaatst wordt. Eén klik in de term is daarom genoeg.

## Uitproberen zonder Word

Open `web/taskpane.html` in een browser (of draai `npm start` en ga naar http://localhost:8080/taskpane.html). Buiten Word start de add-in in demo-modus met een voorbeeldcontract.

Met **Open Word-document** (of door een `.docx` op de pagina te slepen) test je een eigen contract. Het document wordt alleen in de browser gelezen, met de meegeleverde bibliotheek mammoth.js (`web/vendor/`). Er wordt niets verstuurd. De omzetting van Word naar webpagina kan op details afwijken van Word zelf, maar is goed genoeg om de herkenning van definities te testen.

## Installatie op kantoor (voor IT)

De add-in is bedoeld om op een **eigen interne webserver** te draaien. Er is geen database en geen server-side code nodig: het zijn alleen statische bestanden (HTML, JavaScript, CSS en afbeeldingen).

### Eisen aan de server

- **HTTPS met een certificaat dat de werkplekken vertrouwen**, bijvoorbeeld een certificaat van de interne CA dat via groepsbeleid is uitgerold. Office weigert add-ins over HTTP of met een onvertrouwd certificaat.
- **Bereikbaar vanaf de werkplekken**, ook voor thuiswerkers (via VPN of een andere route). Wie Word op het web gebruikt, moet de server vanuit de browser kunnen bereiken.
- **Elke gewone webserver volstaat**: IIS, Apache, nginx of een SharePoint-/intranetomgeving die statische bestanden serveert. Er zijn geen bijzondere instellingen nodig.
- **Internettoegang voor de werkplekken** tot `appsforoffice.microsoft.com`. Daar laadt het paneel de standaard Office-bibliotheek van Microsoft. Bij Microsoft 365 is die toegang er normaal al.
- **Geen contracttekst naar de server.** De server levert alleen de programmacode. Het lezen en herkennen van definities gebeurt in Word op de werkplek zelf.

### Stap 1 – Pakket maken

Hiervoor is Node.js 18 of nieuwer nodig. Geef het adres op waar de add-in komt te staan:

```
npm run build -- https://intranet.kantoor.nl/addins/definities/
```

Dit maakt de map `dist/` met:

- `dist/web/`: deze bestanden komen op de webserver;
- `dist/manifest.xml`: het manifest, met alle adressen al ingevuld voor de eigen server.

### Stap 2 – Webbestanden plaatsen

Kopieer de inhoud van `dist/web/` naar de opgegeven locatie op de server. Open als controle `https://…/taskpane.html` in een browser: je hoort dan de demo met een voorbeeldcontract te zien.

### Stap 3 – Uitrollen via Microsoft 365

1. Ga in het **Microsoft 365-beheercentrum** naar *Settings → Integrated apps → Upload custom apps*.
2. Kies *Office Add-in* en upload `dist/manifest.xml`.
3. Wijs de add-in toe aan de juiste gebruikers of groepen. Start bijvoorbeeld met een kleine testgroep.

Na enige tijd (soms tot 24 uur) verschijnt in Word op het tabblad **Start** de knop **Definities**.

### Updates

- **Nieuwe versie van de webbestanden:** vervang de bestanden op de server. Gebruikers krijgen de nieuwe versie zodra ze het paneel opnieuw openen. Soms is het nodig Word eerst helemaal af te sluiten, omdat Office de bestanden in een cache bewaart.
- **Wijziging in `manifest.xml`** (bijvoorbeeld een ander serveradres of een nieuwe knop): verhoog `<Version>` in het manifest, maak het pakket opnieuw en upload het manifest opnieuw in het beheercentrum.

### Zelf testen vóór uitrol

- **Word op het web**: *Invoegen → Invoegtoepassingen → Mijn invoegtoepassingen → Mijn invoegtoepassing uploaden* en kies `dist/manifest.xml`. Dit kan alleen als het beheerbeleid het toestaat.
- **Word voor Windows**: via een gedeelde netwerkmap als *Vertrouwde catalogus voor invoegtoepassingen* (*Bestand → Opties → Vertrouwenscentrum*). Ook dit kan door IT-beleid geblokkeerd zijn.

### Systeemeisen

Microsoft 365 / Word 2019 of nieuwer (WordApi 1.3) op Windows, Mac of het web.

## Voor ontwikkelaars

```
web/definitions.js   herkenning van definities en termen (pure logica, getest)
web/taskpane.js      paneel + koppeling met de Word-API
web/demo.js          voorbeeldcontract voor demo-modus buiten Word
manifest.xml         add-in-manifest (bron; npm run build vult het serveradres in)
scripts/build.js     maakt het installatiepakket in dist/
test/                tests (npm test, Node 18+)
```

