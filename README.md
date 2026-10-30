# Definities – Word-add-in voor gedefinieerde termen

Klik in een contract op een gedefinieerde term (bijvoorbeeld *Completion Date*). Het zijpaneel toont direct de betekenis. Staat in die definitie weer een gedefinieerde term, dan klik je daarop door. Bovenin het paneel zie je het pad, bijvoorbeeld **Completion Date › Business Day**, zodat je altijd terug kunt.

## Wat het doet

- **Leest de definities uit het document**:
  - definitietabellen met de term links en de betekenis rechts (een kopregel als *Term / Meaning* of *Begrip / Betekenis* wordt overgeslagen);
  - definitiebepalingen, zoals `"Agreement" means …`, `"Encumbrance": …`, `"Koper" betekent …` en `Onder "Aandelen" wordt verstaan: …`;
  - definities in de lopende tekst, zoals `(the "Seller")`, `(hierna: "Koper")` en `each a "Party" and together the "Parties"`.
- **Herkent termen in de tekst**, ook in het meervoud (*Business Days*) of de bezittelijke vorm (*Seller's*). Langere termen gaan voor kortere: *Completion Date* wordt niet als *Completion* gezien.
- **Doorklikken**: termen binnen een definitie zijn klikbaar. Via het pad bovenin ga je terug.
- **Ga naar definitie / Terug naar tekst**: springt naar de definitie in de schedule en daarna weer terug naar waar je was.
- **Zoeken**: het zoekveld filtert de lijst met alle termen in het document.

**Vertrouwelijkheid:** alles gebeurt lokaal in Word. Er wordt geen contracttekst naar een server gestuurd. De server levert alleen de programmacode van het paneel. De add-in vraagt in het manifest lees- en schrijfrechten aan, omdat Word dat nodig heeft om de cursor naar de definitie te verplaatsen. De add-in wijzigt het document nooit.

**Waarom klikken en geen hover?** Word geeft add-ins geen signaal als de muis ergens boven hangt, alleen als de cursor verplaatst wordt. Eén klik in de term is daarom genoeg.

## Uitproberen zonder Word

Open `web/taskpane.html` in een browser (of draai `npm start` en ga naar http://localhost:8080/taskpane.html). Buiten Word start de add-in in demo-modus met een voorbeeldcontract.

## Installatie op kantoor (voor IT)

De add-in bestaat uit:

1. **Webbestanden** (map `web/`): statische HTML/JS/CSS zonder server-side code. Ze moeten via **HTTPS** bereikbaar zijn.
2. **Het manifest** (`manifest.xml`): dit bestand registreert de add-in in Word.

### Stap 1 – Webbestanden hosten

Standaard verwijst het manifest naar GitHub Pages: `https://jornriem.github.io/hoverdefinition/`.

- Zet in de GitHub-repository onder *Settings → Pages* de *Source* op **GitHub Actions**. Bij elke push naar `main` draait de workflow `.github/workflows/pages.yml`: die test de code en publiceert daarna de map `web/`.
- GitHub Pages voor een **privé**-repository vereist een betaald GitHub-abonnement.
- Host je de bestanden liever op een eigen (kantoor)server, Azure Static Web Apps of SharePoint-CDN? Kopieer dan de map `web/` en vervang in `manifest.xml` overal `https://jornriem.github.io/hoverdefinition/` door het nieuwe adres.

### Stap 2 – Uitrollen via Microsoft 365

1. Ga in het **Microsoft 365-beheercentrum** naar *Settings → Integrated apps → Upload custom apps*.
2. Kies *Office Add-in* en upload `manifest.xml`.
3. Wijs de add-in toe aan de juiste gebruikers of groepen.

Na enige tijd (soms tot 24 uur) verschijnt in Word op het tabblad **Start** de knop **Definities**.

### Zelf testen vóór uitrol

- **Word op het web**: *Invoegen → Invoegtoepassingen → Mijn invoegtoepassingen → Mijn invoegtoepassing uploaden* en kies `manifest.xml`. Dit kan alleen als het beheerbeleid het toestaat.
- **Word voor Windows**: via een gedeelde netwerkmap als *Vertrouwde catalogus voor invoegtoepassingen* (*Bestand → Opties → Vertrouwenscentrum*). Ook dit kan door IT-beleid geblokkeerd zijn.

### Systeemeisen

Microsoft 365 / Word 2019 of nieuwer (WordApi 1.3) op Windows, Mac of het web.

## Voor ontwikkelaars

```
web/definitions.js   herkenning van definities en termen (pure logica, getest)
web/taskpane.js      paneel + koppeling met de Word-API
web/demo.js          voorbeeldcontract voor demo-modus buiten Word
manifest.xml         add-in-manifest
test/                tests (npm test, Node 18+)
```

Een update van de webbestanden is direct actief bij alle gebruikers. Alleen bij een wijziging van `manifest.xml` (zoals een nieuwe knop of een nieuw adres) moet IT het manifest opnieuw uploaden. Verhoog dan ook `<Version>`.
