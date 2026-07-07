# Prompt: taakvelden consistent maken (versie 8, totaalprompt)

Gecontroleerd tegen het werkelijke bestand (04_Taakvelden_2027_ná_verwerken_huren.xlsx, 77 tabbladen, ±62 taakveldtabbladen).

**Nieuw t.o.v. versie 3:**

- elke boekingsregel heeft precies één BBV-categorie (herleidbaar via de middenlaag); dubbele categorieën zoals "41100/47501" op één regel zijn niet toegestaan;
- specificatieplicht voor categorieën 42/43/44/45 en 82/83/84/85, en vaste standaardtoelichtingen voor staat-gebonden categorieën;
- de rol van de financieel adviseur is vastgelegd: de onderbouwing (specificatie en toelichting) is de primaire werkplek;
- per taakveld een vast mutatieblok voor salarissen, activa en reserves en voorzieningen: nu handmatig in te vullen, later via een koppeling met de betreffende staten, zonder herbouw.

**Nieuw t.o.v. versie 4:**

- werkwijze mutaties expliciet vastgelegd: de adviseur verwerkt álle mutaties uitsluitend in de onderlaag (het gedetailleerde deel), inclusief het toevoegen van een nieuw grootboeknummer of een nieuwe BBV-categorie. Het "naar boven transporteren" (het aanmaken van de bijbehorende midden- en bovenlaagregels en het bijwerken van de SOM-bereiken) is een aparte stap die daarna wordt verzorgd, niet door de adviseur zelf. Dit geldt op dezelfde wijze voor reguliere posten én voor staat-gebonden mutaties (salarissen, activa, reserves en voorzieningen) met hun standaardomschrijving.

**Nieuw t.o.v. versie 5:**

- jaarovergang (doorschuiffunctie) vastgelegd: na vaststelling van de begroting kan de raming in één keer worden doorgeschoven naar het volgende begrotingsjaar. Een nieuw jaar is één extra kolom rechts, doorgerekend vanuit het voorgaande jaar met exact hetzelfde formulepatroon (Bijlage A.4) en dezelfde indexbron (Indexering) als het basisbestand — geen herbouw. Zie Fase 5.

**Nieuw t.o.v. versie 6:**

- jaarovergang werkt als piepsysteem, niet als verplichte per-post review: de nieuwe jaarkolom rekent standaard automatisch door met de index uit Indexering, zodat de adviseur voor het overgrote deel van de posten niets hoeft te doen. Alle overige bedragen zijn vrij aanpasbaar naar keuze van de adviseur — een hoger bedrag, een lager bedrag of geen bedrag vanaf het nieuwe jaar — uitsluitend waar gewenst. Alleen de indexberekening ligt vast (bron: Indexering); de rest is vrij.

**Nieuw t.o.v. versie 7:**

- T82 wordt voorlopig niet geüniformeerd maar als black box behandeld: de beherende collega blijft er op de bestaande manier in werken (interne opbouw ongewijzigd), terwijl het bestand alleen de uitkomsten leest via een vaste interface — jaarkolommen uitgelijnd met de andere taakvelden (2025–2030 in J–O, nieuw jaar in P) en drie totaalregels per jaar (Totaal lasten, Totaal baten, Saldo) in vaste cellen waar Recap en Inlees CODA per jaar naar verwijzen. Alleen die interface-totalen worden geverifieerd en zo nodig gecorrigeerd (o.a. de bekende baten-SOM die een regel mist); de invoer- en detailregels van de collega blijven ongemoeid. De volledige uniformering van T82 (Bijlage B.1–B.3) blijft geparkeerd tot akkoord met de collega. Zie Bijlage B.0.

Bij deze prompt hoort het voorbeeldbestand Voorbeeld_uniforme_opbouw_T01_T82.xlsx (Leeswijzer, Indexering, T01 als standaardsjabloon, T82 als complex sjabloon). T82 wordt voorlopig niet geüniformeerd maar als black box aangesloten (Bijlage B.0); de volledige uniformering van T82 (Bijlage B.1–B.3) wordt eerst met een collega besproken en pas na akkoord gebouwd.

## Doel

Elk taakveld in de taakveldenraming krijgt dezelfde opbouw: dezelfde volgorde van rijen en kolommen, en dezelfde manier van formules opbouwen. Na de aanpassing blijven de uitkomsten (de waarden) gelijk aan de oude waarden. Kleine verschillen door afronding zijn acceptabel; grotere verschillen worden gemeld, niet stilzwijgend geaccepteerd.

Uitgangspunt: het doel is een degelijk en betrouwbaar bestand, niet het overnemen van de bestaande opzet. Onvolkomenheden worden opgelost, niet gekopieerd. Alleen de waarde moet gelijk blijven; de manier waarop die tot stand komt, wordt verbeterd.

Concreet aangetroffen onvolkomenheden die niét worden overgenomen:

- labels die formules zijn (bijv. =G94 als omschrijving);
- inconsistente formulestijl (=K82 naast =+K118);
- hardcoded indexcijfers binnen formules (bijv. =579*1.026) terwijl er een Indexering-tabblad bestaat;
- ondoorzichtige samengestelde formules waarin brongetal en opslagfactor onlosmakelijk verweven zijn (bijv. =215790*(1+547325/5364835) op T82);
- meerdere BBV-categorieën op één regel;
- harde correctiebedragen zonder markering (bijv. regels "W02-2025", "A03-2025");
- lege resttabbladen (Blad1, Blad2, Blad3) en dubbele inhoudsopgaven.

## Definities en vaste afspraken

- **Tolerantie afronding**: maximaal € 1 verschil per cel en maximaal € 5 per totaalregel (Totaal lasten, Totaal baten, Saldo). Alles daarboven is een afwijking die gemeld wordt.
- **Rekenregel**: baten tellen negatief (-), lasten positief (+). Een tekort op een taakveld is daardoor een positief getal.
- **Notatie negatieve bedragen**: altijd met een minteken vóór het bedrag (bijv. -259.919), nooit tussen haakjes. Dit geldt ook voor de Excel-getalnotatie van alle cellen die worden aangemaakt of aangepast.
- **Categorie**: elke boekingsregel heeft precies één BBV-categorie (kolom E). De categorie is herleidbaar via de middenlaag en de administratie. Correctierijen en grex-invoerregels zijn invoerhulpen bij een boekingsregel, geen zelfstandige mutaties, en hebben daarom geen eigen categorie.
- **Jaren**: begroting 2027, meerjarenraming 2028–2030. De kolommen 2025 en 2026 (historie) blijven staan; zie beslispunt 1. De jaarkolommen zijn uitbreidbaar van opzet: bij een nieuw begrotingsjaar komt er één jaarkolom rechts bij, doorgerekend vanuit het voorgaande jaar (zie Fase 5); zie beslispunt 10 over het al dan niet meegroeien van het aantal kolommen.
- **Herberekening**: "opnieuw doorrekenen" betekent: het bestand headless openen en herberekenen met LibreOffice (soffice), waarden exporteren en vergelijken met de snapshot van de oude waarden. Alleen formuletekst vergelijken is niet voldoende en niet toegestaan. Een taakveld krijgt pas de status "klaar, wacht op controle" nadat deze herberekening is uitgevoerd. (LibreOffice rekent vrijwel identiek aan Excel maar niet gegarandeerd bit-voor-bit gelijk; de tolerantie vangt dit op. Balanceert een verschil precies op de grens: melden.)

## Rol van de financieel adviseur

De onderlaag (onderbouwing middelen: specificaties, toelichtingen, correctierijen, staat-mutatieblokken en incidentele posten) is de primaire en in beginsel enige werkplek van de financieel adviseur. De midden- en bovenlaag bevatten uitsluitend doorkoppelingen en sommen en worden niet handmatig bewerkt. De opzet maakt dit zichtbaar via de kleurcodering (blauw = invoer, gele cel = handmatige correctie, oranje = incidenteel); optioneel kan bladbeveiliging worden toegepast waarbij alleen de invoercellen ontgrendeld zijn (beslispunt 8).

**Werkwijze bij mutaties en nieuwe posten.** De adviseur verwerkt elke mutatie in de onderlaag en nergens anders:

- een **bijstelling** van een bestaand bedrag gaat via de correctierij van de post, of via het staat-mutatieblok van de post;
- een **geheel nieuwe post** (nieuw grootboeknummer en/of nieuwe BBV-categorie) wordt als nieuw blok onder in de onderlaag toegevoegd, op dezelfde vaste manier als de andere posten: basisregel(s), correctierij, eventueel een incidenteel-blok, eventueel een staat-mutatieblok, met indexkeuze (kolom H) en de bijbehorende (standaard)toelichting.

De adviseur werkt daarmee uitsluitend onderin. Het **"naar boven transporteren"** is een aparte stap die door de assistent wordt verzorgd en niet handmatig door de adviseur:

- aanmaken van de bijbehorende **middenlaagregel** (één regel per BBV-categorie, als =-verwijzing naar de totaalregel van de post in de onderlaag);
- aanmaken van de bijbehorende **bovenlaagregel** (per grootboeknummer, als =-verwijzing naar de middenlaag);
- opnemen van de nieuwe regel in de **SOM-bereiken** van Totaal lasten / Totaal baten.

Zolang de vaste blokgrenzen worden aangehouden (zie Bijlage A.2) valt een nieuwe regel binnen het bestaande SOM-bereik. Ligt een nieuwe post buiten die grenzen, dan worden de doorkoppelingen en SOM-bereiken bij het transport bijgewerkt en wordt de wijziging in de oud-nieuw-controle (Fase 1, stap 3) meegenomen.

Deze werkwijze geldt op dezelfde manier voor reguliere posten en voor **staat-gebonden mutaties** — salarissen (41100), activa (473*/873*) en reserves en voorzieningen (471*/871*) — die via hun vaste mutatieblok met standaardomschrijving in de onderlaag worden ingevoerd. Ook daar voegt de adviseur de mutatieregel onderin toe en wordt de doorkoppeling naar boven daarna verzorgd.

## Specificatie- en toelichtingsregels (onderlaag)

- **Specificatieplicht**: elke post met een BBV-categorie die begint met 42, 43, 44 of 45 (lasten) of 82, 83, 84 of 85 (baten) wordt in de onderlaag gespecificeerd via detailregels (omschrijving + bedrag per jaar). Ontbreekt een specificatie in het oude bestand, dan wordt de specificatiestructuur wél klaargezet (het bedrag komt op één benoemde regel, met ruimte om uit te splitsen) en wordt dit als technische bijzonderheid genoteerd.
- **Standaardtoelichtingen** (geen specificatie in de onderlaag; vaste verwijzingstekst als toelichtingsregel bij de post):
  - 41100 → "zie staat van salarissen"
  - 471* en 871* → "zie staat van reserves en voorzieningen"
  - 473* en 873* → "zie staat van activa"
  - 475* en 875* → "verrekening"
- **Overige categorieën** (bijv. 41101): specificatie niet verplicht; toelichting waar zinvol.
- Toelichtingen zijn uitleg, geen vaste waarde: een bedrag dat in een toelichting wordt genoemd is nooit leidend en kan door de adviseur altijd worden overruled via de handmatige correctie. Past een toelichting duidelijk niet meer bij de cijfers, dan wordt doorgerekend met de cijfers en wordt het verschil apart gemeld; cijfers worden niet aangepast aan toelichtingen en toelichtingen worden niet zelf herschreven.

## Staat-gebonden mutaties (salarissen, activa, reserves en voorzieningen)

Posten met een standaardtoelichting ontlenen hun bedragen aan een staat buiten het taakveld (staat van salarissen, staat van activa, staat van reserves en voorzieningen; 475/875 betreft verrekeningen). Voor deze categorieën geldt:

- Per taakveld komt per betrokken categorie een **vast mutatieblok** in de onderlaag: één regel per mutatie, met categorie (E), omschrijving (G) en bedragen per jaar, plus de standaardtoelichting als vaste regel. Zo kunnen álle mutaties in salarissen, activa en reserves en voorzieningen in het taakveld verwerkt worden. De adviseur voegt de mutatieregel onderin toe; de doorkoppeling naar midden- en bovenlaag wordt daarna verzorgd (zie "Werkwijze bij mutaties en nieuwe posten").
- **Fase nu**: de bedragen worden handmatig ingevuld (blauwe invoer).
- **Later**: dezelfde regels worden omgezet naar verwijzingen (koppeling) naar de betreffende staat — als tabblad in dit bestand of als extern bestand (denk aan de bestaande activastaat en het stambestand reserves en voorzieningen). De opzet verandert daarbij niet: alleen de celinhoud wisselt van waarde naar verwijzing. Bij elk blok staat een klein bronveld ("handmatig" / "koppeling: ...") zodat zichtbaar is wat al gekoppeld is.
- Indexkeuze voor deze posten is standaard **"Geen"**: het bedrag volgt uit de staat (bijv. kapitaallasten uit de activastaat, dotaties uit de staat van reserves en voorzieningen), niet uit een prijsindex.
- **Aansluitcontrole**: zodra een staat gekoppeld is, wordt bij Fase 2 per categorie gecontroleerd dat het totaal over alle taakvelden aansluit op het totaal van die staat; tot die tijd is dit een handmatige controle en wordt het als aandachtspunt op het tabblad Voortgang vermeld.

## Beslispunten vooraf (eerst beantwoorden, dan pas beginnen)

1. **Historische kolommen 2025 en 2026**: bevriezen als vaste waarden, of ook als formules meenemen in de nieuwe structuur?
2. **Controlemomenten**: strikt één taakveld per goedkeuring (±62 wachtmomenten), of na het goedgekeurde pilottaakveld T01 verdergaan in batches van ±5 vergelijkbare taakvelden per goedkeuring? De controle-informatie per taakveld blijft in beide gevallen volledig.
3. **Inwonersreeks aanleveren**: de ontwikkeling van het aantal inwoners staat nu nergens in het bestand. Aanleveren als reeks per jaar, inclusief historie (voorstel: 2021 t/m 2030, zelfde jaarbereik als de indexen), met bronvermelding. De indexreeksen zelf zijn bestaand en worden overgenomen uit het huidige Indexering-tabblad.
4. **Inlees CODA**: het bestaande tabblad "Inlees CODA" als vast sjabloon aanhouden en aanvullen met 2030? (Het loopt nu t/m 2029, de taakvelden t/m 2030.)
5. **Overhead op fte-basis (Fase 3)**: wie levert de fte per afdeling aan, en in welke vorm?
6. **Opruimen**: mogen lege/verouderde tabbladen (Blad1/2/3, dubbele inhoudsopgave, oude draaitabellen) na afronding worden verwijderd of gearchiveerd?
7. **Bijlage B (T82)**: T82 wordt voorlopig als black box aangesloten (Bijlage B.0); de volledige uniformering (B.1–B.3) wordt pas opgepakt na akkoord met de collega. Te bevestigen met de collega: (a) de vaste celposities van de drie totaalregels per jaar (Totaal lasten, Totaal baten, Saldo), en (b) of het interface-totaal — met name de baten-SOM die nu een regel mist en de controleregel met harde waarden (rijen 57–58) — mag worden geverifieerd/gecorrigeerd terwijl de interne opbouw ongewijzigd blijft.
8. **Bladbeveiliging**: boven- en middenlaag vergrendelen zodat alleen de invoercellen in de onderlaag bewerkbaar zijn — gewenst, en zo ja met of zonder wachtwoord?
9. **Volgorde van koppelen**: welke staat wordt als eerste gekoppeld (salarissen, activa, of reserves en voorzieningen), en in welke vorm worden die staten aangeleverd (tabblad in dit bestand of extern bestand)?
10. **Kolomgroei bij jaarovergang** (Fase 5): elke jaarovergang voegt één jaarkolom toe. Blijven alle jaren staan (kolommen groeien mee), of wordt bij het toevoegen van een nieuw jaar de oudste historiekolom gearchiveerd/verwijderd zodat het aantal jaarkolommen constant blijft (schuivend venster)? Dit raakt de vaste blok- en SOM-bereiken en de kolomindeling (A.1) en wordt vóór de eerste jaarovergang vastgelegd.

## Fase 0 — Voorbereiding (verplicht, vóór elke aanpassing)

1. **Backup**: kopie van het origineel met datum en volgnummer (..._backup_00_origineel.xlsx). Na elk goedgekeurd taakveld (of batch) opnieuw een genummerde kopie.
2. **Waarden-snapshot van het héle bestand**: alle tabbladen met opgeslagen (cached) waarden uitlezen en bewaren in een apart referentiebestand. Dit is het enige betrouwbare anker voor "de oude waarde": zodra één taakveld wordt herbouwd, kunnen doorverwijzingen elders (Recap, controle-tabbladen) al veranderen. Alle oud-nieuw-vergelijkingen gebeuren tegen deze snapshot, nooit tegen het werkbestand.
3. **Bekende gebreken vastleggen**: op Recap verwijzen ±15 sociaal-domeintaakvelden naar #REF! (o.a. T622, T623, T64, T660, T6711–T6714, T6752, T6762, T6791, T6792, T6821, T6822), bij lasten én baten, terwijl de tabbladen bestaan. Voor deze taakvelden zijn de cached waarden op het taakveldtabblad zelf het anker; de Recap-verwijzing wordt bij de herbouw hersteld en als afwijking gedocumenteerd. Controleer ook of élk bestaand taakveldtabblad (waaronder T6751, T6753, T6761) een regel in Recap heeft.
4. **Tabblad "Voortgang" aanmaken** met per taakveld één regel: naam, status ("nog te doen" / "bezig" / "klaar, wacht op controle" / "klaar en goedgekeurd"), datum laatste bewerking, eerstvolgende stap, en een aandachtspuntenveld (o.a. openstaande aansluitcontroles met de staten).
   - Bij het begin van elke werksessie wordt dit tabblad als eerste gelezen.
   - Na elke stap wordt de status direct bijgewerkt, niet pas aan het einde.
   - Een taakveld met status "klaar, wacht op controle" wordt niet opnieuw aangepast tot er akkoord is; er wordt verdergegaan met het volgende taakveld met status "nog te doen".
5. **Datainvoer: het bestaande tabblad "Indexering" is en blijft de bron.** Er komt géén nieuw, concurrerend tabblad: alle taakvelden verwijzen er nu al naar. Het tabblad wordt beperkt uitgebreid tot precies twee soorten gegevens, elk als herkenbaar benoemd blok met jaarreeks inclusief historie (2021 t/m 2030):
   - **Indexreeksen**: prijsindex, loonindex, gewogen index, sociaal-domeinindex — de bestaande reeksen, opgeschoond en per reeks duidelijk benoemd. Het losse blok "index conform kaderbrief" wordt samengevoegd met of expliciet gekoppeld aan deze reeksen, zodat elk indexcijfer maar op één plek staat.
   - **Aantal inwoners**: één reeks per jaar, inclusief historische jaren (aan te leveren, beslispunt 3), met bronvermelding. Andere gegevens (woningen, bedrijven, WOZ-waarden) worden nu níét opgenomen; de bloksgewijze opzet laat ruimte om ze later toe te voegen zonder herbouw. Hardcoded indexcijfers in formules worden bij de herbouw vervangen door een verwijzing, waarna de uitkomst tegen de oude waarde wordt gecontroleerd.

## Fase 1 — Werkcyclus per taakveld

Werk taakveld voor taakveld. Start met pilottaakveld **T01** volgens Bijlage A; daarna de overige standaardtaakvelden (oplopend in complexiteit). **T82** wordt voorlopig niet geüniformeerd maar als black box aangesloten volgens Bijlage B.0 (volledige uniformering volgens B.1–B.3 pas na akkoord met de collega). Per taakveld:

### Stap 1. Inventariseren

Lees het bestaande tabblad: grootboeknummers, categorieën, verbijzonderingen, formules, toelichtingen, en cellen die hard zijn ingevuld waar een formule verwacht zou worden. Herleid voor elke regel de BBV-categorie via de middenlaag; regels met meerdere of ontbrekende categorieën worden gesplitst respectievelijk aangevuld en als technische bijzonderheid genoteerd.

### Stap 2. Opnieuw opbouwen volgens de vaste opbouw

Drie lagen, van boven naar onder op het tabblad:

1. **Bovenlaag — verzamelstaat**: alle lasten per grootboeknummer opgeteld, apart alle baten, daaronder één saldo-regel. Uitsluitend doorkoppelingen vanuit de middenlaag.
2. **Middenlaag — lijst per jaar** met alle grootboeknummers/categorieën: per grootboeknummer per categorie één regel (één categorie per regel), als doorkoppeling vanuit de onderlaag.
3. **Onderlaag — budgetopbouw per grootboeknummer/categorie**: het meest gedetailleerde niveau en de werkplek van de adviseur, waar een besteller kan checken of een uitgave binnen het budget past.

Nieuwe posten die de adviseur onderin heeft toegevoegd (nieuw grootboeknummer en/of nieuwe categorie) worden bij deze stap naar boven getransporteerd: de bijbehorende midden- en bovenlaagregels worden aangemaakt en de SOM-bereiken bijgewerkt (zie "Werkwijze bij mutaties en nieuwe posten"). De adviseur voegt zelf niets toe aan de midden- of bovenlaag.

De concrete rij/kolom-indeling en formulepatronen staan in Bijlage A (standaard) en Bijlage B (T82) en zijn uitgewerkt in het voorbeeldbestand. Kernregels die overal gelden:

- elke boekingsregel één BBV-categorie in kolom E;
- specificatieplicht en standaardtoelichtingen zoals hierboven vastgelegd;
- staat-gebonden categorieën via het vaste mutatieblok (nu handmatig, later koppeling);
- elk bedrag per jaar volgt het patroon: handmatige correctie indien ingevuld, anders voorgaand jaar × gekozen index (jaren ná een correctie rekenen automatisch op de correctie verder);
- handmatig gecorrigeerde bedragen zijn zichtbaar gemarkeerd (eigen correctierij met celkleur);
- elke post heeft een zichtbaar indexkeuzeveld (Prijzen / Lonen / Gewogen / Sociaal domein / Geen) dat naar Indexering verwijst; nieuwe structurele posten worden nooit automatisch geïndexeerd — de keuze wordt per post expliciet vastgelegd (kapitaallasten en overige staat-posten: "Geen");
- incidentele posten staan in een eigen, als "incidenteel" gemarkeerd blok, tellen niet automatisch door naar een volgend jaar, en kunnen worden toegevoegd/gewijzigd/verwijderd zonder de structurele formules te raken;
- bij elk grootboeknummer is een leeg veld gereserveerd voor het btw/BCF-regime (nu niet invullen);
- de opzet laat ruimte voor extra kolommen/regels en voor latere verwerking van begrotingswijzigingen.

### Stap 3. Herberekenen en vergelijken

LibreOffice-herberekening uitvoeren (zie Definities) en alle waarden van het taakveld vergelijken met de Fase 0-snapshot:

- binnen tolerantie: geen actie;
- groter verschil: opmerking in de betreffende cel met oude waarde, nieuwe waarde en tabblad/taakveld. Niets wordt stilzwijgend overschreven. Controleer hierbij expliciet op **circulaire verwijzingen** (formule die terugverwijst naar zichzelf of naar een hogere laag) — per taakveld, niet achteraf.

### Stap 4. Bijzonderheden noteren

- **Technische bijzonderheden** (formule buiten patroon, vreemde verwijzing, ontbrekende specificatie of categorie, lege cel waar een formule hoort, opmaakverschil): noteren en doorwerken.
- **Financiële bijzonderheden** (onwaarschijnlijk bedrag, categorie die niet bij het taakveld past, vreemde verhouding lasten/baten, vermoedelijke fout in onderliggende cijfers): direct een vraag stellen en wachten op antwoord voordat dit taakveld verdergaat.

### Stap 5. Afronden en melden

- Eén meldingsregel bovenaan het tabblad, bijv. "3 afwijkingen gevonden, zie opmerkingen in cellen B12, D20, D45" of "geen afwijkingen gevonden".
- Status bijwerken op Voortgang, bestand opslaan.
- In gewone taal uitleggen wat er is gedaan (welk taakveld, wat aangepast, wat gevonden, welke indexkeuzes en welke staat-mutatieblokken worden voorgesteld) en om controle vragen.
- Wachten op akkoord (per taakveld of per batch, beslispunt 2) voordat het volgende taakveld/de volgende batch start.

## Fase 2 — Overall-controle (na alle taakvelden)

1. Alle taakvelden naast elkaar leggen: drie lagen, doorkoppelingen, categorieregels en rekenregel (baten -, lasten +) overal identiek toegepast?
2. Alle bijzonderheden bundelen op één plek; patronen benoemen die op meerdere tabbladen voorkomen.
3. Controleren dat elk taakveld voor indexen en inwoners naar Indexering verwijst en nergens meer hardcoded indexcijfers in formules staan.
4. **Aansluitcontrole staten**: per staat-gebonden categorie (41100, 471/871, 473/873, 475/875) het totaal over alle taakvelden opstellen ter aansluiting met de betreffende staat; zolang de koppeling er nog niet is, wordt dit overzicht klaargezet en gemeld als openstaand controlepunt. Voor T82 kan het aandeel in 41100 en 47501 niet automatisch worden opgehaald (categorieën niet één-per-regel in de oude opzet); dat blijft een handmatige controle en wordt als aandachtspunt op Voortgang vermeld (zie B.0).
5. **Recap** herstellen/controleren: per taakveld saldo lasten en saldo baten, één regel begrotingssaldo inclusief reserves en één regel exclusief reserves (exclusief = totaal zonder T010 Mutaties reserves). Recap bevat uitsluitend doorkoppelingen. Voor T82 verwijst Recap naar de vaste interface-cellen per jaar (de drie totaalregels, zie B.0). Alle #REF!-verwijzingen moeten hierna verdwenen zijn.
6. Volledige herberekening van het hele bestand en vergelijking van de totalen met de Fase 0-snapshot.

## Fase 3 — Overheadtoerekening (pas na Fase 2)

- Verdeling op basis van fte per afdeling (hoofdgrondslag), overal op dezelfde manier toegepast en vastgelegd. Benodigde input: fte-gegevens (beslispunt 5).
- Conform BBV blijft overhead centraal op taakveld 0.4 (T04L, T04B, T04W); de verdeling komt daar herkenbaar náást, niet in de plaats van.
- Let op de relatie met T82: dat taakveld kent al een eigen toerekening van loonkosten en overheadopslagen aan de grondexploitaties (BBV-conform). Die interne toerekening blijft binnen T82 bestaan (zie Bijlage B) en mag niet dubbelen tegen de algemene fte-verdeling.
- **Kostendekkendheid van leges en vergelijkbare heffingen** (Gemeentewet art. 229b, opbrengstlimiet): aparte berekening per betrokken taakveld waarin overhead wél in de kostprijs meetelt, inclusief compensabele btw (BCF) en bijdragen aan reserves voor noodzakelijke vervanging van activa. Welke taakvelden: eerst kijken waar al een extracomptabel overzicht aanwezig is (verwachte kandidaten: T72 riolering, T73 afval, T75 begraafplaatsen, leges binnen T02/T81/T83); bij twijfel als bijzonderheid noteren, niet aannemen.
- De gewone verdeling en de kostendekkendheidsberekening blijven strikt gescheiden.

## Fase 4 — Inleesbestand CODA

- Het bestaande tabblad "Inlees CODA" als sjabloon aanhouden (na bevestiging, beslispunt 4), aangevuld met 2030.
- Uitsluitend doorkoppelingen vanuit de taakvelden; geen nieuwe berekeningen.
- Wordt later een afwijkende officiële indeling aangeleverd, dan wordt die het vaste sjabloon.

## Fase 5 — Jaarovergang (na vaststelling van de begroting)

Zodra de begroting is vastgesteld, wordt de raming in één keer doorgeschoven naar het volgende begrotingsjaar. De opzet is hierop voorbereid: een nieuw jaar is één extra kolom rechts, doorgerekend vanuit het voorgaande jaar met exact hetzelfde patroon als Bijlage A.4 en dezelfde indexbron (Indexering). Er wordt niet herbouwd.

1. **Vastgestelde versie vastleggen**: genummerde kopie van het vastgestelde bestand (..._begroting_<jaar>_vastgesteld.xlsx) plus een waarden-snapshot van het héle bestand. Deze snapshot is vanaf nu het anker "oude waarde" voor de nieuwe ronde; de reeds vastgestelde jaren mogen bij het doorschuiven niet wijzigen.
2. **Nieuwe jaarkolom toevoegen**: rechts van de laatste jaarkolom, in exact dezelfde opmaak en op dezelfde positie op élk taakveldtabblad, plus op de hulpblokken (Q–W), op Indexering, op Recap en op Inlees CODA. Voorbeeld bij begroting 2028: kolom 2031 (P) direct na 2030 (O). Bij een schuivend venster (beslispunt 10) wordt tegelijk de oudste historiekolom gearchiveerd/verwijderd en worden de blok- en SOM-bereiken meegeschoven. Voor T82 (black box, B.0) voegt de beherende collega de nieuwe jaarkolom zelf toe op de uitgelijnde positie en op de eigen manier; zodra die kolom op de vaste positie staat, lezen de recaps het nieuwe jaar automatisch mee via de drie totaalregels.
3. **Indexering aanvullen**: per reeks (prijs, loon, gewogen, sociaal domein) en voor de inwonersreeks het cijfer voor het nieuwe jaar toevoegen (aan te leveren uit de kaderbrief). De nieuwe jaarkolom op de taakvelden verwijst uitsluitend naar Indexering, zodat de gehanteerde index gelijk is aan de bron in het basisbestand — nergens een hardcoded of afwijkende index.
4. **Doorrekenen**: de nieuwe jaarkolom volgt per post automatisch het patroon uit Bijlage A.4 — handmatige correctie indien ingevuld, anders voorgaand jaar × gekozen index (kolom H), of het voorgaande jaar bij indexkeuze "Geen". Incidentele posten lopen niet mee (ze werken per definitie niet door naar een volgend jaar) en verschijnen dus niet ten onrechte als structureel in het nieuwe jaar.
5. **Vrije aanpassing naar keuze (piepsysteem — geen verplichte per-post review)**: standaard rekent de nieuwe jaarkolom automatisch door met de index uit Indexering; voor het overgrote deel van de posten hoeft de adviseur dus niets in te vullen. Alleen de indexberekening ligt vast — de rest is vrij aanpasbaar, uitsluitend waar de adviseur dat wil. Alle situaties zijn mogelijk met één handmatige correctie in de kolom van het nieuwe begrotingsjaar: een hoger bedrag, een lager bedrag, of geen bedrag (0). Omdat een correctie standaard doorwerkt naar de latere jaren, geldt de aanpassing meteen voor de hele nieuwe meerjarenraming; wil de adviseur alleen dat ene jaar wijzigen, dan wordt ook in de latere jaren een correctie gezet. De aangepaste posten zijn zichtbaar gemarkeerd (correctierij) en een overzicht/filter toont in één oogopslag welke posten voor het nieuwe jaar afwijken van de pure indexatie.
6. **Herberekenen en vergelijken**: LibreOffice-herberekening; de reeds vastgestelde jaarkolommen moeten exact gelijk blijven aan de snapshot uit stap 1 — een afwijking daar is een fout en wordt gemeld, niet geaccepteerd. Alleen de nieuwe kolom (en, waar de adviseur bewust een post heeft aangepast, die posten) is nieuw.
7. **Afhechten**: de begrotingsjaar-aanduiding en het meerjarenvenster opschuiven (bijv. begrotingsjaar 2027 → 2028, meerjaren 2029–2031), Recap en Inlees CODA het nieuwe jaar meegeven, en de aansluitcontroles met de staten voor het nieuwe jaar bijwerken.

## Geparkeerd (nu niet bouwen, wel niet onmogelijk maken)

- **Koppeling met de staten** (salarissen, activa, reserves en voorzieningen): de mutatieblokken zijn er vanaf de start; de koppeling zelf volgt later (beslispunt 9). De opzet mag een latere koppeling nergens in de weg zitten.
- **Begrotingswijzigingen**: latere ontwikkeling; bij elk taakveld nagaan dat de opzet dit niet blokkeert.
- **Uniformering T82**: T82 blijft voorlopig een black box (Bijlage B.0); de volledige uniformering volgens B.1–B.3 volgt later, na akkoord met de collega, en kan dan zonder de rest te verstoren worden gebouwd omdat de interface (uitgelijnde jaren + drie totaalregels) gelijk blijft.
- **Fusie Oirschot (per 1-1-2028)**: zodra hun bestand er is, zelfde behandeling (taakveld voor taakveld, zelfde vaste opbouw, zelfde oud-nieuw-controle, controlemoment per taakveld) en pas daarna per taakveld samenvoegen. Beleidsverschillen noteren als bijzonderheid, niet zelf oplossen. Dit spoor staat los van het werk aan de eigen raming.

## Praktische risico's en afspraken

- **Backup-discipline**: kopie vóór aanvang en na elk goedgekeurd taakveld/batch, genummerd.
- **Tussentijds opslaan**: na elk afgerond taakveld.
- **Snapshot is leidend**: oud-nieuw-vergelijkingen altijd tegen de Fase 0-snapshot, nooit tegen live waarden in het werkbestand.
- **Geen macro's**: dit bestand is een .xlsx zonder VBA. (Voor de activastaat met macro's blijft gelden: bestaande formules aanpassen, niet herbouwen.)
- **Hervatbaarheid**: het tabblad Voortgang is de eerste en enige plek om te bepalen waar verdergegaan wordt.

## Eerste stap bij uitvoering

1. Beantwoord de beslispunten.
2. Voer Fase 0 volledig uit en meld de uitkomst (inclusief de omvang van de #REF!-problematiek en andere gebreken die bij de snapshot opvallen).
3. Bouw pilottaakveld T01 volgens Bijlage A en wacht op controle. Pas daarna verder.

## Bijlage A — Uitwerking standaardsjabloon: taakveld T01

T01 (Bestuur en ondersteuning) is het pilottaakveld en tegelijk de definitie van het standaardsjabloon; zie het tabblad "T01 (voorbeeld)" in het voorbeeldbestand. Elk regulier taakveld wordt hierop gemodelleerd.

### A.1 Vaste kolomindeling (alle drie de lagen)

| Kolom | Inhoud |
|---|---|
| C | Markeringskolom (op T82: "x" = telt mee in boekwaardemutatie) |
| D | Grootboeknummer |
| E | BBV-categorie — precies één per boekingsregel |
| F | Verbijzondering (leeg bij standaardtaakvelden; complexcode op T82) |
| G | Omschrijving (altijd tekst, nooit een formule als label) |
| H | Indexkeuze per post: Prijzen / Lonen / Gewogen / Sociaal domein / Geen (datavalidatie-lijst) |
| I | Btw/BCF-regime (gereserveerd, nu leeg) |
| J | 2025 (historie) |
| K | 2026 (historie) |
| L–O | 2027, 2028, 2029, 2030 |
| P e.v. | nieuw begrotingsjaar bij jaarovergang (bijv. 2031), rechts aangesloten — zie Fase 5 |

Bovenaan elk tabblad staat één klein hulpblok (kolommen Q–W) met de indexreeksen per jaar, elke cel uitsluitend als verwijzing naar Indexering, plus waar nodig één rij inwoners. Dit hulpblok is op elk tabblad identiek geplaatst. De jaarkolommen zijn zo aangelegd dat een nieuw jaar rechts kan worden aangesloten met exact hetzelfde formulepatroon (A.4) en dezelfde Indexering-verwijzing, zonder de rijstructuur of de blokgrenzen te raken (beslispunt 10 bepaalt of oude kolommen dan meegroeien of meeschuiven).

### A.2 Bovenlaag (verzamelstaat)

- Kop: Programma, Taakveld, meldingsregel afwijkingen.
- Blok "Lasten in € (+)": per grootboeknummer één regel, uitsluitend =-verwijzing naar de totaalregel in de middenlaag; afsluitend Totaal lasten als SOM over het blok.
- Blok "Baten in € (-)": idem. De SOM-bereiken beslaan het volledige blok (vaste blokgrenzen; de oude T82-fout waarbij één regel buiten de SOM viel, is daarmee uitgesloten). Een nieuwe post die onderin wordt toegevoegd, valt bij het transport binnen deze vaste blokgrenzen; ligt hij erbuiten, dan worden de SOM-bereiken bijgewerkt.
- Eén saldo-regel: Totaal lasten + Totaal baten (baten zijn negatief).

### A.3 Middenlaag (per grootboeknummer/categorie)

Per grootboeknummer één blokje: kopregel (D: grootboeknummer, G: omschrijving), daaronder **per BBV-categorie één regel** met per jaar een =-verwijzing naar de totaalregel van die post in de onderlaag, afgesloten met een totaalregel (SOM). Geen berekeningen in deze laag. Via deze laag is voor elke onderlaagregel de categorie herleidbaar. Voegt de adviseur onderin een post met een nieuwe categorie toe, dan krijgt die hier bij het transport een eigen regel (één categorie per regel).

### A.4 Onderlaag (budgetopbouw) — het formulepatroon

Direct onder de kop staat de vaste specificatie-/toelichtingsregel: "Specificatieplicht: categorieën 42/43/44/45 en 82/83/84/85 worden gespecificeerd via detailregels. Standaardtoelichtingen: 41100 zie staat van salarissen · 471/871 zie staat van reserves en voorzieningen · 473/873 zie staat van activa · 475/875 verrekening."

Per post (grootboeknummer × categorie) een vast blok:

1. **Basisregel(s)** — de specificatie: per detailregel omschrijving, categorie (E) en bedrag per jaar. Voor categorieën met specificatieplicht zijn dit de verplichte detailregels; ontbreekt een uitsplitsing, dan één benoemde regel met ruimte om later uit te splitsen (technische bijzonderheid).
2. **Correctierij** ("handmatige correctie", gekleurde invoercellen) — normaal leeg; de adviseur vult hier een bedrag in als de automatische indexering niet klopt. Geen eigen categorie (invoerhulp).
3. **Incidenteel** — indien van toepassing: aparte regels in een blok gemarkeerd "incidenteel", met categorie, per jaar los ingevuld, zonder doorwerking naar volgende jaren.
4. **Staat-mutatieblok** — voor posten met categorie 41100, 471*/871*, 473*/873* of 475*/875*: één regel per mutatie met de standaardtoelichting als vaste regel eronder en een bronveld ("handmatig" nu; "koppeling: <staat>" later). Indexkeuze "Geen".

Formule in de basisregel, per jaar vanaf 2027 (voorbeeld kolom L; L{c} is de correctiecel, K{b} het voorgaande jaar op de basisregel, $H{b} de indexkeuze van de post):

```
=ALS(L{c}<>""; L{c};
   ALS($H{b}="Geen"; K{b};
      AFRONDEN(K{b} * INDEX(hulpblok-indexkolom; VERGELIJKEN($H{b}; reekslabels; 0)); 0)))
```

(In het bestand worden formules met Engelse functienamen opgeslagen — IF/ROUND/INDEX/MATCH — Excel toont ze in het Nederlands.)

Doordat elk jaar op het voorgaande jaar van dezelfde rij rekent, werkt een handmatige correctie automatisch door naar alle latere jaren (correctie € 10.000 in 2028 → index 2029 over € 10.000). Een aflopende reeks zoals wachtgeld wordt met opeenvolgende correcties gemodelleerd (2028 verlaagd, 2029 op 0), zichtbaar in het voorbeeldbestand.

Totaal per post = SOM van basisregels + incidentele regels + regels van het staat-mutatieblok van dat blok; correctierijen tellen nooit mee in een SOM (de correctie werkt via de basisregel).

### A.5 Wat er in T01 concreet verandert (te melden bij de controle)

- Labels die formules zijn (=G94 e.d.) worden vaste teksten.
- Hardcoded aanvullingen zonder markering (W02-2025, A03-2025, elk € 4.000 structureel) worden benoemde structurele regels met indexkeuze Geen (ter bevestiging), met de W/A-referentie in de omschrijving.
- Samengestelde invoer zoals =43940+15751 wordt gesplitst in benoemde regels of één regel met de uitsplitsing in de toelichting.
- Formules zoals =ROUND(K225*T2,0) gaan over op het patroon uit A.4 met indexkeuzeveld in kolom H.
- De post 41101 (uitkering voormalig personeel) valt niet onder de specificatieplicht; het aflopende wachtgeld wordt via correcties gemodelleerd. Zou er een 41100-post zijn, dan geldt het staat-mutatieblok met "zie staat van salarissen".
- Per post wordt een voorstel voor de indexkeuze gedaan op basis van de huidige verwijzing; alle voorstellen worden bij de controle voorgelegd.
- Controle: Totaal lasten, Totaal baten en Saldo per jaar gelijk aan de snapshot, en de Recap-verwijzingen aangepast naar de nieuwe totaalregels.

## Bijlage B — Uitwerking complex sjabloon: taakveld T82 (grondexploitatie)

**Status: T82 wordt voorlopig niet geüniformeerd maar als black box behandeld — zie B.0. De volledige uniformering (B.1–B.3 hieronder) blijft geparkeerd tot er akkoord is met de collega en wordt pas daarna gebouwd.** Zie het tabblad "T82 (voorbeeld)" in het voorbeeldbestand. Bij uniformering volgt T82 dezelfde drie lagen en kolomindeling als Bijlage A, met aanvullende regels omdat bron en rekenlogica anders zijn.

### B.0 Voorlopige behandeling: T82 als black box (interface-contract)

T82 wordt voorlopig niet geüniformeerd. De collega die T82 beheert, blijft er op de bestaande manier in werken; de interne opbouw (grex-doorrekening, eigen toerekening van loonkosten en overhead, sluitpost boekwaardemutatie) blijft ongewijzigd. Het bestand behandelt T82 als een black box waarvan alleen de uitkomsten worden gelezen. Voorwaarde is een vaste, betrouwbare interface:

- **Jaarkolommen uitgelijnd**: de jaarkolommen in T82 staan in exact dezelfde kolommen als bij de standaardtaakvelden (2025 t/m 2030 in J–O, nieuw begrotingsjaar in P e.v.). Zo sluiten de jaren één-op-één aan op de andere taakvelden en op de recaps, en staat bij de jaarovergang de nieuwe kolom vanzelf op de juiste plek.
- **Drie totaalregels per jaar in vaste cellen**: Totaal lasten, Totaal baten en Saldo, per jaarkolom, op bekende en vastgelegde celposities. Recap en Inlees CODA verwijzen uitsluitend naar deze cellen — per jaar, zodat elk jaar bij het juiste jaar in de recap terechtkomt.
- **Alleen de totalen geverifieerd en zo nodig gecorrigeerd**: de drie totaalregels worden gecontroleerd op juistheid, in het bijzonder de bekende baten-SOM die nu één regel buiten de telling laat en de controleregel met hard ingevoerde waarden (rijen 57–58). Waar een interface-totaal aantoonbaar onjuist is, wordt uitsluitend het SOM-bereik van de totaalregel hersteld — niet de invoer- of detailregels van de collega. Het verschil t.o.v. de snapshot wordt als afwijking gedocumenteerd en voorgelegd. Zo kloppen de cijfers die naar Recap en de Raad gaan, zonder de werkwijze van de collega te raken.
- **Index en doorrekening blijven binnen de box**: T82 ontleent zijn bedragen aan de grex-doorrekening (indexkeuze "Geen" conform B.2) en aan een eigen, met de loonindex geïndexeerd loonkostenblok. Dat blijft de verantwoordelijkheid van de collega binnen T82; het bestand herrekent T82 niet van buitenaf. De aansluiting op de rest van het bestand — en dus dat de juiste bedragen per jaar in de recaps worden meegenomen — loopt uitsluitend via de uitgelijnde jaarkolommen en de drie totaalregels.
- **Jaarovergang (Fase 5)**: bij een nieuw begrotingsjaar voegt de collega de nieuwe jaarkolom in T82 zelf toe, op de uitgelijnde positie en op de eigen manier. De recaps lezen de nieuwe jaarkolom automatisch mee zodra die op de vaste positie staat.
- **Bladbeveiliging (beslispunt 8)**: als de andere lagen worden vergrendeld, blijft T82 als enige tab volledig bewerkbaar voor de collega.
- **Aansluitcontrole staten (Fase 2)**: omdat de categorieën in de oude opzet niet één-per-regel zijn, kan T82's aandeel in 41100 en 47501 niet automatisch worden opgehaald; dit blijft een handmatige controle en wordt als aandachtspunt op Voortgang vermeld.

T82 is hiermee de enige, gedocumenteerde niet-uniforme tab. De behandeling is omkeerbaar: omdat de interface (uitgelijnde jaren + drie totaalregels) gelijk blijft, kan T82 later alsnog volgens B.1–B.3 worden geüniformeerd zonder de rest van het bestand te verstoren, op het moment dat de collega eraan toe is.

### B.1 Wat T82 anders maakt (aangetroffen in het bestand)

1. **Extra dimensie per grondexploitatiecomplex** (kolom F): G0101 Dijkstraten, G0201 Aarlesche Erven, G0301 Steegsche Velden Oost, G0401 Centrumplan, G0501 Molenstraat, G0601 Steegsche Velden Noord (en G0901 overig).
2. **De grex-doorrekening is de bron, niet de index.** Bedragen per complex per jaar staan nu hard in formules mét een verweven btw-opslag (bijv. =215790*(1+547325/5364835)); brongetal en factor zijn niet te onderscheiden.
3. **Eigen toerekening binnen het taakveld**: loonkosten per team (wél geïndexeerd met de loonindex, plus een vaste stelpost van -300.000) en overheadopslagen worden doorbelast aan de complexen naar rato; het restant blijft als "niet gedekt" informatief buiten het taakveld.
4. **Sluitpost boekwaardemutatie** (categorie 47501): per complex het negatief van alle toegerekende lasten en baten, plus aparte regels "dotatie voorziening nadelig saldo" per complex.
5. **Losse controleregel** (huidige rijen 57–58) met hard ingevoerde controlewaarden.
6. Kleine gebreken: een baten-SOM die één regel mist, gemengde absolute/relatieve verwijzingen, omschrijvingen als formule.

### B.2 Uniformeringsregels voor T82

- **Zelfde drie lagen, zelfde kolommen** als Bijlage A; kolom F is overal gevuld op boekingsregels per complex. Bovenlaag en middenlaag zijn puur doorkoppeling en som.
- **Categoriesplitsing 652901**: de dekking per complex wordt gesplitst in twee boekingsregels met elk één categorie — het overheaddeel (47501, toelichting "verrekening") berekend als dekking complex ÷ totale dekking × totaal overheadopslagen, en het loonkostendeel (41100, toelichting "zie staat van salarissen") als restant; samen exact de dekking van het complex. De middenlaag toont per categorie één regel. De historische splitsing (2025/2026) wordt overgenomen uit het huidige bestand.
- **Grex-invoerblok in de onderlaag**: per post per complex worden brongetal en opslag gescheiden: (1) grex-bedrag per jaar (invoer, gemarkeerd), (2) btw-opslagfactor per complex per jaar in één centrale factorentabel op het tabblad (invoer, met teller/noemer zichtbaar en bron), (3) berekend bedrag = grex-bedrag × (1 + factor) — de boekingsregel, met categorie. Grex-posten krijgen indexkeuze Geen. Grex-invoerregels zelf zijn invoerhulpen zonder categorie.
- **Loonkostenblok teams**: standaardpatroon uit A.4 met indexkeuze Lonen, inclusief correctierij. De stelpost -300.000 wordt een aparte, benoemde en gemarkeerde structurele regel (indexkeuze Geen). Dit blok is een informatief rekenblok: het telt niet mee in het taakveld, alleen de dekking per complex telt.
- **Overheadopslagen**: apart invoerblok, per jaar hard ingevoerd met bronvermelding, later te koppelen conform het staat-/bronprincipe; relatie met Fase 3 expliciet vastleggen zodat er geen dubbeltelling met de fte-verdeling ontstaat.
- **Sluitpost boekwaardemutatie**: per complex één formule volgens één vast patroon: minus de som van álle aan dat complex toegerekende boekingsregels, via SOMMEN.ALS op de complexcode (kolom F) en de markering (kolom C "x"), zodat een nieuwe regel voor een complex automatisch meetelt. Categorie 47501, toelichting "verrekening".
- **Dotaties voorziening nadelig saldo**: aparte, benoemde regels per complex (invoer vanuit de grex; later koppelbaar aan de staat van reserves en voorzieningen), categorie 47501, indexkeuze Geen.
- **Controleregel formaliseren**: vaste controleregel direct onder het saldo: saldo taakveld minus som dotaties voorzieningen = 0 (binnen tolerantie), beide zijden als formule in plaats van harde controlewaarden. Wijkt het af: melden als financiële bijzonderheid, niet gladstrijken. (In het voorbeeldbestand sluit deze regel op 0 en komt het saldo 2027–2029 vrijwel exact uit op de huidige harde controlewaarden: 313.327 / 319.593 / 325.985.)
- **Specificatie**: 43812- en 83100-posten worden per complex (en waar relevant per component) gespecificeerd; verkoopramingen met samengestelde formules (bijv. =-936000-1950000-107174) worden uitgesplitst in benoemde regels per component (zie vraag 7).
- **Extra aandacht circulariteit**: door de verdeel- en sluitpostformules is het risico op kringverwijzingen hier het grootst; de circulariteitscontrole wordt op T82 op het hele tabblad uitgevoerd vóór de herberekening.
- **Oud-nieuw-controle**: alle totalen per post, per complex en de drie hoofdtotalen per jaar gelijk aan de snapshot; bekende oude fouten (zoals de baten-SOM die een regel mist) worden hersteld en het effect wordt als afwijking met opmerking gedocumenteerd en expliciet voorgelegd.

### B.3 Vragen voor het gesprek met de collega (vóór de bouw beantwoorden)

1. Waar komen de btw-opslagfactoren per complex vandaan en wie actualiseert ze — akkoord met vastlegging als centrale factorentabel met zichtbare teller/noemer per complex per jaar?
2. Blijft de stelpost -300.000 (correctie lagere toerekening loonkosten) structureel bestaan, en onder welke naam moet die zichtbaar zijn?
3. Is de complexenlijst actueel? G0301 en G0501 lijken afgerond (alleen nog historie); als "afgesloten" markeren of opschonen?
4. De dotaties aan de voorziening nadelig saldo staan per complex hard ingevuld: komen die uit de grex-doorrekening en moeten ze als grex-invoer worden behandeld, met later een koppeling aan de staat van reserves en voorzieningen (voorstel: ja)?
5. De controleregel (saldo = dotaties): klopt die interpretatie, en moet een afwijking blokkeren of alleen melden?
6. De overheadopslagen binnen T82 versus de latere algemene overheadverdeling (Fase 3): akkoord dat T82 zijn eigen toerekening houdt en buiten de fte-verdeling blijft?
7. Verkoopramingen (grondverkopen sociaal/vrije sector/overig) bevatten nu samengestelde formules met correctiebedragen: mogen die worden uitgesplitst in benoemde regels per component (woningen, kavels, btw-afdracht)?
8. De splitsing van de dekking in 41100/47501 per complex (B.2): akkoord met de rekenwijze (overheaddeel naar rato van de dekking, loondeel als restant), conform de huidige praktijk in het bestand?
