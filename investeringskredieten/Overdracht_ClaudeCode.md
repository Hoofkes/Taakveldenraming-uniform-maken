# Overdracht: Dashboard investeringskredieten (stand na stap 13, stap 13 nog niet getest in Excel)

Lees dit samen met het functioneel ontwerp (versie 3, `.md`). Dit document bevat de afwijkingen en beslissingen sinds dat ontwerp en de huidige stand van het bestand.

**Gebruiker:** Michiel, financieel adviseur Gemeente Best. Taal: Nederlands. Korte terugkoppeling, alleen actiepunten. Negatieve bedragen met minteken en punten bij duizendtallen, nooit tussen haakjes. Niet-technisch voor VBA: stap-voor-stap uitleg, bij Excel-uitleg de cel noemen. Doe precies wat gevraagd wordt, stel liever een vraag dan te gokken.

## Bestanden
- Werkmap: `Stambestand_investeringskredieten_dashboard_stap4*.xlsm` (laatste versie van Michiel, kopie; origineel staat veilig).
- `Dashboard_AlleMacros_18.bas` (nieuwste): alle macro's in een module (wordt via VBA-editor > Bestand importeren geladen; module eerst verwijderen).
- In de kopie ontbreken de spiegelbladen (Financieel adviseur, Financiële administratie, Kredietoverzicht, Maatwerk, Activa bestaand). Bestaan wel in het origineel.
- `vbaProject.bin` nooit rechtstreeks wijzigen via XML; macro's alleen via importeren of Excel-automatisering (Windows, Excel, "Vertrouwen in VBA-projectmodel").

## Opbouw werkmap
Bladen: Verzamellijst alle kredieten (data, kop rij 1), Voertuigen, Parameters, Dashboard, Controle, Logboek.
- Verzamellijst: kolommen zoals ontwerp. Nieuw: AP controlestatus, AQ door, AR op, AS opmerking, AT naam samengesteld krediet, AU toegevoegd op, AV laatst gewijzigd door, AW op. Formulekolommen E, G, K, AD, AL, AM, AN lopen door tot rij 2468 (`Par_MaxRij`). Op rij 1474 staan SUBTOTAL-regels (AE:AH): aandachtspunt, verplaatsen naar het Dashboard bij stap 12.
- Parameters: B2 `Par_LopendJaar` (hard), B3 `Par_LaatsteRij`, B4 `Par_EersteVrijeRij`, B5 `Par_VolgendVolgnr`, B6 pad uitgavenbestand, B7 peildatum uitgaven, B8 `Par_HuidigeGebruiker`, B9 `Par_HuidigeRol`, B10 `Par_MaxRij`, B11 laatste verversing keuzeopties. Tabellen: Tbl_Rollen (A14:F), Tbl_Programma (Y:AA), Tbl_Taakveld (AC:AF, Iv3 2026), Tbl_Balanscode (AH:AJ), Tbl_BijdrageBron (AL:AR). Keuzelijsten `Lst_...` in kolom H t/m W vanaf rij 14 (worden bij openen ververst).
- Namen: `Krediet_Aantal/Volgnr/J/AL/AM/Data` (OFFSET), `Kaart_Kolommen/Invoer/Rijnummer`.
- Dashboard: invoerkaart rij 9 t/m 50, `KaartRij(i)` en `BijdrageRij(j,k)` bepalen de rijen; veld 4 (kort) verborgen; rijnummer rij 52; knoppen rij 6 (Krediet wijzigen, Scherm leegmaken, Krediet verwijderen), rij 22 (Tijdschrijvers kiezen), rij 54 (Ingevuld? Verwerken!). Teller bovenaan.
- Controle: lijst te controleren/afgekeurd, knoppen Lijst vernieuwen, Regel openen, Bevestigen, Afkeuren.

## Macro's (module)
Intake/rollen (`Auto_Open`, `WieWerktEr`, `GebruikerVerwijderen`, `HeeftRecht`, `LogRegel`), `LijstwaardenNormaliseren`, `WerkbereikVerruimen`, `ControlekolommenAanmaken`, `DashboardAanmaken`, `KeuzelijstenVerversen`, `KredietToevoegen`, `KredietWijzigen`, `WijzigingOpslaan`, `Verwerken`, `KredietVerwijderen`, `TijdschrijversKiezen`, `ControleBladAanmaken`, `ControleLijstVernieuwen`, `RegelOpenenTerControle`, `RegelBevestigen`, `RegelAfkeuren`, `ProefregelWissen`, `VerzamellijstOpschonen`, `KortOmschrijving`.
Werkwijze in elke schrijfmacro: nulmeting vooraf en achteraf (alleen datarijen, nooit hele kolommen, i.v.m. SUBTOTAL), bij verschil terugdraaien. Nooit rijen invoegen/verwijderen.

## Beslissingen sinds ontwerp v3
- Status H: L2026 = krediet loopt in 2026 (bij toekomstige kredieten: jaar van start); A2026 = krediet afgesloten in 2026. V bestaat niet meer (verwijderen = helemaal weg).
- Jaarovergang (stap 14): bij openen in het nieuwe jaar alle L<oud jaar> omzetten naar L<nieuw jaar>; jaren in M en N schuiven dan +1 voor die kredieten (precieze regel nog bevestigen).
- Bij status A moet in de toelichting "Krediet afgesloten" staan.
- Kredietoverzichten (stap 13, ontwerp §7): Meldingenblad blijft apart bestaan; overzicht krijgt kolommen Status en Melding; toelichting breed met terugloop. Extra: kredietoverzicht jaarrekening = situatie 31-12 vorig jaar, alle L en A van dat jaar.
- Kredietnummer (kol C) en activanummer (X) mogen bij intake leeg; alle andere velden verplicht, behalve budgetten. Afschrijftermijn 0 t/m 60.
- Korte omschrijving wordt automatisch afgeleid (max 20, `KortOmschrijving`); alleen te overschrijven door de verborgen rij.
- Programma, taakveld, balanscode worden gekozen als "code - omschrijving" en als code weggeschreven. Exploitatienummer: vrij invoerveld (lijst volgt, Michiel levert die later).
- Samengestelde kredieten: groepsnummer in J wordt automatisch toegekend (keuze nieuw samengesteld krediet of bestaande groep). Budgethouder boekt op een deelkrediet; kosten van de hele groep worden opgeteld en met sleutel K over de deelkredieten verdeeld (ook in overzichten).
- Bijdragen van derden: eigen regel(s) in de verzamellijst met negatief budget, zelfde kredietnummer, toelichting = bron. Op het Dashboard max 2 per jaar, bron per jaarregel te kiezen of typen, regels met dezelfde bron worden samengenomen.
- Raadsoverzicht: regel weglaten als budget (Z of een jaar) negatief is; budgethoudersoverzichten tonen die wel. Statussen V en A van vorige jaren niet tonen.
- Rollen: financieel adviseur toevoegen/wijzigen/verwijderen; financiële administratie wijzigen/bevestigen. Zelfregistratie nieuwe gebruiker in intake. Filter en knop Sort van de verzamellijst zijn verwijderd.
- Overzicht voor beheerder tijdschrijven: apart Excelbestand in SharePoint-map met leesrechten (pad op Parameters); een regel per kredietnummer, tijdschrijvers in een cel, met acties nieuw/gewijzigd/afgesloten op basis van verschil met vorige uitdraai, incl. "toegevoegd op". Kredietnummer is de sleutel. Pad nog opvragen.

## Stap 11 - werkelijke uitgaven bijwerken (`UitgavenBijwerken`, knop Dashboard rij 6)
- Eén vraag: "Wil je het bestand uploaden?" Ja = het nieuwste .xls*-bestand uit de exportmap (administratie/jaarrekening/mva/Stambestand kredieten/export coda) verwerken. Map = map van Parameters B6; eerste keer mapkiezer. Peildatum (B7) = datum van het bestand.
- Kolommen worden zelf herkend (nieuw pakket op komst): kredietnummer = 7 cijfers beginnend met 7; kostensoort begint met 4 (uitgaven) of 8 (inkomsten), 5 of 6 posities (testupload heeft beide); jaarkolommen via jaartal in de kop (niet budget/restant/raming). Lopend jaar -> AC, vorig jaar -> AB. Kolom voor lopend jaar + 1: melding, eerst jaarrol.
- Totaalregels (nummer zonder kostensoort) en eindtotaal: alleen controle op dubbeltelling; wijkt af = stoppen.
- Verdeling: samengesteld krediet (J getal) = saldo groep via K. Anders uitgaven op uitgavenregel(s), inkomsten op bijdrageregel(s) (negatief budget) naar verhouding van budget. Zelf verdeeld = melding.
- Vorig jaar na 31-3: niets schrijven, ALARM "actie/herstel nodig" per verschil. Lopend jaar wordt altijd verwerkt.
- Nummer niet in kolom C met bedrag: ALARM (besluit nodig, raad, budgethouder in gebreke). Krediet niet in uitdraai met bedrag: blijft staan + notitie. Status A waarvan het bedrag wijzigt: wel verwerkt + melding "afgesloten krediet, toch uitgaven/inkomsten verwerkt" + notitie in de cel (bestaande bedragen op A-kredieten zijn goed).
- Harde poort per jaar: uitdraai - onbekende kostensoort - zonder krediet + blijft staan = som kolom. Nulmeting op overige kolommen; afwijking = terugdraaien.
- Blad "Meldingen uitgaven" wordt elke run opnieuw gevuld (ALARM bovenaan, rood tabblad). Voorwaardelijke opmaak: AD rood (overschreden) / groen (meer inkomsten), AJ oranje als toelichting ontbreekt.
- Simulatie testupload (peildatum na 31-3-2026): aansluiting 2026 sluit (0,00); 1 ALARM zonder krediet (7201025, -24.095,73); 6 ALARM 2025 herstel; 33 rijen AC gewijzigd.

## Verwijderen (aangepast)
- `KredietVerwijderen` haalt het krediet helemaal weg: alle regels met hetzelfde volgnummer (incl. bijdrageregels). Alleen bij echte fouten; geweigerd zodra AA, AB of AC ergens niet 0 is (dan afsluiten met A).
- Geen Excel-rij verwijderen: regels eronder schuiven een plek omhoog (alleen invoerkolommen), onderste regel(s) leeg. Alle velden eerst naar het Logboek. Nulmeting; bij verschil terugzetten.
- `VervallenRegelsOpruimen`: eenmalig, haalt bestaande V-regels zonder werkelijke bedragen weg.

## Stap 13 - kredietoverzichten (Dashboard blok 3, rij 57-61)
- `KredietoverzichtMaken` (raad): L t/m lopend jaar + A lopend jaar, zonder regels met negatief budget. Wordt ook automatisch gemaakt na elke upload.
- `KredietoverzichtOpNaam`: rol (C58) en naam (C59) op het Dashboard; incl. bijdrageregels en sectie "Nog niet gestart" (L na lopend jaar); kolom Afschr extra.
- `KredietoverzichtJaarrekening`: uit blad `Verzamellijst t/m <jaar>` (back-up die de jaarrol maakt = stand 31-12): L t/m dat jaar + A dat jaar.
- Kolommen: Volgnr, Omschrijving, Status, (Afschr), Budget t/m jaar, Totaal werkelijk, Restantkrediet, Budget +1..+5, Toelichting (breed, terugloop), Melding (uit blad Meldingen uitgaven).
- Sortering programma > investeringsgroep (I) > J (nummers eerst, dan tekst) > volgnr; groepskoppen, SUBTOTAL per niveau, eindtotaal; aansluitcontrole restantkrediet in A3.
- Bij A altijd "Krediet afgesloten." vooraan in de toelichting (alleen in de overzichten).
- M/N heten nu "Startjaar uitgaven" / "Eindjaar uitgaven" (nooit in het verleden). Jaarrol: L<oud> -> L<nieuw>, M en N +1 en minimaal het nieuwe jaar; werkelijke uitgaven schuiven (AA=AA+AB, AB=AC, AC leeg).
- Opmaak overzichten: blad 'Laatste kredietoverzicht' (raad/jaarrekening) en blad 'Laatste overzicht op naam'; vorige opvraging wordt vervangen, geen rasterlijnen, kop donkerblauw, logo rechtsboven uit 'logo*.png/jpg' in map Stambestand kredieten, geen technische kopregels (aansluiting alleen rood tonen als die niet sluit). Status en vrijval/overschrijding standaard verborgen; status en afschrijving gecentreerd; bedragen rechts, toelichting links.
- Op naam: alleen L-kredieten, kredietnummer i.p.v. volgnr; opgeslagen in 02\<naam>\. Raad: volgnr, geen kredietnummer, A lopend jaar wel tonen.
- Raadsoverzicht wordt bij elke opvraging bewaard in map 03 als "Kredietoverzicht <datum>".
- Raad- en jaarrekeningoverzicht zien er hetzelfde uit (alleen de stand verschilt). Kolom Melding wordt verborgen als er geen meldingen zijn.
- Mappen onder 'Stambestand kredieten' (map boven 'export coda', bewaard in naam `Par_MapStambestand`): 01 Applicatiebeheer Tijdschrijven (`OverzichtTijdschrijven`: lopende kredieten per kredietnummer, actie nieuw/gewijzigd/afgesloten t.o.v. nieuwste vorige bestand), 02 Budgetbeheerders en projectleiders (overzicht op naam, en `KredietoverzichtAlleBudgethouders`), 03 kredietoverzichten ("31-12-<jaar> Kredietoverzicht jaarrekening").
- Afgesloten krediet (A): restant = 0 (formule AD via `RestantFormuleBijwerken`, eenmalig; vervangt handmatige nullen). In de overzichten kolom "Vrijval (+) / overschrijding (-) bij afsluiten" = budget - werkelijk, zodat budget = werkelijk + restant + vrijval altijd sluit.
- `ToelichtingVoertuigenWissen`: eenmalig, toelichting bij voertuigen leeg (25 regels; leaseconstructie verwarrend, apart probleem).

## Ronde feedback overzichten
- Raad en jaarrekening: alleen t/m lopend jaar (geen budgetkolommen +1..+5; regels zonder budget en uitgaven t/m het jaar vallen weg). Jaarrekening zonder back-upblad: PROEF op huidige stand (niet opgeslagen).
- Kolombreedte/rijhoogte via `BladOpmaken` (geen AutoFit): kop minstens langste woord breed, bedragkolommen breder bij ####, rijhoogte via `AantalRegels` (woordterugloop, Arial 10). Geldt voor alle overzichten, uitvraag, tijdschrijven en meldingen. Workbook-standaardlettertype is Aptos Narrow 11 (breedte-eenheid ~6 px).
- Oude overzichtbladen (patroon + titel in A1) worden automatisch opgeruimd.
- Dashboard: blok 3 alleen "Kredietoverzicht op naam" (C58/C59); blok 4 (rij 64-66) voor financieel adviseur: raad, jaarrekening, uitvraag kadernota/begroting, alle budgethouders, tijdschrijven (met rechtencontrole).
- `UitvraagBegroting`: kolommen C, D, O, S (taakveld ingekort tot 30 tekens, geen terugloop), Y, AD, AE-AI, Actualisatie (keuzelijst: budget bevestigen/verhogen/verlagen, krediet vervalt, doorschuiven naar later jaar, terugschuiven naar eerder jaar (niet naar lopend jaar), omschrijving wijzigen, meerdere), Nieuwe omschrijving, AJ toelichting (leeg, in te vullen), alleen L, per budgetbeheerder groep; rij 4 verwijst naar de kolom in de verzamellijst; gele kolommen = in te vullen; bestand per budgetbeheerder in 02\<naam>; blad 'Laatste uitvraag begroting' (onbeveiligd).
- Elke publieke macro die iets wijzigt begint met `WerkStart`: beveiliging uit en via Application.OnTime direct na afloop weer aan (BeveiligingAan). Niet in TabbladenOrdenen (wordt door BeveiligingAan aangeroepen: anders een lus).
- Tabbladen Voertuigen en Controle altijd als laatste.

## Navigatie en bladnamen
- Blad heet nu 'Verzamellijst' (constante BL_VZ; `VerzamellijstHernoemen` bij openen hernoemt het oude blad 'Verzamellijst alle kredieten').
- Tabvolgorde (`TabbladenOrdenen`): Dashboard, Verzamellijst, Laatste overzicht op naam, Laatste kredietoverzicht, Laatste uitvraag begroting, Laatste overzicht tijdschrijven, Meldingen uitgaven, overige, Voertuigen, Controle.
- Op elk zichtbaar blad knop "Terug naar dashboard" (`KnoppenTerugPlaatsen`); Dashboard F2:H2 tabbladkiezer (G2 keuzelijst, knop "Ga naar tabblad"). Knoppen worden uit opgeslagen kopieen verwijderd.
- Overzicht tijdschrijven: kolommen Kredietnummer, Omschrijving, Status, Tijdschrijvers (incl. projectleider), Actie, Toegevoegd op; alleen regels met actie (nieuw/gewijzigd/afgesloten). Vergelijking met verborgen blad "Stand tijdschrijven" (alle lopende kredieten); geen stand = beginstand, alles nieuw. Blad "Laatste overzicht tijdschrijven" + bestand in map 01.

- Alert 7999998: bij elke upload ALARM op Meldingen voor L-kredieten met kredietnummer 7999998 waarvan het statusjaar <= begrotingsjaar (lopend jaar + 1).

## Actualisatie kadernota/begroting
- Uitvraagbestand: L nieuwe omschrijving (invullen), M 'Wijziging (automatisch)' = formule (LET/TEXTJOIN, Formula2, Excel 365) uit was/wordt per jaarschijf: 'Van 2027 naar 2028, bedrag verhoogd met EUR x' / 'krediet vervalt', oranje markering; N onderbouwing (invullen). Blad beveiligd (wachtwoord), alleen gele cellen open. Verborgen O 'Volgnr' (sleutel), P-T 'Was'-bedragen. AJ bij verwerken = M-tekst + '. ' + N.
- Uitvraagbestand heeft verborgen kolom O 'Volgnr' (sleutel) en onderaan 10 lege regels 'Voorstellen nieuwe investeringskredieten' (kolom O = NIEUW).
- Teamleider vult het bestand in zijn map 02\<naam> in. Financieel adviseur controleert en kiest Dashboard C58=Budgetbeheerder, C59=naam, knop "Actualisatie verwerken (C59)" (`ActualisatieVerwerken`): nieuwste uitvraagbestand, samenvatting per jaarschijf, na akkoord AE-AI per jaarschijf, D/F bij nieuwe omschrijving, AJ vervangen als toelichting is ingevuld; AV/AW gezet; geen controle administratie; lopend jaar niet geraakt; nulmeting; Logboek per wijziging; waarschuwing bij dubbel verwerken.
- Elke verwerkte wijziging per jaarschijf (en elk voorstel) komt in register 'Mutaties kadernota-begroting' (aanvullen, nooit leeg; beveiligd, filter mag): budgetbeheerder, volgnr, kredietnummer, omschrijving, taakveld, balanscode, afschr, jaar, oud, nieuw, mutatie, actualisatie, toelichting = basis lastenaanpassingen.
- Voorstellen nieuwe kredieten -> blad 'Voorstellen nieuwe kredieten' (onbeveiligd, kolom 'Verwerkt'); FA voert in via de invoerkaart.

## Bladbeveiliging
- Wachtwoord "Best" (constante WACHTWOORD in de module; VBA-project zelf ook met wachtwoord vergrendelen). `BeveiligingAan` bij elk openen (Auto_Open, UserInterfaceOnly), na upload en na elk overzicht. Verzamellijst, Logboek, Vorige upload, Parameters, Dashboard, Controle, Meldingen en overzichtbladen op slot. Open: Dashboard Kaart_Invoer, bronnen bijdragen (kolom B), C58:C59; Parameters gele tabelkolommen. Filteren mag (autofilter op de verzamellijst), sorteren niet. Geen werkmapstructuur-beveiliging.
- `BeveiligingUit` (alleen financieel adviseur) voor herstel met de hand. Tabellen op Parameters groeien niet op een beveiligd blad: nieuwe rij met de hand = eerst BeveiligingUit (zelfregistratie via intake regelt het zelf).
- Meldingen uitgaven zijn vooral voor de financieel adviseur.

## Openstaande stappen
- 12: controleblok en kengetallen (incl. regels zonder kredietnummer/activanummer, status V, SUBTOTAL-regels verplaatsen).
- 14: jaarrol (laatst, eerst op kopie).
- Later: exploitatienummers met omschrijving, mogelijk "geldig vanaf jaar" bij programma's na de fusie met Oirschot.

## Bekende valkuilen
- Excel is Nederlandstalig: voorwaardelijke opmaak via `FormulaLocal`; VBA-editor Nederlandse menu's.
- Verwijder altijd de oude module voordat `Dashboard_AlleMacros.bas` wordt geïmporteerd; daarna `DashboardAanmaken` (Ja bij opnieuw opbouwen).
- Testen met verzonnen kredieten en `ProefregelWissen`; nooit op echte regels.
