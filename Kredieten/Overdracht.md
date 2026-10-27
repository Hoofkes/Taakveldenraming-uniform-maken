# Overdracht — Dashboard investeringskredieten (stand 27-10-2026)

Dit bestand is bedoeld om het werk in een nieuwe sessie op te pakken zonder alles opnieuw uit te zoeken. Lees daarnaast altijd `Functioneel_ontwerp_v4.md` (§ 0.1 afspraken, § 11 uitvraag, § 12 stappenplan, § 13 indeling blok 2).

## 1. Bestanden (branch `claude/admiring-hypatia-3ulxex`, map `Kredieten/`)

| Bestand | Inhoud |
| --- | --- |
| `Functioneel_ontwerp_v4.md` | Leidend ontwerp. Afspraken in § 0.1 gaan voor oudere tekst. |
| `Kredieten_1_Basis.bas` | Module 1: alle declaraties (Public Const, Public variabelen, Public Type TKaart), rollen, logboek, normaliseren, werkbereik, controlekolommen, Dashboard, krediet toevoegen en wijzigen, controle financiële administratie. |
| `Kredieten_2_Verwerking.bas` | Module 2: tijdschrijvers, bijdragen, afgekeurd/hercontrole, krediet verwijderen, werkelijke uitgaven, kredietoverzichten. Bovenaan eigen `Private mBlok`. |
| `Kredieten_3_Overzichten.bas` | Module 3: uitvraag kadernota/begroting, tabbladen, controles en kengetallen, naam vervangen, jaarovergang, knopuitleg, pdf en mail. |
| `Dashboard_AlleMacros_41.bas` | Oude versie 41 als één module. **Niet meer gebruiken** (te groot, zie § 3). |

Bron van alles: `Dashboard_AlleMacros_40.bas` (geüpload door de gebruiker, werkte goed).

## 2. Stand van de stappen (§ 12 van het functioneel ontwerp)

| Stap | Versie | Onderwerp | Stand |
| --- | --- | --- | --- |
| 1 | 41 | Kleine fouten zonder datarisico | **Klaar en getest**: compileren en opslaan goed. |
| 2 | 42 | Werkelijke uitgaven veilig maken | **Gebouwd, nog niet bevestigd getest.** Na de eerste test gevraagd: rijnummer en omschrijving in "Meldingen uitgaven" (toegevoegd). Open vraag: ook volgnummer (kolom A) tonen? |
| 3 | 43 | Terugdraaien bij een fout + jaarovergang M en N | Te doen |
| 4 | 44 | Beveiliging back-up- en voorstelbladen, proefregel wissen | Te doen |
| 5 | 45 | Kolom J als getal, A-lopend jaar in overzicht op naam | Te doen |
| 6 | 46 | Blok 2 indelen (2a, 2b), overzicht op naam via keuzevragen | Te doen |
| 7 | 47 | Uitvraag kadernota/begroting (2c, § 11) | Te doen |
| 8 | 48 | Dynamische bereiken (§ 2.3), apart uitvoeren | Te doen |
| 9 | 49 | Jaarkoppen, controles § 8.2 nummer 1, 2 en 6, waarschuwing werkbereik | Te doen |

## 3. Harde lessen (niet opnieuw tegenaan lopen)

1. **Groottegrens VBA-module.** Eén module van ruim 435.000 tekens gaf bij opslaan "Er zijn fouten opgetreden tijdens het opslaan van …" (ook lokaal op `C:\Temp`). Elke helft van de wijzigingen apart ging goed, samen niet. Daarom drie modules. Houd elke module ruim onder ±150.000 tekens; splits verder als een module groeit.
2. **Regellengte.** Geen regel langer dan 1.023 tekens. De eerste bas41 had een kopregel van 1.124 tekens.
3. **Na het splitsen** zijn 49 hulpfuncties die tussen modules worden gebruikt van `Private` naar `Public` gezet. Nieuwe hulpfuncties die in een andere module worden aangeroepen: meteen `Public`. Module-brede variabelen horen in module 1 als `Public`.
4. **Werkwijze bij de gebruiker:** kopie maken, de drie oude modules verwijderen (niet exporteren), alle drie de nieuwe importeren, pas dan compileren, dan opslaan. Compileren na één module geeft "Sub of Function is niet gedefinieerd".
5. **Na een mislukte opslag** op SharePoint vraagt Office om "wijzigingen samen te voegen": niet samenvoegen.
6. VBA kan in deze omgeving niet worden gecompileerd of uitgevoerd. Controleer zelf: aantal Sub/Function gelijk aan End Sub/End Function, regellengte, en verwijzingen tussen modules.

## 4. Gebruiker en communicatie

- Nederlands, op het niveau van een 16-jarige, **geen afkortingen**, kort, alleen actiepunten.
- Bedragen: minteken, punten tussen duizendtallen, nooit haakjes.
- Per stap eerst een omschrijving geven, gebruiker bevestigt, dan bouwen.
- Gebruiker wil dat het werkende bestand niet weer "in de soep draait": kleine stappen, eerlijk melden als iets onzeker is.
- **Eerst bas40 zelf nalezen voordat je iets vraagt.** Veel staat al goed in de code.
- Functioneel ontwerp na elke afspraak direct bijwerken.

## 5. Inhoud van de volgende stappen (bevindingen uit de controle van bas40)

Regelnummers zijn bij benadering; zoek op de naam van de procedure.

### Stap 3 (bas43) — terugdraaien bij een fout
- `KredietToevoegen` en `WijzigingOpslaan` (module 1): bij een fout na het schrijven niets teruggedraaid. Vlag "geschreven" bijhouden; in `Fout:` de regel leegmaken (`RuimRij`) of oude waarden terugzetten. `pc`/`tc` (programma, taakveld) vóór het schrijven valideren (`CLng(pc)` kan falen).
- `WijzigingOpslaan`: controleert alleen dat de rij in `Kaart_Rijnummer` niet leeg is. Bij `ToonOpKaart` ook het volgnummer vastleggen en vergelijken met kolom A (zoals `WerkRij` al doet).
- `RegelsWeghalen` (krediet verwijderen, module 2): geen eigen foutafhandeling; half verschoven lijst bij een fout. Logregels "verwijderen (helemaal weg)" pas na een geslaagde nameting schrijven.
- `RestantFormuleBijwerken` (module 2): Calculation herstellen bij een fout; oude AD-formules bewaren en bij een verschil terugzetten.
- `JaarovergangUitvoeren` (module 3): vlag `geschreven` vóór het wegschrijven; in `Fout:` alle kolommen en `Par_LopendJaar` terugzetten en de back-up verwijderen. `KeuzelijstenVerversen` na de jaarovergang apart afvangen.
- **Jaarovergang kolom M en N volgens § 2.7 punt 7**: `L<oud jaar>` wordt `L<nieuw jaar>`; staat er vóór de jaarovergang al een bedrag in Budget lopend jaar + 1 (kolom AE), dan M en N ongewijzigd; anders wordt een waarde gelijk aan het oude jaar in M of N het nieuwe jaar, andere jaartallen blijven staan. Nu doet de code `Max(oud + 1, nieuw jaar)` zonder rekening te houden met AE.
- De uitvraag valt buiten stap 3 (wordt vervangen in stap 7).

### Stap 4 (bas44) — beveiliging
- `BeveiligingTijdelijkUit` haalt de beveiliging van alle bladen af, `BeveiligingAan` zet een vaste lijst terug. Ook `Verzamellijst 31-12-*` en "Voorstellen nieuwe kredieten" weer beveiligen.
- `ProefregelWissen` (module 1): alleen toestaan voor een regel van vandaag (kolom AW/AU) zonder werkelijke uitgaven (AA tot en met AC), inhoud loggen, nulmeting; vaste grens 271 weghalen.

### Stap 5 (bas45) — kolom J en overzicht op naam
- Een groepsnummer in J dat als tekst is opgeslagen ("7") wordt nu niet als groep herkend (`IsGetal`). Hulpfunctie `JGetal` (Trim, IsNumeric, CDbl) maken en overal gebruiken: `SorteerJ`, `GroepKoppen`, groepsbepaling in `UitgavenBijwerken`, `OverzichtTijdschrijven`, `GroepenLijstSchrijven`.
- `OverzichtMaken` met soort "naam": ook `A` + lopend jaar tonen (afspraak § 0.1 punt 6).

### Stap 6 en 7 — zie § 13 en § 11 van het functioneel ontwerp
- Stap 6: blok 2 in subonderdelen 2a (kredietoverzichten) en 2b (omzetting en controle, met "Verwerken werkelijk"); rij 38 en 39 (rol, naam) vervallen; overzicht op naam via keuzevragen zoals bij "Budgethouder vervangen". Oude uitvraagknoppen blijven in stap 6 nog werken. Let op: blok 4 (controles) zichtbaar voor financieel adviseur én financiële administratie (zo is het nu al).
- Stap 7: drie knoppen "1. Uitvraag versturen", "2. Aanleveringen ophalen" (per budgethouder kiezen, aanpassingen van adviseurs niet overschrijven, signaal bij gewijzigde aanlevering), "3. Uitvraag verwerken" (direct status "Te controleren door financiële administratie"; nieuwe kredieten naar "Voorstellen nieuwe kredieten"). Eén blauw tabblad "Uitvraag". Lopend jaar vastleggen; ophalen of verwerken in een ander jaar weigeren; jaarovergang blokkeren bij onverwerkte regels. Daarna oude knoppen, de bladen "Stand uitvraag" en "Verwerking uitvraag", en `ActualisatieVerwerken` (ongebruikt) weghalen. `UitvraagNaarControle` markeert nu ook niet-gevonden volgnummers als aangeboden: vervalt met de knop.

### Stap 8 en 9
- Stap 8: namen `Krediet_Aantal`, `Krediet_J`, `Krediet_AL`, `Krediet_Data` en de formules in K, AL, AM en AN omzetten. Apart uitvoeren met nulmeting en volledige terugzetting.
- Stap 9: eerst in de werkmap nagaan of de jaarkoppen al formules zijn. Controles § 8.2 nummer 1, 2 en 6 bestaan nog niet (het commentaar bij `ControlesVernieuwen` zegt ten onrechte dat de spiegelbladen gecontroleerd worden). Kengetal aantal status O vervalt (status O bestaat niet meer).

## 6. Bewust zo gelaten (niet "herstellen")

Krediet verwijderen via eigen knop; tot en met 31 maart ook vorig jaar (AB) bijwerken; raadsoverzicht zonder budget +1 tot en met +5; overzicht op naam met investeringsnummer en sectie "Nog niet gestart"; vaste bladnamen "Laatste …" plus opgeslagen bestand met datum; alleen status L en A; financiële administratie werkt rechtstreeks in de Verzamellijst; nieuwe gebruiker mag zichzelf toevoegen met direct de rechten van de rol; filteren wel, sorteren niet op de Verzamellijst. Zie § 0.1 van het functioneel ontwerp.

## 7. Nog open

- Uitkomst test bas42 (stap 2) en de vraag of het volgnummer ook in "Meldingen uitgaven" moet.
- Een aparte uitvraag voor het lopende jaar: later besluiten (§ 11.5).
- Verzamelkredieten in Coda: later ontwerpen (§ 6.3).
