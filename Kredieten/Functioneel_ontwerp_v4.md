# Functioneel ontwerp — Dashboard investeringskredieten (versie 4)

**Werkmap:** `Stambestand_investeringskredieten.xlsm`
**Doel:** één tabblad `Dashboard` met invoer, wijziging, rolgestuurde controle en uitdraaien, plus een volledig jaaronafhankelijke (dynamische) opzet van de verzamellijst.
**Status:** specificatie ter uitwerking in Claude Code. Versie 2 verwerkte de beslissingen op de acht openstaande punten uit versie 1; versie 3 verwerkt kolom J als extra sorteerniveau en legt één nieuw beslispunt voor (§ 9.1). Versie 4 legt de afspraken vast uit de controle van macroversie 40 (zie § 0.1); afspraken die afwijken van eerdere tekst gaan voor.
**Taal:** alles Nederlands. Bladnamen en celinhoud nooit vertalen. VBA-editor is Nederlands (Invoegen > Module, Foutopsporing > VBAProject compileren, map Microsoft Excel-objecten); alleen vaste objectnamen zoals `ThisWorkbook` blijven Engels.

---

## 0. Beslissingen uit versie 1 (samengevat)

| # | Punt | Besluit |
| --- | --- | --- |
| 1 | Omschrijving in het overzicht | **Kolom D in plaats van C. Kolom C vervalt in het overzicht.** |
| 2 | Groepering | **Presenteren op B (programma) → I (naam groep kredieten) → J (extra sorteerniveau).** Bevat J een **getal**, dan werkt dat getal als extra sorteervolgorde waardoor de deelkredieten die bij elkaar horen bij elkaar komen te staan. B, I en J sturen de presentatie maar staan **niet zichtbaar** in het overzicht. Doel: een lijst die direct bruikbaar is voor de raad én voor budgethouders. |
| — | Hoofdletters in J | Alles moet kleine letter zijn. Afwijking aangetroffen in **J155:J162** (`Stelpost`), acht stelpostregels riolering. Normaliseren (§ 2.5). |
| 3 | Selectie op status (H) | `L<jaar>` = er volgen nog uitgaven. `A<jaar>` = afgerond, raad heeft conform besloten. **Kredietoverzicht raad:** alle `L` t/m lopend jaar **+** `A` van uitsluitend het lopend jaar. **Kredieten voor een toekomstig jaar niet presenteren** in het raadsoverzicht, **wél** meenemen in het overzicht voor budgethouders. M en N zijn de geschatte doorlooptijd (in welke periode het geld volgens de actuele planning wordt uitgegeven) en sturen de selectie niet. |
| 4 | Controlestatus bestaande regels | Alle bestaande kredieten krijgen `Gecontroleerd door financiële administratie`. Elke regel die vanaf nu wordt toegevoegd krijgt `Te controleren door financiële administratie`, en moet **per regel** door de financiële administratie worden bevestigd. |
| 5 | Gebruiker en rollen | Intakescherm bij openen: gebruiker vult/bevestigt zijn naam, daarna gelden zijn bevoegdheden. Rollentabel is zelf wijzigbaar. Financieel adviseurs (Michiel, Emile, Joost, Monique, Marianne): toevoegen **en** wijzigen. Financiële administratie (Cor, Mariska, Sigrid): wijzigen **en** bevestigen. |
| 6 | Beveiliging | Voorlopig **niet** inbouwen; later mogelijk. |
| 7 | Dynamische bereiken | **Alles dynamisch maken.** Rekening houden met veel extra kredieten na de fusie met Oirschot. |
| 8 | Werkelijke uitgaven | Gebruiker uploadt een bestand van de financiële administratie; de macro werkt daarmee de uitgaven van het **lopend jaar** bij. Kenmerken van dat bestand worden later vastgelegd. Nu al wél geborgd: de uitgaven moeten aan het **juiste investeringskrediet** worden gekoppeld. |
| 9 | Verzamelkredieten in Coda | Coda werkt vaak met verzamelkredieten waarvan de werkelijke uitgaven per 31-12 worden doorgeboekt. **Wordt later apart ontworpen**, nu uitsluitend als aandachtspunt vastgelegd. |

## 0.1 Afspraken uit de controle van macroversie 40 (oktober 2026)

| # | Onderwerp | Afspraak | Gevolg voor de macro's |
| --- | --- | --- | --- |
| 1 | Krediet verwijderen | Een krediet kan **direct** worden verwijderd via de eigen knop **Krediet verwijderen** op het Dashboard, niet via eerst wijzigen. De regels eronder schuiven één regel op (waarden verplaatsen, geen rijen verwijderen in Excel). Vervangt § 3.5 slotalinea. | Werking blijft. Wel: bij een fout tijdens het verschuiven alles terugzetten, en de logregels pas na een geslaagde controle wegschrijven. |
| 2 | Werkelijke uitgaven rond de jaarrekening | **Tot en met 31 maart** van een jaar worden zowel de jaarschijf van het lopende jaar (AC) als die van het voorgaande jaar (AB) bijgewerkt, omdat de jaarrekening dan nog wordt afgesloten. **Na 31 maart** alleen het lopende jaar (AC). Dit geldt elk jaar. Vervangt § 6.2 punt 4. | Werking blijft. Wel: de harde controle (§ 6.2 punt 5) geldt voor **beide jaren samen**. Sluit één jaar niet aan, dan wordt niets weggeschreven, ook niet in het andere jaar. |
| 3 | Krediet niet meer in het uitgavenbestand | Stond er bij de vorige keer inlezen een bedrag bij een krediet en komt dat krediet niet meer voor in het nieuwe bestand, dan wordt dat **verwerkt**: het bedrag wordt 0, met een melding in `Meldingen uitgaven`. Een krediet zonder eerder bedrag dat niet in het bestand staat, is geen probleem. | Aanpassen: nu blijft het oude bedrag staan met een notitie. |
| 4 | Jaarovergang: status en kolom M en N | Zie § 2.7 punt 7 (nieuw). Vervangt de eerdere tekst "status in kolom H niet aanraken". | Aanpassen: nu wordt M en N +1 met minimaal het nieuwe jaar, en wordt geen rekening gehouden met een bedrag in het eerste toekomstige jaar. |
| 5 | Kredietoverzicht voor de raad | Alleen het lopende jaar: geen kolommen Budget +1 tot en met +5. Vervangt § 7.2 regel 6. | Geen: bas40 doet dit al zo. |
| 6 | Kredietoverzicht op naam | Eerste kolom is het investeringsnummer (niet het volgnummer). Lopende kredieten met een toekomstig jaar staan in een aparte sectie "Nog niet gestart". Vervangt § 7.3 punt 2 en 3 voor zover afwijkend. | Geen: bas40 doet dit al zo. |
| 7 | Naam van de overzichten | Het overzicht in de werkmap heet steeds "Laatste kredietoverzicht" of "Laatste overzicht op naam" en wordt bij een nieuw overzicht vervangen. Elk overzicht wordt daarnaast als apart bestand opgeslagen met naam en datum in de bestandsnaam, in de map van het stambestand. Vervangt § 7 eerste alinea (naamgeving). | Geen: bas40 doet dit al zo. |
| 8 | Statusletters | Alleen `L` (lopend) en `A` (afgesloten). Status `O` wordt niet gebruikt. Vervangt `L`/`A`/`O` in § 1.2, § 3.3, § 3.4 en § 8. | Geen: bas40 doet dit al zo. |
| 9 | Controle door de financiële administratie | De financiële administratie opent de regel direct in de Verzamellijst (niet in de invoerkaart). Alleen de invoerkolommen gaan open; werkelijke uitgaven (AA tot en met AC) en formulekolommen blijven op slot. De regel wordt vooraf bewaard op het verborgen blad "Werkregel controle"; elke wijziging komt in het Logboek. Afsluiten met fouten kan alleen na een extra bevestiging. Vervangt § 3.5 en § 4.3 voor zover afwijkend. | Geen: bas40 doet dit al zo. |
| 10 | Nieuwe gebruiker | Iemand die niet in de rollentabel staat, mag zichzelf toevoegen als financieel adviseur of financiële administratie en krijgt direct de rechten van die rol. Dit wordt in het Logboek vastgelegd. Vervangt § 5.1 punt 4. | Geen: bas40 doet dit al zo. |
| 11 | Bladbeveiliging | De beveiliging staat aan (vervangt § 0 punt 6 "voorlopig niet"). Filteren mag op de Verzamellijst, sorteren niet, omdat sorteren de koppeling met de spiegelbladen breekt. Macro's werken op beveiligde bladen doordat de beveiliging "alleen voor de gebruiker" geldt. Vervangt § 10 slotalinea voor zover afwijkend. | Geen: bas40 doet dit al zo. Wel herstellen: na een macro worden ook de back-up "Verzamellijst 31-12-<jaar>", "Verwerking uitvraag" en "Voorstellen nieuwe kredieten" weer beveiligd. |
| 12 | Dashboard per rol | Financieel adviseur ziet blok 1 (invoerkaart), blok 2 (financieel adviseur) en blok 4 (controles). Financiële administratie ziet alleen blok 3 (financiële administratie). Applicatiebeheerder ziet alles. | Aanpassen: in bas40 is blok 4 voor iedereen zichtbaar. |
| 13 | Aanpassen door de financiële administratie | Rechtstreeks in de Verzamellijst, niet via de invoerkaart (zie punt 9). | Geen: bas40 doet dit al zo. |

---

## 1. Huidige situatie (vastgesteld in de werkmap, peildatum 22-10-2026)

### 1.1 Tabbladen

| Blad | Rol |
| --- | --- |
| `Verzamellijst alle kredieten` | **Enige bron.** Kop in rij 1, data vanaf rij 2. Laatste gevulde rij = 270 (volgnr 301). **Eerste vrije rij = 271, nieuw volgnummer = 302.** |
| `Financieel adviseur` | Spiegel, rij-voor-rij: `=@IF(ISBLANK(@INDEX('Verzamellijst alle kredieten'!$A:$A;ROW()));"";INDEX(...;ROW()))`. Kop in rij 1 → **rij n = rij n van de bron**. Formules lopen t/m rij 1500. |
| `Financiële administratie` | Idem spiegel, zelfde `ROW()`-logica, t/m rij 1500. |
| `Kredietoverzicht` | Spiegel met `ROW()-1` (titel rij 1, kop rij 2), t/m rij 1501, plus eigen hulpkolommen N t/m T. |
| `Voertuigen` | Tabel `Tabel33` (A2:T29), los vervangingsschema. |
| `Activa bestaand oud en nieuw` | Hulptabel. |
| `Maatwerk kredieten 23-09-2026` | Bevroren snapshot, harde bedragen, titel rij 1, toelichting rij 2, kop rij 3. |

### 1.2 Kolomindeling `Verzamellijst alle kredieten`

| Kol | Kop | Soort | Rol in het nieuwe dashboard |
| --- | --- | --- | --- |
| A | Volgnr | automatisch | sleutel, zichtbaar in overzichten |
| B | Programma | invoer | **presentatieniveau 1** |
| C | Investering boekhouding | invoer | zoekveld; **niet** in de overzichten |
| D | Omschrijving investeringskrediet | invoer, max 36 | zichtbaar in overzichten |
| E | Lengte lang (36 posities) | formule | controle |
| F | Investeringskrediet kort (20 posities) | invoer, max 20 | reserve |
| G | Lengte kort | formule | controle |
| H | Status (onderweg, lopend of afgesloten) | invoer, `L`/`A`/`O` + jaartal | **selectie** |
| I | Soort investering | invoer | **presentatieniveau 2 — naam van de groep kredieten** |
| J | Samengesteld, standaard, stelpost | invoer | **presentatieniveau 3 — extra sorteerniveau; een getal houdt deelkredieten bij elkaar** |
| K | Verdeelsleutel in % | formule met `$AL`, `$AM`, `$AN` | verdeling samengestelde kredieten |
| L | Laatste besluit | invoer | |
| M | Jaar eerste raming start krediet | invoer | geschatte doorlooptijd (start) |
| N | (Actueel) krediet gereed | invoer | geschatte doorlooptijd (eind) |
| O | Budgetbeheerder | invoer | selectie naamoverzicht |
| P | Projectleider (tevens tijdschrijver) | invoer | selectie naamoverzicht |
| Q | Tijdschrijvers | invoer | selectie naamoverzicht |
| R | Exploitatie boekhouding | invoer | |
| S | Taakveld Bbv | invoer | |
| T | Balanscode Bbv | invoer | |
| U | Organisatie-onderdeel | invoer | |
| V | Budgetbeheerder administratie | invoer | |
| W | Code btw/bcf administratie | invoer | |
| X | Activanr | invoer | zoekveld |
| Y | Afschr (termijn) | invoer | zichtbaar in naamoverzicht |
| Z | Budget beschikbaar tot en met lopend jaar | invoer | zichtbaar |
| AA | Werkelijke uitgaven tot en met lopend jaar − 2 | **alleen via importknop** | opgeteld zichtbaar |
| AB | Werkelijke uitgaven in lopend jaar − 1 | **alleen via importknop** | opgeteld zichtbaar |
| AC | Werkelijke uitgaven in lopend jaar | **alleen via importknop** | opgeteld zichtbaar |
| AD | Nog uit te geven t/m lopend jaar | formule `=Z−AA−AB−AC` | zichtbaar als **Restantkrediet** |
| AE–AI | Budget +1 t/m +5 | invoer | zichtbaar |
| AJ | Toelichting op het investeringskrediet | invoer | zichtbaar, laatste kolom |
| AK | *(leeg, scheiding)* | — | |
| AL | Hulp: positief bedrag regel (Z + AE t/m AI) | formule | |
| AM | Hulp: aandeel binnen nummer (afgerond 0,1%) | formule | |
| AN | Hulp: sluitregel binnen nummer (1 = ja) | formule | |
| **AP–AW** | **Nieuw: controle- en rolvelden** | zie § 5 | |

### 1.3 Harde randvoorwaarden

1. **Nooit rijen tussenvoegen of verwijderen in de verzamellijst.** `Financieel adviseur`, `Financiële administratie` en `Kredietoverzicht` zijn rij-voor-rij gekoppeld via `ROW()`. Nieuwe regels altijd op de eerste vrije rij onderaan.
2. **Nooit kolommen tussenvoegen in A:AJ.** De spiegelbladen verwijzen op kolomletter. Nieuwe velden komen **rechts van AN**, vanaf AP.
3. **Bedragen altijd zwart**, negatief met minteken, nul als `-`, notatie `#,##0;-#,##0;-`. Geen blauwe invoerkleur of andere kleurcodering op cijfers.
4. Cellen verticaal boven uitlijnen; tekstkolommen met terugloop en rijhoogte autofit.
5. Bij elke macro die naar de verzamellijst schrijft: **nulmeting vooraf** (aantal gevulde regels in kolom A plus de sommen van Z, AA, AB, AC en AE t/m AI), na afloop vergelijken. Bij een onverwacht verschil: terugdraaien en melden, niet doorgaan.

---

## 2. Dynamische (jaaronafhankelijke) opzet

### 2.1 Principe jaarkolommen

De kolomstructuur blijft **altijd identiek**; alleen de jaartallen in de koppen en de inhoud schuiven op:

- `Z` = budget beschikbaar **tot en met lopend jaar**
- `AA` = werkelijke uitgaven **tot en met lopend jaar − 2**
- `AB` = werkelijke uitgaven **in lopend jaar − 1**
- `AC` = werkelijke uitgaven **in lopend jaar**
- `AD` = nog uit te geven **t/m lopend jaar** (`= Z − AA − AB − AC`)
- `AE` t/m `AI` = budget **lopend jaar + 1** t/m **lopend jaar + 5**

Bij lopend jaar 2026 geeft dat de huidige situatie (AA = t/m 2024, AE..AI = 2027..2031). Bij lopend jaar 2027: AA = t/m 2025, AB = 2026, AC = 2027, AE..AI = 2028..2032.

### 2.2 Centrale parameters

Nieuw tabblad **`Parameters`**, nooit verbergen.

| Cel | Naam | Inhoud |
| --- | --- | --- |
| B2 | `Par_LopendJaar` | 2026 — **hard getal**, gezet door de jaarrolmacro. Géén `YEAR(TODAY())`: anders rolt de lijst om zonder dat de bedragen zijn verschoven. |
| B3 | `Par_LaatsteRij` | laatste gevulde rij in kolom A van de verzamellijst (dynamisch, § 2.3) |
| B4 | `Par_EersteVrijeRij` | `=Par_LaatsteRij+1` |
| B5 | `Par_VolgendVolgnr` | `=MAX(Krediet_Volgnr)+1` |
| B6 | `Par_PadUitgavenbestand` | pad/bestandsnaam van het laatst ingelezen bestand van de financiële administratie |
| B7 | `Par_PeildatumUitgaven` | datum laatste bijwerking werkelijke uitgaven (hard) |
| B8 | `Par_HuidigeGebruiker` | naam van de gebruiker die nu werkt (§ 5.1) |
| B9 | `Par_HuidigeRol` | afgeleide rol (§ 5.1) |
| B10 | `Par_MaxRij` | hoogste rij waarin formules en spiegelbladen zijn doorgetrokken (nu 1500) |

### 2.3 Dynamische bereiken (vervangt de harde grens 1468)

Benoemde bereiken, allemaal gestuurd door `COUNTA` op kolom A van de verzamellijst:

```
Krediet_Aantal  = COUNTA('Verzamellijst alle kredieten'!$A$2:$A$1048576)
Krediet_Volgnr  = OFFSET('Verzamellijst alle kredieten'!$A$2;0;0;Krediet_Aantal;1)
Krediet_J       = OFFSET('Verzamellijst alle kredieten'!$J$2;0;0;Krediet_Aantal;1)
Krediet_AL      = OFFSET('Verzamellijst alle kredieten'!$AL$2;0;0;Krediet_Aantal;1)
Krediet_Data    = OFFSET('Verzamellijst alle kredieten'!$A$2;0;0;Krediet_Aantal;48)
```

De vaste bereiken `$J$2:$J$1468` en `$AL$2:$AL$1468` in de kolommen K, AL, AM en AN worden vervangen door `Krediet_J` respectievelijk `Krediet_AL`.

**Belangrijk bij het doorvoeren:** dit raakt in één bewerking alle 270 regels in vier kolommen. Daarom:
1. nulmeting vooraf (aantal getalcellen én som per kolom Z, AA, AB, AC, AE t/m AI, plus de som van K per investeringsnummer);
2. formules vervangen;
3. nameting en vergelijking; bij **elk** verschil volledig terugdraaien;
4. apart uitvoeren, niet tegelijk met het bouwen van het dashboard, zodat een eventueel verschil herleidbaar is.

### 2.4 Groei na de fusie met Oirschot

Het datablok loopt nu t/m rij 270; de spiegelbladen hebben formules t/m rij 1500 (`Kredietoverzicht` t/m 1501). Bij een grote instroom aan kredieten is die 1500 de eerste echte grens.

Macro **Werkbereik verruimen** (knop op het Dashboard):
1. vraagt tot welke rij moet worden doorgetrokken (standaard `Par_LaatsteRij` + 1.000);
2. trekt op `Financieel adviseur`, `Financiële administratie` en `Kredietoverzicht` de bestaande spiegelformules door tot die rij, door de laatste gevulde formuleregel te **kopiëren** (nooit opnieuw intypen);
3. trekt op de verzamellijst de formulekolommen E, G, K, AD, AL, AM en AN door tot dezelfde rij, eveneens door kopiëren;
4. werkt `Par_MaxRij` bij;
5. controleert daarna rij-voor-rij of de drie spiegels nog aansluiten (§ 8.2, controle 1 en 2).

Controle op het Dashboard: waarschuwing zodra `Par_LaatsteRij` binnen 100 rijen van `Par_MaxRij` komt.

### 2.5 Normalisatie van lijstwaarden

Kolom J bevat nu `Stelpost` met hoofdletter in **J155:J162** (acht stelpostregels riolering). Alles moet kleine letter zijn.

Macro **Lijstwaarden normaliseren**:
- zet kolommen I en J volledig naar kleine letters (`LCase`), met nulmeting op het aantal gevulde cellen vooraf en achteraf;
- raakt uitsluitend tekstwaarden; de getallen in J (groepsnummers van deelkredieten) blijven ongemoeid, zodat `ISNUMBER($J2)` in kolom K blijft werken;
- rapporteert welke cellen zijn aangepast en schrijft een logregel.

Daarna gegevensvalidatie op I en J, gevoed uit `Parameters`, zodat er geen nieuwe hoofdletters bijkomen.

### 2.6 Koppen als formule

Jaargebonden koppen in de verzamellijst én in de drie spiegelbladen worden formules op `Par_LopendJaar`:

```
Z1   ="Budget beschikbaar tot en met "&Par_LopendJaar
AA1  ="Werkelijke uitgaven tot en met "&(Par_LopendJaar-2)
AB1  ="Werkelijke uitgaven in "&(Par_LopendJaar-1)
AC1  ="Werkelijke uitgaven in "&Par_LopendJaar
AD1  ="Nog uit te geven t/m "&Par_LopendJaar
AE1  ="Budget "&(Par_LopendJaar+1)   ... t/m AI1 = +5
```

De bestaande notitie op `AC1` (bronvermelding van de Coda-uitdraai) moet behouden blijven of opnieuw worden gezet.

### 2.7 Macro `Jaarrol` — knop **Jaarovergang uitvoeren**

1. **Bevestiging vragen** met het expliciete nieuwe jaar: "Jaarovergang van 2026 naar 2027 uitvoeren? Dit verschuift alle bedragkolommen."
2. **Nulmeting + back-up:** kopie van de verzamellijst als `Verzamellijst t/m <oud jaar>`, inclusief de sommen van Z, AA:AC en AE:AI.
3. **Verschuiven per regel, harde waarden, in één bewerking:**
   - `Z_nieuw  = Z_oud + AE_oud`
   - `AA_nieuw = AA_oud + AB_oud`
   - `AB_nieuw = AC_oud`
   - `AC_nieuw = leeg`
   - `AE_nieuw = AF_oud`, `AF_nieuw = AG_oud`, `AG_nieuw = AH_oud`, `AH_nieuw = AI_oud`, `AI_nieuw = leeg`
   - `AD` blijft formule en herrekent zichzelf.
4. `Par_LopendJaar` ophogen met 1; koppen volgen automatisch (§ 2.6).
5. **Controle:** `som(Z_nieuw) = som(Z_oud) + som(AE_oud)`, `som(AA:AC_nieuw) = som(AA:AC_oud)`, `som(AE:AI_nieuw) = som(AF:AI_oud)`. Bij afwijking alles terugzetten uit de back-up en melden.
6. Logregel wegschrijven.
7. **Status en doorlooptijd (afspraak versie 4).** Voorbeeld bij de overgang van 2026 naar 2027. Vóór de jaarovergang hebben alle lopende kredieten minimaal `L2026`; andere jaartallen komen alleen voor bij `A`.
   - `L2026` wordt `L2027`.
     - Staat er vóór de jaarovergang al een bedrag in **Budget 2027** (kolom AE), dan blijven M en N ongewijzigd. Dit komt voor bij een aanvullend krediet dat al in 2027 is geraamd (bijvoorbeeld een brug).
     - Staat er geen bedrag in Budget 2027, dan wordt een **2026 in kolom M of N een 2027**: door de latere uitvoering starten de uitgaven in 2027. Andere jaartallen in M of N blijven staan.
   - `L2027` of later: status, M en N ongewijzigd.
   - `A` (afgesloten): status, M en N ongewijzigd.
   - Het jaar van het besluit staat in kolom L (Laatste besluit) en verandert niet.
8. **Niet aanraken bij de jaarrol:** de controlestatus in AP, en de spiegelbladen (die volgen vanzelf).

---

## 3. Dashboardblok 1 — Invoerkaart nieuw krediet

### 3.1 Vorm: verticale mutatiekaart

Label in kolom B, invoerveld in kolom C, controlemelding in kolom D. **Eén veld per regel**, nadrukkelijk niet horizontaal zoals in de verzamellijst.

```
B                     C                      D
Programma             [keuzelijst]
Investering boekh.    [getal]                bestaat al: nee
Omschrijving lang     [tekst]                28/36 — goed
Omschrijving kort     [tekst]                17/20 — goed
Status                [L/A/O] [jaar]
Soort investering     [keuzelijst]
...
```

Velden: `Programma (B)`, `Investering boekhouding (C)`, `Omschrijving lang (D)`, `Omschrijving kort (F)`, `Status (H)`, `Soort investering (I)`, `Samengesteld/standaard/stelpost (J)`, `Laatste besluit (L)`, `Jaar eerste raming (M)`, `Krediet gereed (N)`, `Budgetbeheerder (O)`, `Projectleider (P)`, `Tijdschrijvers (Q)`, `Exploitatie boekhouding (R)`, `Taakveld Bbv (S)`, `Balanscode Bbv (T)`, `Organisatie-onderdeel (U)`, `Budgetbeheerder administratie (V)`, `Code btw/bcf (W)`, `Activanr (X)`, `Afschr (Y)`, `Budget t/m lopend jaar (Z)`, `Budget +1 t/m +5 (AE..AI)`, `Toelichting (AJ)`.

De koppen van de budgetvelden op de kaart zijn formules op `Par_LopendJaar`, zodat de kaart met het jaar meegaat.

### 3.2 Knop **Krediet toevoegen**

Beschikbaar voor de rol **financieel adviseur** (§ 5).

1. Controles uit § 3.4 uitvoeren. Bij fouten: niets wegschrijven, meldingen in kolom D, afbreken.
2. Rij bepalen via `Par_EersteVrijeRij` (nu 271), volgnummer via `Par_VolgendVolgnr` (nu 302).
3. Formulekolommen `E`, `G`, `K`, `AD`, `AL`, `AM`, `AN` **kopiëren van de regel erboven**.
4. Invoerwaarden wegschrijven in B, C, D, F, H, I, J, L t/m Y, Z, AE t/m AI, AJ.
5. `AA`, `AB`, `AC` leeglaten — uitsluitend via § 6.
6. Controlevelden zetten: `AP = "Te controleren door financiële administratie"`, `AV = Par_HuidigeGebruiker`, `AW = datum/tijd` (hard).
7. Opmaak gelijktrekken met de regel erboven.
8. Logregel wegschrijven, kaart wissen, melden: "Krediet toegevoegd op rij 271, volgnummer 302."

### 3.3 Keuzelijsten

Benoemde bereiken op `Parameters`, gevuld met de unieke bestaande waarden (alles kleine letter):

`Lst_Programma` (B) · `Lst_Soort` (I) · `Lst_SamengesteldStandaardStelpost` (J) · `Lst_Budgetbeheerder` (O) · `Lst_Projectleider` (P) · `Lst_Taakveld` (S) · `Lst_Balanscode` (T) · `Lst_Organisatieonderdeel` (U) · `Lst_BudgetbeheerderAdm` (V) · `Lst_BtwBcf` (W) · `Lst_Statusletter` (L / A / O).

Status wordt in twee velden ingevoerd (letter + jaartal) en samengesteld weggeschreven, zodat `L 2026` nooit als `l2026` of `L26` in de lijst belandt.

Knop **Keuzelijsten verversen** werkt de lijsten bij; na de fusie met Oirschot komen er immers nieuwe programma's, budgetbeheerders en organisatieonderdelen bij.

### 3.4 Controles vóór wegschrijven

| # | Controle | Gevolg |
| --- | --- | --- |
| 1 | `0 < LEN(D) <= 36` | blokkerend |
| 2 | `0 < LEN(F) <= 20` | blokkerend |
| 3 | Investeringsnummer bestaat al in kolom C | waarschuwing + volgnummers tonen (mag bewust bij samengestelde kredieten) |
| 4 | Activanummer bestaat al in kolom X | waarschuwing + volgnummer |
| 5 | Verplicht gevuld: B, C, D, F, H, I, J, O, R, S, T | blokkerend |
| 6 | Status = één letter (L/A/O) + vier cijfers | blokkerend |
| 7 | `M <= N` | blokkerend |
| 8 | `Y` is getal > 0 | blokkerend |
| 9 | Bedragen in Z en AE..AI zijn getal of leeg | blokkerend |
| 10 | I en J in kleine letters | automatisch corrigeren |
| 11 | Bij een getal in J: verdeelsleutel K sluit per investeringsnummer op 100% | waarschuwing, zichtbaar in het controleblok § 8 |

### 3.5 Knop **Krediet wijzigen**

Beschikbaar voor **financieel adviseur** en **financiële administratie**.

1. Zoekveld met keuze waarop gezocht wordt: **investeringsnummer (C)**, **omschrijving (D, deeltekst)**, **activanummer (X)** of **volgnummer (A)**.
2. Bij meerdere treffers: trefferlijst (volgnr, investeringsnummer, omschrijving, status, budget, restantkrediet) met keuze.
3. De gekozen regel wordt in **dezelfde verticale kaart** ingelezen; het rijnummer in een apart veld vastgelegd.
4. **Alle velden wijzigbaar, behalve:** `AA`, `AB`, `AC` (werkelijke uitgaven — alleen via § 6), `A` (volgnr) en de formulekolommen `E`, `G`, `K`, `AD`, `AL`, `AM`, `AN`.
5. Knop **Wijziging opslaan**: dezelfde controles als § 3.4; alleen gewijzigde cellen terugschrijven.
6. Bij opslaan door een financieel adviseur: `AP` terug naar `Te controleren door financiële administratie`. Bij opslaan door de financiële administratie: keuze tussen "opslaan en meteen bevestigen" of "opslaan, nog te controleren".
7. `AV`/`AW` bijwerken en **per gewijzigd veld** een logregel (oud → nieuw).
8. Knop **Kaart wissen** om zonder opslaan te stoppen.

~~Verwijderen staat bewust niet op het Dashboard.~~ **Vervangen in versie 4 (§ 0.1 punt 1):** verwijderen kan direct via de knop **Krediet verwijderen**. Een krediet dat klaar is, wordt nog steeds afgesloten via de status `A<jaar>` in kolom H; verwijderen is bedoeld voor regels die er niet hadden moeten staan.

---

## 4. Dashboardblok 2 — Controle door de financiële administratie

### 4.1 Nieuwe kolommen (rechts van de hulpkolommen)

| Kol | Kop | Inhoud |
| --- | --- | --- |
| AP | Controlestatus | `Te controleren door financiële administratie` / `Gecontroleerd door financiële administratie` / `Afgekeurd` |
| AQ | Gecontroleerd door | naam uit de rollentabel |
| AR | Gecontroleerd op | datum, hard (geen `TODAY()`) |
| AS | Opmerking controle | vrije tekst |
| AT | *(reserve)* | |
| AU | *(reserve)* | |
| AV | Laatst gewijzigd door | naam |
| AW | Laatst gewijzigd op | datum/tijd, hard |

### 4.2 Eenmalige vulling

Alle 269 bestaande regels (rij 2 t/m 270) krijgen `AP = "Gecontroleerd door financiële administratie"`. AQ/AR blijven leeg met in AS de tekst "bestaande regel bij invoering controlemodule". Vanaf dat moment is `Te controleren` uitsluitend nieuw of gewijzigd werk.

### 4.3 Werking

- Elke toegevoegde of door een financieel adviseur gewijzigde regel krijgt automatisch `Te controleren door financiële administratie`.
- Dashboardblok **Te controleren regels**: live lijst (dynamische formule) van alle regels met die status — volgnr, investeringsnummer, omschrijving, budget, restantkrediet, gewijzigd door, gewijzigd op.
- Knop **Regel openen ter controle** zet de gekozen regel in de wijzigkaart.
- Knop **Bevestigen — gecontroleerd** (alleen rol financiële administratie): vult AP, AQ (naam), AR (datum) en schrijft een logregel. **Per regel**, geen bulkbevestiging.
- Knop **Afkeuren met opmerking**: `AP = "Afgekeurd"`, verplichte tekst in AS, logregel. De regel blijft in de lijst staan tot hij is hersteld en opnieuw bevestigd.
- Teller bovenaan het Dashboard: "X regels te controleren, Y afgekeurd", met voorwaardelijke opmaak (oranje bij > 0, rood bij afgekeurd).
- Optioneel: op het blad `Financiële administratie` een extra kolom rechts van Z die `AP` ophaalt met `INDEX(...!$AP:$AP;ROW())`.

---

## 5. Gebruikers en rollen

### 5.1 Intake bij openen

Bij het openen van de werkmap (of via de knop **Wie werkt er?**) verschijnt een intakescherm:
1. De macro leest de Windows-gebruikersnaam (`Environ("USERNAME")`) en zoekt die op in de rollentabel.
2. Is er een treffer, dan toont het scherm de bijbehorende naam en rol ter **bevestiging**. Is er geen treffer, dan kiest de gebruiker zijn naam uit de lijst.
3. Keuze wordt vastgelegd in `Par_HuidigeGebruiker` en `Par_HuidigeRol`; de knoppen op het Dashboard worden daarop in- of uitgeschakeld.
4. Bij een onbekende naam: alleen lezen, geen knoppen actief.

De Windows-naam is dus de **voorgestelde** identiteit, niet de afgedwongen identiteit. Dat is bewust: bij een gedeelde werkplek of een bestand dat iemand namens een ander opent, kan de gebruiker corrigeren. De uiteindelijk gekozen naam **én** de Windows-naam worden allebei in het logboek vastgelegd, zodat achteraf zichtbaar is of die twee afweken.

### 5.2 Rollentabel op `Parameters` (zelf wijzigbaar)

| Naam | Windows-naam | Rol | Toevoegen | Wijzigen | Bevestigen |
| --- | --- | --- | --- | --- | --- |
| Michiel | *(in te vullen)* | financieel adviseur | ja | ja | nee |
| Emile | | financieel adviseur | ja | ja | nee |
| Joost | | financieel adviseur | ja | ja | nee |
| Monique | | financieel adviseur | ja | ja | nee |
| Marianne | | financieel adviseur | ja | ja | nee |
| Cor | | financiële administratie | nee | ja | ja |
| Mariska | | financiële administratie | nee | ja | ja |
| Sigrid | | financiële administratie | nee | ja | ja |

Tabel uitbreidbaar zonder de macro's aan te passen: de knoppen lezen de rechten uit de tabel, niet uit een lijst in de code. Achternamen en Windows-namen bij de uitwerking aanvullen.

### 5.3 Logboek

Nieuw tabblad **`Logboek`**, nooit verbergen:

`Datum/tijd | Gebruiker (gekozen) | Windows-naam | Rol | Actie | Volgnr | Investeringsnummer | Veld | Oude waarde | Nieuwe waarde | Toelichting`

Acties: toevoegen, wijzigen, bevestigen, afkeuren, uitgaven bijwerken, jaarrol, werkbereik verruimen, normaliseren, overzicht gemaakt. Datum/tijd altijd hard, nooit `NOW()`.

---

## 6. Knop **Werkelijke uitgaven bijwerken**

### 6.1 Uitgangspunt

De gebruiker levert een bestand aan van de financiële administratie. **De precieze kenmerken van dat bestand worden later vastgelegd** (kolomindeling, kostensoorten, bestandsnaamconventie). De macro wordt daarom opgezet met de bronherkenning in een apart, vervangbaar onderdeel: een leesfunctie die een lijst oplevert van `investeringsnummer → bedrag lopend jaar`, en een verwerkingsfunctie die daar niets van hoeft te weten.

Wat al wel vaststaat, uit de bestaande notitie bij `AC1`: de uitdraai bevat detailregels en totaalregels (alleen detailregels gebruiken), kostensoort `4…` zijn uitgaven en `8…` inkomsten, en het saldo per investeringsnummer is wat in de verzamellijst hoort.

### 6.2 Verwerking

1. Bestand kiezen (standaard `Par_PadUitgavenbestand`).
2. Lijst `investeringsnummer → bedrag` opbouwen uit de detailregels, uitgaven minus inkomsten.
3. **Koppelen aan het juiste krediet** — de kern van deze macro:
   - één regel in de verzamellijst met dat investeringsnummer → bedrag daarheen;
   - meerdere regels met hetzelfde nummer (samengesteld krediet) → verdelen volgens de verdeelsleutel in kolom K, zodat de som per nummer altijd exact gelijk is aan de bron; de betrokken nummers komen op een **signaallijst** zodat ze kunnen worden nagelopen;
   - nummer komt niet voor in de verzamellijst → **niet wegschrijven**, wel op de verschillenlijst (zoals eerder `7201025`, activanr. 52, −24.095,73);
   - regel in de verzamellijst zonder bedrag in de bron → `AC` op `0`, niet leeg laten, zodat zichtbaar is dat er is gekeken. Stond er bij de vorige keer inlezen wel een bedrag, dan komt er ook een melding in `Meldingen uitgaven` (§ 0.1 punt 3).
4. Na 31 maart alleen `AC` (lopend jaar) overschrijven. Tot en met 31 maart ook `AB` (voorgaand jaar), vanwege het afsluiten van de jaarrekening (§ 0.1 punt 2). `AA` is historie en blijft altijd ongemoeid.
5. **Harde poort:** som van de weggeschreven bedragen = som van de brondetailregels, per jaar. Verschil ≠ 0 in **één** van de jaren → in **geen enkel** jaar iets wegschrijven, verschillenlijst tonen per investeringsnummer. Peildatum, logboek en "Vorige upload" pas bijwerken als alles is weggeschreven.
6. `Par_PeildatumUitgaven` bijwerken en een notitie op `AC1` zetten met bestandsnaam, bronkolom, datum en eventuele aansluitverschillen.
7. Logregel wegschrijven.
8. `AP` (controlestatus) **niet** resetten — bijgewerkte uitgaven zijn geen inhoudelijke wijziging van de kredietgegevens.

### 6.3 Verzamelkredieten — later te ontwerpen

Coda werkt vaak met verzamelkredieten waarvan de werkelijke uitgaven per 31-12 worden doorgeboekt naar de onderliggende kredieten. Dat betekent dat de uitgaven gedurende het jaar op een ander nummer kunnen staan dan waar ze uiteindelijk thuishoren.

**Dit wordt in een later ontwerp uitgewerkt.** Voorbereiding die nu al zinvol is:
- de verschillenlijst uit § 6.2 houdt de nummers vast die niet te koppelen waren; dat is precies de ingang voor het verzamelkredietontwerp;
- een reservekolom op `Parameters` voor een vertaaltabel `verzamelkrediet → onderliggend krediet`, nog niet in gebruik;
- de leesfunctie uit § 6.1 blijft gescheiden van de verwerkingsfunctie, zodat die vertaalslag later tussen beide in kan worden gezet zonder de rest te herbouwen.

---

## 7. Dashboardblok 3 — Uitdraaien

Twee knoppen. Beide schrijven een **nieuw tabblad met harde waarden**. In de kop: titel, peildatum, lopend jaar en de gebruikte selectie. Naamgeving `<soort> <selectie> <dd-mm-jjjj>`, naar het model van `Maatwerk kredieten 23-09-2026`.

### 7.1 Presentatie (geldt voor beide overzichten)

**Sorteervolgorde:** eerst **kolom B (programma)**, dan **kolom I (naam van de groep kredieten)**, dan **kolom J als extra sorteerniveau**.

Kolom J werkt daarbij zo:
- **Bevat J een getal** (de groepsnummers 1 t/m 17), dan sorteert het overzicht binnen de kredietgroep op dat getal. Daardoor komen de deelkredieten die bij elkaar horen automatisch achter elkaar te staan.
- **Bevat J tekst** (`standaard`, `stelpost`), dan sorteert het alfabetisch binnen de kredietgroep.
- Gemengde sortering binnen één kredietgroep: eerst de genummerde deelkredieten op nummer, daarna de tekstwaarden. Zo staan samenhangende deelkredieten bovenaan de groep en de losse posten eronder.

Presentatie:
- B, I en J sturen de volgorde maar staan **niet als kolom in het overzicht**. In plaats daarvan verschijnen ze als **groepskopregels**: een regel met de programmanaam, daaronder een regel met de naam van de kredietgroep uit kolom I, en daaronder de kredietregels in de hierboven beschreven volgorde.
- J krijgt géén eigen groepskopregel — een regel met de tekst "7" zegt de raad niets. Het getal werkt uitsluitend als sorteervolgorde. Wel krijgen genummerde deelkredieten binnen een groep een **tussentotaal per nummer**, met als aanduiding de omschrijving van het eerste deelkrediet of het investeringsnummer dat zij delen.
- Per niveau een **tussentotaal** over alle bedragkolommen, en onderaan een eindtotaal.
- Platte lijst, **geen rijgroepering** — het overzicht moet zonder uitklappen af te drukken en te versturen zijn.
- Groepskoppen vet; bedragen `#,##0;-#,##0;-`, zwart; tekstkolommen met terugloop en autofit rijhoogte.
- Afdrukinstelling: liggend, kop herhalen, passend op breedte.

### 7.2 Knop **Kredietoverzicht maken** (raad / algemeen)

**Selectie, op kolom H:**
- alle regels met `L` + jaartal **≤ lopend jaar** → kredieten waarop nog uitgaven volgen;
- **plus** alle regels met `A` + **exact het lopend jaar** → in dit jaar afgerond, raad heeft conform besloten;
- **niet**: `A` van voorgaande jaren, en **niet**: `L` of `O` van een toekomstig jaar.

**Kolommen, in deze volgorde:**

| # | Bron | Kop in het overzicht |
| --- | --- | --- |
| 1 | A | Volgnr |
| 2 | D | Omschrijving investeringskrediet |
| 3 | Z | Budget beschikbaar tot en met `<lopend jaar>` |
| 4 | AA + AB + AC | Totaal werkelijke uitgaven |
| 5 | AD | **Restantkrediet** |
| 6 | AJ | Toelichting |

**Afspraak versie 4 (§ 0.1 punt 5):** de budgetten vanaf het volgende jaar (AE … AI) staan **niet** in het kredietoverzicht voor de raad. Dat overzicht gaat over het lopende jaar; de budgetten vanaf het volgende jaar horen bij de begroting.

Kolom C (investering boekhouding) komt **niet** in dit overzicht.

### 7.3 Knop **Kredietoverzicht op naam** (budgethouders)

Zelfde presentatie, met drie verschillen:

1. **Selectie op naam** via keuzelijst op het Dashboard, standaard gevuld met de budgetbeheerders uit kolom O, met een schakelaar naar projectleider (P) of tijdschrijver (Q).
2. **Kolom Y (afschrijftermijn) direct vóór Z**: Volgnr, Omschrijving, Afschr, Budget t/m lopend jaar, Totaal werkelijke uitgaven, Restantkrediet, Budget +1 … +5, Toelichting.
3. **Kredieten voor een toekomstig jaar worden hier wél meegenomen** (`L`/`O` met een jaartal na het lopend jaar), zodat een budgethouder ook ziet wat eraan komt. Die regels krijgen een duidelijke aanduiding in de groepskop of een aparte sectie "Nog niet gestart".

### 7.4 Technisch

- Bedragen als **harde waarde** wegschrijven; het overzicht is een momentopname.
- Tussentotalen en eindtotaal wél als `SUM`/`SUBTOTAL` binnen het nieuwe blad, zodat ze controleerbaar zijn.
- **Aansluitcontrole in de kop:** eindtotaal restantkrediet op het overzicht = `SUMIFS` over de verzamellijst met dezelfde selectie. Verschil ≠ 0 → melden en het blad niet als definitief markeren.
- In de kop ook vastleggen: lopend jaar, selectieregel in woorden ("L t/m 2026 plus A2026"), aantal regels, en de gebruiker die het overzicht heeft gemaakt.

---

## 8. Dashboardblok 4 — Controle en kengetallen

### 8.1 Statusregel

Eén cel met `OK — alles sluit aan` of `LET OP`, naar het model van `CC1:CD5` in de activamodule. `LET OP` zodra één van de controles hieronder afwijkt; eerst oplossen, dan verder.

### 8.2 Controles

| # | Controle |
| --- | --- |
| 1 | Aantal gevulde regels in kolom A = aantal gevulde regels op `Financieel adviseur`, `Financiële administratie` en `Kredietoverzicht` |
| 2 | Rij-voor-rij over het hele datablok: investeringsnummer op rij n van de bron = dat op rij n (resp. n+1) van elke spiegel |
| 3 | Aantal `fout` in E en G |
| 4 | Verdeelsleutel K sluit per samengesteld investeringsnummer op 100% |
| 5 | Status `A<jaar>` met restantkrediet in AD ≠ 0 |
| 6 | Werkelijke uitgaven (AA+AB+AC) > budget Z — overschrijding |
| 7 | Regels zonder budgetbeheerder (O), projectleider (P), activanummer (X) of afschrijftermijn (Y) |
| 8 | Dubbele combinatie investeringsnummer + activanummer |
| 9 | Regels met `Te controleren door financiële administratie` of `Afgekeurd` |
| 10 | Hoofdletters of spaties in de lijstkolommen I en J |
| 11 | `Par_LaatsteRij` binnen 100 rijen van `Par_MaxRij` — werkbereik verruimen |
| 12 | Status in H voldoet aan het patroon letter + vier cijfers |

### 8.3 Kengetallen

Aantal regels per statusletter (L / A / O), totaal budget t/m lopend jaar, totaal werkelijke uitgaven, totaal restantkrediet, totaal budget per toekomstig jaar (+1 t/m +5), aantal te controleren en afgekeurde regels. Allemaal als formule (`SUMIFS`/`COUNTIFS`) op de verzamellijst, nooit als harde waarde.

---

## 9. Aandachtspunten bij de uitwerking

### 9.1 Beslispunt: de sortering B → I → J trekt deelkredieten uit elkaar

**Gemeten in de werkmap** (rij 2 t/m 270): kolom J bevat 14 verschillende groepsnummers. Daarvan lopen er **9 dwars door programma's en/of kredietgroepen heen**:

| Nr | Regels | Programma's | Kredietgroepen (kolom I) |
| --- | --- | --- | --- |
| 1 | 5 | 2 en 7 | openb ruimte, wegen, riolering |
| 2 | 3 | 2 en 7 | wegen, riolering |
| 3 | 6 | 4 | school b1, school b2, school b3 |
| 4 | 3 | 4 | school d1, school d2, school d3 |
| 5 | 5 | 4 | school a1, school a2, school a3 |
| 9 | 3 | 2 en 7 | wegen, riolering |
| 11 | 2 | 2 en 7 | wegen, riolering |
| 12 | 6 | 7 en 2 | riolering, wegen |
| 14 | 3 | 4 | school e1, school e2, school e3 |

De overige 5 groepsnummers vallen wél binnen één programma en één kredietgroep.

**Het gevolg:** omdat de sortering begint bij kolom B en daarna I, wordt een deelkrediet met nummer 12 eerst op programma 7 (riolering) en dan op programma 2 (wegen) geplaatst. Het groepsnummer in J komt pas als derde niveau aan bod en kan die scheiding niet meer ongedaan maken. De deelkredieten die bij elkaar horen staan dan juist **niet** bij elkaar.

Dit is geen fout in de data — een straatreconstructie staat terecht deels onder wegen en deels onder riolering. Maar het betekent dat "presenteren per programma" en "deelkredieten bij elkaar" elkaar in negen gevallen uitsluiten. Er moet dus een keuze worden gemaakt:

- **Optie A — programma wint.** Sorteren B → I → J, deelkredieten komen uit elkaar. Wel achter de omschrijving een verwijzing tonen, bijvoorbeeld "(deel van samengesteld krediet 12, zie ook programma 2)". De raad ziet het beeld per programma; de samenhang is herleidbaar maar niet zichtbaar in de volgorde.
- **Optie B — aparte sectie.** De genummerde deelkredieten niet onder de programma's plaatsen, maar in een eigen sectie "Samengestelde kredieten" ná de programma's, daarbinnen gesorteerd op nummer. Programma-indeling blijft zuiver voor de losse kredieten, en samenhang blijft zichtbaar. Nadeel: een deel van het budget valt buiten de programmatotalen.
- **Optie C — sluitregel bepaalt de plek.** De hele groep wordt gepresenteerd bij het programma en de kredietgroep van zijn sluitregel (de regel met `AN = 1`). Samenhang blijft intact en alles valt binnen één programmatotaal, maar het budget van bijvoorbeeld het rioolaandeel wordt dan onder wegen gepresenteerd.

**Nog te beslissen.** De rest van de specificatie staat los van deze keuze; alleen § 7.1 verandert.

### 9.2 Overige aandachtspunten bij de uitwerking

1. **Getal versus tekst in kolom J.** De sortering moet getallen als getal behandelen, niet als tekst — anders komt 10 vóór 2 te staan. Gecontroleerd: op dit moment staan alle 14 groepsnummers als **getal** opgeslagen, geen enkele als tekst. Wel blijven controleren bij nieuwe invoer, want een als tekst opgeslagen `7` sorteert anders én valt buiten `ISNUMBER($J2)` in kolom K.
2. **Teksten in kolom K.** Naast percentages komen in K ook tekstwaarden voor (`nvt, afgerond`, `administratief`). Bij de verdeling van werkelijke uitgaven over samengestelde kredieten (§ 6.2) en bij controle 4 in § 8.2 moeten die regels worden overgeslagen in plaats van als 0% te worden gelezen.

---

## 10. Bouwvolgorde

1. `Parameters`-blad, benoemde bereiken, dynamische koppen (§ 2.1, 2.2, 2.6). Daarna controleren of de spiegelbladen nog aansluiten.
2. Macro **Lijstwaarden normaliseren** (§ 2.5) — J155:J162 naar kleine letters, met nulmeting.
3. Overstap naar dynamische bereiken (§ 2.3) — **apart uitvoeren**, met nulmeting vooraf en nameting achteraf.
4. Macro **Werkbereik verruimen** (§ 2.4) — nodig vóór de instroom uit Oirschot.
5. Nieuwe kolommen AP t/m AW (§ 4.1) plus de eenmalige vulling (§ 4.2).
6. `Logboek`-blad en de rollentabel met intakescherm (§ 5).
7. Dashboard-opmaak: vier blokken, verticale mutatiekaart, keuzelijsten (§ 3.1, 3.3).
8. Macro **Krediet toevoegen** (§ 3.2) — eerst proefregel op rij 271 / volgnr 302, daarna **leegmaken, niet verwijderen**.
9. Macro **Krediet wijzigen** en **Wijziging opslaan** (§ 3.5).
10. Controleknoppen financiële administratie (§ 4.3).
11. Macro **Werkelijke uitgaven bijwerken** (§ 6), met harde poort; bronherkenning als vervangbaar onderdeel.
12. Controleblok en kengetallen (§ 8).
13. Macro's **Kredietoverzicht maken** en **Kredietoverzicht op naam** (§ 7).
14. Macro **Jaarrol** (§ 2.7) — als laatste, want die raakt alle bedragen. Eerst testen op een kopie van de werkmap.

**Later, nog te ontwerpen:** verzamelkredieten in Coda (§ 6.3) en bladbeveiliging. Bij het later alsnog inbouwen van beveiliging geldt: bij het terugzetten altijd `allowEditObjects: true` meegeven, plus autofilter en sorteren toestaan — anders reageren de macroknoppen niet meer.
