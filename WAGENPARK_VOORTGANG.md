# Wagenparkbeheer – voortgang en afspraken

Laatste update: 30-10-2026. Branch: `claude/festive-cannon-37n264`.

## Bestanden

| Bestand | Wat |
|---|---|
| `Wagenpark_beheer.xlsx` | Werkmap (zonder macro's). Gebruiker importeert de .bas en bewaart als `.xlsm`. |
| `Wagenpark_macro.bas` | VBA-module "Wagenpark": jaarovergang, aanschaf, wijzigen, investeringen beheren. |
| Bron | `Overzicht_Voertuigen.xlsx` (oude lijst, blad "Voertuigen"). |

Afspraak: bij wijzigingen alleen nog het **.bas-bestand** opleveren, geen nieuwe Excel (tenzij gevraagd).
De macro's zijn **niet getest in Excel** (geen Excel beschikbaar): altijd eerst op een kopie testen.

## Installeren macro

1. Bestand openen, Alt+F11.
2. Oude module "Wagenpark" verwijderen (rechtsklik → Verwijderen).
3. Bestand → Bestand importeren → `Wagenpark_macro.bas`.
4. Opslaan als `.xlsm`, sluiten, opnieuw openen, "Inhoud inschakelen". Knoppen verschijnen vanzelf (Auto_Open).

## Opbouw werkmap

Alle bladen: titel rij 1, één kopregel op rij 4, data vanaf rij 5, daaronder één lege regel, dan totalen. Geen samengevoegde balken boven de kop. Letters zwart, invulcellen zacht groen (`E2EFDA`), kop donkerblauw met witte tekst.

### Dashboard
- C3 `LopendJaar` (nu 2026).
- C7–C13 invoer aanschaf: `AB_Voertuig` (keuzelijst uit verborgen kolom H), `AB_Kenteken`, `AB_Datum`, `AB_Krediet`, `AB_Activa`, `AB_Kosten`, `AB_DatumUit`.
- C16/D16 doorschuifregel: `DoorschuifVan` = V01, `DoorschuifNaar` = V03.
- Stand (rij 19–22) en Controle (rij 25–36): telt ontbrekende zaken met OK/Actie.
- Knoppen (door macro gemaakt): Jaarovergang nu (E3), Aanschaf bevestigen (E7), Voertuig wijzigen (E10), Voertuig afvoeren (E12), Investering wijzigen (G7), Investering toevoegen (G10), Investering schrappen (G12).

### Voertuigen <jaar> (nu "Voertuigen 2026")
Kolommen A–Q: V-nummer, Kredietnr., Kenteken, Gebruiksdoel, Aanschafwijze, Berijder, Aanschafjaar, Levensduur, Aanschafkosten (werkelijk), Te activeren in jaar, Activakaart, Aanbetaling, Bijkomende kosten, Slotbetaling, Exploitatienr., Leasesom 1 jaar, Status.
- Status "In gebruik" bovenaan, "Te vervangen <jaar>" daaronder (oranje). Sortering: status, V-nummer.
- Koop: alleen kolom I. Lease-koop: kolom I leeg, J en L–N gevuld. Te activeren in jaar = aanschafjaar + 2.
- Cijfers in gebruik = hard (betaald of bestaande afspraak). Uitgaven vanaf 2027 = gemiddelde, overschrijfbaar.
- Totaalregels: "Totaal in gebruik" en "Totaal te vervangen" (SUMIF op status).

### Vervangingsinvesteringen
Kolommen A–P: V-nummer, Gebruiksdoel, Regime, Aanschafjaar, Einde levensduur, Uitgaven jaar+1 t/m jaar+4 (kop = `LopendJaar+1`…), Afschrijftermijn, Activanummer, Te activeren, Jaar activering, Jaarlease (info), Exploitatienr., Opmerking.
- Eén regel per V-nummer = het eerstvolgende voertuig (volgnummer +1). Volgorde: aanschafjaar.
- Te activeren lease-koop = formule `=a+b+o+s` (namen of getallen; macro leest dit terug). Koop = koopprijs (waarde).
- Uitgaven per jaar: lease-koop aanbetaling+belettering+opbouw in aanschafjaar, slotbetaling in jaar activering (aanschafjaar + 2); koop in aanschafjaar.
- Standaardbedragen onder de lijst (namen): `p_Aanbet` 10.000, `p_Belet` 5.000, `p_Opbouw` 7.000, `p_Slot` 35.000, `p_SlotJr` 2, `p_Lev` 8, `p_Afschr` 6, `p_Lease` 12.000.

### Buiten gebruik
Kolommen A–G: V-nummer, Kredietnr., Kenteken, Gebruiksdoel, Aanschafjaar, Datum buiten gebruik, Activakaart. Oude voertuigen zonder volgnummer hebben alleen een jaar.

## Spelregels (van de gebruiker)

- V-nummer + volgnummer: `V05-01` is in gebruik, vervanger is `V05-02`.
- Personen- en bedrijfsauto's = lease-koop. Overige voertuigen/machines = koop, afschrijving volgens eigen levensduur.
- Lease-koop: aanbetaling, belettering, opbouw in aanschafjaar; overname/slotbetaling 2 jaar later = jaar activering; afschrijving 6 jaar. Einde levensduur = aanschafjaar + 8.
- Jaarlease alleen ter info (exploitatie), niet in investeringen.
- Afschrijving per jaar niet tonen (zit in activamodule). Wel activanummer per toekomstige investering.
- V01/V03 doorschuif: bij aanschaf nieuwe V01 wordt oude V01 de nieuwe V03; oude V03 gaat buiten gebruik.
- Lopende lease-contracten lopen volledig via "In gebruik", niet via Vervangingsinvesteringen.
- Negatieve bedragen met minteken, punten tussen duizendtallen. Teksten kort, niveau 16-jarige.

## Wat de macro doet

- **Auto_Open**: maakt knoppen; als huidig jaar > LopendJaar → jaarovergang (eventueel meerdere jaren), meldt resultaat.
- **Jaarovergang**: VI-regels met aanschafjaar = nieuw jaar → regel "Te vervangen <jaar>" op Voertuigen (alle velden gevuld, onbekend = "n.n.b."); VI-regel wordt volgend voertuig (volgnummer +1, aanschafjaar = einde levensduur, standaardbedragen, activanummer leeg). Bij V01 schuift V03-regel mee. Blad wordt hernoemd, LopendJaar +1, sorteren, totalen.
- **Aanschaf bevestigen**: nieuw voertuig → "In gebruik" (kenteken, datum, optioneel krediet/activa/kosten, daarna bedragen aanpassen); oud voertuig zelfde basisnummer → Buiten gebruik met datum (V01-uitzondering).
- **Voertuig wijzigen / afvoeren**, **Investering wijzigen / toevoegen / schrappen**: via invoervensters.

## Open punten / volgende stappen

1. Macro testen in Excel (op kopie): jaarovergang (knop "Jaarovergang nu"), aanschaf, V01-doorschuif, wijzigen.
2. Gebruiker vult in: activanummers (26), activakaarten (11), kredietnummers (4), kentekens (3), datum buiten gebruik V94/V95/V99.
3. Controleren: V96 en V98 staan op 2027 buiten gebruik (toekomst); begroting 2027 rekende V07, V13, V26, V27 nog als koop.
4. Mogelijke uitbreidingen: wijzigingslog, historie per V-nummer, oude buiten-gebruik-regels voorzien van volgnummer.

## Bouwscripts

In de scratchpad van de sessie (niet in de repo): `build18.py` bouwt de Excel uit de bron. Bij een nieuwe sessie: Excel opnieuw bouwen is niet nodig; werk verder in de .bas of laat de gebruiker de huidige `.xlsm` uploaden.
