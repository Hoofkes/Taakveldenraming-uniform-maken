# Werkwijze & sjabloon (interne kennisbank voor uniformering taakvelden)

Doel van dit bestand: alles wat één keer is uitgezocht vastleggen, zodat elk volgend
taakveld sneller gaat. Wordt bijgewerkt zodra er iets nieuws blijkt. Niet vertrouwen
op sessiegeheugen — dit bestand is de bron.

## 0. Omgeving / gereedschap (belangrijke valkuilen)

- **LibreOffice werkt hier NIET voor document-load** (`soffice --convert-to` geeft
  "source file could not be loaded", ook voor csv/ods). Daarom herberekenen we met de
  Python-engine **`formulas`** (pip). Die berekent alle formules opnieuw vanuit de
  bróntekst (negeert cache) → gelijkwaardig aan de vereiste LibreOffice-herberekening,
  en sterker dan alleen formuletekst vergelijken. Documenteer deze substitutie.
- **`pkill -f soffice` NIET gebruiken**: -f matcht de eigen shell-commandline (bevat
  "soffice") → shell killt zichzelf. Gebruik `pkill -9 -x soffice.bin`.
- **Geen foreground `sleep`** in Bash-tool (wordt geblokkeerd → lege exit 1).
- `formulas`-engine loopt vast/traag op het hele bestand (77 tabs, #REF!). Daarom
  altijd een **scoped kopie** maken met alleen het taakveld + `Indexering`.

## 1. Herreken/vergelijk-pijplijn (gevalideerd op originele T01: 0 afwijkingen)

Scripts staan in de repo:
- `scoped_recalc.py <bestand> <prefix>`  → maakt `/tmp/<prefix>_scoped.xlsx` met alleen
  het taakveld + Indexering. (Pas `keep=` aan naar het juiste taakveld.)
- `eval_recalc.py <scoped.xlsx> <prefix>` → herberekent met `formulas`, schrijft
  `/tmp/<prefix>_t01.json` (celwaarden). (Pas de sheet-regex aan het taakveld aan.)
- `compare.py <prefix>` → vergelijkt met `snapshot_fase0_waarden.json` per cel.
  Tolerantie: >€1 per cel telt als afwijking (totaalregels >€5).

Anker = `snapshot_fase0_waarden.json` (Fase 0, cached waarden van hele bestand).
Nooit vergelijken tegen live werkbestand.

## 2. Bekende gebreken (Fase 0 vastgelegd)

- **28 #REF!-rijen in Recap** (14 taakvelden × lasten+baten): T622, T623, T64, T660,
  T6711–T6714, T6752, T6762, T6791, T6792, T6821, T6822. Oorzaak: gebroken externe
  .xlsx-verwijzingen. Anker = cached waarden op het taakveldtabblad zelf; Recap-verwijzing
  bij herbouw herstellen naar `'<TV>'!<totaalcel>` en als afwijking documenteren.
- **Taakvelden zonder Recap-regel**: T6751, T6753, T6761, T6811. (T6791/T6792 hebben een
  regel met #REF! en zonder code in kol B.) Bij herbouw Recap-regel toevoegen.
- Lege resttabbladen: Blad1, Blad2, Blad3. Dubbele inhoudsopgave (Inhoudsopgave 2).
  Pas na afronding opruimen (beslispunt 6).

## 3. Recap-structuur (voor doorkoppeling per taakveld)

- Lastenblok begint r3; batenblok begint ± r68 (na "Saldo lasten" r63). Batenrijen ± r99+.
- Jaar-kolommen: **E=2026, F=2027, G=2028, H=2029, I=2030** (géén 2025 in Recap).
- Lastenregel verwijst naar `'<TV>'!K<TotLasten>` .. `O<TotLasten>` (K=2026..O=2030).
- Batenregel verwijst naar Totaal baten-cel. Saldo excl. reserves = totaal zonder T010.
- Recap bevat UITSLUITEND doorkoppelingen.

## 4. Indexering (bron, ongewijzigd)

Tabblad `Indexering`, jaren in D:M = 2021..2030. Reeksen:
- r4 Prijsindex, r5 Loonindex, r6 Gewogen index, r7 Sociaal domein loonindex.
- "Index conform kaderbrief" (r16–20): later samenvoegen/koppelen (nog te doen).
- Inwonersreeks: gereserveerd blok r23+ (nog aan te leveren, beslispunt 3).

## 5. Sjabloon per taakveld (Bijlage A) — kolom- en laagindeling

Kolommen (alle lagen): C markering · D grootboeknr · E BBV-categorie (één per regel) ·
F verbijzondering · G omschrijving (tekst, nooit formule-label) · H indexkeuze
(Prijzen/Lonen/Gewogen/Sociaal domein/Geen) · I btw/BCF (leeg) · J 2025 · K 2026 ·
L–O 2027–2030 · P e.v. nieuw jaar. Hulpblok Q–W = indexreeksen per jaar (verwijzing
naar Indexering) + reekslabels voor INDEX/VERGELIJKEN.

Drie lagen boven→onder: **Bovenlaag** (verzamelstaat, per grootboek =middenlaag, SOM
Totaal lasten/baten, saldo) · **Middenlaag** (per grootboek × categorie één regel,
=onderlaag) · **Onderlaag** (budgetopbouw, werkplek adviseur).

Onderlaag-formulepatroon (A.4), per jaar vanaf 2027 (L=correctiecel, K=vorig jaar,
$H=indexkeuze van de post):
```
=ALS(L{corr}<>""; L{corr};
   ALS($H{b}="Geen"; K{b};
      AFRONDEN(K{b}*INDEX(hulpblok-jaarkolom; VERGELIJKEN($H{b}; reekslabels;0)); 0)))
```
Correctie werkt via de basisregel → werkt automatisch door naar latere jaren.
Correctierijen tellen NOOIT mee in een SOM.

## 6. Index-classificatie (uit bestaande formules afleiden)

- Formule met `*T2`/`*T$2` (of U2/V2/W2 = rij 2 hulpblok) → **Prijzen**.
- Formule met `*T3`/`*T$3` (rij 3 hulpblok) → **Lonen**.
- Gewogen index rij 4 hulpblok → **Gewogen**.
- Salaris (41100), afschrijving/activa (473*/873*), reserves (471*/871*),
  verrekening (475*/875*) → indexkeuze **Geen** + standaardtoelichting:
  41100 "zie staat van salarissen" · 471/871 "zie staat van reserves en voorzieningen"
  · 473/873 "zie staat van activa" · 475/875 "verrekening". (nu handmatig; later koppeling)
- Inwonergedreven posten (bedrag = inwoners × tarief): indexkeuze **Geen**, bedrag als
  eigen onderbouwing (inwonertal × tarief) laten staan; markeren als technische
  bijzonderheid "koppelen aan inwonersreeks Indexering zodra aangeleverd".

## 7. Vaste onvolkomenheden die NIET worden overgenomen (opsporen & fixen)

- Labels die formules zijn (`=G94`, `="Totaal "&D227`) → vaste tekst.
- Inconsistente stijl (`=+K118` naast `=K82`) → uniform `=K118`.
- Hardcoded index in formule (`=579*1.026`) → verwijzing naar hulpblok/Indexering.
- Samengestelde invoer (`=43940+15751`) → benoemde losse regels (of uitsplitsing in
  toelichting), som gelijk.
- Harde correctiebedragen zonder markering (`W02-2025`, `A03-2025`) → benoemde
  structurele regels met indexkeuze Geen, W/A-referentie in de omschrijving, gemarkeerd.
- Meerdere BBV-categorieën op één regel → splitsen.

## 8. Kleurcodering / markering

blauw = invoer · geel = handmatige correctie · oranje = incidenteel.
I-kolom (btw/BCF) blijft leeg. Bladbeveiliging optioneel (beslispunt 8, nog te bepalen).

## 9. Beslispunt-keuzes (voorlopig, te bevestigen bij controle T01)

1 historie 25/26 = bevroren waarden · 2 ritme = batches ~5 na T01 · 3 inwoners = nog
aanleveren · 4 Inlees CODA + 2030 · 5 fte nog aanleveren · 6 opruimen na afloop ·
7 T82 black box · 8 bladbeveiliging nog bepalen · 9 volgorde koppelen nog bepalen ·
10 kolomgroei nog bepalen.

## 10. Per-taakveld checklist (Fase 1)

1. Inventariseren (grootboek, categorie per regel via middenlaag, formules, harde cellen).
2. Opbouwen volgens sjabloon (§5), fixes §7, index §6, staat-toelichtingen.
3. Scoped herberekenen (§1) + vergelijken met snapshot; circulariteit checken.
4. Bijzonderheden noteren (technisch = doorwerken; financieel = vraag stellen, wachten).
5. Meldingsregel bovenaan tabblad, status op Voortgang, opslaan, om controle vragen.
