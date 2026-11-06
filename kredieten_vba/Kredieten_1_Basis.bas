Attribute VB_Name = "Kredieten_1_Basis"
Option Explicit

' Alle macros voor het Dashboard investeringskredieten in een module.
' Versie 50 (stap 9: controles op de spiegelbladen vervallen omdat die bladen niet meer bestaan; geen controle op overschrijding, want de kredietoverzichten tonen dat al).
' Versie 49 (stap 7: uitvraag kadernota/begroting in een eigen module (6): een blauw blad Uitvraag, drie knoppen (versturen, ophalen, verwerken), lopend jaar vastgelegd, jaarovergang geblokkeerd bij onverwerkte regels, direct Te controleren; oude knoppen en bladen vervallen).
' Versie 48 (stap 6: blok 2 op het Dashboard in 2a, 2b en 2c; overzicht op naam met keuzevragen in plaats van de keuzecellen rol en naam; blok 3 en 4 een rij omlaag).
' Versie 47 (altijd terug op het Dashboard na krediet toevoegen, wijzigen, verwijderen, ook bij foutmeldingen; naam samengesteld krediet geblokkeerd na het eerste deelkrediet).
' Versie 46 (invoerkaart: tekst in de invoercellen past altijd; de rij wordt zo hoog als nodig).
' Versie 45 (stap 5: groepsnummer in kolom J dat als tekst is opgeslagen wordt overal als getal behandeld; overzicht op naam toont ook A + lopend jaar).
' Versie 44 (stap 4: back-up Verzamellijst 31-12-<jaar> en Voorstellen nieuwe kredieten worden na elke macro weer beveiligd).
' Versie 43 (stap 3: terugdraaien bij een fout: krediet toevoegen, wijziging opslaan met controle volgnummer, krediet verwijderen, formule restant, jaarovergang met kolom M en N volgens par. 2.7 punt 7).
' Versie 42a (meldingen uitgaven: volgnummer en rij vooraan).
' Versie 42 (stap 2: werkelijke uitgaven - harde poort over beide jaren samen, vastleggen pas na geslaagde verwerking, notities na de controle, foutwaarde in verdeelsleutel K blokkeert, bedrag dat uit de uitdraai verdwijnt wordt 0; meldingen tonen rijnummer(s) en omschrijving van het krediet).
' Versie 41 (stap 1: keuzelijst samengestelde kredieten zonder groepsnummer, lange uitlegteksten, logregels normaliseren en werkbereik, map aanmaken en overzicht opslaan met nette foutmelding).
' Versie 40 (rijhoogte overzichten 13,5; uitvraag verwerken met de hele lijst; krediet wijzigen en verwijderen op rij 35; uitvraag per budgethouder: inlezen, per investering besluiten, verwerken, stand uitvraag, begroting vastgesteld; knop Uitvraag verwerken voorlopig weg; Dashboard: 1 invoerkaart, 2 financieel adviseur, 3 financiele administratie, 4 controles; toelichting maximaal 1000 tekens; datum keuzeopties in F3; budget in het statusjaar verplicht en geel; fout bij DashboardAanmaken opgelost; nieuwste samengestelde krediet bovenaan in de keuzelijst; applicatiebeheerder ziet alle blokken; blokken per rol, knoppen onder de kaart, budgethouder/projectleider/tijdschrijver vervangen, kortere veldnamen; naam samengesteld krediet direct onder C16; afgekeurd: signaal, herstellen, aanbieden voor hercontrole; kleuren geel = verplicht, lichtblauw = mag leeg; compact Dashboard, kengetallen op eigen tabblad, naam vervangen).

Public Const BL_VZ As String = "Verzamellijst"
Public Const BL_DASH As String = "Dashboard FinAdv"   ' dashboard financieel adviseurs (was 'Dashboard')
Public Const BL_ADM As String = "Dashboard Adm"       ' dashboard financiele administratie
Public Const ADM_FA_KOP As Long = 6         ' Dashboard Adm blok 1 Controle kredieten: kopregel (uitleg 7, knoppen 8)
Public Const ADM_OV_KOP As Long = 10        ' Dashboard Adm blok 2 Kredietoverzichten: kopregel (uitleg 11, knoppen 12)
Public Const ADM_NR_KOP As Long = 14        ' Dashboard Adm blok 3 Eerstvolgende vrije kredietnummers: kopregel (tabel tot en met rij 27)
Public mFinAdm As Boolean                   ' overzicht gemaakt via Dashboard Adm: opslaan in de mappen FinAdm ...
Public Const VL As String = "'Verzamellijst'!"
Public Const eerste As Long = 10         ' eerste kaartrij op het Dashboard (rij 7 kopje, rij 8 leeg, rij 9 kopregel tabel)
Public Const KAART_KNOP As Long = 35     ' rij onder de invoerkaart: Ingevuld? Verwerken! en Scherm leegmaken
Public cfMislukt As Boolean
Public Const aantal As Long = 30
Public Const BL_MELD As String = "Meldingen uitgaven"
Public mMeld As Collection
Public mVorige As Object          ' bedragen per kredietnummer en jaar uit de vorige upload
Public mHeeftVorige As Boolean
Public mAansl As String
Public Const BL_OVERZICHT As String = "Laatste kredietoverzicht"          ' raad en jaarrekening
Public Const BL_OVERZICHT_NAAM As String = "Laatste overzicht op naam"  ' budgethouder / projectleider
Public Const BL_UITVRAAG As String = "Uitvraag"                         ' blauw blad: alle budgetbeheerders, aanleveringen en besluiten
Public Const BL_UVTEMP As String = "Uitvraag tijdelijk"               ' tijdelijk blad waarmee de bestanden per budgetbeheerder worden gemaakt
Public Const BL_TIJD As String = "Laatste overzicht tijdschrijven"
Public Const OPBRENGST As String = "Opbrengst activa"   ' bron/toelichting voor opbrengsten activa (geen subsidie of bijdrage)
Public Const BL_MUTREG As String = "Mutaties kadernota-begroting"   ' register voor de lastenberekening
Public Const WACHTWOORD As String = "Best"      ' bladbeveiliging (zet ook een wachtwoord op het VBA-project)
Public mStil As Boolean
Public Const UITLEG_VERSIE As String = "49e"   ' hoger nummer = uitleg onder de vraagtekens wordt bij het openen opnieuw geplaatst
Public Const CB_KOP As Long = 51            ' Dashboard blok 4 Controles: kopregel (onderaan)
Public Const CB_EERSTE As Long = 54         ' eerste controleregel
Public Const CB_LAATSTE As Long = 69        ' laatste controleregel (16 controles)
Public Const NM_KOP As Long = 37            ' Dashboard blok 2 Financieel adviseur: kopregel; daaronder 2a (38-40), 2b (41-43), 2c (44-45)
Public Const FA_KOP As Long = 47            ' Dashboard blok 3 Financiele administratie: kopregel (uitleg 48, knoppen 49)
Public Const BL_WERK As String = "Werkregel controle"   ' verborgen: regel die de financiele administratie open heeft
Public Const BL_KENG As String = "Kengetallen"
Public Const GEEN_WIJZ As String = "geen wijziging"
Public Const MAX_TOEL As Long = 1000        ' maximale lengte toelichting (past altijd in de kredietoverzichten)
Public Const FA_LIJST As Long = 120         ' kopregel van de lijst te controleren / afgekeurd (regels eronder)
Public mHerbeveiligingGepland As Boolean
Public mActiefVoorMacro As Object        ' blad dat actief was toen de macro begon
Public mBlijfOpDashboard As Boolean      ' na krediet toevoegen, wijzigen of verwijderen altijd terug naar het Dashboard
Public mProef As Boolean             ' proefversie jaarrekening (niet opslaan)             ' geen meldingen tijdens een reeks overzichten
Public Type TKaart
    f(1 To 30) As Variant
    Letter As String
    jaar As String
    grp As Long
    nieuw As Boolean
    jWaarde As Variant
    naam As String
    pc As String
    tc As String
    bc As String
End Type
Public mBlok As Boolean          ' foutwaarde in de verdeelsleutel: werkelijke uitgaven niet wegschrijven

' ======================================================================
' Rollen, intake en logboek (stap 6)
' ======================================================================
' Stap 6 - Logboek, rollentabel en intake
' LogboekEnRollentabelAanmaken : maakt blad Logboek en de rollentabel (op blad Parameters) eenmalig aan.
' WieWerktEr                   : intake. Stelt Par_HuidigeGebruiker in; de rol volgt uit de rollentabel.
' Auto_Open                    : ververst de keuzeopties en start de intake automatisch bij het openen van het bestand.
' HeeftRecht / LogRegel        : hulpfuncties voor de latere stappen.

Public Sub LogboekEnRollentabelAanmaken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wb As Workbook, lg As Worksheet, pa As Worksheet, lo As ListObject, actief As Object
    Dim koppen As Variant, personen As Variant
    Dim i As Long, rol1 As String, rol2 As String
    Dim nieuwLog As Boolean, nieuwTabel As Boolean

    Set wb = ThisWorkbook
    Set actief = ActiveSheet
    On Error GoTo Fout
    Application.ScreenUpdating = False

    ' --- Logboek
    If Not BladBestaat("Logboek") Then
        Set lg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        lg.Name = "Logboek"
        koppen = Array("Datum/tijd", "Gebruiker (gekozen)", "Windows-naam", "Rol", "Actie", "Volgnr", _
                       "Investeringsnummer", "Veld", "Oude waarde", "Nieuwe waarde", "Toelichting")
        For i = 0 To 10
            lg.Cells(1, i + 1).Value = koppen(i)
        Next i
        With lg.Range("A1:K1")
            .Font.Bold = True
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        lg.Range("I:J").NumberFormat = "@"
        lg.Range("A2:A1048576").NumberFormat = "dd-mm-yyyy hh:mm:ss"
        lg.Columns("A").ColumnWidth = 19
        lg.Columns("B:D").ColumnWidth = 22
        lg.Columns("E").ColumnWidth = 22
        lg.Columns("F:G").ColumnWidth = 14
        lg.Columns("H").ColumnWidth = 22
        lg.Columns("I:J").ColumnWidth = 28
        lg.Columns("K").ColumnWidth = 60
        lg.Range("A1:K1").AutoFilter
        lg.Activate
        ActiveWindow.FreezePanes = False
        lg.Range("A2").Select
        ActiveWindow.FreezePanes = True
        nieuwLog = True
    End If

    ' --- Rollentabel
    Set pa = wb.Worksheets("Parameters")
    If Not TabelBestaat() Then
        rol1 = "financieel adviseur"
        rol2 = "financi" & ChrW(235) & "le administratie"
        pa.Range("A12").Value = "Rollentabel: voeg nieuwe personen direct onder de tabel toe, de tabel groeit mee. Vul de Windows-naam in."
        pa.Range("A12").Font.Bold = True
        koppen = Array("Naam", "Windows-naam", "Rol", "Toevoegen", "Wijzigen", "Bevestigen")
        For i = 0 To 5
            pa.Cells(14, i + 1).Value = koppen(i)
        Next i
        personen = Array(Array("Michiel", rol1, "ja", "ja", "nee"), Array("Emile", rol1, "ja", "ja", "nee"), _
                         Array("Joost", rol1, "ja", "ja", "nee"), Array("Monique", rol1, "ja", "ja", "nee"), _
                         Array("Marianne", rol1, "ja", "ja", "nee"), Array("Cor", rol2, "nee", "ja", "ja"), _
                         Array("Mariska", rol2, "nee", "ja", "ja"), Array("Sigrid", rol2, "nee", "ja", "ja"))
        For i = 0 To 7
            pa.Cells(15 + i, 1).Value = personen(i)(0)
            pa.Cells(15 + i, 3).Value = personen(i)(1)
            pa.Cells(15 + i, 4).Value = personen(i)(2)
            pa.Cells(15 + i, 5).Value = personen(i)(3)
            pa.Cells(15 + i, 6).Value = personen(i)(4)
        Next i
        Set lo = pa.ListObjects.Add(xlSrcRange, pa.Range("A14:F22"), , xlYes)
        lo.Name = "Tbl_Rollen"
        lo.ListColumns("Rol").DataBodyRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=rol1 & "," & rol2
        For i = 4 To 6
            lo.ListColumns(i).DataBodyRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="ja,nee"
        Next i
        pa.Columns("B").ColumnWidth = 22
        pa.Columns("D:F").ColumnWidth = 12
        nieuwTabel = True
    End If

    ' Rol is afgeleid uit de rollentabel
    ThisWorkbook.Names("Par_HuidigeRol").RefersToRange.Formula = _
        "=IFERROR(INDEX(Tbl_Rollen[Rol],MATCH(Par_HuidigeGebruiker,Tbl_Rollen[Naam],0)),"""")"

    If Not actief Is Nothing Then actief.Activate
    Application.ScreenUpdating = True
    MsgBox "Klaar." & vbCrLf & _
           "Logboek: " & IIf(nieuwLog, "aangemaakt", "bestond al") & vbCrLf & _
           "Rollentabel: " & IIf(nieuwTabel, "aangemaakt op blad Parameters", "bestond al"), vbInformation
    Exit Sub

Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

Public Sub Auto_Open()
    DashboardNaamZorgen                          ' oud blad Dashboard heet nu Dashboard FinAdv
    KaartIndelingProject                         ' invoerkaart: project met deelkredieten boven de omschrijving
    VerzamellijstHernoemen
    ' beveiliging opnieuw aanzetten (macro's mogen schrijven, gebruikers niet) en oude overzichtbladen opruimen
    BeveiligingAan
    On Error Resume Next
    OudeOverzichtbladenOpruimen
    ControlesVernieuwen
    On Error GoTo 0
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then
        If Not UitlegVersieKlopt() Then
            GebruikerKnopPlaatsen
            KnoppenTerugPlaatsen
            KnoppenUitlegPlaatsen
        End If
    End If
    On Error GoTo 0
    ' keuzeopties verversen bij het openen (niet als het bestand alleen-lezen is geopend)
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then KeuzelijstenVerversen True
    If Not ThisWorkbook.ReadOnly Then KredietnummersBijwerken
    On Error GoTo 0
    If TabelBestaat() Then WieWerktEr
    AfgekeurdSignaal                             ' financieel adviseur: melding als er afgekeurde kredieten zijn
    DashboardEventInstalleren True               ' stil: zet de melding voor de financiele administratie aan als dat nog niet zo is
    ' heeft blok 2 de nieuwe indeling (2a, 2b, 2c)? Zo niet: de knoppen van het Dashboard eenmalig opnieuw plaatsen
    If Not ThisWorkbook.ReadOnly Then DashboardIndelingBijwerken
    JaarovergangSignaleren
    ' altijd op het eigen dashboard beginnen (financieel adviseur of financiele administratie)
    NaarEigenDashboard
End Sub

Public Sub WieWerktEr()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim lo As ListObject, r As ListRow
    Dim windows As String, voorstel As String, rol As String
    Dim lijst As String, antw As String, n As Long, k As Long
    Dim namen() As String

    If Not TabelBestaat() Then
        MsgBox "De rollentabel ontbreekt. Voer eerst LogboekEnRollentabelAanmaken uit.", vbExclamation
        Exit Sub
    End If
    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects("Tbl_Rollen")
    windows = Environ("USERNAME")

    If Len(windows) > 0 Then
        For Each r In lo.ListRows
            If StrComp(Trim$(CStr(r.Range.Cells(1, 2).Value)), windows, vbTextCompare) = 0 Then
                voorstel = CStr(r.Range.Cells(1, 1).Value)
                rol = CStr(r.Range.Cells(1, 3).Value)
                Exit For
            End If
        Next r
    End If

    If Len(voorstel) > 0 Then
        If MsgBox("Windows-gebruiker: " & windows & vbCrLf & "Bent u " & voorstel & " (" & rol & ")?", _
                  vbYesNo + vbQuestion, "Wie werkt er?") = vbYes Then
            ZetGebruiker voorstel
            Exit Sub
        End If
    End If

    For Each r In lo.ListRows
        If Len(CStr(r.Range.Cells(1, 1).Value)) > 0 Then
            n = n + 1
            ReDim Preserve namen(1 To n)
            namen(n) = CStr(r.Range.Cells(1, 1).Value)
            lijst = lijst & n & " = " & namen(n) & vbCrLf
        End If
    Next r

    antw = InputBox("Uw Windows-naam is: " & windows & vbCrLf & vbCrLf & _
                    "Typ het nummer van uw naam." & vbCrLf & "Typ 0 als u er niet bij staat." & vbCrLf & _
                    "Laat leeg om alleen te lezen." & vbCrLf & vbCrLf & lijst, "Wie werkt er?")
    If IsNumeric(antw) And Len(antw) > 0 Then
        k = CLng(antw)
        If k = 0 Then
            If RegistreerNieuweGebruiker(lo, windows) Then Exit Sub
        End If
        If k >= 1 And k <= n Then
            ZetGebruiker namen(k)
            OnthoudWindowsNaam lo, namen(k), windows
            Exit Sub
        End If
    End If
    ZetGebruiker ""
    MsgBox "Geen gebruiker gekozen. U kunt alleen lezen.", vbInformation
End Sub

' Bewaart de Windows-naam bij de gekozen persoon, zodat kiezen de volgende keer niet meer nodig is.
Private Sub OnthoudWindowsNaam(ByVal lo As ListObject, ByVal naam As String, ByVal windows As String)
    Dim r As ListRow
    If Len(windows) = 0 Then Exit Sub
    For Each r In lo.ListRows
        If StrComp(CStr(r.Range.Cells(1, 1).Value), naam, vbTextCompare) = 0 Then
            If Len(Trim$(CStr(r.Range.Cells(1, 2).Value))) = 0 Then
                If MsgBox("Windows-naam " & windows & " opslaan bij " & naam & "?" & vbCrLf & _
                          "Dan hoeft u de volgende keer niet meer te kiezen.", vbYesNo + vbQuestion, "Wie werkt er?") = vbYes Then
                    r.Range.Cells(1, 2).Value = windows
                End If
            End If
            Exit Sub
        End If
    Next r
End Sub

' Voegt een nieuwe medewerker toe aan de rollentabel, met rechten die bij de gekozen functie horen.
Private Function RegistreerNieuweGebruiker(ByVal lo As ListObject, ByVal windows As String) As Boolean
    Dim naam As String, keuze As VbMsgBoxResult, rol As String
    Dim toev As String, wijz As String, bev As String
    Dim r As ListRow, lr As ListRow

    naam = Trim$(InputBox("Typ uw voornaam, zoals collega's u kennen.", "Nieuwe gebruiker"))
    If Len(naam) = 0 Then Exit Function
    For Each r In lo.ListRows
        If StrComp(Trim$(CStr(r.Range.Cells(1, 1).Value)), naam, vbTextCompare) = 0 Then
            MsgBox "Deze naam staat al in de lijst. Kies uw nummer in de lijst.", vbExclamation
            Exit Function
        End If
    Next r

    If Len(windows) > 0 Then
        For Each r In lo.ListRows
            If StrComp(Trim$(CStr(r.Range.Cells(1, 2).Value)), windows, vbTextCompare) = 0 Then
                If MsgBox("De Windows-naam " & windows & " staat al bij " & CStr(r.Range.Cells(1, 1).Value) & "." & vbCrLf & _
                          "Bent u " & CStr(r.Range.Cells(1, 1).Value) & "? Kies dan Nee en kies uw nummer in de lijst." & vbCrLf & vbCrLf & _
                          "Gebruikt u dezelfde computer als " & CStr(r.Range.Cells(1, 1).Value) & " en bent u iemand anders?" & vbCrLf & _
                          "Kies dan Ja om u toch toe te voegen.", vbYesNo + vbQuestion + vbDefaultButton2, "Nieuwe gebruiker") <> vbYes Then Exit Function
                Exit For
            End If
        Next r
    End If

    keuze = MsgBox("Wat is uw functie?" & vbCrLf & vbCrLf & "Ja = financieel adviseur" & vbCrLf & _
                   "Nee = financi" & ChrW(235) & "le administratie" & vbCrLf & "Annuleren = stoppen", _
                   vbYesNoCancel + vbQuestion, "Nieuwe gebruiker")
    If keuze = vbCancel Then Exit Function
    If keuze = vbYes Then
        rol = "financieel adviseur": toev = "ja": wijz = "ja": bev = "nee"
    Else
        rol = "financi" & ChrW(235) & "le administratie": toev = "nee": wijz = "ja": bev = "ja"
    End If

    lo.Parent.Unprotect WACHTWOORD                  ' tabel kan niet groeien op een beveiligd blad
    Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 1).Value = naam
    lr.Range.Cells(1, 2).Value = windows
    lr.Range.Cells(1, 3).Value = rol
    lr.Range.Cells(1, 4).Value = toev
    lr.Range.Cells(1, 5).Value = wijz
    lr.Range.Cells(1, 6).Value = bev
    BladBeveiligen lo.Parent
    ZetGebruiker naam
    LogRegel "gebruiker toegevoegd", , , "Rol", , rol, naam & " heeft zichzelf toegevoegd (Windows-naam " & windows & ")"
    MsgBox naam & " is toegevoegd als " & rol & ".", vbInformation
    RegistreerNieuweGebruiker = True
End Function

Private Sub ZetGebruiker(ByVal naam As String)
    ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value = naam
    Application.Calculate
    DashboardNaarRol                             ' blokken tonen of verbergen voor deze rol
End Sub

' Geeft True als de huidige gebruiker het recht heeft: "Toevoegen", "Wijzigen", "Bevestigen" of "Beheer".
' Ontbreekt de kolom van dat recht in de rollentabel, dan is het antwoord nee (geen foutmelding).
' Een applicatiebeheerder (Beheer = ja) heeft alle rechten.
Public Function HeeftRecht(ByVal recht As String) As Boolean
    Dim lo As ListObject, r As ListRow, naam As String, c As Long, b As Long
    If Not TabelBestaat() Then Exit Function
    naam = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    If Len(naam) = 0 Then Exit Function
    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects("Tbl_Rollen")
    On Error Resume Next
    c = lo.ListColumns(recht).Index
    On Error GoTo 0
    If c = 0 Then Exit Function
    On Error Resume Next
    b = lo.ListColumns("Beheer").Index                 ' applicatiebeheerder: heeft alle rechten
    On Error GoTo 0
    For Each r In lo.ListRows
        If StrComp(CStr(r.Range.Cells(1, 1).Value), naam, vbTextCompare) = 0 Then
            HeeftRecht = (LCase$(CStr(r.Range.Cells(1, c).Value)) = "ja")
            If b > 0 Then
                If LCase$(Trim$(CStr(r.Range.Cells(1, b).Value))) = "ja" Then HeeftRecht = True
            End If
            Exit Function
        End If
    Next r
End Function

' Schrijft een volledige logregel op blad Logboek.
Public Sub LogRegel(ByVal actie As String, Optional ByVal volgnr As Variant, Optional ByVal invnr As Variant, _
                    Optional ByVal veld As String = "", Optional ByVal oud As Variant, _
                    Optional ByVal nieuw As Variant, Optional ByVal toelichting As String = "")
    Dim lg As Worksheet, rij As Long
    If Not BladBestaat("Logboek") Then Exit Sub
    Set lg = ThisWorkbook.Worksheets("Logboek")
    rij = lg.Cells(lg.Rows.Count, "A").End(xlUp).Row + 1
    lg.Cells(rij, 1).Value = Now
    lg.Cells(rij, 2).Value = ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value
    lg.Cells(rij, 3).Value = Environ("USERNAME")
    lg.Cells(rij, 4).Value = ThisWorkbook.Names("Par_HuidigeRol").RefersToRange.Value
    lg.Cells(rij, 5).Value = actie
    If Not IsMissing(volgnr) Then lg.Cells(rij, 6).Value = volgnr
    If Not IsMissing(invnr) Then lg.Cells(rij, 7).Value = invnr
    lg.Cells(rij, 8).Value = veld
    If Not IsMissing(oud) Then lg.Cells(rij, 9).Value = CStr(oud)
    If Not IsMissing(nieuw) Then lg.Cells(rij, 10).Value = CStr(nieuw)
    lg.Cells(rij, 11).Value = toelichting
End Sub

Public Function BladBestaat(ByVal naam As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = naam Then BladBestaat = True: Exit Function
    Next ws
End Function

Private Function TabelBestaat() As Boolean
    Dim lo As ListObject
    If Not BladBestaat("Parameters") Then Exit Function
    For Each lo In ThisWorkbook.Worksheets("Parameters").ListObjects
        If lo.Name = "Tbl_Rollen" Then TabelBestaat = True: Exit Function
    Next lo
End Function

' ======================================================================
' Gebruiker verwijderen (stap 6b)
' ======================================================================
' Stap 6b - Gebruiker verwijderen uit de rollentabel
' Alleen voor gebruikers met het recht Toevoegen (financieel adviseurs). Niet jezelf. Met logregel.
' Gebruikt HeeftRecht en LogRegel uit de module van stap 6.

Public Sub GebruikerVerwijderen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim lo As ListObject
    Dim i As Long, n As Long, k As Long
    Dim antw As String, lijst As String, naam As String, rol As String
    Dim idx() As Long

    On Error GoTo Fout
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen gebruikers verwijderen.", vbExclamation
        Exit Sub
    End If

    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects("Tbl_Rollen")
    For i = 1 To lo.ListRows.Count
        If Len(Trim$(CStr(lo.ListRows(i).Range.Cells(1, 1).Value))) > 0 Then
            n = n + 1
            ReDim Preserve idx(1 To n)
            idx(n) = i
            lijst = lijst & n & " = " & CStr(lo.ListRows(i).Range.Cells(1, 1).Value) & _
                    " (" & CStr(lo.ListRows(i).Range.Cells(1, 3).Value) & ")" & vbCrLf
        End If
    Next i
    If n = 0 Then
        MsgBox "Er staan geen gebruikers in de tabel.", vbInformation
        Exit Sub
    End If

    antw = InputBox("Typ het nummer van de persoon die u wilt verwijderen." & vbCrLf & vbCrLf & lijst, "Gebruiker verwijderen")
    If Len(antw) = 0 Then Exit Sub
    If Not IsNumeric(antw) Then
        MsgBox "Dat is geen nummer. Er is niets verwijderd.", vbExclamation
        Exit Sub
    End If
    k = CLng(antw)
    If k < 1 Or k > n Then
        MsgBox "Dat nummer staat niet in de lijst. Er is niets verwijderd.", vbExclamation
        Exit Sub
    End If

    naam = CStr(lo.ListRows(idx(k)).Range.Cells(1, 1).Value)
    rol = CStr(lo.ListRows(idx(k)).Range.Cells(1, 3).Value)
    If StrComp(naam, CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value), vbTextCompare) = 0 Then
        MsgBox "U kunt uzelf niet verwijderen.", vbExclamation
        Exit Sub
    End If

    ' een applicatiebeheerder kan alleen door een applicatiebeheerder worden verwijderd, en de laatste nooit
    Dim cB As Long, nB As Long, j As Long
    On Error Resume Next
    cB = lo.ListColumns("Beheer").Index
    On Error GoTo Fout
    If cB > 0 Then
        If LCase$(Trim$(CStr(lo.ListRows(idx(k)).Range.Cells(1, cB).Value))) = "ja" Then
            If Not HeeftRecht("Beheer") Then
                MsgBox naam & " is applicatiebeheerder. Alleen een applicatiebeheerder kan " & naam & " verwijderen.", vbExclamation
                Exit Sub
            End If
            For j = 1 To lo.ListRows.Count
                If LCase$(Trim$(CStr(lo.ListRows(j).Range.Cells(1, cB).Value))) = "ja" Then nB = nB + 1
            Next j
            If nB <= 1 Then
                MsgBox naam & " is de enige applicatiebeheerder en kan niet worden verwijderd.", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If MsgBox("Weet u zeker dat u " & naam & " (" & rol & ") wilt verwijderen?", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Gebruiker verwijderen") <> vbYes Then Exit Sub

    lo.Parent.Unprotect WACHTWOORD
    lo.ListRows(idx(k)).Delete
    BladBeveiligen lo.Parent
    LogRegel "gebruiker verwijderd", , , "Rol", rol, "", naam & " verwijderd uit de rollentabel"
    MsgBox naam & " is verwijderd.", vbInformation
    Exit Sub

Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' ======================================================================
' Lijstwaarden normaliseren (stap 2)
' ======================================================================
' Stap 2 - Lijstwaarden normaliseren
' Zet tekst in kolom I en J van 'Verzamellijst alle kredieten' om naar kleine letters.
' Getallen in J blijven ongemoeid. Met nulmeting vooraf en achteraf; bij verschil terugdraaien.

Public Sub LijstwaardenNormaliseren()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const BLAD As String = BL_VZ
    VerzamellijstHernoemen
    Dim ws As Worksheet
    Dim laatste As Long, r As Long, n As Long, i As Long
    Dim kol As Variant, v As Variant, nieuw As String
    Dim voorI As Long, voorJ As Long, voorGetalJ As Long
    Dim naI As Long, naJ As Long, naGetalJ As Long
    Dim adr() As String, oud() As String
    Dim lijst As String, spaties As String

    On Error GoTo Fout
    Set ws = ThisWorkbook.Worksheets(BLAD)
    laatste = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If laatste < 2 Then
        MsgBox "Geen gegevens gevonden.", vbExclamation
        Exit Sub
    End If

    ' Nulmeting vooraf
    voorI = Application.WorksheetFunction.CountA(ws.Range("I2:I" & laatste))
    voorJ = Application.WorksheetFunction.CountA(ws.Range("J2:J" & laatste))
    voorGetalJ = Application.WorksheetFunction.Count(ws.Range("J2:J" & laatste))

    Application.ScreenUpdating = False

    For Each kol In Array("I", "J")
        For r = 2 To laatste
            v = ws.Cells(r, kol).Value
            If VarType(v) = vbString Then
                If Len(v) > 0 Then
                    If v <> Trim$(v) Then spaties = spaties & kol & r & " "
                    nieuw = LCase$(v)
                    If nieuw <> v Then
                        n = n + 1
                        ReDim Preserve adr(1 To n)
                        ReDim Preserve oud(1 To n)
                        adr(n) = kol & r
                        oud(n) = v
                        ws.Cells(r, kol).Value = nieuw
                    End If
                End If
            End If
        Next r
    Next kol

    ' Nameting
    naI = Application.WorksheetFunction.CountA(ws.Range("I2:I" & laatste))
    naJ = Application.WorksheetFunction.CountA(ws.Range("J2:J" & laatste))
    naGetalJ = Application.WorksheetFunction.Count(ws.Range("J2:J" & laatste))

    If naI <> voorI Or naJ <> voorJ Or naGetalJ <> voorGetalJ Then
        For i = 1 To n
            ws.Range(adr(i)).Value = oud(i)
        Next i
        Application.ScreenUpdating = True
        MsgBox "Verschil bij de nameting. Alles is teruggezet. Er is niets gewijzigd." & vbCrLf & _
               "Vooraf I/J/getallen J: " & voorI & "/" & voorJ & "/" & voorGetalJ & vbCrLf & _
               "Achteraf: " & naI & "/" & naJ & "/" & naGetalJ, vbCritical
        Exit Sub
    End If

    For i = 1 To n
        If i <= 25 Then lijst = lijst & adr(i) & ": " & oud(i) & " > " & LCase$(oud(i)) & vbCrLf
    Next i
    If n > 25 Then lijst = lijst & "... en " & (n - 25) & " meer" & vbCrLf

    SchrijfLog n, lijst

    Application.ScreenUpdating = True
    MsgBox "Klaar." & vbCrLf & "Aangepaste cellen: " & n & vbCrLf & lijst & vbCrLf & _
           "Nulmeting klopt (I: " & naI & ", J: " & naJ & ", getallen in J: " & naGetalJ & ")." & _
           IIf(Len(spaties) > 0, vbCrLf & vbCrLf & "Let op, spaties voor of achter de tekst (niet aangepast): " & spaties, ""), _
           vbInformation
    Exit Sub

Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' Schrijft een logregel via LogRegel (met gebruiker en rol).
Private Sub SchrijfLog(ByVal aantal As Long, ByVal toelichting As String)
    On Error Resume Next                         ' een logregel mag de macro nooit laten vastlopen
    LogRegel "normaliseren", , , "kolom I en J", , , _
             Left$(aantal & " cellen naar kleine letters" & IIf(Len(toelichting) > 0, ": " & vbLf & toelichting, ""), 30000)
End Sub

' ======================================================================
' Werkbereik verruimen (stap 4)
' ======================================================================
' Stap 4 - Werkbereik verruimen
' Trekt de formulekolommen E, G, K, AD, AL, AM en AN van 'Verzamellijst alle kredieten'
' door tot een door de gebruiker gekozen rij (door te kopieren, nooit opnieuw typen).
' Bestaan de spiegelbladen (Financieel adviseur, Financiele administratie, Kredietoverzicht),
' dan worden die ook doorgetrokken. Ontbreken ze, dan worden ze overgeslagen.
' Met nulmeting vooraf en achteraf. Bij een verschil worden de nieuwe formules weer gewist.

Public Sub WerkbereikVerruimen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const BLAD As String = BL_VZ
    VerzamellijstHernoemen
    Dim ws As Worksheet, sp As Worksheet
    Dim laatsteRij As Long, maxRij As Long, doelRij As Long, bronRij As Long
    Dim standaard As Long, antw As String
    Dim kol As Variant, naam As Variant
    Dim voor As Variant, na As Variant
    Dim i As Long, verschil As Boolean
    Dim calcOud As Long, spBron As Long, spDoel As Long, kolMax As Long, k2 As Long
    Dim gedaan As String, overgeslagen As String

    calcOud = Application.Calculation
    On Error GoTo Fout
    Set ws = ThisWorkbook.Worksheets(BLAD)

    laatsteRij = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)

    If laatsteRij > maxRij Then
        MsgBox "Er staan gegevens onder rij " & maxRij & " (Par_MaxRij). Dat klopt niet met de formules." & vbCrLf & _
               "Controleer eerst blad Parameters. Er is niets gewijzigd.", vbCritical
        Exit Sub
    End If

    standaard = laatsteRij + 1000
    If standaard <= maxRij Then standaard = maxRij + 1000

    antw = InputBox("Tot welke rij moeten de formules worden doorgetrokken?" & vbCrLf & vbCrLf & _
                    "Nu doorgetrokken tot rij: " & maxRij & vbCrLf & _
                    "Laatste gevulde rij: " & laatsteRij, "Werkbereik verruimen", CStr(standaard))
    If Len(antw) = 0 Then Exit Sub
    If Not IsNumeric(antw) Then
        MsgBox "Dat is geen getal. Er is niets gewijzigd.", vbExclamation
        Exit Sub
    End If
    doelRij = CLng(antw)
    If doelRij <= maxRij Or doelRij > 100000 Then
        MsgBox "Kies een rij groter dan " & maxRij & " en kleiner dan 100.000. Er is niets gewijzigd.", vbExclamation
        Exit Sub
    End If

    bronRij = maxRij

    ' Nulmeting vooraf
    voor = Meting(ws)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each kol In Array("E", "G", "K", "AD", "AL", "AM", "AN")
        ws.Range(kol & bronRij).Copy Destination:=ws.Range(kol & (bronRij + 1) & ":" & kol & doelRij)
    Next kol
    Application.CutCopyMode = False

    ' Nameting
    na = Meting(ws)
    For i = LBound(voor) To UBound(voor)
        If voor(i) <> na(i) Then verschil = True
    Next i

    If verschil Then
        For Each kol In Array("E", "G", "K", "AD", "AL", "AM", "AN")
            ws.Range(kol & (bronRij + 1) & ":" & kol & doelRij).ClearContents
        Next kol
        Application.Calculation = calcOud
        Application.ScreenUpdating = True
        MsgBox "Verschil bij de nameting. De nieuwe formules zijn weer gewist. Er is niets gewijzigd.", vbCritical
        Exit Sub
    End If

    ' Spiegelbladen (alleen als ze bestaan)
    For Each naam In Array("Financieel adviseur", "Financi" & ChrW(235) & "le administratie", "Kredietoverzicht")
        Set sp = Nothing
        On Error Resume Next
        Set sp = ThisWorkbook.Worksheets(CStr(naam))
        On Error GoTo Fout
        If sp Is Nothing Then
            overgeslagen = overgeslagen & naam & ", "
        Else
            spDoel = doelRij
            If CStr(naam) = "Kredietoverzicht" Then spDoel = doelRij + 1
            spBron = sp.Cells(sp.Rows.Count, "A").End(xlUp).Row
            kolMax = sp.Cells(1, sp.Columns.Count).End(xlToLeft).Column
            k2 = sp.Cells(2, sp.Columns.Count).End(xlToLeft).Column
            If k2 > kolMax Then kolMax = k2
            If spBron < spDoel Then
                sp.Range(sp.Cells(spBron, 1), sp.Cells(spBron, kolMax)).Copy _
                    Destination:=sp.Range(sp.Cells(spBron + 1, 1), sp.Cells(spDoel, kolMax))
                Application.CutCopyMode = False
            End If
            gedaan = gedaan & naam & ", "
        End If
    Next naam

    ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value = doelRij
    Application.Calculation = calcOud
    Application.Calculate
    Application.ScreenUpdating = True

    SchrijfLogWerkbereik "Werkbereik verruimd van rij " & maxRij & " naar rij " & doelRij

    antw = "Klaar." & vbCrLf & "Formules doorgetrokken tot rij " & doelRij & " (was " & maxRij & ")." & vbCrLf & _
           "Nulmeting klopt (gevulde regels: " & na(0) & ")." & vbCrLf
    If Len(gedaan) > 0 Then antw = antw & "Spiegelbladen doorgetrokken: " & Left(gedaan, Len(gedaan) - 2) & vbCrLf
    If Len(overgeslagen) > 0 Then antw = antw & "Niet aanwezig, overgeslagen: " & Left(overgeslagen, Len(overgeslagen) - 2) & vbCrLf
    antw = antw & vbCrLf & "Let op: een eventueel ingesteld filter op de verzamellijst loopt nog tot de oude laatste rij."
    MsgBox antw, vbInformation
    Exit Sub

Fout:
    Application.CutCopyMode = False
    Application.Calculation = calcOud
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' Nulmeting: aantal gevulde regels in kolom A en de sommen van Z, AA, AB, AC, AE, AF, AG, AH, AI.
Private Function Meting(ByVal ws As Worksheet) As Variant
    Dim res(0 To 9) As Double, kol As Variant, i As Long
    res(0) = Application.WorksheetFunction.CountA(ws.Range("A2:A" & ws.Rows.Count))
    i = 1
    For Each kol In Array("Z", "AA", "AB", "AC", "AE", "AF", "AG", "AH", "AI")
        res(i) = Application.WorksheetFunction.Sum(ws.Columns(CStr(kol)))
        i = i + 1
    Next kol
    Meting = res
End Function

' Schrijft een logregel via LogRegel (met gebruiker en rol).
Private Sub SchrijfLogWerkbereik(ByVal toelichting As String)
    On Error Resume Next                         ' een logregel mag de macro nooit laten vastlopen
    LogRegel "werkbereik verruimen", , , "", , , toelichting
End Sub

' ======================================================================
' Controlekolommen (stap 5)
' ======================================================================
' Stap 5 - Controlekolommen aanmaken en bestaande regels vullen
' Zet de koppen in AP t/m AW van 'Verzamellijst alle kredieten' en geeft alle bestaande
' regels de controlestatus 'Gecontroleerd door financiele administratie'.
' Met nulmeting vooraf en achteraf. Draait niet opnieuw als AP t/m AW al gevuld zijn.

Public Function CtlGecontroleerd() As String
    CtlGecontroleerd = "Gecontroleerd door financi" & ChrW(235) & "le administratie"
End Function

Public Function CtlTeControleren() As String
    CtlTeControleren = "Te controleren door financi" & ChrW(235) & "le administratie"
End Function

Public Sub ControlekolommenAanmaken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const BLAD As String = BL_VZ
    VerzamellijstHernoemen
    Dim ws As Worksheet
    Dim laatste As Long, maxRij As Long, i As Long
    Dim koppen As Variant
    Dim voorA As Double, voorZ As Double, voorAE As Double
    Dim naA As Double, naZ As Double, naAE As Double
    Dim antw As String

    On Error GoTo Fout
    Set ws = ThisWorkbook.Worksheets(BLAD)
    laatste = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)

    If Application.WorksheetFunction.CountA(ws.Range("AP1:AW" & maxRij)) > 0 Then
        MsgBox "De kolommen AP t/m AW bevatten al gegevens. Er is niets gewijzigd.", vbExclamation
        Exit Sub
    End If

    ' Nulmeting vooraf
    voorA = Application.WorksheetFunction.CountA(ws.Range("A2:A" & laatste))
    voorZ = Application.WorksheetFunction.Sum(ws.Columns("Z"))
    voorAE = Application.WorksheetFunction.Sum(ws.Range("AE:AI"))

    Application.ScreenUpdating = False

    koppen = Array("Controlestatus", "Gecontroleerd door", "Gecontroleerd op", "Opmerking controle", _
                   "(reserve)", "(reserve)", "Laatst gewijzigd door", "Laatst gewijzigd op")
    For i = 0 To 7
        ws.Cells(1, 42 + i).Value = koppen(i)
    Next i

    ' Opmaak: koptekst en cellen gelijk aan kolom AJ
    ws.Range("AJ1").Copy
    ws.Range("AP1:AW1").PasteSpecial xlPasteFormats
    ws.Range("AJ2").Copy
    ws.Range("AP2:AW" & maxRij).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("AR2:AR" & maxRij).NumberFormat = "dd-mm-yyyy"
    ws.Range("AW2:AW" & maxRij).NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Columns("AP").ColumnWidth = 42
    ws.Columns("AQ").ColumnWidth = 22
    ws.Columns("AR").ColumnWidth = 14
    ws.Columns("AS").ColumnWidth = 42
    ws.Columns("AT:AU").ColumnWidth = 10
    ws.Columns("AV").ColumnWidth = 22
    ws.Columns("AW").ColumnWidth = 18

    ' Eenmalige vulling van de bestaande regels
    ws.Range("AP2:AP" & laatste).Value = CtlGecontroleerd()
    ws.Range("AS2:AS" & laatste).Value = "bestaande regel bij invoering controlemodule"

    ' Nameting
    naA = Application.WorksheetFunction.CountA(ws.Range("A2:A" & laatste))
    naZ = Application.WorksheetFunction.Sum(ws.Columns("Z"))
    naAE = Application.WorksheetFunction.Sum(ws.Range("AE:AI"))
    If naA <> voorA Or naZ <> voorZ Or naAE <> voorAE Then
        ws.Range("AP1:AW" & maxRij).Clear
        Application.ScreenUpdating = True
        MsgBox "Verschil bij de nameting. De nieuwe kolommen zijn weer gewist.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = True
    antw = "Klaar." & vbCrLf & "Koppen gezet in AP t/m AW." & vbCrLf & _
           (laatste - 1) & " bestaande regels staan nu op 'Gecontroleerd door financiele administratie'." & vbCrLf & _
           "Nulmeting klopt (gevulde regels: " & naA & ")."
    MsgBox antw, vbInformation
    Exit Sub

Fout:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' ======================================================================
' Dashboard en keuzelijsten (stap 7)
' ======================================================================
' Stap 7 - Dashboard (opmaak, invoerkaart) en keuzelijsten
' KeuzelijstenVerversen : vult de lijsten op blad Parameters (kolom H t/m R) met de unieke waarden
'                         uit de verzamellijst en legt de namen Lst_... vast.
' DashboardAanmaken     : maakt het blad Dashboard (eenmalig) met de verticale invoerkaart en keuzelijsten.


Public Sub KeuzelijstenVerversen(Optional ByVal stil As Boolean = False)
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wsV As Worksheet, wsP As Worksheet
    Dim laatste As Long, i As Long
    Dim cfg As Variant

    On Error GoTo Fout
    Set wsV = Vz()
    Set wsP = ThisWorkbook.Worksheets("Parameters")
    laatste = wsV.Cells(wsV.Rows.Count, "A").End(xlUp).Row
    Application.ScreenUpdating = False

    ProgrammaTabelAanmaken
    TaakveldTabelAanmaken
    BalanscodeTabelAanmaken
    BijdrageBronTabelAanmaken
    wsP.Range("H14:H5000").ClearContents
    wsP.Range("M14:N5000").ClearContents
    wsP.Range("V14:V5000").ClearContents
    On Error Resume Next
    ThisWorkbook.Names("Lst_Exploitatie").Delete
    On Error GoTo Fout
    If Not stil Then ControleCodes wsP, wsV, laatste

    ' doelkolom, naam, bronkolom, naar kleine letters (1) of niet (0)
    cfg = Array(Array("I", "Lst_Soort", "I", 1), _
                Array("K", "Lst_Budgetbeheerder", "O", 0), _
                Array("L", "Lst_Projectleider", "P", 0), _
                Array("O", "Lst_Organisatieonderdeel", "U", 0), _
                Array("P", "Lst_BudgetbeheerderAdm", "V", 0), Array("Q", "Lst_BtwBcf", "W", 0), _
                Array("T", "Lst_LaatsteBesluit", "L", 0), Array("U", "Lst_Tijdschrijver", "Q", 2))
    GroepenLijstSchrijven wsP, wsV, laatste
    For i = 0 To UBound(cfg)
        SchrijfLijst wsP, wsV, laatste, CStr(cfg(i)(0)), CStr(cfg(i)(1)), CStr(cfg(i)(2)), CLng(cfg(i)(3))
    Next i

    ' Vaste lijst voor de statusletter
    wsP.Range("R14:R5000").ClearContents
    wsP.Range("R14").Value = "Lst_Statusletter"
    wsP.Range("R15").Value = "L"
    wsP.Range("R16").Value = "A"
    ThisWorkbook.Names.Add Name:="Lst_Statusletter", RefersTo:="=Parameters!$R$15:$R$16"

    wsP.Range("S14:S5000").ClearContents
    wsP.Range("S14").Value = "Lst_Jaren"
    wsP.Range("S15:S29").Formula = "=Par_LopendJaar-4+ROW()-15"
    ThisWorkbook.Names.Add Name:="Lst_Jaren", RefersTo:="=Parameters!$S$15:$S$29"

    wsP.Range("A11").Value = "Keuzeopties laatst ververst"
    wsP.Range("B11").NumberFormat = "dd-mm-yyyy hh:mm"
    wsP.Range("B11").Value = Now
    wsP.Range("W14:W5000").ClearContents
    wsP.Range("W15:W75").NumberFormat = "General"
    wsP.Range("W14").Value = "Lst_Afschrijving"
    For i = 0 To 60
        wsP.Cells(15 + i, 23).Value = i
    Next i
    ThisWorkbook.Names.Add Name:="Lst_Afschrijving", RefersTo:="=Parameters!$W$15:$W$75"

    wsP.Range("H14:W14").Font.Bold = True
    wsP.Columns("H:W").ColumnWidth = 24
    KleurParameters
    On Error Resume Next
    ControleLijstVernieuwen
    On Error GoTo Fout
    Application.ScreenUpdating = True
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout bij het verversen van de keuzelijsten: " & Err.Description, vbCritical
End Sub

' Maakt (eenmalig) de programmatabel op blad Parameters. Pas die aan na de fusie.
Private Sub ProgrammaTabelAanmaken()
    Dim wsP As Worksheet, lo As ListObject, bestaat As Boolean
    Dim oms As Variant, i As Long
    Set wsP = ThisWorkbook.Worksheets("Parameters")
    For Each lo In wsP.ListObjects
        If lo.Name = "Tbl_Programma" Then bestaat = True
    Next lo
    If Not bestaat Then
        wsP.Range("Y12").Value = "Programmatabel: pas deze aan na de fusie. Code en omschrijving staan samen in de keuzelijst."
        wsP.Range("Y12").Font.Bold = True
        wsP.Range("Y14").Value = "Code"
        wsP.Range("Z14").Value = "Omschrijving"
        wsP.Range("AA14").Value = "Weergave"
        oms = Array("Algemeen bestuur", "Openbare orde en veiligheid", "Verkeer en vervoer", "Economische zaken", _
                    "Onderwijs", "Cultuur, sport en recreatie", "Sociaal domein", "Volksgezondheid en milieu", "Ruimte en bouwen")
        For i = 0 To 8
            wsP.Cells(15 + i, 25).Value = i
            wsP.Cells(15 + i, 26).Value = oms(i)
        Next i
        Set lo = wsP.ListObjects.Add(xlSrcRange, wsP.Range("Y14:AA23"), , xlYes)
        lo.Name = "Tbl_Programma"
        lo.ListColumns("Weergave").DataBodyRange.Formula = "=[@Code]&"" - ""&[@Omschrijving]"
        wsP.Columns("Y").ColumnWidth = 8
        wsP.Columns("Z").ColumnWidth = 36
        wsP.Columns("AA").ColumnWidth = 44
    End If
    ThisWorkbook.Names.Add Name:="Lst_Programma", RefersTo:="=Tbl_Programma[Weergave]"
End Sub

' Meldt codes in de verzamellijst die niet in de bijbehorende tabel op blad Parameters staan.
Private Sub ControleCodes(ByVal wsP As Worksheet, ByVal wsV As Worksheet, ByVal laatste As Long)
    Dim tekst As String
    tekst = tekst & CodesOntbreken(wsP, wsV, laatste, 2, "Tbl_Programma", "Programma")
    tekst = tekst & CodesOntbreken(wsP, wsV, laatste, 19, "Tbl_Taakveld", "Taakveld")
    tekst = tekst & CodesOntbreken(wsP, wsV, laatste, 20, "Tbl_Balanscode", "Balanscode")
    If Len(tekst) > 0 Then
        Application.ScreenUpdating = True
        MsgBox "Deze codes staan in de verzamellijst maar niet in de tabel op blad Parameters:" & vbCrLf & vbCrLf & tekst, vbExclamation
        Application.ScreenUpdating = False
    End If
End Sub

Private Function CodesOntbreken(ByVal wsP As Worksheet, ByVal wsV As Worksheet, ByVal laatste As Long, _
                                ByVal bronKolNr As Long, ByVal tabel As String, ByVal label As String) As String
    Dim cel As Range, r As Long, v As Variant, bekend As String, ontbreekt As String
    For Each cel In wsP.ListObjects(tabel).ListColumns("Code").DataBodyRange.Cells
        bekend = bekend & "|" & CStr(cel.Value) & "|"
    Next cel
    For r = 2 To laatste
        v = wsV.Cells(r, bronKolNr).Value
        If Not IsEmpty(v) And Not IsError(v) Then
            If Len(CStr(v)) > 0 Then
                If InStr(bekend, "|" & CStr(v) & "|") = 0 Then
                    If InStr(ontbreekt, "[" & CStr(v) & "]") = 0 Then ontbreekt = ontbreekt & "[" & CStr(v) & "]"
                End If
            End If
        End If
    Next r
    If Len(ontbreekt) > 0 Then CodesOntbreken = label & ": " & ontbreekt & vbCrLf
End Function

' Taakvelden volgens het Iv3-informatievoorschrift gemeenten, verslagjaar 2026.
Private Sub TaakveldTabelAanmaken()
    Dim wsP As Worksheet, lo As ListObject, bestaat As Boolean
    Dim t(1 To 69) As String, i As Long, n As Long, delen() As String
    Set wsP = ThisWorkbook.Worksheets("Parameters")
    For Each lo In wsP.ListObjects
        If lo.Name = "Tbl_Taakveld" Then bestaat = True
    Next lo
    If Not bestaat Then
        t(1) = "0.1|Bestuur"
        t(2) = "0.2|Burgerzaken"
        t(3) = "0.3|Beheer overige gebouwen en gronden"
        t(4) = "0.4|Overhead"
        t(5) = "0.5|Treasury"
        t(6) = "0.61|OZB woningen"
        t(7) = "0.62|OZB niet-woningen"
        t(8) = "0.63|Parkeerbelasting"
        t(9) = "0.64|Belastingen overig"
        t(10) = "0.7|Algemene uitkering en overige uitkeringen gemeentefonds"
        t(11) = "0.8|Overige baten en lasten"
        t(12) = "0.9|Vennootschapsbelasting (Vpb)"
        t(13) = "0.10|Mutaties reserves"
        t(14) = "0.11|Resultaat van de rekening van baten en lasten"
        t(15) = "1.1|Crisisbeheersing en brandweer"
        t(16) = "1.2|Openbare orde en veiligheid"
        t(17) = "2.1|Verkeer en vervoer"
        t(18) = "2.2|Parkeren"
        t(19) = "2.3|Recreatieve havens"
        t(20) = "2.4|Economische havens en waterwegen"
        t(21) = "2.5|Openbaar vervoer"
        t(22) = "3.1|Economische ontwikkeling"
        t(23) = "3.2|Fysieke bedrijfsinfrastructuur"
        t(24) = "3.3|Bedrijvenloket en bedrijfsregelingen"
        t(25) = "3.4|Economische promotie"
        t(26) = "4.1|Openbaar basisonderwijs"
        t(27) = "4.2|Onderwijshuisvesting"
        t(28) = "4.3|Onderwijsbeleid en leerlingzaken"
        t(29) = "5.1|Sportbeleid en activering"
        t(30) = "5.2|Sportaccommodaties"
        t(31) = "5.3|Cultuurpresentatie, cultuurproductie en cultuurparticipatie"
        t(32) = "5.4|Musea"
        t(33) = "5.5|Cultureel erfgoed"
        t(34) = "5.6|Media"
        t(35) = "5.7|Openbaar groen en (openlucht) recreatie"
        t(36) = "6.1|Samenkracht en burgerparticipatie"
        t(37) = "6.21|Toegang en eerstelijnsvoorzieningen Wmo"
        t(38) = "6.22|Toegang en eerstelijnsvoorzieningen Jeugd"
        t(39) = "6.23|Toegang en eerstelijnsvoorzieningen Integraal"
        t(40) = "6.3|Inkomensregelingen"
        t(41) = "6.4|WSW en beschut werk"
        t(42) = "6.5|Arbeidsparticipatie"
        t(43) = "6.60|Hulpmiddelen en diensten (Wmo)"
        t(44) = "6.711|Huishoudelijke hulp (Wmo)"
        t(45) = "6.712|Begeleiding (Wmo)"
        t(46) = "6.713|Dagbesteding (Wmo)"
        t(47) = "6.714|Overige maatwerkarrangementen (Wmo)"
        t(48) = "6.751|Jeugdhulp ambulant lokaal"
        t(49) = "6.752|Jeugdhulp ambulant regionaal"
        t(50) = "6.753|Jeugdhulp ambulant landelijk"
        t(51) = "6.761|Jeugdhulp met verblijf lokaal"
        t(52) = "6.762|Jeugdhulp met verblijf regionaal"
        t(53) = "6.763|Jeugdhulp met verblijf landelijk"
        t(54) = "6.791|PGB Wmo"
        t(55) = "6.792|PGB Jeugd"
        t(56) = "6.811|Beschermd wonen (Wmo)"
        t(57) = "6.812|Maatschappelijke- en vrouwenopvang (Wmo)"
        t(58) = "6.821|Jeugdbescherming"
        t(59) = "6.822|Jeugdreclassering"
        t(60) = "6.91|Co" & ChrW(246) & "rdinatie en beleid Wmo"
        t(61) = "6.92|Co" & ChrW(246) & "rdinatie en beleid Jeugd"
        t(62) = "7.1|Volksgezondheid"
        t(63) = "7.2|Riolering"
        t(64) = "7.3|Afval"
        t(65) = "7.4|Milieubeheer"
        t(66) = "7.5|Begraafplaatsen en crematoria"
        t(67) = "8.1|Ruimte en leefomgeving"
        t(68) = "8.2|Grondexploitatie (niet-bedrijventerreinen)"
        t(69) = "8.3|Wonen en bouwen"
        n = 69
        wsP.Range("AC12").Value = "Taakvelden BBV (Iv3-informatievoorschrift gemeenten, verslagjaar 2026). De code zonder punt wordt in de verzamellijst gezet."
        wsP.Range("AC12").Font.Bold = True
        wsP.Range("AC14").Value = "Taakveld"
        wsP.Range("AD14").Value = "Omschrijving"
        wsP.Range("AE14").Value = "Code"
        wsP.Range("AF14").Value = "Weergave"
        wsP.Range("AC15:AC" & (14 + n)).NumberFormat = "@"
        For i = 1 To n
            delen = Split(t(i), "|")
            wsP.Cells(14 + i, 29).Value = delen(0)
            wsP.Cells(14 + i, 30).Value = delen(1)
        Next i
        Set lo = wsP.ListObjects.Add(xlSrcRange, wsP.Range("AC14:AF" & (14 + n)), , xlYes)
        lo.Name = "Tbl_Taakveld"
        lo.ListColumns("Code").DataBodyRange.Formula = "=SUBSTITUTE([@Taakveld],""."","""")"
        lo.ListColumns("Weergave").DataBodyRange.Formula = "=[@Taakveld]&"" - ""&[@Omschrijving]"
        wsP.Columns("AC").ColumnWidth = 10
        wsP.Columns("AD").ColumnWidth = 60
        wsP.Columns("AE").ColumnWidth = 8
        wsP.Columns("AF").ColumnWidth = 70
    End If
    ThisWorkbook.Names.Add Name:="Lst_Taakveld", RefersTo:="=Tbl_Taakveld[Weergave]"
End Sub

' Balanscodes: Iv3-balansstanden materiele vaste activa (A121 t/m A129) met nutsindeling volgens art. 35 BBV.
Private Sub BalanscodeTabelAanmaken()
    Dim wsP As Worksheet, lo As ListObject, bestaat As Boolean
    Dim t As Variant, i As Long, delen() As String
    Set wsP = ThisWorkbook.Worksheets("Parameters")
    For Each lo In wsP.ListObjects
        If lo.Name = "Tbl_Balanscode" Then bestaat = True
    Next lo
    If Not bestaat Then
        t = Array("A121EN|Gronden en terreinen (economisch nut)", _
                  "A122EN|Woonruimten (economisch nut)", _
                  "A123EN|Bedrijfsgebouwen (economisch nut)", _
                  "A124EN|Grond-, weg- en waterbouwkundige werken (economisch nut)", _
                  "A124ENH|Grond-, weg- en waterbouwkundige werken (economisch nut, heffing, bijvoorbeeld riolering en afval)", _
                  "A124MN1|Grond-, weg- en waterbouwkundige werken (maatschappelijk nut, openbare ruimte)", _
                  "A125EN|Vervoermiddelen (economisch nut)", _
                  "A126EN|Machines, apparaten en installaties (economisch nut)", _
                  "A129EN|Overige materi" & ChrW(235) & "le vaste activa (economisch nut)")
        wsP.Range("AH12").Value = "Balanscodes BBV (materiele vaste activa, Iv3 A121 t/m A129 met nutsindeling art. 35 BBV). Voeg onderaan codes toe."
        wsP.Range("AH12").Font.Bold = True
        wsP.Range("AH14").Value = "Code"
        wsP.Range("AI14").Value = "Omschrijving"
        wsP.Range("AJ14").Value = "Weergave"
        For i = 0 To UBound(t)
            delen = Split(CStr(t(i)), "|")
            wsP.Cells(15 + i, 34).Value = delen(0)
            wsP.Cells(15 + i, 35).Value = delen(1)
        Next i
        Set lo = wsP.ListObjects.Add(xlSrcRange, wsP.Range("AH14:AJ" & (14 + UBound(t) + 1)), , xlYes)
        lo.Name = "Tbl_Balanscode"
        lo.ListColumns("Weergave").DataBodyRange.Formula = "=[@Code]&"" - ""&[@Omschrijving]"
        wsP.Columns("AH").ColumnWidth = 10
        wsP.Columns("AI").ColumnWidth = 90
        wsP.Columns("AJ").ColumnWidth = 100
    End If
    ThisWorkbook.Names.Add Name:="Lst_Balanscode", RefersTo:="=Tbl_Balanscode[Weergave]"
End Sub

Private Sub SchrijfLijst(ByVal wsP As Worksheet, ByVal wsV As Worksheet, ByVal laatste As Long, _
                         ByVal doelKol As String, ByVal naam As String, ByVal bronKol As String, ByVal modus As Long)
    ' modus: 0 = zoals het staat, 1 = kleine letters, 2 = splitsen op komma, 3 = alleen getallen
    Dim lijst As New Collection
    Dim r As Long, i As Long, n As Long, kolNr As Long, eind As Long, p As Long
    Dim v As Variant, delen As Variant

    For r = 2 To laatste
        v = wsV.Range(bronKol & r).Value
        If Not IsError(v) Then
            If VarType(v) = vbString Then
                v = Trim$(v)
                If modus = 1 Then v = LCase$(v)
                If modus = 2 Then
                    delen = Split(v, ",")
                    For p = LBound(delen) To UBound(delen)
                        VoegToe lijst, Trim$(CStr(delen(p)))
                    Next p
                ElseIf modus <> 3 Then
                    VoegToe lijst, v
                End If
            Else
                VoegToe lijst, v
            End If
        End If
    Next r

    wsP.Range(doelKol & "14:" & doelKol & "5000").ClearContents
    wsP.Range(doelKol & "14").Value = naam
    kolNr = wsP.Range(doelKol & "1").Column
    n = lijst.Count
    For i = 1 To n
        If VarType(lijst(i)) = vbString Then wsP.Cells(14 + i, kolNr).NumberFormat = "@"
        wsP.Cells(14 + i, kolNr).Value = lijst(i)
    Next i
    If n > 1 Then
        wsP.Range(doelKol & "15:" & doelKol & (14 + n)).Sort Key1:=wsP.Range(doelKol & "15"), _
            Order1:=xlAscending, Header:=xlNo
    End If
    eind = 14 + n
    If n = 0 Then eind = 15
    ThisWorkbook.Names.Add Name:=naam, RefersTo:="=Parameters!$" & doelKol & "$15:$" & doelKol & "$" & eind
End Sub

' Keuzelijst voor het veld samengesteld/standaard/stelpost, met de bestaande samengestelde kredieten (groepen).
Private Sub GroepenLijstSchrijven(ByVal wsP As Worksheet, ByVal wsV As Worksheet, ByVal laatste As Long)
    Dim r As Long, n As Long, maxN As Long, rij As Long, j As Variant
    Dim cnt() As Long, nm() As String, eerste() As String
    wsP.Range("J14:J5000").ClearContents
    wsP.Range("J14").Value = "Lst_SamengesteldStandaardStelpost"
    wsP.Range("J15").Value = "standaard"
    wsP.Range("J16").Value = "stelpost"
    wsP.Range("J17").Value = "nieuw project met deelkredieten"
    For r = 2 To laatste
        j = wsV.Cells(r, 10).Value
        If VarType(j) <> vbString And Not IsEmpty(j) And Not IsError(j) Then
            If IsNumeric(j) Then If CLng(j) > maxN Then maxN = CLng(j)
        End If
    Next r
    rij = 18                                     ' ook als er geen groepsnummer is (lijst eindigt dan op J17)
    If maxN > 0 Then
        ReDim cnt(1 To maxN)
        ReDim nm(1 To maxN)
        ReDim eerste(1 To maxN)
        For r = 2 To laatste
            j = wsV.Cells(r, 10).Value
            If VarType(j) <> vbString And Not IsEmpty(j) And Not IsError(j) Then
                If IsNumeric(j) Then
                    n = CLng(j)
                    If n >= 1 Then
                        cnt(n) = cnt(n) + 1
                        If Len(eerste(n)) = 0 Then eerste(n) = CStr(wsV.Cells(r, 4).Value)
                        If Len(nm(n)) = 0 Then nm(n) = Trim$(CStr(wsV.Cells(r, 46).Value))
                    End If
                End If
            End If
        Next r
        For n = maxN To 1 Step -1                ' nieuwste samengestelde krediet (hoogste groepsnummer) bovenaan
            If cnt(n) > 0 Then
                If Len(nm(n)) = 0 Then nm(n) = eerste(n)
                wsP.Cells(rij, 10).NumberFormat = "@"
                wsP.Cells(rij, 10).Value = "groep " & n & " - " & nm(n) & " (" & cnt(n) & " deelkredieten)"
                rij = rij + 1
            End If
        Next n
    End If
    ThisWorkbook.Names.Add Name:="Lst_SamengesteldStandaardStelpost", RefersTo:="=Parameters!$J$15:$J$" & (rij - 1)
End Sub

Private Sub VoegToe(ByRef lijst As Collection, ByVal v As Variant)
    If IsEmpty(v) Then Exit Sub
    If Len(CStr(v)) = 0 Then Exit Sub
    On Error Resume Next
    lijst.Add v, TypeName(v) & "|" & LCase$(CStr(v))
    On Error GoTo 0
End Sub

Private Function veld(ByVal label As String, ByVal kol As String, ByVal lijst As String, _
                      ByVal verplicht As Boolean, ByVal fmt As String) As Variant
    veld = Array(label, kol, lijst, verplicht, fmt)
End Function

Public Sub DashboardAanmaken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, w As Worksheet
    Dim v(1 To 30) As Variant
    Dim i As Long, r As Long, n As Long, kc As Long
    Dim f As String, Leeg As String, lijstNaam As String
    Dim fc As Object
    Dim bj As Long, bk As Long, brr As Long, blst As String
    Dim actief As Object

    For Each w In ThisWorkbook.Worksheets
        If w.Name = BL_DASH Then
            If MsgBox("Het Dashboard bestaat al. Opnieuw opbouwen?" & vbCrLf & _
                      "Wat op de invoerkaart staat, gaat verloren.", vbYesNo + vbQuestion + vbDefaultButton2, "Dashboard") <> vbYes Then Exit Sub
            Application.DisplayAlerts = False
            w.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next w

    On Error GoTo Fout
    Application.ScreenUpdating = False
    KeuzelijstenVerversen
    Application.ScreenUpdating = False

    v(1) = veld("Programma", "B", "Lst_Programma", True, "@")
    v(2) = veld("Kredietnummer", "C", "", False, "")
    v(3) = veld("Omschrijving (deel)krediet (max 36)", "D", "", True, "@")
    v(4) = veld("Omschrijving kort (max 20)", "F", "", False, "@")
    v(5) = veld("Status (L of A)", "H", "Lst_Statusletter", True, "@")
    v(6) = veld("Jaar van de status", "H", "Lst_Jaren", True, "@")
    v(7) = veld("Investeringsgroep", "I", "Lst_Soort", True, "@")
    v(8) = veld("Project met deelkredieten / standaard / stelpost", "J", "Lst_SamengesteldStandaardStelpost", True, "")
    v(9) = veld("Besluit", "L", "Lst_LaatsteBesluit", True, "@")
    v(10) = veld("Startjaar uitgaven", "M", "Lst_Jaren", True, "")
    v(11) = veld("Eindjaar uitgaven", "N", "Lst_Jaren", True, "")
    v(12) = veld("Budgetbeheerder", "O", "Lst_Budgetbeheerder", True, "@")
    v(13) = veld("Projectleider", "P", "Lst_Projectleider", True, "@")
    v(14) = veld("Tijdschrijvers", "Q", "Lst_Tijdschrijver", True, "@")
    v(15) = veld("Exploitatienummer", "R", "", True, "")
    v(16) = veld("Taakveld BBV", "S", "Lst_Taakveld", True, "@")
    v(17) = veld("Balanscode BBV", "T", "Lst_Balanscode", True, "@")
    v(18) = veld("Organisatie-onderdeel", "U", "Lst_Organisatieonderdeel", True, "@")
    v(19) = veld("Budgetbeheerder administratie", "V", "Lst_BudgetbeheerderAdm", True, "@")
    v(20) = veld("Code btw/bcf administratie", "W", "Lst_BtwBcf", True, "@")
    v(21) = veld("Activanummer", "X", "", False, "")
    v(22) = veld("Afschrijftermijn (0-60 jaar)", "Y", "Lst_Afschrijving", True, "")
    v(23) = veld("=IF(Kaart_Rijnummer="""",""Budget ""&Par_LopendJaar,""Budget tot en met ""&Par_LopendJaar)", "Z", "", False, "bedrag")
    v(24) = veld("=""Budget ""&(Par_LopendJaar+1)", "AE", "", False, "bedrag")
    v(25) = veld("=""Budget ""&(Par_LopendJaar+2)", "AF", "", False, "bedrag")
    v(26) = veld("=""Budget ""&(Par_LopendJaar+3)", "AG", "", False, "bedrag")
    v(27) = veld("=""Budget ""&(Par_LopendJaar+4)", "AH", "", False, "bedrag")
    v(28) = veld("=""Budget ""&(Par_LopendJaar+5)", "AI", "", False, "bedrag")
    v(29) = veld("Toelichting op het investeringskrediet", "AJ", "", False, "@")
    v(30) = veld("Naam project met deelkredieten", "AT", "", False, "@")

    Set actief = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = BL_DASH
    ' Namen voor de latere macro's
    ThisWorkbook.Names.Add Name:="Kaart_Kolommen", RefersTo:="='" & BL_DASH & "'!$A$" & eerste & ":$A$" & (eerste + 23)
    ThisWorkbook.Names.Add Name:="Kaart_Invoer", RefersTo:="='" & BL_DASH & "'!$C$" & eerste & ":$C$" & (eerste + 23) & _
                                                           ",'" & BL_DASH & "'!$F$" & (eerste + 4) & ":$F$" & (eerste + 21)
    ThisWorkbook.Names.Add Name:="Kaart_Rijnummer", RefersTo:="='" & BL_DASH & "'!$F$" & (eerste + 1)

    ws.Tab.Color = RGB(0, 112, 192)
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C").ColumnWidth = 44
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 36
    ws.Columns("F").ColumnWidth = 18

    ' Kop
    ws.Range("B1").Value = "Dashboard investeringskredieten"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B2").Value = "Gebruiker"
    ws.Range("C2").Formula = "=Par_HuidigeGebruiker"
    ws.Range("C2").Formula = "=Par_HuidigeGebruiker&IF(Par_HuidigeRol="""","""","" (""&Par_HuidigeRol&"")"")"
    ws.Range("B3").Value = "Lopend jaar"
    ws.Range("C3").Formula = "=Par_LopendJaar"
    ws.Range("C3").HorizontalAlignment = xlLeft
    ws.Range("E3").Value = "Keuzeopties ververst op"
    ws.Range("E3").HorizontalAlignment = xlRight
    ws.Range("F3").Formula = "=Parameters!B11"
    ws.Range("F3").NumberFormat = "dd-mm-yyyy hh:mm"           ' datum apart, anders past het niet (####)
    ws.Range("F3").HorizontalAlignment = xlLeft
    ws.Range("B4").Value = "Controlestatus"
    ws.Range("C4").Value = ""                                 ' statusregel: zie ControleblokPlaatsen (stap 12)
    f = "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)&"" regels te controleren, ""&COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")&"" afgekeurd"""
    ws.Range("B5").Formula = f
    ws.Range("B2:B4").Font.Bold = True
    Set fc = VoegCF(ws.Range("B5"), "=COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")>0")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 150)
    Set fc = VoegCF(ws.Range("B5"), "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)>0")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 200, 120)

    ' Blok 1: invoerkaart. Links de gewone velden (B = veld, C = invoer), rechts de budgetten met bijdragen
    ' (E = veld/bron, F = bedrag). Verplichte velden hebben een * en worden gecontroleerd bij 'Ingevuld? Verwerken!'.
    ws.Range("B7").Value = "1. Invoerkaart kredieten"
    ws.Range("B7").Font.Bold = True
    ws.Range("B7").Font.Size = 12
    ' rij 8 blijft leeg
    ws.Range("B9").Value = "Veld (* = verplicht)"
    ws.Range("C9").Value = "Invoer (geel = verplicht, lichtblauw = mag leeg)"
    ws.Range("E9").Value = "Budgetten en bijdragen van derden"
    ws.Range("F9").Value = "Bedrag"
    ws.Range("B9:C9,E9:F9").Font.Bold = True
    ws.Range("B9:C9,E9:F9").Borders(xlEdgeBottom).LineStyle = xlContinuous

    For i = 1 To 30
        r = KaartRij(i)
        kc = KaartKol(i)
        If Left$(CStr(v(i)(0)), 1) = "=" Then
            ws.Cells(r, kc - 1).Formula = CStr(v(i)(0))
        Else
            ws.Cells(r, kc - 1).Value = v(i)(0) & IIf(v(i)(3), " *", "")
        End If
        If kc = 3 Then
            ws.Cells(r, 1).Value = v(i)(1)
            ws.Cells(r, 1).Font.Size = 8
            ws.Cells(r, 1).Font.Color = RGB(150, 150, 150)
        End If
        With ws.Cells(r, kc)
            If v(i)(4) = "@" Then
                .NumberFormat = "@"
            ElseIf v(i)(4) = "bedrag" Then
                .NumberFormat = "#,##0;-#,##0;-"
            End If
            .Interior.Color = IIf(v(i)(3), RGB(255, 255, 204), RGB(221, 235, 247))   ' geel = verplicht, lichtblauw = mag leeg
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
        End With
        ws.Cells(r, kc - 1).VerticalAlignment = xlTop
        ws.Cells(r, kc - 1).ShrinkToFit = True                 ' veldnaam loopt nooit door in de invoercel
        If kc = 6 Then ws.Cells(r, kc - 1).Font.Bold = True
        lijstNaam = CStr(v(i)(2))
        If Len(lijstNaam) > 0 Then
            With ws.Cells(r, kc).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & lijstNaam
                .IgnoreBlank = True
                .InCellDropdown = True
                ' vaste keuzes (status, samengesteld, afschrijving, programma, taakveld, balanscode) alleen uit de lijst;
                ' alle andere velden: kiezen uit de lijst of zelf een nieuwe naam typen
                .ShowError = (lijstNaam = "Lst_Statusletter" Or lijstNaam = "Lst_Afschrijving" Or _
                              lijstNaam = "Lst_SamengesteldStandaardStelpost" Or lijstNaam = "Lst_Programma" Or _
                              lijstNaam = "Lst_Taakveld" Or lijstNaam = "Lst_Balanscode")
                If .ShowError Then
                    .ErrorTitle = "Kies uit de lijst"
                    .ErrorMessage = "Kies een waarde uit de keuzelijst."
                End If
            End With
        End If
        If i = 3 Or i = 4 Then
            With ws.Cells(r, kc).Validation
                .Delete
                .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlLessEqual, _
                     Formula1:=IIf(i = 3, "36", "20")
                .IgnoreBlank = True
                .ErrorTitle = "Te lang"
                .ErrorMessage = "Maximaal " & IIf(i = 3, "36", "20") & " posities."
            End With
        End If
    Next i
    ' budget in het jaar van de status wordt geel (verplicht) bij een lopend krediet: L2026 -> budget t/m 2026, L2027 -> budget 2027, enz.
    Dim cS As String, cJ As String, bk2 As Long
    cS = "$C$" & KaartRij(5): cJ = "IFERROR(--$C$" & KaartRij(6) & ",0)"
    Set fc = VoegCF(ws.Cells(KaartRij(23), 6), "=AND(" & cS & "=""L""," & cJ & ">0," & cJ & "<=Par_LopendJaar)")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 255, 204)
    For bk2 = 1 To 5
        Set fc = VoegCF(ws.Cells(KaartRij(23 + bk2), 6), "=AND(" & cS & "=""L""," & cJ & "=Par_LopendJaar+" & bk2 & ")")
        If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 255, 204)
    Next bk2
    Set fc = VoegCF(ws.Cells(KaartRij(30), 3), "=LEFT($C$" & KaartRij(8) & ",5)=""nieuw""")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 255, 204)
    ws.Cells(KaartRij(30), 2).Value = "Naam project met deelkredieten"
    ws.Cells(eerste + 1, 5).ShrinkToFit = True
    ws.Rows(eerste & ":" & (eerste + 23)).RowHeight = 17
    ws.Cells(KaartRij(29), 3).WrapText = True
    ws.Rows(KaartRij(29)).RowHeight = 54
    With ws.Cells(KaartRij(29), 3).Validation
        .Delete
        .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlLessEqual, Formula1:=CStr(MAX_TOEL)
        .IgnoreBlank = True
        .ErrorTitle = "Toelichting te lang"
        .ErrorMessage = "Maximaal " & MAX_TOEL & " tekens, zodat de toelichting altijd past in de kredietoverzichten."
    End With

    ' Bijdragen van derden: onder elk budgetjaar twee regels. In kolom E kiest u de bron (of typt u die zelf),
    ' in kolom F staat het bedrag. Alleen de jaren met een bijdrage worden ingevuld.
    For bj = 0 To 5
        For bk = 1 To 2
            brr = BijdrageRij(bj, bk)
            ws.Cells(brr, 5).Formula = "=""Bijdrage " & bk & " - ""&(Par_LopendJaar+" & bj & ")"
            With ws.Cells(brr, 5)
                .VerticalAlignment = xlTop
                .HorizontalAlignment = xlLeft
                .Interior.Color = RGB(221, 235, 247)
                .IndentLevel = 1
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(191, 191, 191)
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Lst_BijdrageJaar" & bj
                .Validation.IgnoreBlank = True
                .Validation.InCellDropdown = True
                .Validation.ShowError = False                 ' bron kiezen of zelf typen
            End With
            With ws.Cells(brr, 6)
                .NumberFormat = "#,##0;-#,##0;-"
                .Interior.Color = RGB(221, 235, 247)
                .Font.Color = RGB(0, 0, 0)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(191, 191, 191)
            End With
        Next bk
    Next bj

    ' Rijnummer van een geladen regel (bij wijzigen), rechtsboven op de kaart
    ws.Cells(eerste + 1, 5).Value = "Geladen krediet: rij in de Verzamellijst"
    ws.Cells(eerste + 1, 5).Font.Italic = True
    ws.Cells(eerste + 1, 6).Interior.Color = RGB(217, 217, 217)
    ws.Cells(eerste + 1, 6).HorizontalAlignment = xlLeft
    ws.Cells(eerste + 1, 6).Borders.LineStyle = xlContinuous
    ws.Cells(eerste + 1, 6).Borders.Color = RGB(191, 191, 191)

    ' De korte omschrijving wordt automatisch afgeleid; de rij blijft bestaan maar is verborgen.
    ws.Rows(eerste + 3).Hidden = True

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    ws.Range("C9").Select
    On Error Resume Next
    Application.Run "DashboardKnoppenPlaatsen"
    Application.Run "ControleBladAanmaken"
    Application.Run "VerzamellijstOpschonen"
    On Error GoTo Fout
    Application.ScreenUpdating = True
    n = ThisWorkbook.Names("Lst_Programma").RefersToRange.Cells.Count
    MsgBox "Klaar. Het blad Dashboard staat vooraan." & vbCrLf & "Keuzelijsten gevuld (programma's: " & n & ")." & _
           IIf(cfMislukt, vbCrLf & "Let op: de kleurmarkering (alleen opmaak) is niet gelukt.", ""), vbInformation
    Exit Sub

Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' Bouwt de controleformule voor kaartrij r (veldnummer i). leeg = tekst als het veld nog leeg is.
Private Function ControleFormule(ByVal i As Long, ByVal r As Long, ByVal Leeg As String) As String
    Dim c As String, k As String
    c = "C" & r
    If i = 30 Then
        ControleFormule = "=IF(LEFT($C$" & KaartRij(8) & ",5)=""nieuw"",IF(" & c & "="""",""verplicht bij een nieuw project met deelkredieten"",""goed""),IF(LEFT($C$" & KaartRij(8) & ",6)=""groep "",""wordt overgenomen van de groep"",""""))"
        Exit Function
    End If
    Select Case i
        Case 1
            k = "IF(ISNUMBER(MATCH(" & c & ",Lst_Programma,0)),""goed"",""kies een programma uit de lijst"")"
        Case 2
            k = "IF(COUNTIF(" & VL & "$C:$C," & c & ")>0,""waarschuwing: bestaat al (volgnr ""&INDEX(" & VL & "$A:$A,MATCH(" & c & "," & VL & "$C:$C,0))&"")"",""bestaat al: nee"")"
        Case 3
            k = "LEN(" & c & ")&""/36 - ""&IF(LEN(" & c & ")<=36,""goed, kort: ""&IF(C" & (r + 1) & "<>"""",C" & (r + 1) & ",KortOmschrijving(" & c & ")),""te lang"")"
        Case 4
            k = "LEN(" & c & ")&""/20 - ""&IF(LEN(" & c & ")<=20,""goed"",""te lang"")"
        Case 5
            k = "IF(OR(" & c & "=""L""," & c & "=""A""),""goed"",""kies L of A"")"
        Case 6
            k = "IF(IFERROR(AND(LEN(" & c & ")=4,ISNUMBER(--" & c & ")),FALSE),""goed"",""vier cijfers nodig"")"
        Case 8
            k = "IF(LEFT(" & c & ",5)=""nieuw"",""nieuw: krijgt groepsnummer ""&(MAX(" & VL & "$J:$J)+1),IF(OR(" & c & "=""standaard""," & c & "=""stelpost"",LEFT(" & c & ",6)=""groep ""),""goed"",""kies standaard, stelpost of samengesteld""))"
        Case 11
            k = "IF(OR(C" & (r - 1) & "="""",NOT(ISNUMBER(C" & (r - 1) & ")),NOT(ISNUMBER(" & c & "))),""goed"",IF(C" & (r - 1) & "<=" & c & ",""goed"",""start na einde""))"
        Case 16
            k = "IF(ISNUMBER(MATCH(" & c & ",Lst_Taakveld,0)),""goed"",""kies een taakveld uit de lijst"")"
        Case 17
            k = "IF(ISNUMBER(MATCH(" & c & ",Lst_Balanscode,0)),""goed"",""kies een balanscode uit de lijst"")"
        Case 21
            k = "IF(COUNTIF(" & VL & "$X:$X," & c & ")>0,""waarschuwing: bestaat al (volgnr ""&INDEX(" & VL & "$A:$A,MATCH(" & c & "," & VL & "$X:$X,0))&"")"",""bestaat al: nee"")"
        Case 22
            k = "IF(AND(ISNUMBER(" & c & ")," & c & ">=0," & c & "<=60),""goed"",""kies 0 t/m 60 jaar"")"
        Case 23 To 28
            k = "IF(ISNUMBER(" & c & "),""goed"",""moet een getal zijn"")"
        Case Else
            k = """"""
    End Select
    ControleFormule = "=IF(" & c & "="""",""" & Leeg & """," & k & ")"
End Function

' Voegt voorwaardelijke opmaak toe. De formule wordt omgezet naar de taal van uw Excel en
' de eerste cel van het bereik wordt eerst geselecteerd (relatieve verwijzingen).
Public Function VoegCF(ByVal rng As Range, ByVal formule As String) As Object
    Dim ws As Worksheet, t As Range, lokaal As String
    On Error GoTo Mis
    Set ws = rng.Worksheet
    ws.Activate
    rng.Cells(1, 1).Select
    Set t = ws.Range("Z1")
    t.Formula = formule
    lokaal = t.FormulaLocal
    t.Clear
    Set VoegCF = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=lokaal)
    Exit Function
Mis:
    cfMislukt = True
    Set VoegCF = Nothing
End Function

' ======================================================================
' Krediet toevoegen (stap 8)
' ======================================================================
' Stap 8 - Krediet toevoegen
' KredietToevoegen          : controleert de invoerkaart op het Dashboard en schrijft een nieuwe regel onderaan de verzamellijst.
' KaartWissen               : maakt de invoerkaart leeg.
' DashboardKnoppenPlaatsen  : zet de knoppen Krediet toevoegen en Scherm leegmaken op het Dashboard.
' Samengestelde kredieten: de keuze op de kaart bepaalt het groepsnummer in kolom J. Na het opslaan van een
' deelkrediet kan direct het volgende deelkrediet van dezelfde groep worden ingevoerd.


Public Sub KredietToevoegen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    BlijfOpDashboard                             ' ook bij een foutmelding terug op het Dashboard
    Dim wd As Worksheet, wv As Worksheet
    Dim f(1 To aantal) As Variant
    Dim fouten As String, waarsch As String
    Dim i As Long, r As Long, laatste As Long, maxRij As Long, volgnr As Long
    Dim Letter As String, jaar As String, j8 As String
    Dim nieuw As Boolean, grp As Long, jWaarde As Variant, naam As String
    Dim pc As String, tc As String, bc As String
    Dim voor As Variant, na As Variant, verschil As Boolean
    Dim kolom As Variant, som As Double, x As Long, cel As Range, kn As Variant
    Dim nB As Long, kk As Long, jj As Long, bijtxt As String
    Dim bron(1 To 12) As String, bed(1 To 12, 0 To 5) As Variant, sumB(0 To 5) As Double
    Dim geschreven As Boolean, klaar As Boolean, foutTekst As String

    On Error GoTo Fout
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen kredieten toevoegen.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    Set wv = Vz()
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then
        MsgBox "Er is een regel geladen om te wijzigen. Gebruik eerst Scherm leegmaken.", vbExclamation
        Exit Sub
    End If
    If wv.Range("AT1").Value = "(reserve)" Then wv.Range("AT1").Value = "Samengesteld krediet (naam)"
    If wv.Range("AU1").Value = "(reserve)" Then wv.Range("AU1").Value = "Toegevoegd op"

    For i = 1 To aantal
        f(i) = wd.Cells(KaartRij(i), KaartKol(i)).Value
    Next i

    ' --- verplichte velden
    For i = 1 To 22
        If i <> 4 And i <> 2 And i <> 21 And Leeg(f(i)) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), KaartKol(i) - 1).Value & " is verplicht" & vbCrLf
    Next i
    ' toelichting (veld 29) is bij een nieuw krediet niet verplicht

    ' --- lengte omschrijvingen
    If Not Leeg(f(3)) Then If Len(Trim$(CStr(f(3)))) > 36 Then fouten = fouten & "- omschrijving lang is langer dan 36 posities" & vbCrLf
    If Not Leeg(f(4)) Then If Len(Trim$(CStr(f(4)))) > 20 Then fouten = fouten & "- omschrijving kort is langer dan 20 posities" & vbCrLf
    If Len(Trim$(CStr(f(29)))) > MAX_TOEL Then fouten = fouten & "- toelichting is langer dan " & MAX_TOEL & " tekens" & vbCrLf

    ' --- status
    Letter = UCase$(Trim$(CStr(f(5))))
    jaar = Trim$(CStr(f(6)))
    If Not Leeg(f(5)) Then If Letter <> "L" And Letter <> "A" Then fouten = fouten & "- status moet L of A zijn" & vbCrLf
    If Not Leeg(f(6)) Then If Len(jaar) <> 4 Or Not IsNumeric(jaar) Then fouten = fouten & "- eerste jaar krediet beschikbaar moet uit vier cijfers bestaan" & vbCrLf

    ' --- budget in het jaar van de status is verplicht bij een lopend krediet (L2026 -> budget t/m 2026, L2027 -> budget 2027)
    If StatusBudgetVeld(Letter, jaar) > 0 Then
        If Leeg(f(StatusBudgetVeld(Letter, jaar))) Then fouten = fouten & "- " & wd.Cells(KaartRij(StatusBudgetVeld(Letter, jaar)), 5).Value & _
                                                      " is verplicht bij status " & Letter & jaar & vbCrLf
    End If

    ' --- getallen
    If Not Leeg(f(2)) Then If Not IsNumeric(f(2)) Then fouten = fouten & "- kredietnummer moet een getal zijn" & vbCrLf
    If Not Leeg(f(10)) Then If Not IsNumeric(f(10)) Then fouten = fouten & "- startjaar uitgaven moet een jaartal zijn" & vbCrLf
    If Not Leeg(f(11)) Then If Not IsNumeric(f(11)) Then fouten = fouten & "- laatste jaar uitgaven moet een jaartal zijn" & vbCrLf
    If IsNumeric(f(10)) And IsNumeric(f(11)) And Not Leeg(f(10)) And Not Leeg(f(11)) Then
        If CDbl(f(10)) > CDbl(f(11)) Then fouten = fouten & "- startjaar uitgaven ligt na het laatste jaar uitgaven" & vbCrLf
    End If
    If Not Leeg(f(15)) Then If Not IsNumeric(f(15)) Then fouten = fouten & "- exploitatienummer moet een getal zijn" & vbCrLf
    If Not Leeg(f(22)) Then
        If Not IsNumeric(f(22)) Then
            fouten = fouten & "- afschrijftermijn moet een getal zijn" & vbCrLf
        ElseIf CDbl(f(22)) < 0 Or CDbl(f(22)) > 60 Then
            fouten = fouten & "- afschrijftermijn moet tussen 0 en 60 jaar liggen" & vbCrLf
        End If
    End If
    For i = 23 To 28
        If Not Leeg(f(i)) Then If Not IsNumeric(f(i)) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), KaartKol(i) - 1).Value & " moet een getal zijn" & vbCrLf
    Next i

    ' --- codes uit de tabellen
    pc = VoorCode(CStr(f(1)))
    tc = Replace(VoorCode(CStr(f(16))), ".", "")
    bc = VoorCode(CStr(f(17)))
    If Not Leeg(f(1)) Then If Not CodeInTabel("Tbl_Programma", pc) Then fouten = fouten & "- programma staat niet in de programmatabel (blad Parameters)" & vbCrLf
    If Not Leeg(f(16)) Then If Not CodeInTabel("Tbl_Taakveld", tc) Then fouten = fouten & "- taakveld staat niet in de taakveldentabel (blad Parameters)" & vbCrLf
    If Not Leeg(f(17)) Then If Not CodeInTabel("Tbl_Balanscode", bc) Then fouten = fouten & "- balanscode staat niet in de balanscodetabel (blad Parameters)" & vbCrLf
    If Not Leeg(f(1)) Then If Not IsNumeric(pc) Then fouten = fouten & "- programma moet met een getal beginnen" & vbCrLf
    If Not Leeg(f(16)) Then If Not IsNumeric(tc) Then fouten = fouten & "- taakveld moet met een getal beginnen" & vbCrLf

    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)

    ' --- samengesteld, standaard of stelpost
    j8 = LCase$(Trim$(CStr(f(8))))
    If Not Leeg(f(8)) Then
        If j8 = "standaard" Or j8 = "stelpost" Then
            jWaarde = j8
        ElseIf Left$(j8, 5) = "nieuw" Then
            nieuw = True
            grp = CLng(Application.WorksheetFunction.Max(wv.Range("J2:J" & laatste))) + 1
            jWaarde = grp
            naam = Trim$(CStr(f(30)))
            If Len(naam) = 0 Then fouten = fouten & "- naam project is verplicht bij een nieuw project met deelkredieten" & vbCrLf
            ' bestaat dit project al? Dan geen tweede projectnummer: het volgende deel hoort bij de bestaande groep
            If Len(naam) > 0 Then
                If ProjectGroep(naam) > 0 Then fouten = fouten & "- project '" & naam & "' bestaat al (groep " & ProjectGroep(naam) & _
                                                       "). Kies bij 'Project met deelkredieten' die groep uit de lijst" & vbCrLf
            End If
        ElseIf Left$(j8, 6) = "groep " Then
            grp = CLng(Val(Mid$(j8, 7)))
            If grp < 1 Then
                fouten = fouten & "- kies de groep uit de lijst" & vbCrLf
            ElseIf Application.WorksheetFunction.CountIf(wv.Range("J2:J" & laatste), grp) = 0 Then
                fouten = fouten & "- groep " & grp & " bestaat niet" & vbCrLf
            Else
                jWaarde = grp
                For Each cel In wv.Range("J2:J" & laatste).Cells
                    If JGroep(cel.Value) = grp Then
                        If Len(Trim$(CStr(wv.Cells(cel.Row, 46).Value))) > 0 Then
                            naam = Trim$(CStr(wv.Cells(cel.Row, 46).Value))
                            Exit For
                        End If
                    End If
                Next cel
            End If
        Else
            fouten = fouten & "- kies bij samengesteld/standaard/stelpost een waarde uit de lijst" & vbCrLf
        End If
    End If

    nB = LeesBijdragen(bron, bed, fouten)
    For kk = 1 To nB
        For jj = 0 To 5
            sumB(jj) = sumB(jj) + Nul(bed(kk, jj))
        Next jj
    Next kk

    If Len(fouten) > 0 Then
        MsgBox "Niet opgeslagen. Dit moet nog worden aangepast:" & vbCrLf & vbCrLf & fouten, vbExclamation
        Exit Sub
    End If

    ' --- kredietnummer leeg: eerstvolgende vrije nummer van de balansrubriek (module 9)
    If Leeg(f(2)) Then
        kn = VolgendKredietnummer(bc)
        If Len(CStr(kn)) > 0 Then f(2) = kn
    End If

    ' --- waarschuwingen
    If Not Leeg(f(2)) Then
        If Application.WorksheetFunction.CountIf(wv.Range("C2:C" & laatste), CDbl(f(2))) > 0 Then
            waarsch = waarsch & "- kredietnummer " & f(2) & " bestaat al in de verzamellijst" & vbCrLf
        End If
    End If
    If Not Leeg(f(21)) Then
        If IsNumeric(f(21)) Then
            If Application.WorksheetFunction.CountIf(wv.Range("X2:X" & laatste), CDbl(f(21))) > 0 Then waarsch = waarsch & "- activanummer " & f(21) & " bestaat al" & vbCrLf
        End If
    End If
    If Len(waarsch) > 0 Then
        If MsgBox("Let op:" & vbCrLf & vbCrLf & waarsch & vbCrLf & "Toch toevoegen?", vbYesNo + vbExclamation + vbDefaultButton2, "Krediet toevoegen") <> vbYes Then Exit Sub
    End If

    ' --- rij en volgnummer
    r = CLng(ThisWorkbook.Names("Par_EersteVrijeRij").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    volgnr = CLng(ThisWorkbook.Names("Par_VolgendVolgnr").RefersToRange.Value)
    If r > maxRij Then
        MsgBox "De formules lopen tot rij " & maxRij & ". Voer eerst WerkbereikVerruimen uit.", vbExclamation
        Exit Sub
    End If
    If Not IsEmpty(wv.Cells(r, 1).Value) Or r <> laatste + 1 Then
        MsgBox "De eerste vrije rij klopt niet (rij " & r & "). Er is niets weggeschreven. Controleer blad Parameters.", vbCritical
        Exit Sub
    End If

    If r + nB > maxRij Then
        MsgBox "De formules lopen tot rij " & maxRij & ". Voer eerst WerkbereikVerruimen uit.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    voor = MeetStap8(wv, r + nB)
    geschreven = True                            ' vanaf hier wordt bij een fout alles teruggezet

    ' --- wegschrijven
    wv.Cells(r, 1).Value = volgnr
    wv.Cells(r, 2).Value = CLng(pc)
    If Not Leeg(f(2)) Then wv.Cells(r, 3).Value = CDbl(f(2))
    wv.Cells(r, 4).Value = Trim$(CStr(f(3)))
    If Leeg(f(4)) Then
        wv.Cells(r, 6).Value = KortOmschrijving(Trim$(CStr(f(3))))
    Else
        wv.Cells(r, 6).Value = Trim$(CStr(f(4)))
    End If
    wv.Cells(r, 8).Value = Letter & jaar
    wv.Cells(r, 9).Value = LCase$(Trim$(CStr(f(7))))
    wv.Cells(r, 10).Value = jWaarde
    wv.Cells(r, 12).Value = GetalOfTekst(f(9))
    wv.Cells(r, 13).Value = CDbl(f(10))
    wv.Cells(r, 14).Value = CDbl(f(11))
    wv.Cells(r, 15).Value = Trim$(CStr(f(12)))
    wv.Cells(r, 16).Value = Trim$(CStr(f(13)))
    wv.Cells(r, 17).Value = Trim$(CStr(f(14)))
    wv.Cells(r, 18).Value = CDbl(f(15))
    If Left$(tc, 1) = "0" Then
        wv.Cells(r, 19).NumberFormat = "@"
        wv.Cells(r, 19).Value = tc
    Else
        wv.Cells(r, 19).Value = CLng(tc)
    End If
    wv.Cells(r, 20).Value = bc
    wv.Cells(r, 21).Value = Trim$(CStr(f(18)))
    wv.Cells(r, 22).Value = Trim$(CStr(f(19)))
    wv.Cells(r, 23).Value = Trim$(CStr(f(20)))
    If Not Leeg(f(21)) Then wv.Cells(r, 24).Value = GetalOfTekst(f(21))
    wv.Cells(r, 25).Value = CDbl(f(22))
    If Not Leeg(f(23)) Then wv.Cells(r, 26).Value = CDbl(f(23))
    For i = 24 To 28
        If Not Leeg(f(i)) Then wv.Cells(r, 31 + i - 24).Value = CDbl(f(i))
    Next i
    wv.Cells(r, 36).Value = Trim$(CStr(f(29)))
    wv.Cells(r, 42).Value = CtlTeControleren()
    If Len(naam) > 0 Then wv.Cells(r, 46).Value = naam
    wv.Cells(r, 48).Value = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    wv.Cells(r, 49).Value = Now
    wv.Cells(r, 47).NumberFormat = "dd-mm-yyyy"
    wv.Cells(r, 47).Value = Date

    ' bijdragen van derden: aparte regels met negatieve bedragen
    For kk = 1 To nB
        SchrijfBijdrageRegel wv, r, r + kk, volgnr + kk, bron(kk), bed, kk
    Next kk

    ' formules controleren (zo nodig kopieren van de regel erboven)
    For kk = 0 To nB
        ControleerFormules wv, r + kk
    Next kk
    Application.Calculate

    ' --- nameting
    na = MeetStap8(wv, r + nB)
    If na(0) <> voor(0) + 1 + nB Then verschil = True
    If Abs(na(1) - (voor(1) + Nul(f(23)) + sumB(0))) > 0.005 Then verschil = True
    For i = 2 To 4
        If Abs(na(i) - voor(i)) > 0.005 Then verschil = True
    Next i
    For i = 24 To 28
        If Abs(na(5 + i - 24) - (voor(5 + i - 24) + Nul(f(i)) + sumB(i - 23))) > 0.005 Then verschil = True
    Next i
    If verschil Then
        For kk = 0 To nB
            RuimRij wv, r + kk
        Next kk
        Application.ScreenUpdating = True
        MsgBox "Verschil bij de nameting. De nieuwe regel is weer gewist. Er is niets toegevoegd.", vbCritical
        Exit Sub
    End If
    For kk = 0 To nB                             ' opmaak exact gelijk aan voorbeeldregel 272 (module 9)
        RegelOpmaakAlsVoorbeeld wv, r + kk
    Next kk

    ' --- controle sleutel samengesteld krediet
    waarsch = ""
    If grp > 0 Then
        For x = 2 To r
            If JGroep(wv.Cells(x, 10).Value) = grp Then
                If VarType(wv.Cells(x, 11).Value) <> vbString And Not IsError(wv.Cells(x, 11).Value) Then som = som + wv.Cells(x, 11).Value
            End If
        Next x
        If Abs(som - 1) > 0.002 Then waarsch = "Let op: de verdeelsleutel van groep " & grp & " komt uit op " & Format(som * 100, "0.0") & "% in plaats van 100%. Controleer de budgetten van de deelkredieten."
    End If

    If Leeg(f(2)) Then kn = "" Else kn = CDbl(f(2))
    LogRegel "toevoegen", volgnr, kn, "", , , "rij " & r & IIf(grp > 0, ", groep " & grp, "")
    For kk = 1 To nB
        LogRegel "toevoegen bijdrage", volgnr + kk, kn, "", , , "rij " & (r + kk) & ", bron " & bron(kk)
    Next kk
    If nB > 0 Then bijtxt = " en " & nB & " bijdrage(n) van derden"
    klaar = True                                 ' het krediet is gelukt; fouten hierna draaien het niet meer terug
    KredietnummerGebruikt bc, f(2)               ' eerstvolgende vrije nummer van deze balansrubriek schuift door
    KeuzelijstenVerversen True
    Application.ScreenUpdating = True

    BlijfOpDashboard
    If grp > 0 Then
        If MsgBox("Krediet toegevoegd op rij " & r & ", volgnummer " & volgnr & ", groep " & grp & bijtxt & "." & vbCrLf & _
                  IIf(Len(waarsch) > 0, waarsch & vbCrLf, "") & vbCrLf & _
                  "Nog een deelkrediet bij dit project toevoegen?", vbYesNo + vbQuestion, "Krediet toevoegen") = vbYes Then
            VolgendeDeelVoorbereiden grp
        Else
            KaartWissen
        End If
    Else
        KaartWissen
        MsgBox "Krediet toegevoegd op rij " & r & ", volgnummer " & volgnr & bijtxt & ".", vbInformation
    End If
    Exit Sub

Fout:
    foutTekst = Err.Description
    Application.CutCopyMode = False
    If geschreven And Not klaar Then
        On Error Resume Next
        For kk = 0 To nB
            RuimRij wv, r + kk
        Next kk
        Application.Calculate
        Application.ScreenUpdating = True
        MsgBox "Fout: " & foutTekst & vbCrLf & "De nieuwe regel is weer gewist. Er is niets toegevoegd." & vbCrLf & _
               "Controleer voor de zekerheid het Logboek.", vbCritical
    ElseIf klaar Then
        Application.ScreenUpdating = True
        MsgBox "Het krediet is toegevoegd (rij " & r & "), maar daarna ging iets mis: " & foutTekst & vbCrLf & _
               "Ververs de keuzelijsten met de knop op het Dashboard.", vbExclamation
    Else
        Application.ScreenUpdating = True
        MsgBox "Fout: " & foutTekst & vbCrLf & "Er is niets weggeschreven.", vbCritical
    End If
End Sub

Public Sub KaartWissen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    wd.Range("Kaart_Invoer").ClearContents
    HerstelBijdrageLabels wd
    wd.Range("Kaart_Rijnummer").ClearContents
    WisKaartVolgnr
    KaartRijenPassendMaken
End Sub

' Zorgt dat de gebruiker na afloop op het Dashboard staat (ook nadat de beveiliging een seconde later weer aan gaat).
Public Sub BlijfOpDashboard()
    mBlijfOpDashboard = True
    On Error Resume Next
    ThisWorkbook.Worksheets(BL_DASH).Activate
End Sub

' Het volgnummer van de geladen regel wordt vastgelegd, zodat bij opslaan blijkt of de regel nog dezelfde is.
Public Sub ZetKaartVolgnr(ByVal v As Variant)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="Kaart_VolgnrGeladen", RefersTo:="=" & CStr(CLng(v)), Visible:=False
End Sub

Public Sub WisKaartVolgnr()
    On Error Resume Next
    ThisWorkbook.Names("Kaart_VolgnrGeladen").Delete
End Sub

Public Function KaartVolgnr() As Long
    Dim t As String
    On Error Resume Next
    t = ThisWorkbook.Names("Kaart_VolgnrGeladen").RefersTo
    If Len(t) > 1 Then KaartVolgnr = CLng(Mid$(t, 2))
End Function

' Maakt alleen de velden per deelkrediet leeg en zet de groep klaar voor het volgende deelkrediet.
Private Sub VolgendeDeelVoorbereiden(ByVal grp As Long)
    Dim wd As Worksheet, i As Variant, cel As Range, zoek As String, bj As Long, bk As Long
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    For Each i In Array(2, 3, 4, 17, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30)
        wd.Cells(KaartRij(CLng(i)), KaartKol(CLng(i))).ClearContents
    Next i
    For bk = 1 To 2
        For bj = 0 To 5
            wd.Cells(BijdrageRij(bj, bk), 6).ClearContents
        Next bj
    Next bk
    HerstelBijdrageLabels wd
    KeuzelijstenVerversen True
    zoek = "groep " & grp & " - "
    For Each cel In ThisWorkbook.Names("Lst_SamengesteldStandaardStelpost").RefersToRange.Cells
        If Left$(CStr(cel.Value), Len(zoek)) = zoek Then
            wd.Cells(KaartRij(8), KaartKol(8)).Value = cel.Value
            Exit For
        End If
    Next cel
    KaartRijenPassendMaken
    BladBeveiligen wd                            ' naam samengesteld krediet (C18) is nu geblokkeerd
    wd.Activate
    wd.Cells(eerste + 1, 3).Select
End Sub
