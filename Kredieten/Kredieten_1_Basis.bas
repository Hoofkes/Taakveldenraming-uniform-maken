Attribute VB_Name = "Kredieten_1_Basis"
Option Explicit

' Alle macros voor het Dashboard investeringskredieten in een module.
' Versie 41 (stap 1: keuzelijst samengestelde kredieten zonder groepsnummer, lange uitlegteksten, logregels normaliseren en werkbereik, map aanmaken en overzicht opslaan met nette foutmelding).
' Versie 40 (rijhoogte overzichten 13,5; uitvraag verwerken met de hele lijst; krediet wijzigen en verwijderen op rij 35; uitvraag per budgethouder: inlezen, per investering besluiten, verwerken, stand uitvraag, begroting vastgesteld; knop Uitvraag verwerken voorlopig weg; Dashboard: 1 invoerkaart, 2 financieel adviseur, 3 financiele administratie, 4 controles; toelichting maximaal 1000 tekens; datum keuzeopties in F3; budget in het statusjaar verplicht en geel; fout bij DashboardAanmaken opgelost; nieuwste samengestelde krediet bovenaan in de keuzelijst; applicatiebeheerder ziet alle blokken; blokken per rol, knoppen onder de kaart, budgethouder/projectleider/tijdschrijver vervangen, kortere veldnamen; naam samengesteld krediet direct onder C16; afgekeurd: signaal, herstellen, aanbieden voor hercontrole; kleuren geel = verplicht, lichtblauw = mag leeg; compact Dashboard, kengetallen op eigen tabblad, naam vervangen).

Public Const BL_VZ As String = "Verzamellijst"
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
Public Const BL_UITVRAAG As String = "Laatste uitvraag begroting"
Public Const BL_TIJD As String = "Laatste overzicht tijdschrijven"
Public Const OPBRENGST As String = "Opbrengst activa"   ' bron/toelichting voor opbrengsten activa (geen subsidie of bijdrage)
Public Const BL_MUTREG As String = "Mutaties kadernota-begroting"   ' register voor de lastenberekening
Public Const WACHTWOORD As String = "Best"      ' bladbeveiliging (zet ook een wachtwoord op het VBA-project)
Public mStil As Boolean
Public Const CB_KOP As Long = 50            ' Dashboard blok 4 Controles: kopregel (onderaan)
Public Const CB_EERSTE As Long = 53         ' eerste controleregel
Public Const CB_LAATSTE As Long = 68        ' laatste controleregel (16 controles)
Public Const NM_KOP As Long = 37            ' Dashboard blok 2 Financieel adviseur: kopregel
Public Const NM_ROL As Long = 38            ' cel C38: rol
Public Const NM_NAAM As Long = 39           ' cel C39: naam (knoppen op rij 40)
Public Const ADV_KNOP As Long = 41          ' blok 2: eerste van vier rijen knoppen (41 t/m 44)
Public Const FA_KOP As Long = 46            ' Dashboard blok 3 Financiele administratie: kopregel (uitleg 47, knoppen 48)
Public Const BL_WERK As String = "Werkregel controle"   ' verborgen: regel die de financiele administratie open heeft
Public Const BL_KENG As String = "Kengetallen"
Public Const BL_UVVERW As String = "Verwerking uitvraag"
Public Const BL_UVSTAND As String = "Stand uitvraag"
Public Const GEEN_WIJZ As String = "geen wijziging"
Public Const MAX_TOEL As Long = 1000        ' maximale lengte toelichting (past altijd in de kredietoverzichten)
Public Const FA_LIJST As Long = 120         ' kopregel van de lijst te controleren / afgekeurd (regels eronder)
Public mHerbeveiligingGepland As Boolean
Public mActiefVoorMacro As Object        ' blad dat actief was toen de macro begon
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
    VerzamellijstHernoemen
    ' beveiliging opnieuw aanzetten (macro's mogen schrijven, gebruikers niet) en oude overzichtbladen opruimen
    BeveiligingAan
    On Error Resume Next
    OudeOverzichtbladenOpruimen
    ControlesVernieuwen
    On Error GoTo 0
    ' keuzeopties verversen bij het openen (niet als het bestand alleen-lezen is geopend)
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then KeuzelijstenVerversen True
    On Error GoTo 0
    If TabelBestaat() Then WieWerktEr
    AfgekeurdSignaal                             ' financieel adviseur: melding als er afgekeurde kredieten zijn
    DashboardEventInstalleren True               ' stil: zet de melding voor de financiele administratie aan als dat nog niet zo is
    JaarovergangSignaleren
    ' altijd op het Dashboard beginnen
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Activate
    ThisWorkbook.Worksheets("Dashboard").Range("A1").Select
    On Error GoTo 0
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

' Geeft True als de huidige gebruiker het recht heeft: "Toevoegen", "Wijzigen" of "Bevestigen".
Public Function HeeftRecht(ByVal recht As String) As Boolean
    Dim lo As ListObject, r As ListRow, naam As String
    If Not TabelBestaat() Then Exit Function
    naam = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    If Len(naam) = 0 Then Exit Function
    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects("Tbl_Rollen")
    For Each r In lo.ListRows
        If StrComp(CStr(r.Range.Cells(1, 1).Value), naam, vbTextCompare) = 0 Then
            HeeftRecht = (LCase$(CStr(r.Range.Cells(1, lo.ListColumns(recht).Index).Value)) = "ja")
            Exit Function
        End If
    Next r
End Function

' Schrijft een volledige logregel op blad Logboek.
Public Sub LogRegel(ByVal actie As String, Optional ByVal volgnr As Variant, Optional ByVal invnr As Variant, _
                    Optional ByVal Veld As String = "", Optional ByVal oud As Variant, _
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
    lg.Cells(rij, 8).Value = Veld
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
    wsP.Range("J17").Value = "nieuw samengesteld krediet"
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

Private Function Veld(ByVal label As String, ByVal kol As String, ByVal lijst As String, _
                      ByVal verplicht As Boolean, ByVal fmt As String) As Variant
    Veld = Array(label, kol, lijst, verplicht, fmt)
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
        If w.Name = "Dashboard" Then
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

    v(1) = Veld("Programma", "B", "Lst_Programma", True, "@")
    v(2) = Veld("Kredietnummer", "C", "", False, "")
    v(3) = Veld("Omschrijving lang (max 36)", "D", "", True, "@")
    v(4) = Veld("Omschrijving kort (max 20)", "F", "", False, "@")
    v(5) = Veld("Status (L of A)", "H", "Lst_Statusletter", True, "@")
    v(6) = Veld("Jaar van de status", "H", "Lst_Jaren", True, "@")
    v(7) = Veld("Investeringsgroep", "I", "Lst_Soort", True, "@")
    v(8) = Veld("Samengesteld/standaard/stelpost", "J", "Lst_SamengesteldStandaardStelpost", True, "")
    v(9) = Veld("Besluit", "L", "Lst_LaatsteBesluit", True, "@")
    v(10) = Veld("Startjaar uitgaven", "M", "Lst_Jaren", True, "")
    v(11) = Veld("Eindjaar uitgaven", "N", "Lst_Jaren", True, "")
    v(12) = Veld("Budgetbeheerder", "O", "Lst_Budgetbeheerder", True, "@")
    v(13) = Veld("Projectleider", "P", "Lst_Projectleider", True, "@")
    v(14) = Veld("Tijdschrijvers", "Q", "Lst_Tijdschrijver", True, "@")
    v(15) = Veld("Exploitatienummer", "R", "", True, "")
    v(16) = Veld("Taakveld BBV", "S", "Lst_Taakveld", True, "@")
    v(17) = Veld("Balanscode BBV", "T", "Lst_Balanscode", True, "@")
    v(18) = Veld("Organisatie-onderdeel", "U", "Lst_Organisatieonderdeel", True, "@")
    v(19) = Veld("Budgetbeheerder administratie", "V", "Lst_BudgetbeheerderAdm", True, "@")
    v(20) = Veld("Code btw/bcf administratie", "W", "Lst_BtwBcf", True, "@")
    v(21) = Veld("Activanummer", "X", "", False, "")
    v(22) = Veld("Afschrijftermijn (0-60 jaar)", "Y", "Lst_Afschrijving", True, "")
    v(23) = Veld("=IF(Kaart_Rijnummer="""",""Budget ""&Par_LopendJaar,""Budget tot en met ""&Par_LopendJaar)", "Z", "", False, "bedrag")
    v(24) = Veld("=""Budget ""&(Par_LopendJaar+1)", "AE", "", False, "bedrag")
    v(25) = Veld("=""Budget ""&(Par_LopendJaar+2)", "AF", "", False, "bedrag")
    v(26) = Veld("=""Budget ""&(Par_LopendJaar+3)", "AG", "", False, "bedrag")
    v(27) = Veld("=""Budget ""&(Par_LopendJaar+4)", "AH", "", False, "bedrag")
    v(28) = Veld("=""Budget ""&(Par_LopendJaar+5)", "AI", "", False, "bedrag")
    v(29) = Veld("Toelichting op het investeringskrediet", "AJ", "", False, "@")
    v(30) = Veld("Naam samengesteld krediet", "AT", "", False, "@")

    Set actief = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = "Dashboard"
    ' Namen voor de latere macro's
    ThisWorkbook.Names.Add Name:="Kaart_Kolommen", RefersTo:="=Dashboard!$A$" & eerste & ":$A$" & (eerste + 23)
    ThisWorkbook.Names.Add Name:="Kaart_Invoer", RefersTo:="=Dashboard!$C$" & eerste & ":$C$" & (eerste + 23) & _
                                                           ",Dashboard!$F$" & (eerste + 4) & ":$F$" & (eerste + 21)
    ThisWorkbook.Names.Add Name:="Kaart_Rijnummer", RefersTo:="=Dashboard!$F$" & (eerste + 1)

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
    ws.Cells(KaartRij(30), 2).Value = "Naam samengesteld krediet"
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
        ControleFormule = "=IF(LEFT($C$16,5)=""nieuw"",IF(" & c & "="""",""verplicht bij een nieuw samengesteld krediet"",""goed""),IF(LEFT($C$16,6)=""groep "",""wordt overgenomen van de groep"",""""))"
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
            k = "IF(" & c & "=""nieuw samengesteld krediet"",""nieuw: krijgt groepsnummer ""&(MAX(" & VL & "$J:$J)+1),IF(OR(" & c & "=""standaard""," & c & "=""stelpost"",LEFT(" & c & ",6)=""groep ""),""goed"",""kies standaard, stelpost of samengesteld""))"
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
    Dim wd As Worksheet, wv As Worksheet
    Dim f(1 To aantal) As Variant
    Dim fouten As String, waarsch As String
    Dim i As Long, r As Long, laatste As Long, maxRij As Long, volgnr As Long
    Dim Letter As String, jaar As String, j8 As String
    Dim nieuw As Boolean, grp As Long, jWaarde As Variant, naam As String
    Dim pc As String, tc As String, bc As String
    Dim voor As Variant, na As Variant, verschil As Boolean
    Dim kolom As Variant, som As Double, x As Long, cel As Range, kn As Variant
    Dim nb As Long, kk As Long, jj As Long, bijtxt As String
    Dim bron(1 To 12) As String, bed(1 To 12, 0 To 5) As Variant, sumB(0 To 5) As Double

    On Error GoTo Fout
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen kredieten toevoegen.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets("Dashboard")
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

    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)

    ' --- samengesteld, standaard of stelpost
    j8 = LCase$(Trim$(CStr(f(8))))
    If Not Leeg(f(8)) Then
        If j8 = "standaard" Or j8 = "stelpost" Then
            jWaarde = j8
        ElseIf j8 = "nieuw samengesteld krediet" Then
            nieuw = True
            grp = CLng(Application.WorksheetFunction.Max(wv.Range("J2:J" & laatste))) + 1
            jWaarde = grp
            naam = Trim$(CStr(f(30)))
            If Len(naam) = 0 Then fouten = fouten & "- naam samengesteld krediet is verplicht bij een nieuw samengesteld krediet" & vbCrLf
        ElseIf Left$(j8, 6) = "groep " Then
            grp = CLng(Val(Mid$(j8, 7)))
            If grp < 1 Then
                fouten = fouten & "- kies de groep uit de lijst" & vbCrLf
            ElseIf Application.WorksheetFunction.CountIf(wv.Range("J2:J" & laatste), grp) = 0 Then
                fouten = fouten & "- groep " & grp & " bestaat niet" & vbCrLf
            Else
                jWaarde = grp
                For Each cel In wv.Range("J2:J" & laatste).Cells
                    If VarType(cel.Value) <> vbString Then
                        If Not IsEmpty(cel.Value) Then
                            If CLng(cel.Value) = grp Then
                                If Len(Trim$(CStr(wv.Cells(cel.Row, 46).Value))) > 0 Then
                                    naam = Trim$(CStr(wv.Cells(cel.Row, 46).Value))
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next cel
            End If
        Else
            fouten = fouten & "- kies bij samengesteld/standaard/stelpost een waarde uit de lijst" & vbCrLf
        End If
    End If

    nb = LeesBijdragen(bron, bed, fouten)
    For kk = 1 To nb
        For jj = 0 To 5
            sumB(jj) = sumB(jj) + Nul(bed(kk, jj))
        Next jj
    Next kk

    If Len(fouten) > 0 Then
        MsgBox "Niet opgeslagen. Dit moet nog worden aangepast:" & vbCrLf & vbCrLf & fouten, vbExclamation
        Exit Sub
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

    If r + nb > maxRij Then
        MsgBox "De formules lopen tot rij " & maxRij & ". Voer eerst WerkbereikVerruimen uit.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    voor = MeetStap8(wv, r + nb)

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
    For kk = 1 To nb
        SchrijfBijdrageRegel wv, r, r + kk, volgnr + kk, bron(kk), bed, kk
    Next kk

    ' formules controleren (zo nodig kopieren van de regel erboven)
    For kk = 0 To nb
        ControleerFormules wv, r + kk
    Next kk
    Application.Calculate

    ' --- nameting
    na = MeetStap8(wv, r + nb)
    If na(0) <> voor(0) + 1 + nb Then verschil = True
    If Abs(na(1) - (voor(1) + Nul(f(23)) + sumB(0))) > 0.005 Then verschil = True
    For i = 2 To 4
        If Abs(na(i) - voor(i)) > 0.005 Then verschil = True
    Next i
    For i = 24 To 28
        If Abs(na(5 + i - 24) - (voor(5 + i - 24) + Nul(f(i)) + sumB(i - 23))) > 0.005 Then verschil = True
    Next i
    If verschil Then
        For kk = 0 To nb
            RuimRij wv, r + kk
        Next kk
        Application.ScreenUpdating = True
        MsgBox "Verschil bij de nameting. De nieuwe regel is weer gewist. Er is niets toegevoegd.", vbCritical
        Exit Sub
    End If

    ' --- controle sleutel samengesteld krediet
    waarsch = ""
    If grp > 0 Then
        For x = 2 To r
            If VarType(wv.Cells(x, 10).Value) <> vbString And Not IsEmpty(wv.Cells(x, 10).Value) Then
                If CLng(wv.Cells(x, 10).Value) = grp Then
                    If VarType(wv.Cells(x, 11).Value) <> vbString And Not IsError(wv.Cells(x, 11).Value) Then som = som + wv.Cells(x, 11).Value
                End If
            End If
        Next x
        If Abs(som - 1) > 0.002 Then waarsch = "Let op: de verdeelsleutel van groep " & grp & " komt uit op " & Format(som * 100, "0.0") & "% in plaats van 100%. Controleer de budgetten van de deelkredieten."
    End If

    If Leeg(f(2)) Then kn = "" Else kn = CDbl(f(2))
    LogRegel "toevoegen", volgnr, kn, "", , , "rij " & r & IIf(grp > 0, ", groep " & grp, "")
    For kk = 1 To nb
        LogRegel "toevoegen bijdrage", volgnr + kk, kn, "", , , "rij " & (r + kk) & ", bron " & bron(kk)
    Next kk
    If nb > 0 Then bijtxt = " en " & nb & " bijdrage(n) van derden"
    KeuzelijstenVerversen True
    Application.ScreenUpdating = True

    If grp > 0 Then
        If MsgBox("Krediet toegevoegd op rij " & r & ", volgnummer " & volgnr & ", groep " & grp & bijtxt & "." & vbCrLf & _
                  IIf(Len(waarsch) > 0, waarsch & vbCrLf, "") & vbCrLf & _
                  "Nog een deelkrediet bij dit samengestelde krediet toevoegen?", vbYesNo + vbQuestion, "Krediet toevoegen") = vbYes Then
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
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description & vbCrLf & "Controleer in de verzamellijst of er een halve regel is weggeschreven.", vbCritical
End Sub

Public Sub KaartWissen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    wd.Range("Kaart_Invoer").ClearContents
    HerstelBijdrageLabels wd
    wd.Range("Kaart_Rijnummer").ClearContents
End Sub

' Maakt alleen de velden per deelkrediet leeg en zet de groep klaar voor het volgende deelkrediet.
Private Sub VolgendeDeelVoorbereiden(ByVal grp As Long)
    Dim wd As Worksheet, i As Variant, cel As Range, zoek As String, bj As Long, bk As Long
    Set wd = ThisWorkbook.Worksheets("Dashboard")
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
            wd.Cells(eerste + 7, 3).Value = cel.Value
            Exit For
        End If
    Next cel
    wd.Activate
    wd.Cells(eerste + 1, 3).Select
End Sub

Public Sub DashboardKnoppenPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, naam As Variant, i As Long
    Dim namen As Variant, titels As Variant, acties As Variant
    VerzamellijstHernoemen
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    For Each naam In Array("btnToevoegen", "btnWijzigen", "btnOpslaan", "btnKaartWissen", "btnLijsten", "btnVerwerken", "btnTijdschrijvers", _
                           "spnAfschrijving", "btnVerwijderen", "btnUitgaven", "btnNaamVervangen", "btnBhVervangen", "btnPlVervangen", "btnTsVervangen")
        ws.Shapes(CStr(naam)).Delete                 ' bestaat de knop niet, dan gewoon door (On Error Resume Next)
        ws.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    ws.Rows(6).RowHeight = 6                          ' de knoppen staan onder de invoerkaart en in blok 2
    ' knop naast het veld Tijdschrijvers
    Set b = ws.Buttons.Add(ws.Range("B" & KaartRij(14)).Left + ws.Range("B" & KaartRij(14)).Width - 112, ws.Range("B" & KaartRij(14)).Top + 1, 110, 16)
    b.Name = "btnTijdschrijvers"
    b.Caption = "Tijdschrijvers kiezen"
    b.OnAction = "TijdschrijversKiezen"
    ' onder de invoerkaart: Ingevuld? Verwerken! en Scherm leegmaken
    ws.Rows(KAART_KNOP).RowHeight = 38
    Set b = ws.Buttons.Add(ws.Range("B" & KAART_KNOP).Left, ws.Range("B" & KAART_KNOP).Top + 4, 230, 30)
    b.Name = "btnVerwerken"
    b.Caption = "Ingevuld? Verwerken!"
    b.OnAction = "Verwerken"
    Set b = ws.Buttons.Add(ws.Range("B" & KAART_KNOP).Left + 250, ws.Range("B" & KAART_KNOP).Top + 8, 160, 22)
    b.Name = "btnKaartWissen"
    b.Caption = "Scherm leegmaken"
    b.OnAction = "KaartWissen"
    Set b = ws.Buttons.Add(ws.Range("B" & KAART_KNOP).Left + 430, ws.Range("B" & KAART_KNOP).Top + 8, 160, 22)
    b.Name = "btnWijzigen"
    b.Caption = "Krediet wijzigen"
    b.OnAction = "KredietWijzigen"
    Set b = ws.Buttons.Add(ws.Range("B" & KAART_KNOP).Left + 610, ws.Range("B" & KAART_KNOP).Top + 8, 160, 22)
    b.Name = "btnVerwijderen"
    b.Caption = "Krediet verwijderen"
    b.OnAction = "KredietVerwijderen"
    UitdraaiBlokPlaatsen
    ControleblokPlaatsen
    FinAdmBlokPlaatsen
    KoppenStartEindjaar
    TabbladkiezerPlaatsen
    KnoppenTerugPlaatsen
    KnoppenUitlegPlaatsen
    TabbladenOrdenen
    DashboardEventInstalleren                    ' melding voor de financiele administratie bij invoer op een lege kaart
    DashboardNaarRol
    ws.Activate
End Sub

' Een knop voor toevoegen en wijzigen: is er een regel geladen om te wijzigen, dan wordt de wijziging opgeslagen,
' anders wordt een nieuw krediet toegevoegd.
Public Sub Verwerken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    If Len(CStr(ThisWorkbook.Worksheets("Dashboard").Range("Kaart_Rijnummer").Value)) > 0 Then
        WijzigingOpslaan
    Else
        KredietToevoegen
    End If
End Sub

' Maakt de laatste regel van de verzamellijst leeg (niet verwijderen). Alleen voor een zojuist toegevoegde, nog niet gecontroleerde regel.
Public Sub ProefregelWissen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, vandaag As Boolean
    On Error GoTo Fout
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    r = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    If r < 271 Then
        MsgBox "Er is geen nieuw toegevoegde regel om te wissen.", vbExclamation
        Exit Sub
    End If
    If CStr(wv.Cells(r, 42).Value) <> CtlTeControleren() And CStr(wv.Cells(r, 42).Value) <> "Afgekeurd" Then
        vandaag = False
        If IsDate(wv.Cells(r, 47).Value) Then vandaag = (CDate(wv.Cells(r, 47).Value) = Date)
        If Not vandaag Then
            MsgBox "De laatste regel is al gecontroleerd en is niet vandaag toegevoegd. Er is niets gewist.", vbExclamation
            Exit Sub
        End If
    End If
    If MsgBox("Regel " & r & " leegmaken?" & vbCrLf & "Volgnummer " & wv.Cells(r, 1).Value & ": " & wv.Cells(r, 4).Value, _
              vbYesNo + vbQuestion + vbDefaultButton2, "Proefregel wissen") <> vbYes Then Exit Sub
    RuimRij wv, r
    Application.Calculate
    LogRegel "regel leeggemaakt", , , "", , , "rij " & r
    MsgBox "Regel " & r & " is leeggemaakt.", vbInformation
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

Public Function Leeg(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        Leeg = True
    ElseIf IsError(v) Then
        Leeg = True
    Else
        Leeg = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Public Function Nul(ByVal v As Variant) As Double
    If Leeg(v) Then
        Nul = 0
    ElseIf IsNumeric(v) Then
        Nul = CDbl(v)
    End If
End Function

Private Function VoorCode(ByVal t As String) As String
    Dim p As Long
    p = InStr(t, " - ")
    If p > 0 Then VoorCode = Trim$(Left$(t, p - 1)) Else VoorCode = Trim$(t)
End Function

Private Function GetalOfTekst(ByVal v As Variant) As Variant
    If Not Leeg(v) And IsNumeric(v) Then
        GetalOfTekst = CDbl(v)
    Else
        GetalOfTekst = Trim$(CStr(v))
    End If
End Function

Private Function CodeInTabel(ByVal tabel As String, ByVal code As String) As Boolean
    Dim cel As Range
    For Each cel In ThisWorkbook.Worksheets("Parameters").ListObjects(tabel).ListColumns("Code").DataBodyRange.Cells
        If CStr(cel.Value) = code Then CodeInTabel = True: Exit Function
    Next cel
End Function

Public Function MeetStap8(ByVal ws As Worksheet, ByVal laatsteRij As Long) As Variant
    Dim res(0 To 9) As Double, kol As Variant, i As Long
    ' alleen de gegevensrijen (rij 2 t/m laatsteRij): onderaan staan som-regels die niet mogen meetellen
    res(0) = Application.WorksheetFunction.CountA(ws.Range("A2:A" & laatsteRij))
    i = 1
    For Each kol In Array("Z", "AA", "AB", "AC", "AE", "AF", "AG", "AH", "AI")
        res(i) = Application.WorksheetFunction.Sum(ws.Range(CStr(kol) & "2:" & CStr(kol) & laatsteRij))
        i = i + 1
    Next kol
    MeetStap8 = res
End Function

Private Sub RuimRij(ByVal wv As Worksheet, ByVal r As Long)
    Dim c As Variant
    For Each c In Array("A", "B", "C", "D", "F", "H", "I", "J", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", _
                        "X", "Y", "Z", "AE", "AF", "AG", "AH", "AI", "AJ", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW")
        wv.Range(c & r).ClearContents
    Next c
End Sub

' ======================================================================
' Korte omschrijving (maximaal 20 posities) afleiden uit de lange omschrijving
' ======================================================================
Public Function KortOmschrijving(ByVal lang As String) As String
    Dim s As String, t As String, w As Variant, pre As Variant
    Dim p1 As Long, p2 As Long, p As Long, i As Long, g As Long, lw As Long, lk As Long
    Dim code As String, rest As String
    Dim afk() As String, par() As String, tok() As String

    s = LCase$(Trim$(lang))
    Do While InStr(s, "(") > 0
        p1 = InStr(s, "(")
        p2 = InStr(p1, s, ")")
        If p2 > 0 Then
            s = Trim$(Left$(s, p1 - 1) & Mid$(s, p2 + 1))
        Else
            s = Trim$(Left$(s, p1 - 1))
        End If
    Loop
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)

    ' voertuig v07 - alg dienst  ->  V07-alg dienst
    If Left$(s, 9) = "voertuig " Then
        rest = Mid$(s, 10)
        p = InStr(rest, "-")
        If p > 1 Then
            code = Trim$(Left$(rest, p - 1))
            rest = Trim$(Mid$(rest, p + 1))
            If Len(code) >= 2 And Len(rest) > 0 Then
                If Left$(code, 1) = "v" And IsNumeric(Mid$(code, 2)) Then s = UCase$(code) & "-" & rest
            End If
        End If
    End If

    ' voorvoegsels die niets toevoegen
    For Each pre In Array("weg ", "or ", "ikc ")
        If Left$(s, Len(pre)) = pre Then s = Mid$(s, Len(pre) + 1)
    Next pre
    s = Trim$(s)
    If Len(s) <= 20 Then
        KortOmschrijving = s
        Exit Function
    End If

    ' overbodige woorden
    For Each w In Array("stelpost", "investeringen", "de", "en", "aan", "van", "naar")
        If Len(s) > 20 Then
            t = " " & s & " "
            If InStr(t, " " & w & " ") > 1 Then s = Trim$(Replace(t, " " & w & " ", " "))
        End If
    Next w

    ' afkortingen
    afk = Split(AfkortingenTekst(), "|")
    For i = 0 To UBound(afk)
        If Len(s) > 20 And Len(afk(i)) > 0 Then
            par = Split(afk(i), ">")
            If InStr(s, par(0)) > 0 Then s = Replace(s, par(0), par(1))
        End If
    Next i

    ' jaartallen inkorten (2026 -> 26)
    If Len(s) > 20 Then
        tok = Split(s, " ")
        For i = 0 To UBound(tok)
            If tok(i) Like "20##" Then tok(i) = Mid$(tok(i), 3)
        Next i
        s = Join(tok, " ")
    End If

    ' langste woord steeds met een letter inkorten
    Do While Len(s) > 20 And g < 100
        g = g + 1
        tok = Split(s, " ")
        lw = 0
        lk = 0
        For i = 0 To UBound(tok)
            If Len(tok(i)) > lw Then
                lw = Len(tok(i))
                lk = i
            End If
        Next i
        If lw <= 4 Then Exit Do
        tok(lk) = Left$(tok(lk), lw - 1)
        s = Join(tok, " ")
    Loop
    If Len(s) > 20 Then s = RTrim$(Left$(s, 20))
    KortOmschrijving = s
End Function

Private Function AfkortingenTekst() As String
    Dim a As String
    a = "verduurzaming>verduurz|verduurzamen>verduurz|installaties>instal|installatie>instal|"
    a = a & "inspectie>insp|inspec>insp|huisvesting>huisv|aankoop>aank|aanpak>aanp|communicatie>comm|"
    a = a & "kunstgrasveld>kunstgrv|kunstveld>kunstv|toplaag>topl|ondergrondse>og|ondergrond>onderg|"
    a = a & "wilhelmina>wilh|brandweerkazerne>bwkazerne|ruwbouw>ruwb|afbouw>afb|verbouwing>verb|"
    a = a & "renovatie>ren|kruising>kruis|verkeersveiligheid>verkeersveil|stationsomgeving>stationsomg|"
    a = a & "electrotechnische>electr tech|gemalen>gem|fietspad>fietsp|standplaatsen>standpl|"
    a = a & "deklagen>dekl|deklaag>dekl|salderes>sald|gymzaal>gymz|kleedlokalen>kleedl|"
    a = a & "meubilair>meub|onderhoud>onderh|riolering>riol|voorziening>voorz"
    AfkortingenTekst = a
End Function

' ======================================================================
' Krediet wijzigen (stap 9) en gedeelde controle van de invoerkaart
' ======================================================================
' Leest de invoerkaart en controleert alles. Geeft True als er mag worden opgeslagen.
' Veld op de kaart (23 t/m 28) met het budget dat verplicht is bij deze status: L met jaar t/m het lopend jaar ->
' veld 23 (budget t/m lopend jaar), L lopend jaar + 1 -> veld 24, ... L lopend jaar + 5 -> veld 28. Anders 0.
Private Function StatusBudgetVeld(ByVal stLetter As String, ByVal stJaar As String) As Long
    Dim lj As Long, j As Long
    If UCase$(Trim$(stLetter)) <> "L" Or Not IsNumeric(stJaar) Or Len(Trim$(stJaar)) <> 4 Then Exit Function
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    j = CLng(stJaar)
    If j <= lj Then
        StatusBudgetVeld = 23
    ElseIf j <= lj + 5 Then
        StatusBudgetVeld = 23 + (j - lj)
    End If
End Function

' True als het veld van de kaart bij het geladen krediet (rij wijzigRij) in de verzamellijst leeg was.
Private Function WasLeeg(ByVal wv As Worksheet, ByVal wijzigRij As Long, ByVal i As Long) As Boolean
    Dim kol As Variant
    If wijzigRij < 2 Then Exit Function
    kol = Array(0, 2, 3, 4, 6, 8, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 31, 32, 33, 34, 35, 36, 46)
    If i < 1 Or i > 30 Then Exit Function
    WasLeeg = Leeg(wv.Cells(wijzigRij, CLng(kol(i))).Value)
End Function

Private Function ControleerKaart(ByRef k As TKaart, ByVal wijzigRij As Long) As Boolean
    Dim wd As Worksheet, wv As Worksheet
    Dim fouten As String, waarsch As String, j8 As String
    Dim i As Long, laatste As Long, nDubbel As Long
    Dim cel As Range

    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set wv = Vz()
    For i = 1 To aantal
        k.f(i) = wd.Cells(KaartRij(i), KaartKol(i)).Value
    Next i

    ' --- verplichte velden
    ' Bij wijzigen van een bestaand (ook oud) krediet: een veld dat in de verzamellijst al leeg was, mag leeg blijven.
    For i = 1 To 22
        If i <> 4 And i <> 2 And i <> 21 And Leeg(k.f(i)) Then
            If Not WasLeeg(wv, wijzigRij, i) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), KaartKol(i) - 1).Value & " is verplicht" & vbCrLf
        End If
    Next i
    If Leeg(k.f(29)) And Not WasLeeg(wv, wijzigRij, 29) Then fouten = fouten & "- " & wd.Cells(KaartRij(29), 2).Value & " is verplicht bij wijzigen" & vbCrLf

    ' --- lengte omschrijvingen
    If Not Leeg(k.f(3)) Then If Len(Trim$(CStr(k.f(3)))) > 36 Then fouten = fouten & "- omschrijving lang is langer dan 36 posities" & vbCrLf
    If Not Leeg(k.f(4)) Then If Len(Trim$(CStr(k.f(4)))) > 20 Then fouten = fouten & "- omschrijving kort is langer dan 20 posities" & vbCrLf
    If Len(Trim$(CStr(k.f(29)))) > MAX_TOEL Then fouten = fouten & "- toelichting is langer dan " & MAX_TOEL & " tekens" & vbCrLf

    ' --- status
    k.Letter = UCase$(Trim$(CStr(k.f(5))))
    k.jaar = Trim$(CStr(k.f(6)))
    If Not Leeg(k.f(5)) Then If k.Letter <> "L" And k.Letter <> "A" Then fouten = fouten & "- status moet L of A zijn" & vbCrLf
    If Not Leeg(k.f(6)) Then If Len(k.jaar) <> 4 Or Not IsNumeric(k.jaar) Then fouten = fouten & "- eerste jaar krediet beschikbaar moet uit vier cijfers bestaan" & vbCrLf

    ' --- budget in het jaar van de status is verplicht bij een lopend krediet (behalve als het al leeg was)
    If StatusBudgetVeld(k.Letter, k.jaar) > 0 Then
        If Leeg(k.f(StatusBudgetVeld(k.Letter, k.jaar))) And Not WasLeeg(wv, wijzigRij, StatusBudgetVeld(k.Letter, k.jaar)) Then
            fouten = fouten & "- " & wd.Cells(KaartRij(StatusBudgetVeld(k.Letter, k.jaar)), 5).Value & " is verplicht bij status " & k.Letter & k.jaar & vbCrLf
        End If
    End If

    ' --- getallen
    If Not Leeg(k.f(2)) Then If Not IsNumeric(k.f(2)) Then fouten = fouten & "- kredietnummer moet een getal zijn" & vbCrLf
    If Not Leeg(k.f(10)) Then If Not IsNumeric(k.f(10)) Then fouten = fouten & "- startjaar uitgaven moet een jaartal zijn" & vbCrLf
    If Not Leeg(k.f(11)) Then If Not IsNumeric(k.f(11)) Then fouten = fouten & "- laatste jaar uitgaven moet een jaartal zijn" & vbCrLf
    If IsNumeric(k.f(10)) And IsNumeric(k.f(11)) And Not Leeg(k.f(10)) And Not Leeg(k.f(11)) Then
        If CDbl(k.f(10)) > CDbl(k.f(11)) Then fouten = fouten & "- startjaar uitgaven ligt na het laatste jaar uitgaven" & vbCrLf
    End If
    If Not Leeg(k.f(15)) Then If Not IsNumeric(k.f(15)) Then fouten = fouten & "- exploitatienummer moet een getal zijn" & vbCrLf
    If Not Leeg(k.f(22)) Then
        If Not IsNumeric(k.f(22)) Then
            fouten = fouten & "- afschrijftermijn moet een getal zijn" & vbCrLf
        ElseIf CDbl(k.f(22)) < 0 Or CDbl(k.f(22)) > 60 Then
            fouten = fouten & "- afschrijftermijn moet tussen 0 en 60 jaar liggen" & vbCrLf
        End If
    End If
    For i = 23 To 28
        If Not Leeg(k.f(i)) Then If Not IsNumeric(k.f(i)) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), KaartKol(i) - 1).Value & " moet een getal zijn" & vbCrLf
    Next i

    ' --- codes uit de tabellen
    k.pc = VoorCode(CStr(k.f(1)))
    k.tc = Replace(VoorCode(CStr(k.f(16))), ".", "")
    k.bc = VoorCode(CStr(k.f(17)))
    If Not Leeg(k.f(1)) Then If Not CodeInTabel("Tbl_Programma", k.pc) Then fouten = fouten & "- programma staat niet in de programmatabel (blad Parameters)" & vbCrLf
    If Not Leeg(k.f(16)) Then If Not CodeInTabel("Tbl_Taakveld", k.tc) Then fouten = fouten & "- taakveld staat niet in de taakveldentabel (blad Parameters)" & vbCrLf
    If Not Leeg(k.f(17)) Then If Not CodeInTabel("Tbl_Balanscode", k.bc) Then fouten = fouten & "- balanscode staat niet in de balanscodetabel (blad Parameters)" & vbCrLf

    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)

    ' --- samengesteld, standaard of stelpost
    j8 = LCase$(Trim$(CStr(k.f(8))))
    If Not Leeg(k.f(8)) Then
        If j8 = "standaard" Or j8 = "stelpost" Then
            k.jWaarde = j8
        ElseIf j8 = "nieuw samengesteld krediet" Then
            k.nieuw = True
            k.grp = CLng(Application.WorksheetFunction.Max(wv.Range("J2:J" & laatste))) + 1
            k.jWaarde = k.grp
            k.naam = Trim$(CStr(k.f(30)))
            If Len(k.naam) = 0 Then fouten = fouten & "- naam samengesteld krediet is verplicht bij een nieuw samengesteld krediet" & vbCrLf
        ElseIf Left$(j8, 6) = "groep " Then
            k.grp = CLng(Val(Mid$(j8, 7)))
            If k.grp < 1 Then
                fouten = fouten & "- kies de groep uit de lijst" & vbCrLf
            ElseIf Application.WorksheetFunction.CountIf(wv.Range("J2:J" & laatste), k.grp) = 0 Then
                fouten = fouten & "- groep " & k.grp & " bestaat niet" & vbCrLf
            Else
                k.jWaarde = k.grp
                For Each cel In wv.Range("J2:J" & laatste).Cells
                    If VarType(cel.Value) <> vbString Then
                        If Not IsEmpty(cel.Value) Then
                            If CLng(cel.Value) = k.grp Then
                                If Len(Trim$(CStr(wv.Cells(cel.Row, 46).Value))) > 0 Then
                                    k.naam = Trim$(CStr(wv.Cells(cel.Row, 46).Value))
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next cel
            End If
        Else
            fouten = fouten & "- kies bij samengesteld/standaard/stelpost een waarde uit de lijst" & vbCrLf
        End If
    End If

    If Len(fouten) > 0 Then
        MsgBox "Niet opgeslagen. Dit moet nog worden aangepast:" & vbCrLf & vbCrLf & fouten, vbExclamation
        Exit Function
    End If

    ' --- waarschuwingen
    If Not Leeg(k.f(2)) Then
        nDubbel = Application.WorksheetFunction.CountIf(wv.Range("C2:C" & laatste), CDbl(k.f(2)))
        If wijzigRij > 0 Then
            If IsNumeric(wv.Cells(wijzigRij, 3).Value) Then If CDbl(wv.Cells(wijzigRij, 3).Value) = CDbl(k.f(2)) Then nDubbel = nDubbel - 1
        End If
        If nDubbel > 0 Then waarsch = waarsch & "- kredietnummer " & k.f(2) & " bestaat al in de verzamellijst" & vbCrLf
    End If
    If Not Leeg(k.f(21)) Then
        If IsNumeric(k.f(21)) Then
            nDubbel = Application.WorksheetFunction.CountIf(wv.Range("X2:X" & laatste), CDbl(k.f(21)))
            If wijzigRij > 0 Then
                If IsNumeric(wv.Cells(wijzigRij, 24).Value) Then If CDbl(wv.Cells(wijzigRij, 24).Value) = CDbl(k.f(21)) Then nDubbel = nDubbel - 1
            End If
            If nDubbel > 0 Then waarsch = waarsch & "- activanummer " & k.f(21) & " bestaat al" & vbCrLf
        End If
    End If
    If Len(waarsch) > 0 Then
        If MsgBox("Let op:" & vbCrLf & vbCrLf & waarsch & vbCrLf & "Toch doorgaan?", vbYesNo + vbExclamation + vbDefaultButton2, "Controle") <> vbYes Then Exit Function
    End If

    ControleerKaart = True
End Function

' Zet de waarden van de kaart om naar celwaarden voor de verzamellijst (kolomnummer + waarde).
Private Function KaartNaarWaarden(ByRef k As TKaart, ByRef cols() As Long, ByRef vals() As Variant) As Long
    Dim n As Long, i As Long, tv As Variant, ev As Variant
    ReDim cols(1 To 45)
    ReDim vals(1 To 45)
    If Len(k.pc) = 0 Then ev = Empty Else ev = CLng(k.pc)
    VoegWaarde cols, vals, n, 2, ev
    If Leeg(k.f(2)) Then ev = Empty Else ev = CDbl(k.f(2))
    VoegWaarde cols, vals, n, 3, ev
    VoegWaarde cols, vals, n, 4, Trim$(CStr(k.f(3)))
    VoegWaarde cols, vals, n, 8, k.Letter & k.jaar
    VoegWaarde cols, vals, n, 9, LCase$(Trim$(CStr(k.f(7))))
    VoegWaarde cols, vals, n, 10, k.jWaarde
    VoegWaarde cols, vals, n, 12, GetalOfTekst(k.f(9))
    If Leeg(k.f(10)) Then ev = Empty Else ev = CDbl(k.f(10))
    VoegWaarde cols, vals, n, 13, ev
    If Leeg(k.f(11)) Then ev = Empty Else ev = CDbl(k.f(11))
    VoegWaarde cols, vals, n, 14, ev
    VoegWaarde cols, vals, n, 15, Trim$(CStr(k.f(12)))
    VoegWaarde cols, vals, n, 16, Trim$(CStr(k.f(13)))
    VoegWaarde cols, vals, n, 17, Trim$(CStr(k.f(14)))
    If Leeg(k.f(15)) Then ev = Empty Else ev = CDbl(k.f(15))
    VoegWaarde cols, vals, n, 18, ev
    If Len(k.tc) = 0 Then
        tv = Empty
    ElseIf Left$(k.tc, 1) = "0" Then
        tv = k.tc
    Else
        tv = CLng(k.tc)
    End If
    VoegWaarde cols, vals, n, 19, tv
    VoegWaarde cols, vals, n, 20, k.bc
    VoegWaarde cols, vals, n, 21, Trim$(CStr(k.f(18)))
    VoegWaarde cols, vals, n, 22, Trim$(CStr(k.f(19)))
    VoegWaarde cols, vals, n, 23, Trim$(CStr(k.f(20)))
    If Leeg(k.f(21)) Then ev = Empty Else ev = GetalOfTekst(k.f(21))
    VoegWaarde cols, vals, n, 24, ev
    If Leeg(k.f(22)) Then ev = Empty Else ev = CDbl(k.f(22))
    VoegWaarde cols, vals, n, 25, ev
    If Leeg(k.f(23)) Then ev = Empty Else ev = CDbl(k.f(23))
    VoegWaarde cols, vals, n, 26, ev
    For i = 24 To 28
        If Leeg(k.f(i)) Then ev = Empty Else ev = CDbl(k.f(i))
        VoegWaarde cols, vals, n, 31 + i - 24, ev
    Next i
    VoegWaarde cols, vals, n, 36, Trim$(CStr(k.f(29)))
    If Len(k.naam) > 0 Then VoegWaarde cols, vals, n, 46, k.naam
    KaartNaarWaarden = n
End Function

Private Sub VoegWaarde(ByRef cols() As Long, ByRef vals() As Variant, ByRef n As Long, ByVal c As Long, ByVal v As Variant)
    n = n + 1
    cols(n) = c
    vals(n) = v
End Sub

Private Sub ZetCel(ByVal wv As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    If IsEmpty(v) Then
        wv.Cells(r, c).ClearContents
    Else
        If VarType(v) = vbString Then
            If IsNumeric(v) Then wv.Cells(r, c).NumberFormat = "@"
        End If
        wv.Cells(r, c).Value = v
    End If
End Sub

Private Function GelijkWaarde(ByVal a As Variant, ByVal b As Variant) As Boolean
    If Leeg(a) And Leeg(b) Then
        GelijkWaarde = True
    ElseIf IsEmpty(a) Or IsEmpty(b) Or IsError(a) Or IsError(b) Then
        GelijkWaarde = False
    ElseIf VarType(a) <> vbString And VarType(b) <> vbString Then
        GelijkWaarde = (Abs(CDbl(a) - CDbl(b)) < 0.000001)
    Else
        GelijkWaarde = (CStr(a) = CStr(b))
    End If
End Function

' Zoekt een krediet en zet het op de invoerkaart.
Public Sub KredietWijzigen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, wd As Worksheet
    Dim keuze As String, zoek As String, lijst As String, antw As String
    Dim r As Long, laatste As Long, n As Long, kk As Long, tonen As Long
    Dim treffers(1 To 10) As Long, meer As Boolean, hit As Boolean

    On Error GoTo Fout
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "Alleen financieel adviseurs en de financi" & ChrW(235) & "le administratie mogen kredieten wijzigen.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then
        If MsgBox("Er staat al een regel op de kaart. Vervangen?", vbYesNo + vbQuestion, "Krediet wijzigen") <> vbYes Then Exit Sub
    End If

    keuze = Trim$(InputBox("Zoeken op:" & vbCrLf & "1 = kredietnummer" & vbCrLf & "2 = omschrijving (een deel van de tekst)" & vbCrLf & _
                           "3 = activanummer" & vbCrLf & "4 = volgnummer", "Krediet wijzigen", "1"))
    If Len(keuze) = 0 Then Exit Sub
    If keuze <> "1" And keuze <> "2" And keuze <> "3" And keuze <> "4" Then
        MsgBox "Typ 1, 2, 3 of 4.", vbExclamation
        Exit Sub
    End If
    zoek = Trim$(InputBox("Wat zoekt u?", "Krediet wijzigen"))
    If Len(zoek) = 0 Then Exit Sub

    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        hit = False
        Select Case keuze
            Case "1"
                hit = (CStr(wv.Cells(r, 3).Value) = zoek)
            Case "2"
                hit = (InStr(1, CStr(wv.Cells(r, 4).Value), zoek, vbTextCompare) > 0)
            Case "3"
                hit = (CStr(wv.Cells(r, 24).Value) = zoek)
            Case "4"
                hit = (CStr(wv.Cells(r, 1).Value) = zoek)
        End Select
        If hit Then
            If n < 10 Then
                n = n + 1
                treffers(n) = r
            Else
                meer = True
            End If
        End If
    Next r

    If n = 0 Then
        MsgBox "Niets gevonden.", vbInformation
        Exit Sub
    End If
    If n = 1 Then
        r = treffers(1)
    Else
        For kk = 1 To n
            r = treffers(kk)
            lijst = lijst & kk & " = vnr " & wv.Cells(r, 1).Value & " | " & wv.Cells(r, 3).Value & " | " & _
                    Left$(CStr(wv.Cells(r, 4).Value), 30) & " | " & wv.Cells(r, 8).Value & vbCrLf
        Next kk
        antw = InputBox("Meerdere treffers" & IIf(meer, " (eerste 10, verfijn de zoekterm)", "") & ":" & vbCrLf & vbCrLf & lijst & vbCrLf & _
                        "Typ het nummer:", "Krediet wijzigen")
        If Len(antw) = 0 Then Exit Sub
        If Not IsNumeric(antw) Then
            MsgBox "Dat is geen nummer.", vbExclamation
            Exit Sub
        End If
        kk = CLng(antw)
        If kk < 1 Or kk > n Then
            MsgBox "Dat nummer staat niet in de lijst.", vbExclamation
            Exit Sub
        End If
        r = treffers(kk)
    End If

    ToonOpKaart r
    MsgBox "Volgnummer " & wv.Cells(r, 1).Value & " (rij " & r & ") staat op de kaart." & vbCrLf & _
           "Pas de gegevens aan en klik op Wijziging opslaan.", vbInformation
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

Public Sub ToonOpKaart(ByVal r As Long)
    Dim wd As Worksheet, wv As Worksheet, h As String, v As Variant
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set wv = Vz()
    KaartWissen
    wd.Range("Kaart_Rijnummer").Value = r
    ZetKaartVeld wd, 1, CodeNaarWeergave("Tbl_Programma", wv.Cells(r, 2).Value)
    ZetKaartVeld wd, 2, wv.Cells(r, 3).Value
    ZetKaartVeld wd, 3, wv.Cells(r, 4).Value
    h = CStr(wv.Cells(r, 8).Value)
    ZetKaartVeld wd, 5, UCase$(Left$(h, 1))
    ZetKaartVeld wd, 6, Mid$(h, 2)
    ZetKaartVeld wd, 7, wv.Cells(r, 9).Value
    v = wv.Cells(r, 10).Value
    If VarType(v) = vbString Then
        ZetKaartVeld wd, 8, v
    ElseIf Not IsEmpty(v) Then
        ZetKaartVeld wd, 8, GroepWeergave(CLng(v))
    End If
    ZetKaartVeld wd, 9, wv.Cells(r, 12).Value
    ZetKaartVeld wd, 10, wv.Cells(r, 13).Value
    ZetKaartVeld wd, 11, wv.Cells(r, 14).Value
    ZetKaartVeld wd, 12, wv.Cells(r, 15).Value
    ZetKaartVeld wd, 13, wv.Cells(r, 16).Value
    ZetKaartVeld wd, 14, wv.Cells(r, 17).Value
    ZetKaartVeld wd, 15, wv.Cells(r, 18).Value
    ZetKaartVeld wd, 16, CodeNaarWeergave("Tbl_Taakveld", wv.Cells(r, 19).Value)
    ZetKaartVeld wd, 17, CodeNaarWeergave("Tbl_Balanscode", wv.Cells(r, 20).Value)
    ZetKaartVeld wd, 18, wv.Cells(r, 21).Value
    ZetKaartVeld wd, 19, wv.Cells(r, 22).Value
    ZetKaartVeld wd, 20, wv.Cells(r, 23).Value
    ZetKaartVeld wd, 21, wv.Cells(r, 24).Value
    ZetKaartVeld wd, 22, wv.Cells(r, 25).Value
    ZetKaartVeld wd, 23, wv.Cells(r, 26).Value
    ZetKaartVeld wd, 24, wv.Cells(r, 31).Value
    ZetKaartVeld wd, 25, wv.Cells(r, 32).Value
    ZetKaartVeld wd, 26, wv.Cells(r, 33).Value
    ZetKaartVeld wd, 27, wv.Cells(r, 34).Value
    ZetKaartVeld wd, 28, wv.Cells(r, 35).Value
    ZetKaartVeld wd, 29, wv.Cells(r, 36).Value
    wd.Activate
    wd.Cells(eerste, 3).Select
End Sub

Private Sub ZetKaartVeld(ByVal wd As Worksheet, ByVal i As Long, ByVal v As Variant)
    If IsEmpty(v) Then Exit Sub
    If IsError(v) Then Exit Sub
    wd.Cells(KaartRij(i), KaartKol(i)).Value = v
End Sub

Public Function CodeNaarWeergave(ByVal tabel As String, ByVal code As Variant) As Variant
    Dim lo As ListObject, i As Long
    If IsEmpty(code) Then Exit Function
    CodeNaarWeergave = CStr(code)
    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects(tabel)
    For i = 1 To lo.ListRows.Count
        If CStr(lo.ListColumns("Code").DataBodyRange.Cells(i, 1).Value) = CStr(code) Then
            CodeNaarWeergave = CStr(lo.ListColumns("Weergave").DataBodyRange.Cells(i, 1).Value)
            Exit Function
        End If
    Next i
End Function

Private Function GroepWeergave(ByVal n As Long) As String
    Dim cel As Range, zoek As String
    zoek = "groep " & n & " - "
    GroepWeergave = "groep " & n
    For Each cel In ThisWorkbook.Names("Lst_SamengesteldStandaardStelpost").RefersToRange.Cells
        If Left$(CStr(cel.Value), Len(zoek)) = zoek Then
            GroepWeergave = CStr(cel.Value)
            Exit For
        End If
    Next cel
End Function

' Slaat de gewijzigde gegevens van de geladen regel op. Alleen gewijzigde cellen worden weggeschreven.
Public Sub WijzigingOpslaan()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, wv As Worksheet
    Dim k As TKaart
    Dim r As Long, i As Long, n As Long, nw As Long, x As Long
    Dim cols() As Long, vals() As Variant
    Dim wijzig(1 To 45) As Boolean, oudW(1 To 45) As Variant
    Dim voor As Variant, na As Variant, verschil As Boolean
    Dim dlt(0 To 9) As Double, som As Double, waarsch As String
    Dim keuze As VbMsgBoxResult, bevestig As Boolean, gebr As String, herAanbieden As Boolean
    Dim nb As Long, kk As Long, jj As Long, fb As String, laatste As Long, rn As Long, volgnr0 As Long, maxRij As Long
    Dim bron(1 To 12) As String, bed(1 To 12, 0 To 5) As Variant, sumB(0 To 5) As Double, ext(0 To 9) As Double

    On Error GoTo Fout
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "Alleen financieel adviseurs en de financi" & ChrW(235) & "le administratie mogen kredieten wijzigen.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set wv = Vz()
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) = 0 Then
        MsgBox "Er is geen regel geladen. Gebruik eerst Krediet wijzigen.", vbExclamation
        Exit Sub
    End If
    r = CLng(wd.Range("Kaart_Rijnummer").Value)
    If r < 2 Or IsEmpty(wv.Cells(r, 1).Value) Then
        MsgBox "De geladen rij bestaat niet meer. Gebruik Scherm leegmaken en laad het krediet opnieuw.", vbCritical
        Exit Sub
    End If

    If Not ControleerKaart(k, r) Then Exit Sub
    n = KaartNaarWaarden(k, cols, vals)
    If Not Leeg(k.f(4)) Then
        VoegWaarde cols, vals, n, 6, Trim$(CStr(k.f(4)))
    ElseIf Not GelijkWaarde(wv.Cells(r, 4).Value, Trim$(CStr(k.f(3)))) Then
        VoegWaarde cols, vals, n, 6, KortOmschrijving(Trim$(CStr(k.f(3))))
    End If

    For i = 1 To n
        If Not GelijkWaarde(wv.Cells(r, cols(i)).Value, vals(i)) Then
            nw = nw + 1
            wijzig(i) = True
        End If
    Next i

    ' bijdragen van derden (nieuwe regels)
    nb = LeesBijdragen(bron, bed, fb)
    If Len(fb) > 0 Then
        MsgBox "Niet opgeslagen. Dit moet nog worden aangepast:" & vbCrLf & vbCrLf & fb, vbExclamation
        Exit Sub
    End If
    If nw = 0 And nb = 0 Then
        If CStr(wv.Cells(r, 42).Value) = "Afgekeurd" And Not HeeftRecht("Bevestigen") Then
            AanbiedenTerControle r                ' niets gewijzigd: eventueel toch opnieuw aanbieden
            KaartWissen
        Else
            MsgBox "Er is niets gewijzigd.", vbInformation
        End If
        Exit Sub
    End If
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    rn = CLng(ThisWorkbook.Names("Par_EersteVrijeRij").RefersToRange.Value)
    volgnr0 = CLng(ThisWorkbook.Names("Par_VolgendVolgnr").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    If nb > 0 Then
        If rn + nb - 1 > maxRij Or rn <> laatste + 1 Then
            MsgBox "Er is geen ruimte voor de bijdragen (formules lopen tot rij " & maxRij & "). Voer eerst WerkbereikVerruimen uit.", vbExclamation
            Exit Sub
        End If
    End If
    For kk = 1 To nb
        For jj = 0 To 5
            sumB(jj) = sumB(jj) + Nul(bed(kk, jj))
        Next jj
    Next kk
    ext(1) = sumB(0)
    For jj = 1 To 5
        ext(4 + jj) = sumB(jj)
    Next jj

    If HeeftRecht("Bevestigen") Then
        keuze = MsgBox(nw & " wijziging(en)" & IIf(nb > 0, " en " & nb & " bijdrage(n)", "") & " opslaan." & vbCrLf & vbCrLf & "Ja = opslaan en meteen als gecontroleerd bevestigen" & vbCrLf & _
                       "Nee = opslaan, nog te controleren" & vbCrLf & "Annuleren = niet opslaan", vbYesNoCancel + vbQuestion, "Wijziging opslaan")
        If keuze = vbCancel Then Exit Sub
        bevestig = (keuze = vbYes)
    End If

    Application.ScreenUpdating = False
    voor = MeetStap8(wv, laatste + nb)
    For i = 1 To n
        If wijzig(i) Then
            oudW(i) = wv.Cells(r, cols(i)).Value
            ZetCel wv, r, cols(i), vals(i)
            If cols(i) = 26 Then dlt(1) = dlt(1) + Nul(vals(i)) - Nul(oudW(i))
            If cols(i) >= 31 And cols(i) <= 35 Then dlt(5 + cols(i) - 31) = dlt(5 + cols(i) - 31) + Nul(vals(i)) - Nul(oudW(i))
        End If
    Next i
    For kk = 1 To nb
        SchrijfBijdrageRegel wv, r, rn + kk - 1, volgnr0 + kk - 1, bron(kk), bed, kk
        ControleerFormules wv, rn + kk - 1
    Next kk
    Application.Calculate

    na = MeetStap8(wv, laatste + nb)
    If na(0) <> voor(0) + nb Then verschil = True
    For i = 1 To 9
        If Abs(na(i) - (voor(i) + dlt(i) + ext(i))) > 0.005 Then verschil = True
    Next i
    If verschil Then
        For i = 1 To n
            If wijzig(i) Then ZetCel wv, r, cols(i), oudW(i)
        Next i
        For kk = 1 To nb
            RuimRij wv, rn + kk - 1
        Next kk
        Application.Calculate
        Application.ScreenUpdating = True
        MsgBox "Verschil bij de nameting. De wijzigingen zijn teruggezet. Er is niets gewijzigd.", vbCritical
        Exit Sub
    End If

    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    If nw > 0 Then
        If bevestig Then
            wv.Cells(r, 42).Value = CtlGecontroleerd()
            wv.Cells(r, 43).Value = gebr
            wv.Cells(r, 44).NumberFormat = "dd-mm-yyyy"
            wv.Cells(r, 44).Value = Date
        ElseIf CStr(wv.Cells(r, 42).Value) = "Afgekeurd" And Not HeeftRecht("Bevestigen") Then
            ' afgekeurd krediet: blijft afgekeurd tot de adviseur het opnieuw aanbiedt (hieronder of later met de knop)
            herAanbieden = True
        Else
            wv.Cells(r, 42).Value = CtlTeControleren()
            wv.Cells(r, 43).ClearContents
            wv.Cells(r, 44).ClearContents
        End If
        wv.Cells(r, 48).Value = gebr
        wv.Cells(r, 49).Value = Now
    End If

    For i = 1 To n
        If wijzig(i) Then LogRegel "wijzigen", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, CStr(wv.Cells(1, cols(i)).Value), oudW(i), vals(i), "rij " & r
    Next i
    For kk = 1 To nb
        LogRegel "toevoegen bijdrage", volgnr0 + kk - 1, wv.Cells(r, 3).Value, "", , , "rij " & (rn + kk - 1) & ", bron " & bron(kk) & ", bij volgnummer " & wv.Cells(r, 1).Value
    Next kk

    If k.grp > 0 Then
        For x = 2 To CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
            If VarType(wv.Cells(x, 10).Value) <> vbString And Not IsEmpty(wv.Cells(x, 10).Value) Then
                If CLng(wv.Cells(x, 10).Value) = k.grp Then
                    If VarType(wv.Cells(x, 11).Value) <> vbString And Not IsError(wv.Cells(x, 11).Value) Then som = som + wv.Cells(x, 11).Value
                End If
            End If
        Next x
        If Abs(som - 1) > 0.002 Then waarsch = vbCrLf & vbCrLf & "Let op: de verdeelsleutel van groep " & k.grp & " komt uit op " & Format(som * 100, "0.0") & "% in plaats van 100%."
    End If

    KeuzelijstenVerversen True
    KaartWissen
    Application.ScreenUpdating = True
    MsgBox "Opgeslagen (" & nw & " veld(en) gewijzigd" & IIf(nb > 0, ", " & nb & " bijdrage(n) toegevoegd", "") & ", rij " & r & ")." & waarsch, vbInformation
    If herAanbieden Then AanbiedenTerControle r
    Exit Sub

Fout:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description & vbCrLf & "Controleer in de verzamellijst rij " & r & ".", vbCritical
End Sub

' Knop op het Dashboard: ververst alle keuzelijsten.
Public Sub ListenVerversen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    KeuzelijstenVerversen False
    MsgBox "De keuzeopties zijn bijgewerkt.", vbInformation
End Sub

' ======================================================================
' Controle door de financiele administratie (stap 10)
' ======================================================================
' Maakt het tabblad Controle met de lijst van te controleren en afgekeurde regels en de knoppen.
Public Sub ControleBladAanmaken()
    FinAdmBlokPlaatsen                           ' de controle staat voortaan op het Dashboard (blok 3)
End Sub

' Blok 4 op het Dashboard: Financiele administratie. Alleen knoppen (de regels zelf worden in de Verzamellijst
' geopend). Standaard: Volgend krediet controleren. Incidenteel: Bestaand krediet aanpassen.
Public Sub FinAdmBlokPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, naam As Variant
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("A" & FA_KOP & ":H" & (CB_KOP - 1)).Clear
    ws.Range("B" & FA_KOP).Value = "3. Financi" & ChrW(235) & "le administratie"
    ws.Range("B" & FA_KOP).Font.Bold = True
    ws.Range("B" & FA_KOP).Font.Size = 12
    ws.Range("C" & FA_KOP).Formula = "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)&"" te controleren, ""&COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")&"" afgekeurd"""
    ws.Range("C" & FA_KOP).Font.Bold = True
    ws.Range("B" & (FA_KOP + 1)).Value = "Standaard: nieuwe kredieten controleren. Incidenteel: een bestaand (al goedgekeurd) krediet aanpassen. " & _
                                          "Beide in de Verzamellijst; daarna bovenaan op 'Klaar: goedkeuren'."
    ws.Range("B" & (FA_KOP + 1)).Font.Italic = True
    ws.Range("B" & (FA_KOP + 1)).Font.Color = RGB(89, 89, 89)
    On Error Resume Next
    For Each naam In Array("btnFaVernieuwen", "btnFaOpenen", "btnFaBevestigen", "btnFaAfkeuren", "btnFaAanpassen", "btnFaVolgende")
        ws.Shapes(CStr(naam)).Delete
        ws.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    ws.Rows(FA_KOP + 2).RowHeight = 26
    Set b = ws.Buttons.Add(ws.Range("B" & (FA_KOP + 2)).Left, ws.Range("B" & (FA_KOP + 2)).Top + 2, 210, 22)
    b.Name = "btnFaVolgende"
    b.Caption = "Volgend krediet controleren"
    b.OnAction = "FaVolgende"
    Set b = ws.Buttons.Add(ws.Range("B" & (FA_KOP + 2)).Left + 220, ws.Range("B" & (FA_KOP + 2)).Top + 2, 210, 22)
    b.Name = "btnFaAanpassen"
    b.Caption = "Bestaand krediet aanpassen"
    b.OnAction = "GoedgekeurdKredietAanpassen"
    KnopUitleg ws, "btnFaVolgende", "Volgend krediet controleren (standaard, alleen financi" & ChrW(235) & "le administratie): gaat naar het eerste " & _
               "krediet met 'Te controleren' in de Verzamellijst en zet die regel open. Pas zo nodig aan en klik bovenaan op " & _
               "'Klaar: goedkeuren' (of 'Afkeuren'). De teller gaat omlaag. Niets te controleren? Dan krijgt u een melding."
    KnopUitleg ws, "btnFaAanpassen", "Bestaand krediet aanpassen (incidenteel, bijv. andere wetgeving of uitgangspunten): zoek op kredietnummer, " & _
               "activanummer of een deel van de omschrijving. U gaat naar die regel in de Verzamellijst; pas aan en klik bovenaan op " & _
               "'Klaar: goedkeuren'. Wijzigingen gaan naar het Logboek."
    ' het oude tabblad Controle is niet meer nodig
    If BladBestaatStil("Controle") Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("Controle").Unprotect WACHTWOORD
        ThisWorkbook.Worksheets("Controle").Delete
        Application.DisplayAlerts = True
    End If
End Sub

' Bestond voor de lijst in het oude blok 6; de teller in blok 4 is een formule en rekent zelf. Blijft bestaan omdat
' andere macro's hem aanroepen.
Public Sub ControleLijstVernieuwen()
End Sub

' Knop Volgend krediet controleren: opent het eerste krediet met 'Te controleren' in de Verzamellijst.
Public Sub FaVolgende()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long
    If Not FaRecht() Then Exit Sub
    If WerkRij() > 0 Then
        FaRegelOpenen WerkRij(), "controle"       ' er staat nog een regel open: daar verder
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        If CStr(wv.Cells(r, 42).Value) = CtlTeControleren() Then
            FaRegelOpenen r, "controle"
            Exit Sub
        End If
    Next r
    MsgBox "Er zijn geen kredieten te controleren.", vbInformation, "Volgend krediet controleren"
End Sub

' Rij van de open regel; zonder open regel een melding (de lijst op het Dashboard bestaat niet meer).
Private Function ControleGekozenRij() As Long
    MsgBox "Er staat geen regel open. Klik op het Dashboard in blok 3 op 'Volgend krediet controleren' of 'Bestaand krediet aanpassen'.", vbExclamation
End Function

' Blok 4 is alleen voor de financiele administratie (recht Bevestigen); financieel adviseurs zijn daartoe niet bevoegd.
Private Function FaRecht() As Boolean
    FaRecht = HeeftRecht("Bevestigen")
    If Not FaRecht Then MsgBox "Blok 3 is alleen voor de financi" & ChrW(235) & "le administratie.", vbExclamation
End Function

' Knop Lijst vernieuwen in blok 6.
Public Sub FaLijstVernieuwen()
    If Not FaRecht() Then Exit Sub
    ControleLijstVernieuwen
End Sub

' ----------------------------------------------------------------------
' Werkwijze financiele administratie: regel openen in de Verzamellijst, daar aanpassen, daarna Klaar.
'   Standaard  : knop 'Volgend krediet controleren' (blok 4).
'   Incidenteel: knop 'Bestaand krediet aanpassen' (zoeken op kredietnummer, activanummer of omschrijving).
' De invoercellen van alleen die ene regel gaan open. Bij openen wordt de regel vastgelegd op het verborgen blad
' 'Werkregel controle'; bij Klaar/Afkeuren/Sluiten worden de gewijzigde velden in het Logboek gezet (oud -> nieuw)
' en gaat de regel weer op slot.
' ----------------------------------------------------------------------

' Invoerkolommen van de verzamellijst die de financiele administratie in de regel mag aanpassen
' (niet: volgnr, formulekolommen E, G, K, AD, AL t/m AN, werkelijke uitgaven AA t/m AC en de controlekolommen).
Private Function FaInvoerKolommen() As Variant
    FaInvoerKolommen = Array(2, 3, 4, 6, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 31, 32, 33, 34, 35, 36, 46)
End Function

Private Function WerkBlad(ByVal maken As Boolean) As Worksheet
    If BladBestaatStil(BL_WERK) Then
        Set WerkBlad = ThisWorkbook.Worksheets(BL_WERK)
    ElseIf maken Then
        Set WerkBlad = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        WerkBlad.Name = BL_WERK
        WerkBlad.Visible = xlSheetVeryHidden
    End If
End Function

' Rij van de regel die nu open staat (0 = geen).
Private Function WerkRij() As Long
    Dim ww As Worksheet, wv As Worksheet, v As Variant
    Set ww = WerkBlad(False)
    If ww Is Nothing Then Exit Function
    If IsEmpty(ww.Range("B1").Value) Then Exit Function
    Set wv = Vz()
    If wv.Cells(CLng(ww.Range("A1").Value), 1).Value = ww.Range("B1").Value Then
        WerkRij = CLng(ww.Range("A1").Value)
    Else
        v = Application.Match(ww.Range("B1").Value, wv.Columns(1), 0)
        If Not IsError(v) Then WerkRij = CLng(v)
    End If
End Function

' Zet een regel open in de Verzamellijst. soort = "controle" (nieuw krediet) of "aanpassen" (bestaand krediet).
Private Sub FaRegelOpenen(ByVal r As Long, ByVal soort As String)
    Dim wv As Worksheet, ww As Worksheet, c As Variant, open0 As Long
    Set wv = Vz()
    open0 = WerkRij()
    If open0 > 0 And open0 <> r Then
        MsgBox "Volgnummer " & wv.Cells(open0, 1).Value & " staat nog open in de Verzamellijst (rij " & open0 & ")." & vbCrLf & _
               "Rond die eerst af met 'Klaar: goedkeuren', 'Afkeuren' of 'Sluiten zonder goedkeuren'.", vbExclamation
        Application.Goto wv.Cells(open0, 2), True
        Exit Sub
    End If
    If open0 = 0 Then
        Set ww = WerkBlad(True)
        ww.Cells.Clear
        ww.Range("A1").Value = r
        ww.Range("B1").Value = wv.Cells(r, 1).Value
        ww.Range("C1").Value = soort
        ww.Range("A2").Resize(1, 49).Value = wv.Range("A" & r & ":AW" & r).Value
    End If
    wv.Unprotect WACHTWOORD
    For Each c In FaInvoerKolommen()
        wv.Cells(r, CLng(c)).Locked = False
    Next c
    FaKnoppenVerzamellijst True
    BladBeveiligen wv
    wv.Activate
    Application.Goto wv.Cells(r, 2), True
    LogRegel IIf(soort = "controle", "controle geopend", "aanpassen geopend"), wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "", , , "rij " & r
    MsgBox "Volgnummer " & wv.Cells(r, 1).Value & " (rij " & r & ") staat open. Pas de regel hier in de Verzamellijst aan." & vbCrLf & vbCrLf & _
           "Klaar? Klik bovenaan op 'Klaar: goedkeuren'." & vbCrLf & _
           "Werkelijke uitgaven (AA t/m AC) en formulekolommen blijven op slot.", vbInformation, "Regel open"
End Sub

' Knoppen bovenaan de Verzamellijst zolang er een regel open staat.
Private Sub FaKnoppenVerzamellijst(ByVal tonen As Boolean)
    Dim wv As Worksheet, b As Object, naam As Variant, x As Double, i As Long, namen As Variant, titels As Variant, acties As Variant
    Set wv = Vz()
    On Error Resume Next
    For Each naam In Array("btnVzKlaar", "btnVzAfkeuren", "btnVzSluiten")
        wv.Shapes(CStr(naam)).Delete
        wv.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    If Not tonen Then Exit Sub
    x = wv.Range("A1").Left + 140
    namen = Array("btnVzKlaar", "btnVzAfkeuren", "btnVzSluiten")
    titels = Array("Klaar: goedkeuren", "Afkeuren", "Sluiten zonder goedkeuren")
    acties = Array("RegelBevestigen", "RegelAfkeuren", "FaRegelSluiten")
    For i = 0 To 2
        Set b = wv.Buttons.Add(x, wv.Range("A1").Top + 1, IIf(i = 2, 170, 130), 18)
        b.Name = namen(i)
        b.Caption = titels(i)
        b.OnAction = acties(i)
        b.Placement = xlFreeFloating
        b.PrintObject = False
        x = x + b.Width + 6
    Next i
    KnopUitleg wv, "btnVzKlaar", "Klaar: goedkeuren. De regel die open staat is gecontroleerd. De wijzigingen gaan naar het Logboek, " & _
               "de regel gaat weer op slot en de teller in blok 3 (Dashboard) gaat omlaag."
    KnopUitleg wv, "btnVzAfkeuren", "Afkeuren: de regel die open staat klopt niet. Typ de reden; de financieel adviseur moet hem herstellen. " & _
               "De regel blijft als afgekeurd staan (zie de teller in blok 3)."
    KnopUitleg wv, "btnVzSluiten", "Sluiten zonder goedkeuren: de regel gaat weer op slot. Gemaakte wijzigingen blijven staan en gaan naar " & _
               "het Logboek; de controlestatus verandert niet."
End Sub

' Controle van de aangepaste regel (zelfde regels als de invoerkaart, voor zover die op een regel te controleren zijn).
Private Function FaRegelFouten(ByVal wv As Worksheet, ByVal r As Long) As String
    Dim f As String, st As String, c As Variant
    If Len(Trim$(CStr(wv.Cells(r, 4).Value))) = 0 Then f = f & "- omschrijving (kolom D) is leeg" & vbCrLf
    If Len(Trim$(CStr(wv.Cells(r, 4).Value))) > 36 Then f = f & "- omschrijving (kolom D) is langer dan 36 posities" & vbCrLf
    If Len(Trim$(CStr(wv.Cells(r, 6).Value))) > 20 Then f = f & "- korte omschrijving (kolom F) is langer dan 20 posities" & vbCrLf
    If Len(Trim$(CStr(wv.Cells(r, 36).Value))) > MAX_TOEL Then f = f & "- toelichting (kolom AJ) is langer dan " & MAX_TOEL & " tekens" & vbCrLf
    st = UCase$(Trim$(CStr(wv.Cells(r, 8).Value)))
    If Not (st Like "[LA]####") Then f = f & "- status (kolom H) moet L of A plus een jaartal zijn, bijv. L" & Year(Date) & vbCrLf
    If Not Leeg(wv.Cells(r, 25).Value) Then
        If Not IsNumeric(wv.Cells(r, 25).Value) Then
            f = f & "- afschrijftermijn (kolom Y) moet een getal zijn" & vbCrLf
        ElseIf wv.Cells(r, 25).Value < 0 Or wv.Cells(r, 25).Value > 60 Then
            f = f & "- afschrijftermijn (kolom Y) moet tussen 0 en 60 liggen" & vbCrLf
        End If
    End If
    For Each c In Array(26, 31, 32, 33, 34, 35)
        If Not Leeg(wv.Cells(r, CLng(c)).Value) And Not IsNumeric(wv.Cells(r, CLng(c)).Value) Then
            f = f & "- kolom " & Split(wv.Cells(1, CLng(c)).Address(True, False), "$")(0) & " moet een bedrag zijn" & vbCrLf
        End If
    Next c
    If IsNumeric(wv.Cells(r, 13).Value) And IsNumeric(wv.Cells(r, 14).Value) And Not Leeg(wv.Cells(r, 13).Value) And Not Leeg(wv.Cells(r, 14).Value) Then
        If wv.Cells(r, 13).Value > wv.Cells(r, 14).Value Then f = f & "- startjaar (M) ligt na het eindjaar (N)" & vbCrLf
    End If
    FaRegelFouten = f
End Function

' Sluit de open regel: gewijzigde velden naar het Logboek, regel op slot, werkregel weg. Geeft het aantal wijzigingen.
Private Function FaRegelAfsluiten(ByVal r As Long, ByVal actie As String) As Long
    Dim wv As Worksheet, ww As Worksheet, c As Variant, oudW As Variant, nieuwW As Variant, n As Long
    Set wv = Vz()
    Set ww = WerkBlad(False)
    If Not ww Is Nothing Then
        For Each c In FaInvoerKolommen()
            oudW = ww.Cells(2, CLng(c)).Value
            nieuwW = wv.Cells(r, CLng(c)).Value
            If Not GelijkWaarde(oudW, nieuwW) Then
                n = n + 1
                LogRegel actie, wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, CStr(wv.Cells(1, CLng(c)).Value), oudW, nieuwW, "rij " & r & ", direct in de Verzamellijst"
            End If
        Next c
        ww.Cells.Clear
    End If
    wv.Unprotect WACHTWOORD
    wv.Range("A" & r & ":AW" & r).Locked = True
    If n > 0 Then
        wv.Cells(r, 48).Value = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
        wv.Cells(r, 49).Value = Now
    End If
    FaKnoppenVerzamellijst False
    BladBeveiligen wv
    FaRegelAfsluiten = n
End Function

' Vanuit het blad Dashboard (Worksheet_FollowHyperlink): klik op een volgnummer in de lijst van blok 6.
Public Sub FaHyperlink(ByVal h As Object)
    Dim wd As Worksheet, cel As Range, v As Variant
    On Error GoTo Klaar
    Set cel = h.Range
    Set wd = cel.Worksheet
    If wd.Name <> "Dashboard" Or cel.Column <> 1 Or cel.Row <= FA_LIJST Then Exit Sub
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    If Not FaRecht() Then Exit Sub
    v = Application.Match(cel.Value, Vz().Columns(1), 0)
    If IsError(v) Then
        MsgBox "Deze regel is niet meer gevonden. Klik op Lijst vernieuwen.", vbExclamation
        Exit Sub
    End If
    FaRegelOpenen CLng(v), "controle"
Klaar:
End Sub

' Oude knopnaam (blijft werken): opent de in blok 6 gekozen regel in de Verzamellijst.
Public Sub RegelOpenenTerControle()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim r As Long
    If Not FaRecht() Then Exit Sub
    r = ControleGekozenRij()
    If r = 0 Then Exit Sub
    FaRegelOpenen r, "controle"
End Sub

' Incidenteel: een bestaand (al goedgekeurd) krediet aanpassen, direct in de Verzamellijst.
Public Sub GoedgekeurdKredietAanpassen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim r As Long
    If Not FaRecht() Then Exit Sub
    If WerkRij() > 0 Then
        FaRegelOpenen WerkRij(), "aanpassen"      ' toont welke regel nog open staat
        Exit Sub
    End If
    r = ZoekKredietRij("Bestaand krediet aanpassen")
    If r = 0 Then Exit Sub
    FaRegelOpenen r, "aanpassen"
End Sub

' Knop 'Sluiten zonder goedkeuren' (bovenaan de Verzamellijst).
Public Sub FaRegelSluiten()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim r As Long, n As Long, fout As String
    If Not FaRecht() Then Exit Sub
    r = WerkRij()
    If r = 0 Then
        FaKnoppenVerzamellijst False
        MsgBox "Er staat geen regel open.", vbInformation
        Exit Sub
    End If
    fout = FaRegelFouten(Vz(), r)
    If Len(fout) > 0 Then
        If MsgBox("Let op, in de regel klopt nog iets niet:" & vbCrLf & vbCrLf & fout & vbCrLf & "Toch sluiten?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Sluiten") <> vbYes Then Exit Sub
    End If
    n = FaRegelAfsluiten(r, "wijzigen (financi" & ChrW(235) & "le administratie)")
    ControleLijstVernieuwen
    NaarBlok6
    MsgBox "Regel gesloten (" & n & " wijziging(en) in het Logboek). De controlestatus is niet veranderd.", vbInformation
End Sub

Private Sub NaarBlok6()
    With ThisWorkbook.Worksheets("Dashboard")
        .Activate
        Application.Goto .Range("A" & FA_KOP), True
    End With
End Sub

' Knop 'Klaar: goedkeuren' (Verzamellijst): keurt de regel goed die open staat.
Public Sub RegelBevestigen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, st As String, gebr As String, n As Long, fout As String, isOpen As Boolean
    On Error GoTo Fout
    If Not FaRecht() Then Exit Sub
    Set wv = Vz()
    r = WerkRij()
    isOpen = (r > 0)
    If Not isOpen Then r = ControleGekozenRij()
    If r = 0 Then Exit Sub
    fout = FaRegelFouten(wv, r)
    If Len(fout) > 0 Then
        MsgBox "Niet goedgekeurd. Dit moet nog worden aangepast in de Verzamellijst (rij " & r & "):" & vbCrLf & vbCrLf & fout, vbExclamation
        Exit Sub
    End If
    st = CStr(wv.Cells(r, 42).Value)
    If st = "Afgekeurd" Then
        If MsgBox("Deze regel is afgekeurd. Toch goedkeuren?", vbYesNo + vbExclamation + vbDefaultButton2, "Goedkeuren") <> vbYes Then Exit Sub
    End If
    If MsgBox("Volgnummer " & wv.Cells(r, 1).Value & ": " & wv.Cells(r, 4).Value & vbCrLf & vbCrLf & "Goedkeuren als gecontroleerd?", _
              vbYesNo + vbQuestion, "Goedkeuren") <> vbYes Then Exit Sub
    If isOpen Then n = FaRegelAfsluiten(r, "wijzigen (financi" & ChrW(235) & "le administratie)")
    wv.Unprotect WACHTWOORD
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    wv.Cells(r, 42).Value = CtlGecontroleerd()
    wv.Cells(r, 43).Value = gebr
    wv.Cells(r, 44).NumberFormat = "dd-mm-yyyy"
    wv.Cells(r, 44).Value = Date
    wv.Cells(r, 45).ClearContents
    BladBeveiligen wv
    LogRegel "goedkeuren", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "", , , "rij " & r & IIf(n > 0, ", " & n & " veld(en) aangepast", "")
    ControleLijstVernieuwen
    NaarBlok6
    MsgBox "Goedgekeurd" & IIf(n > 0, " (" & n & " veld(en) aangepast, staan in het Logboek)", "") & "." & vbCrLf & _
           "Nog te controleren: " & Application.WorksheetFunction.CountIf(wv.Range("AP:AP"), CtlTeControleren()) & ".", vbInformation
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

Public Sub RegelAfkeuren()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, reden As String, gebr As String, isOpen As Boolean
    On Error GoTo Fout
    If Not FaRecht() Then Exit Sub
    Set wv = Vz()
    r = WerkRij()
    isOpen = (r > 0)
    If Not isOpen Then r = ControleGekozenRij()
    If r = 0 Then Exit Sub
    reden = Trim$(InputBox("Volgnummer " & wv.Cells(r, 1).Value & ": " & wv.Cells(r, 4).Value & vbCrLf & vbCrLf & _
                           "Reden van afkeuring (verplicht):", "Afkeuren"))
    If Len(reden) = 0 Then
        MsgBox "Zonder reden wordt er niet afgekeurd.", vbExclamation
        Exit Sub
    End If
    If isOpen Then FaRegelAfsluiten r, "wijzigen (financi" & ChrW(235) & "le administratie)"
    wv.Unprotect WACHTWOORD
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    wv.Cells(r, 42).Value = "Afgekeurd"
    wv.Cells(r, 43).Value = gebr
    wv.Cells(r, 44).NumberFormat = "dd-mm-yyyy"
    wv.Cells(r, 44).Value = Date
    wv.Cells(r, 45).Value = reden
    BladBeveiligen wv
    LogRegel "afkeuren", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "", , , reden
    ControleLijstVernieuwen
    NaarBlok6
    MsgBox "Afgekeurd. De regel blijft in de lijst staan tot de financieel adviseur hem heeft hersteld.", vbInformation
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

