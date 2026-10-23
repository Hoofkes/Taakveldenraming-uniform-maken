Option Explicit

' Alle macros voor het Dashboard investeringskredieten in een module.
' Versie 19 (stap 12: controle en kengetallen). Controle: zoek met Ctrl+F op 'ControleblokPlaatsen'.

Private Const BL_VZ As String = "Verzamellijst"
Private Const VL As String = "'Verzamellijst'!"
Private Const EERSTE As Long = 9          ' eerste kaartrij op het Dashboard
Private cfMislukt As Boolean
Private Const AANTAL As Long = 30
Private Const BL_MELD As String = "Meldingen uitgaven"
Private mMeld As Collection
Private mVorige As Object          ' bedragen per kredietnummer en jaar uit de vorige upload
Private mHeeftVorige As Boolean
Private mAansl As String
Private Const BL_OVERZICHT As String = "Laatste kredietoverzicht"          ' raad en jaarrekening
Private Const BL_OVERZICHT_NAAM As String = "Laatste overzicht op naam"  ' budgethouder / projectleider
Private Const BL_UITVRAAG As String = "Laatste uitvraag begroting"
Private Const BL_TIJD As String = "Laatste overzicht tijdschrijven"
Private Const OPBRENGST As String = "Opbrengst activa"   ' bron/toelichting voor opbrengsten activa (geen subsidie of bijdrage)
Private Const BL_MUTREG As String = "Mutaties kadernota-begroting"   ' register voor de lastenberekening
Private Const WACHTWOORD As String = "Best"      ' bladbeveiliging (zet ook een wachtwoord op het VBA-project)
Private mStil As Boolean
Private Const CB_EERSTE As Long = 72         ' eerste controleregel
Private Const CB_LAATSTE As Long = 87        ' laatste controleregel (16 controles)
Private mHerbeveiligingGepland As Boolean
Private mActiefVoorMacro As Object        ' blad dat actief was toen de macro begon
Private mProef As Boolean             ' proefversie jaarrekening (niet opslaan)             ' geen meldingen tijdens een reeks overzichten

Private Type TKaart
    f(1 To 30) As Variant
    letter As String
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

Private Function BladBestaat(ByVal naam As String) As Boolean
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

' Schrijft een logregel als het blad Logboek bestaat (komt in stap 6).
Private Sub SchrijfLog(ByVal aantal As Long, ByVal toelichting As String)
    Dim lg As Worksheet, rij As Long
    On Error Resume Next
    Set lg = ThisWorkbook.Worksheets("Logboek")
    On Error GoTo 0
    If lg Is Nothing Then Exit Sub
    rij = lg.Cells(lg.Rows.Count, "A").End(xlUp).Row + 1
    lg.Cells(rij, 1).Value = Now
    lg.Cells(rij, 3).Value = Environ("USERNAME")
    lg.Cells(rij, 5).Value = "normaliseren"
    lg.Cells(rij, 11).Value = aantal & " cellen in kolom I/J naar kleine letters"
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

' Schrijft een logregel als het blad Logboek bestaat (komt in stap 6).
Private Sub SchrijfLogWerkbereik(ByVal toelichting As String)
    Dim lg As Worksheet, rij As Long
    On Error Resume Next
    Set lg = ThisWorkbook.Worksheets("Logboek")
    On Error GoTo 0
    If lg Is Nothing Then Exit Sub
    rij = lg.Cells(lg.Rows.Count, "A").End(xlUp).Row + 1
    lg.Cells(rij, 1).Value = Now
    lg.Cells(rij, 3).Value = Environ("USERNAME")
    lg.Cells(rij, 5).Value = "werkbereik verruimen"
    lg.Cells(rij, 11).Value = toelichting
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
        rij = 18
        For n = 1 To maxN
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
    Dim i As Long, r As Long, n As Long
    Dim f As String, leeg As String, lijstNaam As String
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
    v(4) = Veld("Omschrijving kort (automatisch, max 20; invullen = overschrijven)", "F", "", False, "@")
    v(5) = Veld("Status - lopend of afgesloten (L of A)", "H", "Lst_Statusletter", True, "@")
    v(6) = Veld("Jaar van de status (loopt in / afgesloten in)", "H", "Lst_Jaren", True, "@")
    v(7) = Veld("Investeringsgroep", "I", "Lst_Soort", True, "@")
    v(8) = Veld("Samengesteld, standaard of stelpost", "J", "Lst_SamengesteldStandaardStelpost", True, "")
    v(9) = Veld("Besluit", "L", "Lst_LaatsteBesluit", True, "@")
    v(10) = Veld("Startjaar uitgaven", "M", "Lst_Jaren", True, "")
    v(11) = Veld("Eindjaar uitgaven", "N", "Lst_Jaren", True, "")
    v(12) = Veld("Budgetbeheerder", "O", "Lst_Budgetbeheerder", True, "@")
    v(13) = Veld("Projectleider (tevens tijdschrijver)", "P", "Lst_Projectleider", True, "@")
    v(14) = Veld("Tijdschrijvers", "Q", "Lst_Tijdschrijver", True, "@")
    v(15) = Veld("Exploitatienummer boekhouding", "R", "", True, "")
    v(16) = Veld("Taakveld BBV", "S", "Lst_Taakveld", True, "@")
    v(17) = Veld("Balanscode BBV", "T", "Lst_Balanscode", True, "@")
    v(18) = Veld("Organisatie-onderdeel", "U", "Lst_Organisatieonderdeel", True, "@")
    v(19) = Veld("Budgetbeheerder administratie", "V", "Lst_BudgetbeheerderAdm", True, "@")
    v(20) = Veld("Code btw/bcf administratie", "W", "Lst_BtwBcf", True, "@")
    v(21) = Veld("Activanummer", "X", "", False, "")
    v(22) = Veld("Afschrijftermijn (0 t/m 60 jaar)", "Y", "Lst_Afschrijving", True, "")
    v(23) = Veld("=IF(Kaart_Rijnummer="""",""Budget ""&Par_LopendJaar,""Budget tot en met ""&Par_LopendJaar)", "Z", "", False, "bedrag")
    v(24) = Veld("=""Budget ""&(Par_LopendJaar+1)", "AE", "", False, "bedrag")
    v(25) = Veld("=""Budget ""&(Par_LopendJaar+2)", "AF", "", False, "bedrag")
    v(26) = Veld("=""Budget ""&(Par_LopendJaar+3)", "AG", "", False, "bedrag")
    v(27) = Veld("=""Budget ""&(Par_LopendJaar+4)", "AH", "", False, "bedrag")
    v(28) = Veld("=""Budget ""&(Par_LopendJaar+5)", "AI", "", False, "bedrag")
    v(29) = Veld("Toelichting op het investeringskrediet", "AJ", "", True, "@")
    v(30) = Veld("Naam samengesteld krediet (alleen bij nieuw samengesteld krediet)", "AT", "", False, "@")

    Set actief = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = "Dashboard"
    ' Namen voor de latere macro's
    ThisWorkbook.Names.Add Name:="Kaart_Kolommen", RefersTo:="=Dashboard!$A$9:$A$50"
    ThisWorkbook.Names.Add Name:="Kaart_Invoer", RefersTo:="=Dashboard!$C$9:$C$50"
    ThisWorkbook.Names.Add Name:="Kaart_Rijnummer", RefersTo:="=Dashboard!$C$52"

    ws.Tab.Color = RGB(0, 112, 192)
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 64
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 48
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 22

    ' Kop
    ws.Range("B1").Value = "Dashboard investeringskredieten"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B2").Value = "Gebruiker"
    ws.Range("C2").Formula = "=Par_HuidigeGebruiker"
    ws.Range("D2").Formula = "=Par_HuidigeRol"
    ws.Range("B3").Value = "Lopend jaar"
    ws.Range("C3").Formula = "=Par_LopendJaar"
    ws.Range("C3").HorizontalAlignment = xlLeft
    ws.Range("D3").Formula = "=Parameters!B11"
    ws.Range("D3").NumberFormat = """keuzeopties ververst op ""dd-mm-yyyy hh:mm"
    ws.Range("D3").HorizontalAlignment = xlLeft
    ws.Range("B4").Value = "Controlestatus"
    ws.Range("C4").Value = ""                                 ' statusregel: zie ControleblokPlaatsen (stap 12)
    f = "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)&"" regels te controleren, ""&COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")&"" afgekeurd"""
    ws.Range("B5").Formula = f
    ws.Range("B2:B4").Font.Bold = True
    Set fc = VoegCF(ws.Range("B5"), "=COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")>0")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 150)
    Set fc = VoegCF(ws.Range("B5"), "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)>0")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 200, 120)

    ' Blok 1: invoerkaart
    ws.Range("B7").Value = "1. Invoerkaart krediet"
    ws.Range("B7").Font.Bold = True
    ws.Range("B7").Font.Size = 12
    ws.Range("B8").Value = "Veld"
    ws.Range("C8").Value = "Invoer"
    ws.Range("D8").Value = "Controle"
    ws.Range("B8:D8").Font.Bold = True
    ws.Range("B8:D8").Borders(xlEdgeBottom).LineStyle = xlContinuous

    For i = 1 To 30
        r = KaartRij(i)
        If Left$(CStr(v(i)(0)), 1) = "=" Then
            ws.Cells(r, 2).Formula = CStr(v(i)(0))
        Else
            ws.Cells(r, 2).Value = v(i)(0)
        End If
        ws.Cells(r, 1).Value = v(i)(1)
        ws.Cells(r, 1).Font.Size = 8
        ws.Cells(r, 1).Font.Color = RGB(150, 150, 150)
        With ws.Cells(r, 3)
            If v(i)(4) = "@" Then
                .NumberFormat = "@"
            ElseIf v(i)(4) = "bedrag" Then
                .NumberFormat = "#,##0;-#,##0;-"
            End If
            .Interior.Color = RGB(255, 255, 204)
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
        End With
        ws.Cells(r, 2).VerticalAlignment = xlTop
        ws.Cells(r, 4).VerticalAlignment = xlTop
        ws.Cells(r, 4).WrapText = True
        lijstNaam = CStr(v(i)(2))
        If Len(lijstNaam) > 0 Then
            With ws.Cells(r, 3).Validation
                .Delete
                If lijstNaam = "Lst_Statusletter" Or lijstNaam = "Lst_Afschrijving" Then
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & lijstNaam
                ElseIf lijstNaam = "Lst_SamengesteldStandaardStelpost" Then
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & lijstNaam
                Else
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=" & lijstNaam
                End If
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        End If
        If i = 3 Or i = 4 Then
            With ws.Cells(r, 3).Validation
                .Delete
                .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlLessEqual, _
                     Formula1:=IIf(i = 3, "36", "20")
                .IgnoreBlank = True
                .ErrorTitle = "Te lang"
                .ErrorMessage = "Maximaal " & IIf(i = 3, "36", "20") & " posities."
            End With
        End If
        leeg = ""
        If v(i)(3) Then leeg = "verplicht"
        If i = 2 Or i = 21 Then leeg = "mag later worden aangevuld"
        If i = 14 Then leeg = "verplicht (meer namen? klik op de knop rechts)"
        ws.Cells(r, 4).Formula = ControleFormule(i, r, leeg)
    Next i

    ' Bijdragen van derden: onder elk budgetjaar twee regels. In kolom B kiest u de bron (of typt u die zelf),
    ' in kolom C staat het bedrag. Alleen de jaren met een bijdrage worden ingevuld.
    For bj = 0 To 5
        For bk = 1 To 2
            brr = BijdrageRij(bj, bk)
            ws.Cells(brr, 2).Formula = "=""Bijdrage " & bk & " - ""&(Par_LopendJaar+" & bj & ")"
            With ws.Cells(brr, 2)
                .VerticalAlignment = xlTop
                .HorizontalAlignment = xlLeft
                .Interior.Color = RGB(255, 255, 204)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(191, 191, 191)
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=Lst_BijdrageJaar" & bj
                .Validation.IgnoreBlank = True
                .Validation.InCellDropdown = True
            End With
            With ws.Cells(brr, 3)
                .NumberFormat = "#,##0;-#,##0;-"
                .Interior.Color = RGB(255, 255, 204)
                .Font.Color = RGB(0, 0, 0)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(191, 191, 191)
            End With
            ws.Cells(brr, 4).Formula = "=IF(C" & brr & "="""","""",IF(NOT(ISNUMBER(C" & brr & ")),""moet een getal zijn"",IF(OR(B" & brr & "="""",B" & brr & "=""Bijdrage " & bk & " - ""&(Par_LopendJaar+" & bj & ")),""moet een bron hebben (kies links bij de bijdrage)"",""goed - wordt opgeslagen als bijdrage van derden"")))"
            ws.Cells(brr, 4).WrapText = True
            ws.Cells(brr, 4).VerticalAlignment = xlTop
        Next bk
    Next bj

    ' Rijnummer van een geladen regel (bij wijzigen)
    ws.Cells(52, 2).Value = "Rijnummer geladen regel (alleen bij wijzigen)"
    ws.Cells(52, 1).Value = "rij"
    ws.Cells(52, 1).Font.Size = 8
    ws.Cells(52, 1).Font.Color = RGB(150, 150, 150)
    ws.Cells(52, 3).Interior.Color = RGB(217, 217, 217)
    ws.Cells(52, 3).HorizontalAlignment = xlLeft
    ws.Cells(52, 3).Borders.LineStyle = xlContinuous
    ws.Cells(52, 3).Borders.Color = RGB(191, 191, 191)

    ' De korte omschrijving wordt automatisch afgeleid; de rij blijft bestaan maar is verborgen.
    ws.Rows(EERSTE + 3).Hidden = True

    ' Meldingen in kolom D: kleur
    Set fc = VoegCF(ws.Range("D9:D52"), "=LEFT($D9,12)=""waarschuwing""")
    If Not fc Is Nothing Then fc.Font.Color = RGB(197, 90, 17)
    If Not fc Is Nothing Then fc.Font.Bold = True
    Set fc = VoegCF(ws.Range("D9:D52"), "=$D9=""verplicht""")
    If Not fc Is Nothing Then fc.Font.Color = RGB(150, 150, 150)
    Set fc = VoegCF(ws.Range("D9:D52"), "=OR(ISNUMBER(SEARCH(""te lang"",$D9)),ISNUMBER(SEARCH(""moet"",$D9)),ISNUMBER(SEARCH(""kies"",$D9)),ISNUMBER(SEARCH(""vier"",$D9)),ISNUMBER(SEARCH(""start na"",$D9)),ISNUMBER(SEARCH(""invullen"",$D9)))")
    If Not fc Is Nothing Then fc.Font.Color = RGB(192, 0, 0)
    If Not fc Is Nothing Then fc.Font.Bold = True

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
Private Function ControleFormule(ByVal i As Long, ByVal r As Long, ByVal leeg As String) As String
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
    ControleFormule = "=IF(" & c & "="""",""" & leeg & """," & k & ")"
End Function

' Voegt voorwaardelijke opmaak toe. De formule wordt omgezet naar de taal van uw Excel en
' de eerste cel van het bereik wordt eerst geselecteerd (relatieve verwijzingen).
Private Function VoegCF(ByVal rng As Range, ByVal formule As String) As Object
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
    Dim f(1 To AANTAL) As Variant
    Dim fouten As String, waarsch As String
    Dim i As Long, r As Long, laatste As Long, maxRij As Long, volgnr As Long
    Dim letter As String, jaar As String, j8 As String
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

    For i = 1 To AANTAL
        f(i) = wd.Cells(KaartRij(i), 3).Value
    Next i

    ' --- verplichte velden
    For i = 1 To 22
        If i <> 4 And i <> 2 And i <> 21 And Leeg(f(i)) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), 2).Value & " is verplicht" & vbCrLf
    Next i
    If Leeg(f(29)) Then fouten = fouten & "- " & wd.Cells(KaartRij(29), 2).Value & " is verplicht" & vbCrLf

    ' --- lengte omschrijvingen
    If Not Leeg(f(3)) Then If Len(Trim$(CStr(f(3)))) > 36 Then fouten = fouten & "- omschrijving lang is langer dan 36 posities" & vbCrLf
    If Not Leeg(f(4)) Then If Len(Trim$(CStr(f(4)))) > 20 Then fouten = fouten & "- omschrijving kort is langer dan 20 posities" & vbCrLf

    ' --- status
    letter = UCase$(Trim$(CStr(f(5))))
    jaar = Trim$(CStr(f(6)))
    If Not Leeg(f(5)) Then If letter <> "L" And letter <> "A" Then fouten = fouten & "- status moet L of A zijn" & vbCrLf
    If Not Leeg(f(6)) Then If Len(jaar) <> 4 Or Not IsNumeric(jaar) Then fouten = fouten & "- eerste jaar krediet beschikbaar moet uit vier cijfers bestaan" & vbCrLf

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
        If Not Leeg(f(i)) Then If Not IsNumeric(f(i)) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), 2).Value & " moet een getal zijn" & vbCrLf
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
    wv.Cells(r, 8).Value = letter & jaar
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
    wd.Range("C" & EERSTE & ":C" & KaartRij(30)).ClearContents
    HerstelBijdrageLabels wd
    wd.Range("Kaart_Rijnummer").ClearContents
End Sub

' Maakt alleen de velden per deelkrediet leeg en zet de groep klaar voor het volgende deelkrediet.
Private Sub VolgendeDeelVoorbereiden(ByVal grp As Long)
    Dim wd As Worksheet, i As Variant, cel As Range, zoek As String, bj As Long, bk As Long
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    For Each i In Array(2, 3, 4, 17, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30)
        wd.Cells(KaartRij(CLng(i)), 3).ClearContents
    Next i
    For bk = 1 To 2
        For bj = 0 To 5
            wd.Cells(BijdrageRij(bj, bk), 3).ClearContents
        Next bj
    Next bk
    HerstelBijdrageLabels wd
    KeuzelijstenVerversen True
    zoek = "groep " & grp & " - "
    For Each cel In ThisWorkbook.Names("Lst_SamengesteldStandaardStelpost").RefersToRange.Cells
        If Left$(CStr(cel.Value), Len(zoek)) = zoek Then
            wd.Cells(EERSTE + 7, 3).Value = cel.Value
            Exit For
        End If
    Next cel
    wd.Activate
    wd.Cells(EERSTE + 1, 3).Select
End Sub

Public Sub DashboardKnoppenPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, naam As Variant, i As Long
    Dim namen As Variant, titels As Variant, acties As Variant
    VerzamellijstHernoemen
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    For Each naam In Array("btnToevoegen", "btnWijzigen", "btnOpslaan", "btnKaartWissen", "btnLijsten", "btnVerwerken", "btnTijdschrijvers", "spnAfschrijving", "btnVerwijderen", "btnUitgaven")
        ws.Shapes(CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    ws.Rows(6).RowHeight = 26
    namen = Array("btnWijzigen", "btnKaartWissen", "btnVerwijderen", "btnUitgaven")
    titels = Array("Krediet wijzigen", "Scherm leegmaken", "Krediet verwijderen", "Werkelijke uitgaven bijwerken")
    acties = Array("KredietWijzigen", "KaartWissen", "KredietVerwijderen", "UitgavenBijwerken")
    For i = 0 To 3
        Set b = ws.Buttons.Add(ws.Range("B6").Left + i * 130, ws.Range("B6").Top + 2, 125, 22)
        b.Name = namen(i)
        b.Caption = titels(i)
        b.OnAction = acties(i)
    Next i
    ' knop naast het veld Tijdschrijvers
    Set b = ws.Buttons.Add(ws.Range("F" & (EERSTE + 13)).Left + 2, ws.Range("F" & (EERSTE + 13)).Top + 1, 130, 20)
    b.Name = "btnTijdschrijvers"
    b.Caption = "Tijdschrijvers kiezen"
    b.OnAction = "TijdschrijversKiezen"
    ' knop onder de invoerkaart
    ws.Rows(54).RowHeight = 38
    Set b = ws.Buttons.Add(ws.Range("B54").Left, ws.Range("B54").Top + 4, 230, 30)
    b.Name = "btnVerwerken"
    b.Caption = "Ingevuld? Verwerken!"
    b.OnAction = "Verwerken"
    UitdraaiBlokPlaatsen
    ControleblokPlaatsen
    KoppenStartEindjaar
    TabbladkiezerPlaatsen
    KnoppenTerugPlaatsen
    TabbladenOrdenen
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

Private Function Leeg(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        Leeg = True
    ElseIf IsError(v) Then
        Leeg = True
    Else
        Leeg = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Private Function Nul(ByVal v As Variant) As Double
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

Private Function MeetStap8(ByVal ws As Worksheet, ByVal laatsteRij As Long) As Variant
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
Private Function ControleerKaart(ByRef k As TKaart, ByVal wijzigRij As Long) As Boolean
    Dim wd As Worksheet, wv As Worksheet
    Dim fouten As String, waarsch As String, j8 As String
    Dim i As Long, laatste As Long, aantal As Long
    Dim cel As Range

    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set wv = Vz()
    For i = 1 To AANTAL
        k.f(i) = wd.Cells(KaartRij(i), 3).Value
    Next i

    ' --- verplichte velden
    For i = 1 To 22
        If i <> 4 And i <> 2 And i <> 21 And Leeg(k.f(i)) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), 2).Value & " is verplicht" & vbCrLf
    Next i
    If Leeg(k.f(29)) Then fouten = fouten & "- " & wd.Cells(KaartRij(29), 2).Value & " is verplicht" & vbCrLf

    ' --- lengte omschrijvingen
    If Not Leeg(k.f(3)) Then If Len(Trim$(CStr(k.f(3)))) > 36 Then fouten = fouten & "- omschrijving lang is langer dan 36 posities" & vbCrLf
    If Not Leeg(k.f(4)) Then If Len(Trim$(CStr(k.f(4)))) > 20 Then fouten = fouten & "- omschrijving kort is langer dan 20 posities" & vbCrLf

    ' --- status
    k.letter = UCase$(Trim$(CStr(k.f(5))))
    k.jaar = Trim$(CStr(k.f(6)))
    If Not Leeg(k.f(5)) Then If k.letter <> "L" And k.letter <> "A" Then fouten = fouten & "- status moet L of A zijn" & vbCrLf
    If Not Leeg(k.f(6)) Then If Len(k.jaar) <> 4 Or Not IsNumeric(k.jaar) Then fouten = fouten & "- eerste jaar krediet beschikbaar moet uit vier cijfers bestaan" & vbCrLf

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
        If Not Leeg(k.f(i)) Then If Not IsNumeric(k.f(i)) Then fouten = fouten & "- " & wd.Cells(KaartRij(i), 2).Value & " moet een getal zijn" & vbCrLf
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
        aantal = Application.WorksheetFunction.CountIf(wv.Range("C2:C" & laatste), CDbl(k.f(2)))
        If wijzigRij > 0 Then
            If IsNumeric(wv.Cells(wijzigRij, 3).Value) Then If CDbl(wv.Cells(wijzigRij, 3).Value) = CDbl(k.f(2)) Then aantal = aantal - 1
        End If
        If aantal > 0 Then waarsch = waarsch & "- kredietnummer " & k.f(2) & " bestaat al in de verzamellijst" & vbCrLf
    End If
    If Not Leeg(k.f(21)) Then
        If IsNumeric(k.f(21)) Then
            aantal = Application.WorksheetFunction.CountIf(wv.Range("X2:X" & laatste), CDbl(k.f(21)))
            If wijzigRij > 0 Then
                If IsNumeric(wv.Cells(wijzigRij, 24).Value) Then If CDbl(wv.Cells(wijzigRij, 24).Value) = CDbl(k.f(21)) Then aantal = aantal - 1
            End If
            If aantal > 0 Then waarsch = waarsch & "- activanummer " & k.f(21) & " bestaat al" & vbCrLf
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
    VoegWaarde cols, vals, n, 2, CLng(k.pc)
    If Leeg(k.f(2)) Then ev = Empty Else ev = CDbl(k.f(2))
    VoegWaarde cols, vals, n, 3, ev
    VoegWaarde cols, vals, n, 4, Trim$(CStr(k.f(3)))
    VoegWaarde cols, vals, n, 8, k.letter & k.jaar
    VoegWaarde cols, vals, n, 9, LCase$(Trim$(CStr(k.f(7))))
    VoegWaarde cols, vals, n, 10, k.jWaarde
    VoegWaarde cols, vals, n, 12, GetalOfTekst(k.f(9))
    VoegWaarde cols, vals, n, 13, CDbl(k.f(10))
    VoegWaarde cols, vals, n, 14, CDbl(k.f(11))
    VoegWaarde cols, vals, n, 15, Trim$(CStr(k.f(12)))
    VoegWaarde cols, vals, n, 16, Trim$(CStr(k.f(13)))
    VoegWaarde cols, vals, n, 17, Trim$(CStr(k.f(14)))
    VoegWaarde cols, vals, n, 18, CDbl(k.f(15))
    If Left$(k.tc, 1) = "0" Then tv = k.tc Else tv = CLng(k.tc)
    VoegWaarde cols, vals, n, 19, tv
    VoegWaarde cols, vals, n, 20, k.bc
    VoegWaarde cols, vals, n, 21, Trim$(CStr(k.f(18)))
    VoegWaarde cols, vals, n, 22, Trim$(CStr(k.f(19)))
    VoegWaarde cols, vals, n, 23, Trim$(CStr(k.f(20)))
    If Leeg(k.f(21)) Then ev = Empty Else ev = GetalOfTekst(k.f(21))
    VoegWaarde cols, vals, n, 24, ev
    VoegWaarde cols, vals, n, 25, CDbl(k.f(22))
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

Private Sub ToonOpKaart(ByVal r As Long)
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
    wd.Cells(EERSTE, 3).Select
End Sub

Private Sub ZetKaartVeld(ByVal wd As Worksheet, ByVal i As Long, ByVal v As Variant)
    If IsEmpty(v) Then Exit Sub
    If IsError(v) Then Exit Sub
    wd.Cells(KaartRij(i), 3).Value = v
End Sub

Private Function CodeNaarWeergave(ByVal tabel As String, ByVal code As Variant) As Variant
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
    Dim wijzig(1 To 45) As Boolean, oudw(1 To 45) As Variant
    Dim voor As Variant, na As Variant, verschil As Boolean
    Dim dlt(0 To 9) As Double, som As Double, waarsch As String
    Dim keuze As VbMsgBoxResult, bevestig As Boolean, gebr As String
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
        MsgBox "Er is niets gewijzigd.", vbInformation
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
            oudw(i) = wv.Cells(r, cols(i)).Value
            ZetCel wv, r, cols(i), vals(i)
            If cols(i) = 26 Then dlt(1) = dlt(1) + Nul(vals(i)) - Nul(oudw(i))
            If cols(i) >= 31 And cols(i) <= 35 Then dlt(5 + cols(i) - 31) = dlt(5 + cols(i) - 31) + Nul(vals(i)) - Nul(oudw(i))
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
            If wijzig(i) Then ZetCel wv, r, cols(i), oudw(i)
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
        Else
            wv.Cells(r, 42).Value = CtlTeControleren()
            wv.Cells(r, 43).ClearContents
            wv.Cells(r, 44).ClearContents
        End If
        wv.Cells(r, 48).Value = gebr
        wv.Cells(r, 49).Value = Now
    End If

    For i = 1 To n
        If wijzig(i) Then LogRegel "wijzigen", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, CStr(wv.Cells(1, cols(i)).Value), oudw(i), vals(i), "rij " & r
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
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, i As Long, vorig As Object
    Dim titels As Variant, acties As Variant, namen As Variant, naam As Variant
    Set vorig = ActiveSheet
    If Not BladBestaat("Controle") Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Dashboard"))
        ws.Name = "Controle"
        ws.Tab.Color = RGB(237, 125, 49)
    Else
        Set ws = ThisWorkbook.Worksheets("Controle")
    End If
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "Controle door de financi" & ChrW(235) & "le administratie"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Formula = "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)&"" regels te controleren, ""&COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")&"" afgekeurd"""
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "Selecteer " & ChrW(233) & ChrW(233) & "n regel in de lijst en klik op een knop. Bevestigen en afkeuren gaat per regel."
    ws.Range("A4").Font.Italic = True
    ws.Range("A5").Value = "Volgnr"
    ws.Range("B5").Value = "Kredietnummer"
    ws.Range("C5").Value = "Omschrijving"
    ws.Range("D5").Formula = "=""Budget tot en met ""&Par_LopendJaar"
    ws.Range("E5").Value = "Restantkrediet"
    ws.Range("F5").Value = "Status"
    ws.Range("G5").Value = "Laatst gewijzigd door"
    ws.Range("H5").Value = "Laatst gewijzigd op"
    ws.Range("I5").Value = "Opmerking controle"
    With ws.Range("A5:I5")
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 44
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 40
    ws.Columns("G").ColumnWidth = 22
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I").ColumnWidth = 50
    ws.Range("D6:E2000").NumberFormat = "#,##0;-#,##0;-"
    ws.Range("H6:H2000").NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Range("A6:A2000").HorizontalAlignment = xlLeft
    ws.Range("B6:B2000").HorizontalAlignment = xlLeft

    On Error Resume Next
    For Each naam In Array("btnVernieuwen", "btnOpenen", "btnBevestigen", "btnAfkeuren")
        ws.Shapes(CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    ws.Rows(2).RowHeight = 26
    namen = Array("btnVernieuwen", "btnOpenen", "btnBevestigen", "btnAfkeuren")
    titels = Array("Lijst vernieuwen", "Regel openen ter controle", "Bevestigen: gecontroleerd", "Afkeuren met opmerking")
    acties = Array("ControleLijstVernieuwen", "RegelOpenenTerControle", "RegelBevestigen", "RegelAfkeuren")
    For i = 0 To 3
        Set b = ws.Buttons.Add(ws.Range("A2").Left + i * 160, ws.Range("A2").Top + 2, 155, 22)
        b.Name = namen(i)
        b.Caption = titels(i)
        b.OnAction = acties(i)
    Next i

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A6").Select
    ControleLijstVernieuwen
    If Not vorig Is Nothing Then vorig.Activate
End Sub

' Vult de lijst op het tabblad Controle met alle regels die te controleren of afgekeurd zijn.
Public Sub ControleLijstVernieuwen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wc As Worksheet, wv As Worksheet, r As Long, laatste As Long, rij As Long, st As String
    If Not BladBestaat("Controle") Then Exit Sub
    Set wc = ThisWorkbook.Worksheets("Controle")
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    wc.Range("A6:I2000").ClearContents
    wc.Range("A6:I2000").Font.Color = RGB(0, 0, 0)
    rij = 6
    For r = 2 To laatste
        st = CStr(wv.Cells(r, 42).Value)
        If st = CtlTeControleren() Or st = "Afgekeurd" Then
            wc.Cells(rij, 1).Value = wv.Cells(r, 1).Value
            wc.Cells(rij, 2).Value = wv.Cells(r, 3).Value
            wc.Cells(rij, 3).Value = wv.Cells(r, 4).Value
            wc.Cells(rij, 4).Value = wv.Cells(r, 26).Value
            wc.Cells(rij, 5).Value = wv.Cells(r, 30).Value
            wc.Cells(rij, 6).Value = st
            wc.Cells(rij, 7).Value = wv.Cells(r, 48).Value
            wc.Cells(rij, 8).Value = wv.Cells(r, 49).Value
            wc.Cells(rij, 9).Value = wv.Cells(r, 45).Value
            If st = "Afgekeurd" Then wc.Range(wc.Cells(rij, 1), wc.Cells(rij, 9)).Font.Color = RGB(192, 0, 0)
            rij = rij + 1
            If rij > 2000 Then Exit For
        End If
    Next r
End Sub

' Zoekt de verzamellijst-rij bij de regel die op het tabblad Controle is geselecteerd (0 = niets gekozen).
Private Function ControleGekozenRij() As Long
    Dim wc As Worksheet, wv As Worksheet, x As Long, v As Variant
    If ActiveSheet.Name <> "Controle" Then
        MsgBox "Ga naar het tabblad Controle en selecteer een regel in de lijst.", vbExclamation
        Exit Function
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecteer een regel in de lijst.", vbExclamation
        Exit Function
    End If
    If Selection.Rows.Count <> 1 Then
        MsgBox "Selecteer " & ChrW(233) & ChrW(233) & "n regel. Bevestigen en afkeuren gaat per regel.", vbExclamation
        Exit Function
    End If
    Set wc = ThisWorkbook.Worksheets("Controle")
    Set wv = Vz()
    x = Selection.Row
    If x < 6 Then
        MsgBox "Selecteer een regel in de lijst.", vbExclamation
        Exit Function
    End If
    If IsEmpty(wc.Cells(x, 1).Value) Then
        MsgBox "Selecteer een regel in de lijst.", vbExclamation
        Exit Function
    End If
    v = Application.Match(wc.Cells(x, 1).Value, wv.Columns(1), 0)
    If IsError(v) Then
        MsgBox "Deze regel is niet meer gevonden. Klik op Lijst vernieuwen.", vbExclamation
        Exit Function
    End If
    ControleGekozenRij = CLng(v)
End Function

Public Sub RegelOpenenTerControle()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim r As Long
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "U mag kredieten niet openen om te wijzigen.", vbExclamation
        Exit Sub
    End If
    r = ControleGekozenRij()
    If r = 0 Then Exit Sub
    ToonOpKaart r
    MsgBox "De regel staat op het Dashboard. Pas zo nodig aan met Wijziging opslaan." & vbCrLf & _
           "Ga daarna terug naar het tabblad Controle om te bevestigen of af te keuren.", vbInformation
End Sub

Public Sub RegelBevestigen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, st As String, gebr As String
    On Error GoTo Fout
    If Not HeeftRecht("Bevestigen") Then
        MsgBox "Alleen de financi" & ChrW(235) & "le administratie mag regels bevestigen.", vbExclamation
        Exit Sub
    End If
    r = ControleGekozenRij()
    If r = 0 Then Exit Sub
    Set wv = Vz()
    st = CStr(wv.Cells(r, 42).Value)
    If st = CtlGecontroleerd() Then
        MsgBox "Deze regel is al gecontroleerd.", vbInformation
        Exit Sub
    End If
    If st = "Afgekeurd" Then
        If MsgBox("Deze regel is afgekeurd. Toch bevestigen?", vbYesNo + vbExclamation + vbDefaultButton2, "Bevestigen") <> vbYes Then Exit Sub
    End If
    If MsgBox("Volgnummer " & wv.Cells(r, 1).Value & ": " & wv.Cells(r, 4).Value & vbCrLf & vbCrLf & "Bevestigen als gecontroleerd?", _
              vbYesNo + vbQuestion, "Bevestigen") <> vbYes Then Exit Sub
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    wv.Cells(r, 42).Value = CtlGecontroleerd()
    wv.Cells(r, 43).Value = gebr
    wv.Cells(r, 44).NumberFormat = "dd-mm-yyyy"
    wv.Cells(r, 44).Value = Date
    wv.Cells(r, 45).ClearContents
    LogRegel "bevestigen", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "", , , "rij " & r
    ControleLijstVernieuwen
    MsgBox "Bevestigd.", vbInformation
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

Public Sub RegelAfkeuren()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, reden As String, gebr As String
    On Error GoTo Fout
    If Not HeeftRecht("Bevestigen") Then
        MsgBox "Alleen de financi" & ChrW(235) & "le administratie mag regels afkeuren.", vbExclamation
        Exit Sub
    End If
    r = ControleGekozenRij()
    If r = 0 Then Exit Sub
    Set wv = Vz()
    If CStr(wv.Cells(r, 42).Value) = CtlGecontroleerd() Then
        If MsgBox("Deze regel is al gecontroleerd. Toch afkeuren?", vbYesNo + vbExclamation + vbDefaultButton2, "Afkeuren") <> vbYes Then Exit Sub
    End If
    reden = Trim$(InputBox("Volgnummer " & wv.Cells(r, 1).Value & ": " & wv.Cells(r, 4).Value & vbCrLf & vbCrLf & _
                           "Reden van afkeuring (verplicht):", "Afkeuren"))
    If Len(reden) = 0 Then
        MsgBox "Zonder reden wordt er niet afgekeurd.", vbExclamation
        Exit Sub
    End If
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    wv.Cells(r, 42).Value = "Afgekeurd"
    wv.Cells(r, 43).Value = gebr
    wv.Cells(r, 44).NumberFormat = "dd-mm-yyyy"
    wv.Cells(r, 44).Value = Date
    wv.Cells(r, 45).Value = reden
    LogRegel "afkeuren", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "", , , reden
    ControleLijstVernieuwen
    MsgBox "Afgekeurd. De regel blijft op de lijst staan tot hij is hersteld en opnieuw bevestigd.", vbInformation
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' ======================================================================
' Meerdere tijdschrijvers kiezen
' ======================================================================
' Toont de lijst met namen (20 per pagina). U typt de nummers van de gewenste tijdschrijvers, gescheiden door komma's.
' Namen die nu in het veld staan, zijn voorgevuld. Namen die niet in de lijst staan, blijven behouden.
Public Sub TijdschrijversKiezen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const PER As Long = 20
    Dim wd As Worksheet, cel As Range, lst As Range
    Dim namen() As String, sel() As Boolean
    Dim n As Long, i As Long, p As Long, pagina As Long, paginas As Long, van As Long, tot As Long
    Dim delen() As String, deel As Variant, extra As String, resultaat As String
    Dim prompt As String, standaard As String, antw As String, tok() As String, t As Variant
    Dim fout As Boolean, nr As Long

    On Error GoTo Fout
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set cel = wd.Cells(EERSTE + 13, 3)
    Set lst = ThisWorkbook.Names("Lst_Tijdschrijver").RefersToRange
    For i = 1 To lst.Cells.Count
        If Len(Trim$(CStr(lst.Cells(i, 1).Value))) > 0 Then
            n = n + 1
            ReDim Preserve namen(1 To n)
            ReDim Preserve sel(1 To n)
            namen(n) = Trim$(CStr(lst.Cells(i, 1).Value))
        End If
    Next i
    ' huidige namen voorselecteren
    If Len(Trim$(CStr(cel.Value))) > 0 Then
        delen = Split(CStr(cel.Value), ",")
        For Each deel In delen
            If Len(Trim$(CStr(deel))) > 0 Then
                p = 0
                For i = 1 To n
                    If StrComp(namen(i), Trim$(CStr(deel)), vbTextCompare) = 0 Then p = i: Exit For
                Next i
                If p > 0 Then
                    sel(p) = True
                Else
                    If Len(extra) > 0 Then extra = extra & ", "
                    extra = extra & Trim$(CStr(deel))
                End If
            End If
        Next deel
    End If

    paginas = 0
    If n > 0 Then paginas = (n - 1) \ PER + 1
    pagina = 1
    Do While pagina <= paginas
        van = (pagina - 1) * PER + 1
        tot = pagina * PER
        If tot > n Then tot = n
        prompt = "Pagina " & pagina & " van " & paginas & ". Typ de nummers van de tijdschrijvers, gescheiden door komma's." & vbCrLf & vbCrLf
        standaard = ""
        For i = van To tot
            prompt = prompt & IIf(sel(i), "* ", "   ") & i & " = " & namen(i) & vbCrLf
            If sel(i) Then standaard = standaard & IIf(Len(standaard) > 0, ", ", "") & i
        Next i
        antw = InputBox(prompt, "Tijdschrijvers kiezen", standaard)
        If StrPtr(antw) = 0 Then Exit Sub
        tok = Split(Replace(Replace(antw, ";", ","), " ", ""), ",")
        fout = False
        For Each t In tok
            If Len(CStr(t)) > 0 Then
                If Not IsNumeric(t) Then
                    fout = True
                Else
                    nr = CLng(t)
                    If nr < van Or nr > tot Then fout = True
                End If
            End If
        Next t
        If fout Then
            MsgBox "Typ alleen nummers van deze pagina (" & van & " t/m " & tot & "), gescheiden door komma's.", vbExclamation
        Else
            For i = van To tot
                sel(i) = False
            Next i
            For Each t In tok
                If Len(CStr(t)) > 0 Then sel(CLng(t)) = True
            Next t
            pagina = pagina + 1
        End If
    Loop

    ' eigen namen: staat iemand niet in de lijst, dan typt u de naam hier
    antw = InputBox("Staat er een tijdschrijver niet in de lijst? Typ de naam hier, meerdere namen gescheiden door komma's." & vbCrLf & _
                    "Laat leeg als dat niet nodig is.", "Tijdschrijvers kiezen", extra)
    If StrPtr(antw) = 0 Then Exit Sub
    extra = ""
    For Each deel In Split(antw, ",")
        If Len(Trim$(CStr(deel))) > 0 Then
            If Len(extra) > 0 Then extra = extra & ", "
            extra = extra & Trim$(CStr(deel))
        End If
    Next deel

    For i = 1 To n
        If sel(i) Then
            If Len(resultaat) > 0 Then resultaat = resultaat & ", "
            resultaat = resultaat & namen(i)
        End If
    Next i
    If Len(extra) > 0 Then
        If Len(resultaat) > 0 Then resultaat = resultaat & ", "
        resultaat = resultaat & extra
    End If
    cel.NumberFormat = "@"
    cel.Value = resultaat
    wd.Activate
    cel.Select
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' ======================================================================
' Bijdragen van derden
' ======================================================================
' Leest het bijdragenraster op het Dashboard. Geeft het aantal in te voeren bijdragen terug.
' Bedragen worden altijd negatief opgeslagen.
Private Function LeesBijdragen(ByRef bron() As String, ByRef bed() As Variant, ByRef fouten As String) As Long
    Dim wd As Worksheet, j As Long, k As Long, n As Long, g As Long, p As Long, jj As Long
    Dim v As Variant, tekst As String, standaard As String, jaar As String
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    For j = 0 To 5
        jaar = CStr(CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value) + j)
        For k = 1 To 2
            v = wd.Cells(BijdrageRij(j, k), 3).Value
            If Not Leeg(v) Then
                If Not IsNumeric(v) Then
                    fouten = fouten & "- bijdrage " & k & " in " & jaar & ": een bedrag moet een getal zijn" & vbCrLf
                Else
                    tekst = Trim$(CStr(wd.Cells(BijdrageRij(j, k), 2).Value))
                    standaard = "Bijdrage " & k & " - " & jaar
                    If Len(tekst) = 0 Or StrComp(tekst, standaard, vbTextCompare) = 0 Then
                        fouten = fouten & "- bijdrage " & k & " in " & jaar & ": kies de bron van de bijdrage" & vbCrLf
                    Else
                        If tekst Like "* - ####" Then tekst = Trim$(Left$(tekst, Len(tekst) - 7))
                        g = 0
                        For p = 1 To n
                            If StrComp(bron(p), tekst, vbTextCompare) = 0 Then g = p: Exit For
                        Next p
                        If g = 0 Then
                            If n >= 12 Then
                                fouten = fouten & "- er zijn te veel verschillende bijdragen (maximaal 12)" & vbCrLf
                            Else
                                n = n + 1
                                g = n
                                bron(g) = tekst
                            End If
                        End If
                        If g > 0 Then
                            If Not IsEmpty(bed(g, j)) Then
                                fouten = fouten & "- de bron " & tekst & " staat twee keer in " & jaar & vbCrLf
                            Else
                                bed(g, j) = -Abs(CDbl(v))
                            End If
                        End If
                    End If
                End If
            End If
        Next k
    Next j
    LeesBijdragen = n
End Function

' Schrijft een bijdrageregel: kopie van het krediet (rij rBron) met negatieve bedragen.
Private Sub SchrijfBijdrageRegel(ByVal wv As Worksheet, ByVal rBron As Long, ByVal rNieuw As Long, ByVal volgnr As Long, _
                                 ByVal bron As String, ByRef bed() As Variant, ByVal k As Long)
    Dim c As Variant, j As Long, gebr As String
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    wv.Cells(rNieuw, 1).Value = volgnr
    For Each c In Array(2, 3, 4, 6, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 25)
        wv.Cells(rNieuw, CLng(c)).NumberFormat = wv.Cells(rBron, CLng(c)).NumberFormat
        If Not IsEmpty(wv.Cells(rBron, CLng(c)).Value) Then wv.Cells(rNieuw, CLng(c)).Value = wv.Cells(rBron, CLng(c)).Value
    Next c
    If Not IsEmpty(bed(k, 0)) Then wv.Cells(rNieuw, 26).Value = bed(k, 0)
    For j = 1 To 5
        If Not IsEmpty(bed(k, j)) Then wv.Cells(rNieuw, 30 + j).Value = bed(k, j)
    Next j
    wv.Cells(rNieuw, 36).Value = bron
    wv.Cells(rNieuw, 42).Value = CtlTeControleren()
    If Len(CStr(wv.Cells(rBron, 46).Value)) > 0 Then wv.Cells(rNieuw, 46).Value = wv.Cells(rBron, 46).Value
    wv.Cells(rNieuw, 47).NumberFormat = "dd-mm-yyyy"
    wv.Cells(rNieuw, 47).Value = Date
    wv.Cells(rNieuw, 48).Value = gebr
    wv.Cells(rNieuw, 49).Value = Now
End Sub

' Controleert of de formulekolommen in de rij staan; zo niet, kopieer ze van de rij erboven.
Private Sub ControleerFormules(ByVal wv As Worksheet, ByVal r As Long)
    Dim kolom As Variant
    For Each kolom In Array("E", "G", "K", "AD", "AL", "AM", "AN")
        If Not wv.Range(kolom & r).HasFormula Then wv.Range(kolom & (r - 1)).Copy wv.Range(kolom & r)
    Next kolom
    Application.CutCopyMode = False
End Sub

' ======================================================================
' Krediet verwijderen
' ======================================================================
' Een krediet wordt helemaal weggehaald (alleen bij echte fouten, zonder werkelijke bedragen); het Logboek houdt de gegevens vast.
Public Sub KredietVerwijderen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, wv As Worksheet, r As Long, i As Long, laatste As Long, volgnr As Variant
    Dim rijen As New Collection, uitg As Double, reden As String, grpTxt As String, rv As Variant

    On Error GoTo Fout
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen kredieten verwijderen.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set wv = Vz()
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) = 0 Then
        MsgBox "Laad eerst het krediet met Krediet wijzigen. Klik daarna op Krediet verwijderen.", vbExclamation
        Exit Sub
    End If
    r = CLng(wd.Range("Kaart_Rijnummer").Value)
    If r < 2 Or IsEmpty(wv.Cells(r, 1).Value) Then
        MsgBox "De geladen rij bestaat niet meer. Gebruik Scherm leegmaken en laad het krediet opnieuw.", vbCritical
        Exit Sub
    End If
    ' het krediet = alle regels met hetzelfde volgnummer (uitgavenregel plus bijdrageregels)
    volgnr = wv.Cells(r, 1).Value
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For i = 2 To laatste
        If wv.Cells(i, 1).Value = volgnr Then
            rijen.Add i
            uitg = uitg + Abs(Nul(wv.Cells(i, 27).Value)) + Abs(Nul(wv.Cells(i, 28).Value)) + Abs(Nul(wv.Cells(i, 29).Value))
        End If
    Next i
    If uitg <> 0 Then
        MsgBox "Er zijn werkelijke uitgaven of inkomsten geboekt. Een krediet met werkelijke bedragen mag nooit weg." & vbCrLf & _
               "Sluit het af: wijzig de status in A.", vbExclamation
        Exit Sub
    End If
    If IsGetal(wv.Cells(r, 10).Value) Then
        grpTxt = vbCrLf & "Let op: dit is een deelkrediet van groep " & CLng(wv.Cells(r, 10).Value) & ". De verdeelsleutel van de andere deelkredieten verandert."
    End If
    If MsgBox("Volgnummer " & volgnr & ": " & wv.Cells(r, 4).Value & vbCrLf & rijen.Count & " regel(s), inclusief bijdrageregels." & grpTxt & vbCrLf & vbCrLf & _
              "Helemaal verwijderen? Alleen het Logboek houdt de gegevens vast. Dit is alleen bedoeld voor echte fouten.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Krediet verwijderen") <> vbYes Then Exit Sub
    reden = Trim$(InputBox("Reden van verwijderen (verplicht):", "Krediet verwijderen"))
    If Len(reden) = 0 Then
        MsgBox "Zonder reden wordt er niets verwijderd.", vbExclamation
        Exit Sub
    End If
    If Not RegelsWeghalen(wv, rijen, reden) Then Exit Sub
    KeuzelijstenVerversen True
    KaartWissen
    If BladBestaat("Controle") Then ControleLijstVernieuwen
    MsgBox "Krediet verwijderd (" & rijen.Count & " regel(s)). De gegevens staan in het Logboek.", vbInformation
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout: " & Err.Description, vbCritical
End Sub

' Eenmalig: haalt de regels met status V (oude werkwijze verwijderen) helemaal weg, als er geen werkelijke bedragen op staan.
Public Sub VervallenRegelsOpruimen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, i As Long, laatste As Long, rijen As New Collection, overslaan As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For i = 2 To laatste
        If Not Leeg(wv.Cells(i, 1).Value) And UCase$(Left$(CStr(wv.Cells(i, 8).Value), 1)) = "V" Then
            If Nul(wv.Cells(i, 27).Value) <> 0 Or Nul(wv.Cells(i, 28).Value) <> 0 Or Nul(wv.Cells(i, 29).Value) <> 0 Then
                overslaan = overslaan & vbCrLf & "rij " & i & " (" & wv.Cells(i, 3).Value & ")"
            Else
                rijen.Add i
            End If
        End If
    Next i
    If rijen.Count = 0 Then
        MsgBox "Geen regels met status V zonder werkelijke bedragen gevonden." & IIf(Len(overslaan) > 0, vbCrLf & "Met bedragen (niet weggehaald):" & overslaan, ""), vbInformation
        Exit Sub
    End If
    If MsgBox(rijen.Count & " regel(s) met status V helemaal weghalen? De gegevens gaan naar het Logboek.", vbYesNo + vbQuestion + vbDefaultButton2, _
              "Vervallen regels opruimen") <> vbYes Then Exit Sub
    If RegelsWeghalen(wv, rijen, "opruimen vervallen regel (status V)") Then
        KeuzelijstenVerversen True
        If BladBestaat("Controle") Then ControleLijstVernieuwen
        MsgBox rijen.Count & " regel(s) weggehaald." & IIf(Len(overslaan) > 0, vbCrLf & "Niet weggehaald (werkelijke bedragen):" & overslaan, ""), vbInformation
    End If
End Sub

' Haalt regels helemaal weg zonder Excel-rijen te verwijderen: de regels eronder schuiven een plek omhoog
' (alleen invoerkolommen, formulekolommen blijven staan) en de onderste regels worden leeggemaakt.
' Alle gegevens van de weggehaalde regels gaan eerst naar het Logboek. Nulmeting vooraf en achteraf; bij verschil terugzetten.
Private Function RegelsWeghalen(ByVal wv As Worksheet, ByVal rijen As Collection, ByVal reden As String) As Boolean
    Dim laatste As Long, eerste As Long, n As Long, i As Long, c As Long, k As Long, doel As Long
    Dim kols As Variant, weg() As Boolean, oudW() As Variant, oudF() As Variant, rv As Variant
    Dim voor As Variant, na As Variant, verwacht(0 To 9) As Double, tekst As String, kolMeet As Variant, tmp() As String

    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    kols = Array(1, 2, 3, 4, 6, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, _
                 31, 32, 33, 34, 35, 36, 42, 43, 44, 45, 46, 47, 48, 49)
    kolMeet = Array(0, 26, 27, 28, 29, 31, 32, 33, 34, 35)
    ReDim weg(2 To laatste)
    eerste = laatste
    For Each rv In rijen
        weg(rv) = True
        If rv < eerste Then eerste = rv
        n = n + 1
    Next rv

    ' logboek: alle gevulde invoervelden per regel
    For Each rv In rijen
        tekst = ""
        For k = 0 To UBound(kols)
            If Not Leeg(wv.Cells(rv, kols(k)).Value) Then
                tekst = tekst & Split(wv.Cells(1, kols(k)).Address(True, False), "$")(0) & "=" & CStr(wv.Cells(rv, kols(k)).Value) & "; "
            End If
        Next k
        LogRegel "verwijderen (helemaal weg)", wv.Cells(rv, 1).Value, wv.Cells(rv, 3).Value, "rij " & rv, , , reden & " | " & tekst
    Next rv

    voor = MeetStap8(wv, laatste)
    For k = 0 To 9
        verwacht(k) = voor(k)
    Next k
    For Each rv In rijen
        verwacht(0) = verwacht(0) - 1
        For k = 1 To 9
            verwacht(k) = verwacht(k) - Nul(wv.Cells(rv, kolMeet(k)).Value)
        Next k
    Next rv

    Application.ScreenUpdating = False
    ' reservekopie van waarden en notatie (eerste t/m laatste)
    ReDim oudW(0 To UBound(kols)): ReDim oudF(0 To UBound(kols))
    For k = 0 To UBound(kols)
        oudW(k) = wv.Range(wv.Cells(eerste, kols(k)), wv.Cells(laatste, kols(k))).Value
        ReDim tmp(eerste To laatste)
        For i = eerste To laatste
            tmp(i) = wv.Cells(i, kols(k)).NumberFormat
        Next i
        oudF(k) = tmp
    Next k

    ' schuiven: elke blijvende regel naar de eerstvolgende vrije plek
    doel = eerste
    For i = eerste To laatste
        If Not weg(i) Then
            If doel <> i Then
                For k = 0 To UBound(kols)
                    wv.Cells(doel, kols(k)).NumberFormat = wv.Cells(i, kols(k)).NumberFormat
                    wv.Cells(doel, kols(k)).Value = wv.Cells(i, kols(k)).Value
                Next k
            End If
            doel = doel + 1
        End If
    Next i
    For i = doel To laatste
        For k = 0 To UBound(kols)
            wv.Cells(i, kols(k)).ClearContents
        Next k
    Next i
    wv.Range(wv.Cells(eerste, 28), wv.Cells(laatste, 29)).ClearComments   ' notities uit stap 11 kloppen niet meer na schuiven
    Application.Calculate

    na = MeetStap8(wv, laatste)
    For k = 0 To 9
        If Round(na(k) - verwacht(k), 2) <> 0 Then
            For c = 0 To UBound(kols)
                wv.Range(wv.Cells(eerste, kols(c)), wv.Cells(laatste, kols(c))).Value = oudW(c)
                For i = eerste To laatste
                    wv.Cells(i, kols(c)).NumberFormat = oudF(c)(i)
                Next i
            Next c
            Application.Calculate
            Application.ScreenUpdating = True
            LogRegel "verwijderen teruggedraaid", , , "", , , "nulmeting wijkt af (meting " & k & ")"
            MsgBox "Verschil bij de nameting. Alles is teruggezet. Er is niets verwijderd.", vbCritical
            Exit Function
        End If
    Next k
    Application.ScreenUpdating = True
    RegelsWeghalen = True
End Function

' ======================================================================
' Verzamellijst opschonen
' ======================================================================
' Haalt het filter en de knop Sort van het tabblad Verzamellijst alle kredieten. Werken gaat via het Dashboard.
Public Sub VerzamellijstOpschonen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet
    Set ws = Vz()
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    ws.Buttons.Delete
    On Error GoTo 0
End Sub

' Rij op het Dashboard van veld i (1 t/m 30). Onder elk budgetjaar staan twee bijdrageregels.
Private Function KaartRij(ByVal i As Long) As Long
    If i <= 22 Then
        KaartRij = EERSTE + i - 1
    ElseIf i <= 28 Then
        KaartRij = 31 + 3 * (i - 23)
    ElseIf i = 29 Then
        KaartRij = 49
    Else
        KaartRij = 50
    End If
End Function

' Rij van bijdrage k (1 of 2) in jaar j (0 = lopend jaar t/m 5 = lopend jaar + 5).
Private Function BijdrageRij(ByVal j As Long, ByVal k As Long) As Long
    BijdrageRij = 31 + 3 * j + k
End Function


' Bronnen van bijdragen van derden (tabel op blad Parameters, zelf aan te vullen).
Private Sub BijdrageBronTabelAanmaken()
    Dim wsP As Worksheet, lo As ListObject, bestaat As Boolean, t As Variant, i As Long, j As Long
    Dim kol As ListColumn, heeft As Boolean
    Set wsP = ThisWorkbook.Worksheets("Parameters")
    For Each lo In wsP.ListObjects
        If lo.Name = "Tbl_BijdrageBron" Then bestaat = True
    Next lo
    If Not bestaat Then
        t = Array("Bijdrage rijk", "Bijdrage provincie", "Bijdrage SPUK", "Bijdrage DUMAVA", "Bijdrage gemeente", _
                  "Bijdrage gemeenschappelijke regeling", "Bijdrage derden")
        wsP.Range("AL12").Value = "Bronnen van bijdragen van derden. Voeg onderaan nieuwe bronnen toe (kolom Bron). De kolommen Jaar worden automatisch gevuld."
        wsP.Range("AL12").Font.Bold = True
        wsP.Range("AL14").Value = "Bron"
        For i = 0 To UBound(t)
            wsP.Cells(15 + i, 38).Value = t(i)
        Next i
        Set lo = wsP.ListObjects.Add(xlSrcRange, wsP.Range("AL14:AL" & (15 + UBound(t))), , xlYes)
        lo.Name = "Tbl_BijdrageBron"
        wsP.Columns("AL").ColumnWidth = 40
    End If
    Set lo = wsP.ListObjects("Tbl_BijdrageBron")
    ' vaste bron voor opbrengsten activa (verkoop, directe mva-boekwaardemutatie): wel zichtbaar in de raadsoverzichten
    heeft = False
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.ListColumns("Bron").DataBodyRange.Cells(i, 1).Value), OPBRENGST, vbTextCompare) = 0 Then heeft = True
    Next i
    If Not heeft Then lo.ListRows.Add.Range.Cells(1, lo.ListColumns("Bron").Index).Value = OPBRENGST
    For j = 0 To 5
        heeft = False
        For Each kol In lo.ListColumns
            If kol.Name = "Jaar " & j Then heeft = True
        Next kol
        If Not heeft Then
            Set kol = lo.ListColumns.Add
            kol.Name = "Jaar " & j
        End If
        lo.ListColumns("Jaar " & j).DataBodyRange.Formula = "=[@Bron]&"" - ""&(Par_LopendJaar+" & j & ")"
        ThisWorkbook.Names.Add Name:="Lst_BijdrageJaar" & j, RefersTo:="=Tbl_BijdrageBron[Jaar " & j & "]"
        wsP.Columns(lo.ListColumns("Jaar " & j).Range.Column).ColumnWidth = 44
    Next j
    ThisWorkbook.Names.Add Name:="Lst_BijdrageBron", RefersTo:="=Tbl_BijdrageBron[Bron]"
End Sub

' Zet de standaardteksten van de bijdrageregels terug (na het kiezen van een bron staat daar de gekozen tekst).
Private Sub HerstelBijdrageLabels(ByVal wd As Worksheet)
    Dim j As Long, k As Long
    For j = 0 To 5
        For k = 1 To 2
            wd.Cells(BijdrageRij(j, k), 2).Formula = "=""Bijdrage " & k & " - ""&(Par_LopendJaar+" & j & ")"
        Next k
    Next j
End Sub

' Geeft de zelf aan te passen cellen op blad Parameters een gele kleur en de automatische kolommen een grijze kleur.
Private Sub KleurParameters()
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets("Parameters")
    KleurKolommen wsP, "Tbl_Rollen", Array("Naam", "Windows-naam", "Rol", "Toevoegen", "Wijzigen", "Bevestigen"), Empty
    KleurKolommen wsP, "Tbl_Programma", Array("Code", "Omschrijving"), Array("Weergave")
    KleurKolommen wsP, "Tbl_Taakveld", Array("Taakveld", "Omschrijving"), Array("Code", "Weergave")
    KleurKolommen wsP, "Tbl_Balanscode", Array("Code", "Omschrijving"), Array("Weergave")
    KleurKolommen wsP, "Tbl_BijdrageBron", Array("Bron"), Array("Jaar 0", "Jaar 1", "Jaar 2", "Jaar 3", "Jaar 4", "Jaar 5")
    wsP.Range("A13").Value = "Geel = zelf aan te passen. Grijs = automatisch, niet aanpassen."
    wsP.Range("A13").Font.Italic = True
End Sub

Private Sub KleurKolommen(ByVal ws As Worksheet, ByVal tabel As String, ByVal invoer As Variant, ByVal auto As Variant)
    Dim lo As ListObject, c As Variant
    On Error Resume Next
    Set lo = ws.ListObjects(tabel)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    For Each c In invoer
        lo.ListColumns(CStr(c)).DataBodyRange.Interior.Color = RGB(255, 255, 204)
    Next c
    If Not IsEmpty(auto) Then
        For Each c In auto
            lo.ListColumns(CStr(c)).DataBodyRange.Interior.Color = RGB(242, 242, 242)
        Next c
    End If
End Sub

' ======================================================================
' Werkelijke uitgaven bijwerken (stap 11)
' ======================================================================
' Stap 11 - Werkelijke uitgaven en inkomsten bijwerken vanuit de uitdraai van de financiele administratie
' UitgavenBijwerken      : vraagt 'Wil je het bestand uploaden?'; bij Ja wordt het nieuwste bestand uit de
'                          exportmap verwerkt. Herkent de kolommen zelf en verwerkt
'                          lopend jaar (AC) en vorig jaar (AB). Meldingen komen op blad 'Meldingen uitgaven'.
' UitgavenOpmaakAanmaken : kleurt AD (rood = uitgaven overschreden, groen = meer inkomsten dan geraamd),
'                          en AJ (toelichting nodig).
' Regels:
' - Sleutel 1: kredietnummer (7 cijfers, begint met 7) <-> kolom C. Sleutel 2: kostensoort begint met 4 (uitgaven) of 8 (inkomsten).
' - Totaalregels (wel nummer, geen kostensoort) tellen niet mee; per nummer en in totaal wordt nagerekend.
' - Nummers met budget 0 en in beide jaren geen bedragen die niet in kolom C staan, worden overgeslagen.
' - Samengesteld krediet (J is een getal): saldo van de hele groep via sleutel K over de deelkredieten.
' - Anders: uitgaven op de uitgavenregel, inkomsten op de bijdrageregel(s) (negatief budget) naar verhouding van het budget.
' - Peildatum = datum van het bestand. Vorig jaar na 31-3: niets meer wegschrijven, wel alarm bij elk verschil.
' - Uitgaven zonder krediet: alarm, telt mee in de aansluiting. Krediet niet in de uitdraai: bedrag blijft staan, met notitie.
' - Afgesloten krediet (A) waarvan het bedrag wijzigt: wel verwerkt, melding en notitie.
' - Harde poort: aansluiting per jaar exact 0,00. Anders wordt dat jaar niet weggeschreven.


Public Sub UitgavenBijwerken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, wsP As Worksheet, pad As String, peil As Date, lj As Long, laatste As Long
    Dim wbU As Workbook, d As Variant, nRij As Long, nKol As Long, kop As Long
    Dim cNr As Long, cKs As Long, cBud As Long, cJ(0 To 1) As Long, j As Long, y As Long
    Dim r As Long, i As Long, k As Long, nr As String, ks As String, v As Double, t As String
    Dim nrIdx As Object, nrs() As String, bU() As Double, bI() As Double, bO() As Double, bTot() As Double, heeftTot() As Boolean
    Dim bud() As Double, skip() As Boolean, nN As Long, det(0 To 1) As Double, eind(0 To 1) As Double, heeftEind As Boolean
    Dim geenNr(0 To 1) As Double, vS As Variant, nS As Long, stamRij As Object, rijen As Collection
    Dim grpRij As Object, oud() As Double, nieuw() As Double, gezet() As Boolean, schrijven As Boolean
    Dim zk As Double, ni As Double, onb As Double, verwacht As Double, som As Double, kol As Long
    Dim m0 As Variant, m1 As Variant, arr() As Variant, rng As Range, tekst As String, alarm As Long
    Dim geschreven(0 To 1) As Boolean, oudKol(0 To 1) As Variant, sOud(0 To 1) As Double, sNieuw(0 To 1) As Double
    Dim rv As Variant, grpBedrag As Object, g As Variant, vAD As Variant, rij0 As Long, bew() As Variant, nb As Long

    On Error GoTo Fout
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "U heeft geen recht om de werkelijke uitgaven bij te werken.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    Set wsP = ThisWorkbook.Worksheets("Parameters")
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    Set mMeld = New Collection
    mAansl = ""
    VorigeUploadLezen

    ' 1. Bestand kiezen: nieuwste bestand uit de exportmap
    pad = KiesUitdraai(CStr(ThisWorkbook.Names("Par_PadUitgavenbestand").RefersToRange.Value))
    If Len(pad) = 0 Then Exit Sub
    peil = DateValue(FileDateTime(pad))               ' peildatum = datum van het bestand

    ' beveiliging tijdelijk uit; aan het eind (ook bij een fout) weer aan
    BeveiligingTijdelijkUit

    ' 2. Uitdraai inlezen (alleen lezen, daarna sluiten)
    Application.ScreenUpdating = False
    Set wbU = Workbooks.Open(Filename:=pad, ReadOnly:=True, UpdateLinks:=0)
    d = wbU.Worksheets(1).UsedRange.Value
    rij0 = wbU.Worksheets(1).UsedRange.Row - 1
    wbU.Close SaveChanges:=False
    Set wbU = Nothing
    ThisWorkbook.Activate
    If Not IsArray(d) Then
        MsgBox "De uitdraai is leeg.", vbExclamation
        GoTo Einde
    End If
    nRij = UBound(d, 1): nKol = UBound(d, 2)

    ' 3. Kolommen herkennen
    kop = 0
    For r = 1 To Application.Min(10, nRij)
        For k = 1 To nKol
            j = JaarUitKop(Tkst(d(r, k)))
            If j = lj Or j = lj - 1 Then kop = r: Exit For
        Next k
        If kop > 0 Then Exit For
    Next r
    If kop = 0 Then
        MsgBox "In de eerste 10 regels van de uitdraai staat geen kolom met jaar " & lj & " of " & (lj - 1) & "." & vbCrLf & _
               "Er is niets verwerkt.", vbExclamation
        GoTo Einde
    End If
    cNr = BesteKolom(d, kop, 1, 0)
    cKs = BesteKolom(d, kop, 2, cNr)
    If cNr = 0 Or cKs = 0 Then
        MsgBox "Geen kolom met kredietnummers (7 cijfers, begint met 7) of kostensoorten (begint met 4 of 8) gevonden." & vbCrLf & _
               "Er is niets verwerkt.", vbExclamation
        GoTo Einde
    End If
    For y = 0 To 1
        cJ(y) = 0
        For k = 1 To nKol
            t = LCase$(Tkst(d(kop, k)))
            If JaarUitKop(t) = lj - y And InStr(t, "budget") = 0 And InStr(t, "restant") = 0 And InStr(t, "raming") = 0 Then
                If cJ(y) > 0 Then
                    MsgBox "Er zijn twee kolommen voor jaar " & (lj - y) & ": " & d(kop, cJ(y)) & " en " & d(kop, k) & "." & vbCrLf & _
                           "Er is niets verwerkt.", vbExclamation
                    GoTo Einde
                End If
                cJ(y) = k
            End If
        Next k
    Next y
    For k = 1 To nKol
        t = LCase$(Tkst(d(kop, k)))
        If JaarUitKop(t) = lj + 1 And InStr(t, "budget") = 0 Then
            MsgBox "De uitdraai heeft al een kolom voor " & (lj + 1) & ". Die wordt pas verwerkt na de jaarrol." & vbCrLf & _
                   "Nu worden alleen " & lj & " en " & (lj - 1) & " verwerkt.", vbInformation
        End If
        If cBud = 0 And InStr(t, "budget") > 0 And JaarUitKop(t) = 0 Then cBud = k
    Next k
    If cJ(0) = 0 Then
        MsgBox "De uitdraai heeft geen kolom voor het lopend jaar " & lj & ". Er is niets verwerkt.", vbExclamation
        GoTo Einde
    End If
    ' 4. Uitdraai optellen per kredietnummer
    Set nrIdx = CreateObject("Scripting.Dictionary")
    ReDim nrs(1 To nRij): ReDim bU(1 To nRij, 0 To 1): ReDim bI(1 To nRij, 0 To 1): ReDim bO(1 To nRij, 0 To 1)
    ReDim bTot(1 To nRij, 0 To 1): ReDim heeftTot(1 To nRij): ReDim bud(1 To nRij): ReDim skip(1 To nRij)
    For r = kop + 1 To nRij
        nr = Trim$(Tkst(d(r, cNr)))
        ks = Trim$(Tkst(d(r, cKs)))
        If Len(ks) = 0 Then
            If Len(nr) = 0 Then
                ' eindtotaal (of lege regel)
                For y = 0 To 1
                    If cJ(y) > 0 Then
                        If Not Leeg(d(r, cJ(y))) Then eind(y) = eind(y) + Nul(d(r, cJ(y))): heeftEind = True
                    End If
                Next y
            ElseIf IsNummer7(nr) Then
                i = NrIndex(nrIdx, nrs, nN, nr)
                heeftTot(i) = True
                For y = 0 To 1
                    If cJ(y) > 0 Then bTot(i, y) = bTot(i, y) + Nul(d(r, cJ(y)))
                Next y
            End If
        Else
            For y = 0 To 1
                If cJ(y) > 0 Then
                    v = Nul(d(r, cJ(y)))
                    det(y) = det(y) + v
                    If Not IsNummer7(nr) Then
                        If v <> 0 Then
                            geenNr(y) = geenNr(y) + v
                            Meld 1, "ALARM: regel zonder geldig kredietnummer", lj - y, nr, ks, 0, Empty, v, _
                                 "Rij " & (r + rij0) & " in de uitdraai. Niet verwerkt, besluit nodig."
                        End If
                    Else
                        i = NrIndex(nrIdx, nrs, nN, nr)
                        If Left$(ks, 1) = "4" Then
                            bU(i, y) = bU(i, y) + v
                        ElseIf Left$(ks, 1) = "8" Then
                            bI(i, y) = bI(i, y) + v
                        ElseIf v <> 0 Then
                            bO(i, y) = bO(i, y) + v
                            Meld 1, "ALARM: onbekende kostensoort", lj - y, nr, ks, 0, Empty, v, _
                                 "Rij " & (r + rij0) & " in de uitdraai. Kostensoort begint niet met 4 of 8. Niet verwerkt."
                        End If
                    End If
                End If
            Next y
            If IsNummer7(nr) And cBud > 0 Then
                i = NrIndex(nrIdx, nrs, nN, nr)
                bud(i) = bud(i) + Nul(d(r, cBud))
            End If
        End If
    Next r

    ' Nacontrole op dubbeltelling: totaalregels en eindtotaal
    For i = 1 To nN
        If heeftTot(i) Then
            For y = 0 To 1
                If cJ(y) > 0 Then
                    If Abs(bU(i, y) + bI(i, y) + bO(i, y) - bTot(i, y)) > 0.005 Then
                        MsgBox "Kredietnummer " & nrs(i) & ", jaar " & (lj - y) & ": de detailregels (" & Format$(bU(i, y) + bI(i, y) + bO(i, y), "#,##0.00") & _
                               ") sluiten niet op de totaalregel (" & Format$(bTot(i, y), "#,##0.00") & ")." & vbCrLf & "Er is niets verwerkt.", vbCritical
                        GoTo Einde
                    End If
                End If
            Next y
        End If
        skip(i) = (bud(i) = 0 And bU(i, 0) = 0 And bI(i, 0) = 0 And bO(i, 0) = 0 And bU(i, 1) = 0 And bI(i, 1) = 0 And bO(i, 1) = 0)
    Next i
    If heeftEind Then
        For y = 0 To 1
            If cJ(y) > 0 Then
                If Abs(det(y) - eind(y)) > 0.005 Then
                    MsgBox "Jaar " & (lj - y) & ": de som van de detailregels (" & Format$(det(y), "#,##0.00") & _
                           ") sluit niet op het eindtotaal (" & Format$(eind(y), "#,##0.00") & ")." & vbCrLf & "Er is niets verwerkt.", vbCritical
                    GoTo Einde
                End If
            End If
        Next y
    End If

    ' 5. Verzamellijst inlezen
    If laatste < 2 Then GoTo Einde
    vS = wv.Range("A2:AJ" & laatste).Value
    nS = UBound(vS, 1)
    Set stamRij = CreateObject("Scripting.Dictionary")
    Set grpRij = CreateObject("Scripting.Dictionary")
    For i = 1 To nS
        If Not Leeg(vS(i, 1)) Then
            nr = NrTekst(vS(i, 3))
            If Len(nr) > 0 And nr <> "7999998" Then
                If Not stamRij.Exists(nr) Then stamRij.Add nr, New Collection
                stamRij(nr).Add i
            End If
            If IsGetal(vS(i, 10)) Then
                If Not grpRij.Exists(CStr(vS(i, 10))) Then grpRij.Add CStr(vS(i, 10)), New Collection
                grpRij(CStr(vS(i, 10))).Add i
            End If
        End If
    Next i

    ' 6. Per jaar verwerken
    m0 = MeetStap8(wv, laatste)
    For y = 0 To 1
        If cJ(y) = 0 Then GoTo VolgendJaar
        kol = IIf(y = 0, 29, 28)                       ' AC = lopend jaar, AB = vorig jaar
        schrijven = (y = 0) Or (peil <= DateSerial(lj, 3, 31))
        ' formules in de kolom (bijv. handmatige verdeling =bedrag*40%) worden vervangen door de waarde uit de uitdraai
        ReDim oud(1 To nS): ReDim nieuw(1 To nS): ReDim gezet(1 To nS)
        For i = 1 To nS
            If Not Leeg(vS(i, 1)) Then oud(i) = Nul(vS(i, kol))
            nieuw(i) = oud(i)
        Next i
        zk = 0: ni = 0: onb = 0
        Set grpBedrag = CreateObject("Scripting.Dictionary")

        For i = 1 To nN
            onb = onb + bO(i, y)
            If skip(i) And Not stamRij.Exists(nrs(i)) Then GoTo VolgendNr
            If Not stamRij.Exists(nrs(i)) Then
                If bU(i, y) <> 0 Or bI(i, y) <> 0 Then
                    zk = zk + bU(i, y) + bI(i, y)
                    If Gemuteerd(nrs(i), lj - y, bU(i, y) + bI(i, y)) Then Meld 1, "ALARM: uitgaven zonder krediet", lj - y, nrs(i), "", 0, Empty, bU(i, y) + bI(i, y), _
                         "Kredietnummer staat niet in kolom C. Besluit nodig; rapporteren aan de raad, budgethouder in gebreke."
                End If
                GoTo VolgendNr
            End If
            Set rijen = stamRij(nrs(i))
            g = Empty
            For Each rv In rijen
                If IsGetal(vS(rv, 10)) Then g = CStr(vS(rv, 10)): Exit For
            Next rv
            If Not IsEmpty(g) Then
                grpBedrag(g) = grpBedrag(g) + bU(i, y) + bI(i, y)
                For Each rv In rijen
                    ' regels van dit nummer buiten de groep (bijv. bijdrageregel zonder groepsnummer) krijgen 0
                    If Not IsGetal(vS(rv, 10)) Then nieuw(rv) = 0: gezet(rv) = True
                Next rv
            Else
                VerdeelNummer vS, rijen, bU(i, y), bI(i, y), nieuw, gezet, nrs(i), lj - y
            End If
VolgendNr:
        Next i
        For Each g In grpBedrag.Keys
            VerdeelGroep vS, grpRij(g), CDbl(grpBedrag(g)), nieuw, gezet, CStr(g), lj - y
        Next g

        ' Regels die niet in de uitdraai staan: bedrag blijft staan
        For i = 1 To nS
            If Not gezet(i) And oud(i) <> 0 And Not Leeg(vS(i, 1)) Then
                ni = ni + oud(i)
            End If
        Next i

        If schrijven Then
            ' Harde poort: aansluiting
            som = 0
            For i = 1 To nS
                nieuw(i) = Round(nieuw(i), 2)
                som = som + nieuw(i)
            Next i
            verwacht = det(y) - onb - geenNr(y) - zk + ni
            If Round(som - verwacht, 2) <> 0 Then
                Meld 1, "ALARM: aansluiting sluit niet", lj - y, "", "", 0, verwacht, som, _
                     "Verschil " & Format$(som - verwacht, "#,##0.00") & ". Jaar " & (lj - y) & " is NIET bijgewerkt."
                GoTo VolgendJaar
            End If
            mAansl = mAansl & "Aansluiting " & (lj - y) & ": uitdraai " & Format$(det(y), "#,##0.00") & ", verzamellijst " & _
                     Format$(som, "#,##0.00") & IIf(Round(som - det(y), 2) <> 0, " (verschil verklaard door uitgaven zonder krediet / bedragen die blijven staan)", "") & ". Sluit.  "
            ReDim arr(1 To nS, 1 To 1)
            For i = 1 To nS
                If Leeg(vS(i, 1)) Then arr(i, 1) = vS(i, kol) Else arr(i, 1) = nieuw(i)
            Next i
            Set rng = wv.Range(wv.Cells(2, kol), wv.Cells(laatste, kol))
            oudKol(y) = rng.Formula                     ' formules bewaren voor terugzetten
            rng.Value = arr
            geschreven(y) = True
            sOud(y) = 0: sNieuw(y) = som
            For i = 1 To nS
                sOud(y) = sOud(y) + oud(i)
            Next i
            ' notities: oude weg, nieuwe bij regels die niet in de uitdraai staan
            rng.ClearComments
            For i = 1 To nS
                If Not gezet(i) And oud(i) <> 0 Then
                    wv.Cells(i + 1, kol).AddComment "Niet in uitdraai " & Dir(pad) & " (peildatum " & Format$(peil, "dd-mm-yyyy") & "). Bedrag blijft staan."
                End If
                If gezet(i) And Round(nieuw(i) - oud(i), 2) <> 0 And UCase$(Left$(CStr(vS(i, 8)), 1)) = "A" Then
                    Meld 2, "Afgesloten krediet, toch uitgaven/inkomsten verwerkt", lj - y, NrTekst(vS(i, 3)), CStr(vS(i, 4)), i + 1, oud(i), nieuw(i), _
                         "Status " & vS(i, 8) & ". Bedrag gewijzigd en wel verwerkt. Gemarkeerd met een notitie in kolom '" & wv.Cells(1, kol).Text & "'."
                    wv.Cells(i + 1, kol).AddComment "Afgesloten krediet (" & vS(i, 8) & "): bedrag gewijzigd van " & Format$(oud(i), "#,##0.00") & _
                                                    " naar " & Format$(nieuw(i), "#,##0.00") & " (uitdraai " & Format$(peil, "dd-mm-yyyy") & ")."
                End If
            Next i
        Else
            ' Vorig jaar na 31-3: alleen controleren
            For Each g In stamRij.Keys
                som = 0
                For Each rv In stamRij(g)
                    som = som + oud(rv)
                Next rv
                v = 0
                If nrIdx.Exists(g) Then
                    i = nrIdx(g)
                    v = bU(i, y) + bI(i, y)
                End If
                If Round(som - v, 2) <> 0 And Gemuteerd(CStr(g), lj - y, v) Then
                    Meld 1, "ALARM: " & (lj - y) & " gewijzigd na 31-3, actie/herstel nodig", lj - y, CStr(g), CStr(vS(stamRij(g)(1), 4)), _
                         stamRij(g)(1) + 1, som, v, "Jaarrekening is definitief. Kolom '" & wv.Cells(1, kol).Text & "' is NIET aangepast."
                End If
            Next g
        End If
VolgendJaar:
    Next y

    ' 7. Nulmeting achteraf: alleen AB/AC mogen veranderd zijn
    Application.Calculate
    m1 = MeetStap8(wv, laatste)
    For j = 0 To 9
        If j <> 3 And j <> 4 Then
            If Round(m0(j) - m1(j), 2) <> 0 Then Err.Raise vbObjectError + 11, , "Nulmeting wijkt af in meting " & j & ". Alles is teruggedraaid."
        End If
    Next j

    Application.Calculate

    ' 9. Vastleggen
    ThisWorkbook.Names("Par_PadUitgavenbestand").RefersToRange.Value = pad
    ThisWorkbook.Names("Par_PeildatumUitgaven").RefersToRange.Value = peil
    For y = 0 To 1
        If geschreven(y) Then
            kol = IIf(y = 0, 29, 28)
            wv.Cells(1, kol).ClearComments
            wv.Cells(1, kol).AddComment "Bijgewerkt uit " & Dir(pad) & ", kolom '" & d(kop, cJ(y)) & "', peildatum " & _
                                        Format$(peil, "dd-mm-yyyy") & ", op " & Format$(Now, "dd-mm-yyyy hh:mm") & "."
            LogRegel "werkelijke uitgaven bijwerken", , , wv.Cells(1, kol).Text, Format$(sOud(y), "0.00"), Format$(sNieuw(y), "0.00"), Dir(pad) & " (" & (lj - y) & ")"
        End If
    Next y
    ' bedragen van deze upload bewaren als vergelijking voor de volgende keer
    ReDim bew(1 To 2 * nN + 1, 1 To 3)
    For i = 1 To nN
        For y = 0 To 1
            If cJ(y) > 0 Then
                If Round(bU(i, y) + bI(i, y) + bO(i, y), 2) <> 0 Then
                    nb = nb + 1
                    bew(nb, 1) = nrs(i): bew(nb, 2) = lj - y: bew(nb, 3) = Round(bU(i, y) + bI(i, y) + bO(i, y), 2)
                End If
            End If
        Next y
    Next i
    VorigeUploadSchrijven bew, nb, Dir(pad)
    UitgavenOpmaakAanmaken
    Alert7999998 vS, lj
    alarm = MeldingenSchrijven(Dir(pad), peil)
    RaadsoverzichtMaken
    ControlesVernieuwen
    BeveiligingAan
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(BL_MELD).Activate
    tekst = "Verwerkt: " & Dir(pad) & vbCrLf & _
            lj & ": " & IIf(geschreven(0), "bijgewerkt", "NIET bijgewerkt") & vbCrLf
    If cJ(1) > 0 Then tekst = tekst & (lj - 1) & ": " & IIf(geschreven(1), "bijgewerkt", "alleen gecontroleerd") & vbCrLf
    tekst = tekst & vbCrLf & mMeld.Count & " meldingen, waarvan " & alarm & " ALARM." & vbCrLf & "Zie blad '" & BL_MELD & "'." & vbCrLf & "Het kredietoverzicht is bijgewerkt (nieuw tabblad achteraan)."
    MsgBox tekst, IIf(alarm > 0, vbCritical, vbInformation), IIf(alarm > 0, "LET OP: actie nodig", "Uitgaven bijgewerkt")
    Exit Sub

Einde:
    BeveiligingAan
    Application.ScreenUpdating = True
    Exit Sub
Fout:
    Dim foutTekst As String
    foutTekst = Err.Description
    On Error Resume Next
    If Not wbU Is Nothing Then wbU.Close SaveChanges:=False
    For y = 0 To 1
        If geschreven(y) Then wv.Range(wv.Cells(2, IIf(y = 0, 29, 28)), wv.Cells(laatste, IIf(y = 0, 29, 28))).Formula = oudKol(y)
    Next y
    BeveiligingAan
    Application.ScreenUpdating = True
    MsgBox "Fout: " & foutTekst & vbCrLf & "Er is niets gewijzigd.", vbCritical
End Sub

' Verdeelt een kredietnummer zonder groep: uitgaven op de uitgavenregel(s), inkomsten op de bijdrageregel(s).
Private Sub VerdeelNummer(ByVal vS As Variant, ByVal rijen As Collection, ByVal uit As Double, ByVal ink As Double, _
                          ByRef nieuw() As Double, ByRef gezet() As Boolean, ByVal nr As String, ByVal jaar As Long)
    Dim r As Variant, uR As New Collection, bR As New Collection, b As Double
    For Each r In rijen
        nieuw(r) = 0
        gezet(r) = True
        b = BudgetRegel(vS, r)
        If b < 0 Then bR.Add r Else uR.Add r
    Next r
    If bR.Count = 0 Then
        Verdeel vS, uR, uit + ink, nieuw, False
    ElseIf uR.Count = 0 Then
        Verdeel vS, bR, uit + ink, nieuw, True
    Else
        Verdeel vS, uR, uit, nieuw, False
        Verdeel vS, bR, ink, nieuw, True
    End If
End Sub

' Verdeelt het saldo van een samengesteld krediet via sleutel K over alle regels van de groep.
Private Sub VerdeelGroep(ByVal vS As Variant, ByVal rijen As Collection, ByVal bedrag As Double, _
                         ByRef nieuw() As Double, ByRef gezet() As Boolean, ByVal grp As String, ByVal jaar As Long)
    Dim r As Variant, somK As Double, rest As Double, deel As Double, laatsteR As Long
    For Each r In rijen
        nieuw(r) = 0
        gezet(r) = True
        ' alleen uitgavenregels; een bijdrageregel (negatief budget) telt niet mee in de verdeelsleutel
        If Nul(vS(r, 11)) > 0 And BudgetRegel(vS, r) >= 0 Then somK = somK + Nul(vS(r, 11)): laatsteR = r
    Next r
    If somK <= 0 Then
        nieuw(rijen(1)) = bedrag
        Exit Sub
    End If
    rest = bedrag
    For Each r In rijen
        If Nul(vS(r, 11)) > 0 And BudgetRegel(vS, r) >= 0 And r <> laatsteR Then
            deel = Round(bedrag * Nul(vS(r, 11)) / somK, 2)
            nieuw(r) = deel
            rest = rest - deel
        End If
    Next r
    nieuw(laatsteR) = Round(rest, 2)
End Sub

' Verdeelt een bedrag over regels naar verhouding van het budget (bij bijdrageregels het negatieve budget).
' Laatste regel krijgt de afrondingsrest. Zonder budget gaat alles naar de eerste regel.
Private Sub Verdeel(ByVal vS As Variant, ByVal rijen As Collection, ByVal bedrag As Double, ByRef nieuw() As Double, ByVal negatief As Boolean)
    Dim r As Variant, w As Double, som As Double, rest As Double, deel As Double, n As Long
    If rijen.Count = 0 Then Exit Sub
    For Each r In rijen
        w = BudgetRegel(vS, r)
        If negatief Then w = -w
        If w > 0 Then som = som + w
    Next r
    If som <= 0 Or rijen.Count = 1 Then
        nieuw(rijen(1)) = nieuw(rijen(1)) + bedrag
        Exit Sub
    End If
    rest = bedrag
    For Each r In rijen
        n = n + 1
        w = BudgetRegel(vS, r)
        If negatief Then w = -w
        If w < 0 Then w = 0
        If n = rijen.Count Then
            nieuw(r) = nieuw(r) + Round(rest, 2)
        Else
            deel = Round(bedrag * w / som, 2)
            nieuw(r) = nieuw(r) + deel
            rest = rest - deel
        End If
    Next r
End Sub

' Totaal budget van een regel: Z plus AE t/m AI.
Private Function BudgetRegel(ByVal vS As Variant, ByVal r As Long) As Double
    Dim c As Long, som As Double
    som = Nul(vS(r, 26))
    For c = 31 To 35
        som = som + Nul(vS(r, c))
    Next c
    BudgetRegel = som
End Function

' Kiest de uitdraai: het nieuwste bestand in de exportmap (map van Par_PadUitgavenbestand). Een vraag: uploaden ja/nee.
' Is de map nog niet bekend of niet bereikbaar, dan wordt eenmalig de map gevraagd.
Private Function KiesUitdraai(ByVal vorigPad As String) As String
    Dim map As String, f As String, nieuwst As String, dt As Date
    If Len(vorigPad) > 0 Then map = Left$(vorigPad, InStrRev(vorigPad, "\"))
    If Not MapBestaat(map) Then
        With Application.FileDialog(4)                 ' 4 = mapkiezer
            .Title = "Kies de map 'export coda' (Stambestand kredieten)"
            If .Show <> -1 Then Exit Function
            map = .SelectedItems(1)
        End With
        If Right$(map, 1) <> "\" Then map = map & "\"
    End If
    f = Dir(map & "*.xls*")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then
            If Len(nieuwst) = 0 Or FileDateTime(map & f) > dt Then nieuwst = map & f: dt = FileDateTime(map & f)
        End If
        f = Dir()
    Loop
    If Len(nieuwst) = 0 Then
        MsgBox "Geen Excelbestand gevonden in:" & vbCrLf & map, vbExclamation
        Exit Function
    End If
    If MsgBox("Wil je het bestand uploaden?" & vbCrLf & vbCrLf & Dir(nieuwst) & vbCrLf & _
              "datum " & Format$(dt, "dd-mm-yyyy hh:mm"), vbYesNo + vbQuestion, "Werkelijke uitgaven bijwerken") = vbYes Then
        KiesUitdraai = nieuwst
    End If
End Function

Private Function MapBestaat(ByVal map As String) As Boolean
    If Len(map) = 0 Then Exit Function
    On Error Resume Next
    MapBestaat = (Len(Dir(map, vbDirectory)) > 0)
End Function

' Zoekt de kolom met de meeste kredietnummers (soort 1) of kostensoorten (soort 2) onder de kopregel.
Private Function BesteKolom(ByVal d As Variant, ByVal kop As Long, ByVal soort As Long, ByVal niet As Long) As Long
    Dim k As Long, r As Long, n As Long, best As Long, t As String
    For k = 1 To UBound(d, 2)
        If k <> niet Then
            n = 0
            For r = kop + 1 To Application.Min(UBound(d, 1), kop + 300)
                t = Trim$(Tkst(d(r, k)))
                If soort = 1 Then
                    If IsNummer7(t) Then n = n + 1
                Else
                    If (Left$(t, 1) = "4" Or Left$(t, 1) = "8") And (Len(t) = 5 Or Len(t) = 6) And IsAlleenCijfers(t) Then n = n + 1
                End If
            Next r
            If n > best Then best = n: BesteKolom = k
        End If
    Next k
End Function

' Geeft het eerste jaartal (2000 t/m 2099) in een koptekst, anders 0.
Private Function JaarUitKop(ByVal t As String) As Long
    Dim p As Long
    For p = 1 To Len(t) - 3
        If Mid$(t, p, 2) = "20" And IsAlleenCijfers(Mid$(t, p, 4)) Then
            If p + 4 > Len(t) Or Not IsAlleenCijfers(Mid$(t, p + 4, 1)) Then
                If p = 1 Or Not IsAlleenCijfers(Mid$(t, IIf(p > 1, p - 1, 1), 1)) Then JaarUitKop = CLng(Mid$(t, p, 4)): Exit Function
            End If
        End If
    Next p
End Function

Private Function Tkst(ByVal v As Variant) As String
    If Not IsError(v) Then Tkst = CStr(v)
End Function

' True als minstens een cel in het bereik een formule heeft.
Private Function HeeftFormules(ByVal rng As Range) As Boolean
    Dim hf As Variant
    hf = rng.HasFormula
    If IsNull(hf) Then HeeftFormules = True Else HeeftFormules = CBool(hf)
End Function

Private Function IsAlleenCijfers(ByVal t As String) As Boolean
    Dim p As Long
    If Len(t) = 0 Then Exit Function
    For p = 1 To Len(t)
        If Mid$(t, p, 1) < "0" Or Mid$(t, p, 1) > "9" Then Exit Function
    Next p
    IsAlleenCijfers = True
End Function

Private Function IsNummer7(ByVal t As String) As Boolean
    IsNummer7 = (Len(t) = 7 And Left$(t, 1) = "7" And IsAlleenCijfers(t))
End Function

Private Function IsGetal(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: IsGetal = True
    End Select
End Function

' Kredietnummer uit kolom C als tekst (getal zonder decimalen).
Private Function NrTekst(ByVal v As Variant) As String
    If Leeg(v) Then Exit Function
    If IsGetal(v) Then NrTekst = Format$(v, "0") Else NrTekst = Trim$(CStr(v))
End Function

Private Function NrIndex(ByVal dict As Object, ByRef nrs() As String, ByRef nN As Long, ByVal nr As String) As Long
    If Not dict.Exists(nr) Then
        nN = nN + 1
        nrs(nN) = nr
        dict.Add nr, nN
    End If
    NrIndex = dict(nr)
End Function

Private Function Letter(ByVal k As Long) As String
    Letter = Split(ThisWorkbook.Worksheets(1).Cells(1, k).Address(True, False), "$")(0)
End Function

' Leest de bedragen van de vorige upload (verborgen blad 'Vorige upload'). Geen blad = eerste upload.
Private Sub VorigeUploadLezen()
    Dim ws As Worksheet, r As Long, laatste As Long
    Set mVorige = CreateObject("Scripting.Dictionary")
    mHeeftVorige = BladBestaat("Vorige upload")
    If Not mHeeftVorige Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Vorige upload")
    laatste = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To laatste
        mVorige(CStr(ws.Cells(r, 1).Value) & "|" & CStr(ws.Cells(r, 2).Value)) = Nul(ws.Cells(r, 3).Value)
    Next r
End Sub

' True als het bedrag anders is dan bij de vorige upload. Eerste upload: alleen het lopend jaar telt
' (vorige jaren zijn al beoordeeld en worden de beginstand).
Private Function Gemuteerd(ByVal nr As String, ByVal jaar As Long, ByVal bedrag As Double) As Boolean
    Dim vorigBedrag As Double
    If Not mHeeftVorige Then
        Gemuteerd = (jaar = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value))
        Exit Function
    End If
    If mVorige.Exists(nr & "|" & jaar) Then vorigBedrag = mVorige(nr & "|" & jaar)
    Gemuteerd = (Round(bedrag - vorigBedrag, 2) <> 0)
End Function

Private Sub VorigeUploadSchrijven(ByRef bew() As Variant, ByVal nb As Long, ByVal bestand As String)
    Dim ws As Worksheet
    If Not BladBestaat("Vorige upload") Then
        Set ws = NieuwBlad()
        ws.Name = "Vorige upload"
    Else
        Set ws = ThisWorkbook.Worksheets("Vorige upload")
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "Bedragen per kredietnummer en jaar uit de laatste upload (" & bestand & "). Niet wijzigen: hiermee worden mutaties herkend."
    ws.Range("A2:C2").Value = Array("Kredietnummer", "Jaar", "Bedrag")
    ws.Columns("A").NumberFormat = "@"
    If nb > 0 Then ws.Range("A3").Resize(nb, 3).Value = bew
    ws.Visible = xlSheetHidden
End Sub

' Meldingen: 1 = ALARM, 2 = niet toegestaan, 3 = toelichting nodig, 4 = let op, 5 = info.
Private Sub Meld(ByVal prio As Long, ByVal soort As String, ByVal jaar As Long, ByVal nr As String, ByVal oms As String, _
                 ByVal rij As Long, ByVal oud As Variant, ByVal nieuw As Variant, ByVal toel As String)
    If mMeld Is Nothing Then Set mMeld = New Collection
    mMeld.Add Array(prio, soort, jaar, nr, oms, rij, oud, nieuw, toel)
End Sub

' Schrijft blad 'Meldingen uitgaven' opnieuw, gesorteerd op ernst. Geeft het aantal ALARM-meldingen terug.
Private Function MeldingenSchrijven(ByVal bestand As String, ByVal peil As Date) As Long
    Dim ws As Worksheet, rij As Long, p As Long, m As Variant, kleur As Long, alarm As Long
    Dim wv As Worksheet, stat As Object, r As Long, laatste As Long, nrT As String
    ' status (kolom H) per kredietnummer, voor de kolom Status
    Set wv = Vz()
    Set stat = CreateObject("Scripting.Dictionary")
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        nrT = NrTekst(wv.Cells(r, 3).Value)
        If Len(nrT) > 0 And Not stat.Exists(nrT) Then stat.Add nrT, CStr(wv.Cells(r, 8).Value)
    Next r
    If Not BladBestaat(BL_MELD) Then
        Set ws = NieuwBlad()
        ws.Name = BL_MELD
    Else
        Set ws = ThisWorkbook.Worksheets(BL_MELD)
    End If
    ws.Cells.Clear
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "Meldingen werkelijke uitgaven"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Uitdraai: " & bestand & " - peildatum " & Format$(peil, "dd-mm-yyyy") & " - verwerkt op " & _
                           Format$(Now, "dd-mm-yyyy hh:mm") & " door " & ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value
    ws.Range("A2").Font.Italic = True
    ws.Range("A4:J4").Value = Array("Soort", "Jaar", "Kredietnummer", "Omschrijving", "Status", "Rij", "Oud / verwacht", "Nieuw / uitdraai", "Verschil", "Toelichting")
    With ws.Range("A4:J4")
        .Font.Bold = True
        .Interior.Color = RGB(208, 208, 208)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    rij = 5
    For p = 1 To 5
        For Each m In mMeld
            If m(0) = p Then
                ws.Cells(rij, 1).Value = m(1)
                ws.Cells(rij, 2).Value = m(2)
                ws.Cells(rij, 3).NumberFormat = "@"
                ws.Cells(rij, 3).Value = m(3)
                ws.Cells(rij, 4).Value = m(4)
                If m(5) > 1 Then
                    ws.Cells(rij, 5).Value = wv.Cells(m(5), 8).Value
                ElseIf stat.Exists(CStr(m(3))) Then
                    ws.Cells(rij, 5).Value = stat(CStr(m(3)))
                ElseIf Len(CStr(m(3))) > 0 Then
                    ws.Cells(rij, 5).Value = "geen krediet"
                End If
                If m(5) > 0 Then ws.Cells(rij, 6).Value = m(5)
                If Not IsEmpty(m(6)) Then ws.Cells(rij, 7).Value = m(6)
                If Not IsEmpty(m(7)) Then ws.Cells(rij, 8).Value = m(7)
                If Not IsEmpty(m(6)) And Not IsEmpty(m(7)) Then ws.Cells(rij, 9).Value = m(7) - m(6)
                ws.Cells(rij, 10).Value = m(8)
                Select Case p
                    Case 1: kleur = RGB(255, 150, 150): alarm = alarm + 1
                        ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 10)).Font.Bold = True
                    Case 2: kleur = RGB(255, 190, 190)
                    Case 3: kleur = RGB(255, 200, 120)
                    Case 4: kleur = RGB(255, 255, 204)
                    Case Else: kleur = xlNone
                End Select
                If kleur <> xlNone Then ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 10)).Interior.Color = kleur
                rij = rij + 1
            End If
        Next m
    Next p
    If rij = 5 Then ws.Range("A5").Value = "Geen meldingen."
    ws.Range("A3").Value = alarm & " ALARM-meldingen, " & (mMeld.Count - alarm) & " overige meldingen.  " & mAansl
    ws.Range("A3").Font.Bold = True
    If alarm > 0 Then
        ws.Range("A3").Interior.Color = RGB(255, 0, 0)
        ws.Range("A3").Font.Color = RGB(255, 255, 255)
        ws.Tab.Color = RGB(255, 0, 0)
    Else
        ws.Tab.Color = RGB(0, 176, 80)
    End If
    ws.Range("G5:I" & Application.Max(5, rij)).NumberFormat = "#,##0.00;-#,##0.00;-"
    ws.Range("B:C,E:F").HorizontalAlignment = xlLeft
    ws.Columns("A").ColumnWidth = 44
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 36
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 7
    ws.Columns("G:I").ColumnWidth = 16
    ws.Columns("J").ColumnWidth = 80
    ws.Range("A4:J" & rij).VerticalAlignment = xlTop
    ws.Range("A4:J4").WrapText = True
    ws.Range("A5:A" & rij & ",D5:D" & rij & ",J5:J" & rij).WrapText = True
    BladOpmaken ws, 4, 5, rij
    KnopTerug ws, ws.Range("D1")
    MeldingenSchrijven = alarm
End Function

' Voorwaardelijke opmaak op de verzamellijst. Wordt eenmalig aangemaakt (herkend aan de regel op AD2).
Public Sub UitgavenOpmaakAanmaken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, maxRij As Long, fc As Object, f As Object, vorig As Object
    Set wv = Vz()
    For Each f In wv.Range("AD2").FormatConditions
        If f.Type = xlExpression Then
            If InStr(f.Formula1, "$AD2") > 0 And InStr(f.Formula1, "$Z2") > 0 Then Exit Sub
        End If
    Next f
    Set vorig = ActiveSheet
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    ' rood: uitgavenregel (budget niet negatief) met nog uit te geven onder nul
    Set fc = VoegCFVerzamellijst(wv.Range("AD2:AD" & maxRij), "=AND($A2<>"""",ISNUMBER($AD2),N($Z2)>=0,$AD2<-0.005)")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 150): fc.SetFirstPriority
    ' groen: bijdrageregel (budget negatief) met meer inkomsten dan geraamd
    Set fc = VoegCFVerzamellijst(wv.Range("AD2:AD" & maxRij), "=AND($A2<>"""",ISNUMBER($AD2),N($Z2)<0,$AD2>0.005)")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(146, 208, 80): fc.SetFirstPriority
    ' oranje: toelichting nodig
    Set fc = VoegCFVerzamellijst(wv.Range("AJ2:AJ" & maxRij), _
             "=AND($A2<>"""",ISNUMBER($AD2),OR(AND(N($Z2)>=0,$AD2<-0.005),AND(N($Z2)<0,$AD2>0.005)),LEN(TRIM($AJ2))=0)")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 200, 120): fc.SetFirstPriority
    If Not vorig Is Nothing Then vorig.Activate
End Sub

' Zoals VoegCF, maar vertaalt de formule via een hulpcel op Parameters (op de verzamellijst is Z1 een kopcel).
Private Function VoegCFVerzamellijst(ByVal rng As Range, ByVal formule As String) As Object
    Dim t As Range, lokaal As String
    On Error GoTo Mis
    Set t = ThisWorkbook.Worksheets("Parameters").Range("XFD1")
    If Not Leeg(t.Value) Or t.HasFormula Then GoTo Mis
    t.Formula = formule
    lokaal = t.FormulaLocal
    t.Clear
    rng.Worksheet.Activate
    rng.Cells(1, 1).Select
    Set VoegCFVerzamellijst = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=lokaal)
    Exit Function
Mis:
    cfMislukt = True
    Set VoegCFVerzamellijst = Nothing
End Function

' ======================================================================
' Kredietoverzichten (stap 13)
' ======================================================================
' Stap 13 - Uitdraaien (functioneel ontwerp par. 7, aangevuld met latere afspraken)
' KredietoverzichtMaken         : overzicht raad: L t/m lopend jaar plus A van het lopend jaar, zonder regels met negatief budget.
' KredietoverzichtOpNaam        : overzicht budgethouder/projectleider/tijdschrijver (keuze op het Dashboard), incl. toekomstige
'                                 kredieten (aparte sectie "Nog niet gestart") en bijdrageregels.
' KredietoverzichtJaarrekening  : stand per 31-12 vorig jaar uit het blad 'Verzamellijst t/m <jaar>' (ontstaat bij de jaarovergang):
'                                 alle L en A van dat jaar.
' UitdraaiBlokPlaatsen          : blok 3 op het Dashboard (rol, naam en drie knoppen).
' Alle overzichten: nieuw blad met harde waarden, sortering programma > investeringsgroep > samengesteld/standaard,
' groepskoppen en tussentotalen (SUBTOTAL), kolom Status, bij A altijd "Krediet afgesloten" in de toelichting,
' kolom Melding uit het blad 'Meldingen uitgaven', aansluitcontrole op het restantkrediet in de kop.

Public Sub KredietoverzichtMaken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Dit overzicht wordt gemaakt door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    RaadsoverzichtMaken
End Sub

' Ook aangeroepen na het bijwerken van de werkelijke uitgaven (dan zonder rechtencontrole).
Private Sub RaadsoverzichtMaken()
    Dim lj As Long, ws As Worksheet, map As String
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    Set ws = OverzichtMaken("raad", Vz(), lj, 0, "")
    If ws Is Nothing Then Exit Sub
    ' elke opvraging bewaren in de map met kredietoverzichten
    map = MapStambestand("03 kredietoverzichten")
    If Len(map) > 0 Then BladOpslaan ws, map, "Kredietoverzicht " & Format$(Date, "dd-mm-yyyy")
    ws.Activate
End Sub

Public Sub KredietoverzichtOpNaam()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, rol As String, naam As String, kol As Long, lj As Long, ws As Worksheet, map As String
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    rol = Trim$(CStr(wd.Range("C58").Value))
    naam = Trim$(CStr(wd.Range("C59").Value))
    Select Case LCase$(rol)
        Case "budgetbeheerder": kol = 15
        Case "projectleider": kol = 16
        Case "tijdschrijver": kol = 17
        Case Else
            MsgBox "Kies op het Dashboard in cel C58 de rol: Budgetbeheerder, Projectleider of Tijdschrijver.", vbExclamation
            Exit Sub
    End Select
    If Len(naam) = 0 Then
        MsgBox "Kies op het Dashboard in cel C59 een naam.", vbExclamation
        Exit Sub
    End If
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    Set ws = OverzichtMaken("naam", Vz(), lj, kol, naam)
    If ws Is Nothing Then Exit Sub
    map = MapStambestand("02 Budgetbeheerders en projectleiders")
    If Len(map) > 0 Then
        map = SubMap(map, naam)
        BladOpslaan ws, map, "Kredietoverzicht " & naam & " " & Format$(Date, "dd-mm-yyyy")
        ws.Activate
        MsgBox "Opgeslagen in:" & vbCrLf & map, vbInformation
    End If
End Sub

Public Sub KredietoverzichtJaarrekening()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, jaar As Long, best As Worksheet, j As Long, map As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Dit overzicht wordt gemaakt door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Verzamellijst t/m ####" Then
            j = CLng(Right$(ws.Name, 4))
            If j > jaar Then jaar = j: Set best = ws
        End If
    Next ws
    If best Is Nothing Then
        If MsgBox("Er is nog geen blad 'Verzamellijst t/m <jaar>' (ontstaat bij de jaarovergang, stand per 31-12)." & vbCrLf & vbCrLf & _
                  "Een PROEF maken op basis van de huidige verzamellijst, om de opmaak te bekijken? (wordt niet opgeslagen)", _
                  vbYesNo + vbQuestion, "Kredietoverzicht jaarrekening") <> vbYes Then Exit Sub
        mProef = True
        OverzichtMaken "jaarrekening", Vz(), _
                       CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value), 0, ""
        mProef = False
        Exit Sub
    End If
    Set ws = OverzichtMaken("jaarrekening", best, jaar, 0, "")
    If ws Is Nothing Then Exit Sub
    map = MapStambestand("03 kredietoverzichten")
    If Len(map) > 0 Then
        BladOpslaan ws, map, "31-12-" & jaar & " Kredietoverzicht jaarrekening"
        ws.Activate
        MsgBox "Opgeslagen in:" & vbCrLf & map, vbInformation
    End If
End Sub

' Maakt het overzicht. soort = "raad", "naam" of "jaarrekening"; jaar = het jaar waarop de selectie is gebaseerd.
Private Function OverzichtMaken(ByVal soort As String, ByVal bron As Worksheet, ByVal jaar As Long, ByVal naamKol As Long, ByVal naam As String) As Worksheet
    Dim v As Variant, laatste As Long, r As Long, n As Long, i As Long, k As Long, c As Long
    Dim sel() As Long, sleutel() As String, toek() As Boolean, nSel As Long
    Dim stLetter As String, sj As Long, neem As Boolean, toekomst As Boolean, b As Double
    Dim ws As Worksheet, rij As Long, eersteData As Long, kols As Variant, koppen As Variant, nK As Long
    Dim cBud As Long, cWerk As Long, cRest As Long, cAfs As Long, somAfs As Double, somBud As Double, somWerk As Double, cToel As Long, cMeld As Long, cStat As Long, cEerste As Long, cLaatste As Long
    Dim prog As String, grp As String, jNr As String, vProg As String, vGrp As String, vJ As String
    Dim startProg As Long, startGrp As Long, startJ As Long, labelJ As String, sectie As Long, vSectie As Long
    Dim meldRij As Object, meldNr As Object, toel As String, somRest As Double, somBron As Double, tekst As String
    Dim bladNaam As String, selTekst As String

    On Error GoTo Fout
    Application.ScreenUpdating = False
    ' laatste regel met een volgnummer
    laatste = bron.Cells(bron.Rows.Count, 1).End(xlUp).Row
    Do While laatste > 1 And Not IsNumeric(bron.Cells(laatste, 1).Value)
        laatste = laatste - 1
    Loop
    Do While laatste > 1
        If Not Leeg(bron.Cells(laatste, 1).Value) And IsNumeric(bron.Cells(laatste, 1).Value) Then Exit Do
        laatste = laatste - 1
    Loop
    If laatste < 2 Then
        MsgBox "Geen kredietregels gevonden op blad " & bron.Name & ".", vbExclamation
        GoTo Einde
    End If
    v = bron.Range("A2:AW" & laatste).Value
    n = UBound(v, 1)

    ' 1. Selectie
    ReDim sel(1 To n): ReDim sleutel(1 To n): ReDim toek(1 To n)
    For r = 1 To n
        If Not Leeg(v(r, 1)) And IsNumeric(v(r, 1)) Then
            stLetter = UCase$(Left$(Trim$(CStr(v(r, 8))), 1))
            sj = Val(Mid$(Trim$(CStr(v(r, 8))), 2))
            neem = False: toekomst = False
            If soort = "naam" Then
                If NaamKomtVoor(CStr(v(r, naamKol)), naam) Then
                    ' budgethouders en projectleiders: alleen lopende kredieten (L), ook als ze nog niet gestart zijn
                    If stLetter = "L" And sj <= jaar Then neem = True
                    If stLetter = "L" And sj > jaar Then neem = True: toekomst = True
                End If
            Else
                If stLetter = "L" And sj <= jaar Then neem = True
                If stLetter = "A" And sj = jaar Then neem = True
                ' raad en jaarrekening: alleen t/m het jaar; regels zonder budget en uitgaven t/m het jaar vallen weg
                If neem And Nul(v(r, 26)) = 0 And Nul(v(r, 27)) + Nul(v(r, 28)) + Nul(v(r, 29)) = 0 Then neem = False
                ' raad en jaarrekening: regels met een negatief budget weglaten als het een subsidie of bijdrage is;
                ' opbrengsten activa (toelichting begint met "Opbrengst activa") blijven staan
                If neem And Not IsOpbrengstActiva(v(r, 36)) Then
                    If Nul(v(r, 26)) < 0 Then neem = False
                    For c = 31 To 35
                        If Nul(v(r, c)) < 0 Then neem = False
                    Next c
                End If
            End If
            If neem Then
                nSel = nSel + 1
                sel(nSel) = r
                toek(nSel) = toekomst
                sleutel(nSel) = IIf(toekomst, "1", "0") & Chr$(1) & SorteerDeel(v(r, 2)) & Chr$(1) & LCase$(Trim$(CStr(v(r, 9)))) & Chr$(1) & _
                                SorteerJ(v(r, 10)) & Chr$(1) & Format$(v(r, 1), "000000000")
            End If
        End If
    Next r
    If nSel = 0 Then
        If Not mStil Then MsgBox "Geen kredieten gevonden voor deze selectie.", vbInformation
        GoTo Einde
    End If
    SorteerSleutels sleutel, sel, toek, 1, nSel

    ' 2. Meldingen van de laatste upload (alleen voor de actuele verzamellijst)
    Set meldRij = CreateObject("Scripting.Dictionary")
    Set meldNr = CreateObject("Scripting.Dictionary")
    If soort <> "jaarrekening" And BladBestaat(BL_MELD) Then
        With ThisWorkbook.Worksheets(BL_MELD)
            For r = 5 To .Cells(.Rows.Count, 1).End(xlUp).Row
                If IsNumeric(.Cells(r, 6).Value) And Not Leeg(.Cells(r, 6).Value) Then
                    meldRij(CLng(.Cells(r, 6).Value)) = CStr(.Cells(r, 1).Value)
                ElseIf Not Leeg(.Cells(r, 3).Value) Then
                    meldNr(CStr(.Cells(r, 3).Value)) = CStr(.Cells(r, 1).Value)
                End If
            Next r
        End With
    End If

    ' 3. Kolommen
    If soort = "naam" Then
        kols = Array("Kredietnummer", "Omschrijving investeringskrediet", "Status", "Afschr", "Budget t/m " & jaar, "Totaal werkelijke uitgaven", "Restantkrediet", _
                     "Vrijval (+) / overschrijding (-) bij afsluiten")
    Else
        kols = Array("Volgnr", "Omschrijving investeringskrediet", "Status", "Budget t/m " & jaar, "Totaal werkelijke uitgaven", "Restantkrediet", _
                     "Vrijval (+) / overschrijding (-) bij afsluiten")
    End If
    nK = UBound(kols) + 1
    ReDim koppen(1 To nK + 7)
    For k = 1 To nK
        koppen(k) = kols(k - 1)
    Next k
    For k = 1 To 5
        koppen(nK + k) = "Budget " & (jaar + k)
    Next k
    koppen(nK + 6) = "Toelichting"
    koppen(nK + 7) = "Melding"
    cStat = 3
    cBud = nK - 3: cWerk = nK - 2: cRest = nK - 1: cAfs = nK
    cEerste = cBud: cLaatste = nK + 5
    cToel = nK + 6: cMeld = nK + 7

    ' 4. Blad: een blad voor het laatste overzicht op naam en een voor het laatste kredietoverzicht; de vorige opvraging wordt vervangen
    Select Case soort
        Case "raad": selTekst = "L t/m " & jaar & " plus A" & jaar & "; zonder regels met negatief budget"
        Case "naam": selTekst = RolNaam(naamKol) & ": " & naam & "; alleen lopende kredieten"
        Case Else:   selTekst = "stand 31-12-" & jaar & " (blad " & bron.Name & "): L t/m " & jaar & " plus A" & jaar
    End Select
    bladNaam = IIf(soort = "naam", BL_OVERZICHT_NAAM, BL_OVERZICHT)
    OudeOverzichtbladenOpruimen
    If BladBestaatStil(bladNaam) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(bladNaam).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = NieuwBlad()
    ws.Name = bladNaam
    ws.Tab.Color = RGB(0, 73, 118)
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10

    ' titel en ondertitel (zonder technische regels: die zijn niet voor lezers buiten financien)
    ws.Range("A1").Value = "Kredietoverzicht investeringskredieten"
    With ws.Range("A1").Font
        .Bold = True
        .Size = 16
        .Color = RGB(0, 73, 118)
    End With
    Select Case soort
        Case "jaarrekening": ws.Range("A2").Value = IIf(mProef, "PROEF (huidige stand) - ", "") & "Jaarrekening " & jaar & " - stand 31-12-" & jaar
        Case "raad":         ws.Range("A2").Value = "Stand " & Format$(Date, "dd-mm-yyyy")
        Case Else:           ws.Range("A2").Value = RolNaam(naamKol) & ": " & naam & " - stand " & Format$(Date, "dd-mm-yyyy")
    End Select
    ws.Range("A2").Font.Size = 11
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Rows(1).RowHeight = 26
    ws.Rows(4).RowHeight = 6
    ws.Range("A5").Resize(1, nK + 7).Value = koppen
    With ws.Range("A5").Resize(1, nK + 7)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 73, 118)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(5).RowHeight = 30

    ' 5. Regels met groepskoppen en tussentotalen
    rij = 6: eersteData = 6
    vProg = Chr$(0): vGrp = Chr$(0): vJ = Chr$(0): vSectie = -1
    For i = 1 To nSel
        r = sel(i)
        sectie = IIf(toek(i), 1, 0)
        prog = CStr(CodeNaarWeergave("Tbl_Programma", v(r, 2)))
        grp = Trim$(CStr(v(r, 9)))
        jNr = IIf(IsGetal(v(r, 10)), CStr(v(r, 10)), "")
        ' groep sluiten bij wissel
        If vJ <> Chr$(0) And Len(vJ) > 0 And (jNr <> vJ Or grp <> vGrp Or prog <> vProg Or sectie <> vSectie) Then
            Tussentotaal ws, rij, startJ, "Subtotaal " & labelJ, cEerste, cLaatste, 3
        End If
        If vGrp <> Chr$(0) And (grp <> vGrp Or prog <> vProg Or sectie <> vSectie) Then
            Tussentotaal ws, rij, startGrp, "Totaal " & vGrp, cEerste, cLaatste, 2
        End If
        If vProg <> Chr$(0) And (prog <> vProg Or sectie <> vSectie) Then
            Tussentotaal ws, rij, startProg, "Totaal " & vProg, cEerste, cLaatste, 1
        End If
        If sectie <> vSectie And sectie = 1 Then
            ws.Cells(rij, 1).Value = "Nog niet gestart (start na " & jaar & ")"
            ws.Cells(rij, 1).Font.Bold = True
            ws.Cells(rij, 1).Font.Size = 12
            rij = rij + 1
        End If
        If prog <> vProg Or sectie <> vSectie Then
            ws.Cells(rij, 1).Value = prog
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Font.Bold = True
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Interior.Color = RGB(214, 228, 240)
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Font.Color = RGB(0, 73, 118)
            rij = rij + 1: startProg = rij
            vGrp = Chr$(0)
        End If
        If grp <> vGrp Then
            ws.Cells(rij, 2).Value = grp
            ws.Cells(rij, 2).Font.Bold = True
            ws.Cells(rij, 2).Font.Italic = True
            rij = rij + 1: startGrp = rij
            vJ = Chr$(0)
        End If
        If jNr <> vJ And Len(jNr) > 0 Then
            startJ = rij
            labelJ = Trim$(CStr(v(r, 46)))
            If Len(labelJ) = 0 Then labelJ = Trim$(CStr(v(r, 4)))
        End If
        vProg = prog: vGrp = grp: vJ = jNr: vSectie = sectie

        ' kredietregel
        If soort = "naam" Then
            ws.Cells(rij, 1).NumberFormat = "@"
            ws.Cells(rij, 1).Value = NrTekst(v(r, 3))
        Else
            ws.Cells(rij, 1).Value = v(r, 1)
        End If
        ws.Cells(rij, 2).Value = v(r, 4)
        ws.Cells(rij, 3).Value = v(r, 8)
        If soort = "naam" Then ws.Cells(rij, 4).Value = v(r, 25)
        ws.Cells(rij, cBud).Value = Nul(v(r, 26))
        ws.Cells(rij, cWerk).Value = Nul(v(r, 27)) + Nul(v(r, 28)) + Nul(v(r, 29))
        ' afgesloten krediet: restant is 0, het verschil staat als vrijval/overschrijding bij afsluiten
        If UCase$(Left$(Trim$(CStr(v(r, 8))), 1)) = "A" Then
            ws.Cells(rij, cRest).Value = 0
            ws.Cells(rij, cAfs).Value = Nul(v(r, 26)) - Nul(v(r, 27)) - Nul(v(r, 28)) - Nul(v(r, 29))
        Else
            ws.Cells(rij, cRest).Value = Nul(v(r, 26)) - Nul(v(r, 27)) - Nul(v(r, 28)) - Nul(v(r, 29))
            ws.Cells(rij, cAfs).Value = 0
        End If
        somBud = somBud + Nul(v(r, 26))
        somWerk = somWerk + ws.Cells(rij, cWerk).Value
        somAfs = somAfs + ws.Cells(rij, cAfs).Value
        somRest = somRest + ws.Cells(rij, cRest).Value
        If IsNumeric(v(r, 30)) And Not Leeg(v(r, 30)) Then somBron = somBron + CDbl(v(r, 30)) Else somBron = somBron + ws.Cells(rij, cRest).Value
        For k = 1 To 5
            ws.Cells(rij, nK + k).Value = Nul(v(r, 30 + k))
        Next k
        toel = Trim$(CStr(v(r, 36)))
        If UCase$(Left$(Trim$(CStr(v(r, 8))), 1)) = "A" And InStr(1, toel, "krediet afgesloten", vbTextCompare) = 0 Then
            toel = "Krediet afgesloten." & IIf(Len(toel) > 0, " " & toel, "")
        End If
        ws.Cells(rij, cToel).Value = toel
        If meldRij.Exists(r + 1) Then
            ws.Cells(rij, cMeld).Value = meldRij(r + 1)
        ElseIf meldNr.Exists(NrTekst(v(r, 3))) Then
            ws.Cells(rij, cMeld).Value = meldNr(NrTekst(v(r, 3)))
        End If
        If Len(CStr(ws.Cells(rij, cMeld).Value)) > 0 Then ws.Cells(rij, cMeld).Interior.Color = RGB(255, 150, 150)
        rij = rij + 1
    Next i
    If Len(vJ) > 0 And vJ <> Chr$(0) Then Tussentotaal ws, rij, startJ, "Subtotaal " & labelJ, cEerste, cLaatste, 3
    Tussentotaal ws, rij, startGrp, "Totaal " & vGrp, cEerste, cLaatste, 2
    Tussentotaal ws, rij, startProg, "Totaal " & vProg, cEerste, cLaatste, 1
    rij = rij + 1
    ws.Cells(rij, 1).Value = "Eindtotaal"
    For c = cEerste To cLaatste
        ws.Cells(rij, c).Formula = "=SUBTOTAL(9," & ws.Range(ws.Cells(eersteData, c), ws.Cells(rij - 1, c)).Address(False, False) & ")"
    Next c
    With ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7))
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With

    ' 6. Opmaak
    With ws.Range(ws.Cells(eersteData, cEerste), ws.Cells(rij, cLaatste))
        .NumberFormat = "#,##0;-#,##0;-"
        .Font.Color = RGB(0, 0, 0)
    End With
    ws.Range(ws.Cells(5, cEerste), ws.Cells(rij, cLaatste)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(5, 1), ws.Cells(rij, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(5, cStat), ws.Cells(rij, cStat)).HorizontalAlignment = xlCenter
    If soort = "naam" Then ws.Range(ws.Cells(5, 4), ws.Cells(rij, 4)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(5, cToel), ws.Cells(rij, cMeld)).HorizontalAlignment = xlLeft
    ws.Columns(1).ColumnWidth = IIf(soort = "naam", 16, 9)
    ws.Columns(2).ColumnWidth = 45
    ws.Columns(3).ColumnWidth = 9
    If soort = "naam" Then ws.Columns(4).ColumnWidth = 8
    ws.Range(ws.Columns(cEerste), ws.Columns(cLaatste)).ColumnWidth = 15
    ws.Columns(cToel).ColumnWidth = 80
    ws.Columns(cMeld).ColumnWidth = 28
    ws.Range(ws.Cells(eersteData, 2), ws.Cells(rij, 2)).WrapText = True
    ws.Range(ws.Cells(eersteData, cToel), ws.Cells(rij, cMeld)).WrapText = True
    ws.Range(ws.Cells(eersteData, 1), ws.Cells(rij, nK + 7)).VerticalAlignment = xlTop
    ' standaard verborgen: status (bij A staat "Krediet afgesloten" in de toelichting) en vrijval/overschrijding
    ws.Columns(cStat).Hidden = True
    ws.Columns(cAfs).Hidden = True
    LogoPlaatsen ws, cToel

    ' kolom Melding alleen tonen als er meldingen zijn
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(eersteData, cMeld), ws.Cells(rij, cMeld))) = 0 Then ws.Columns(cMeld).Hidden = True

    ' 7. Aansluitcontrole restantkrediet
    If Round(somRest - somBron, 2) <> 0 Or Round(somBud - somWerk - somRest - somAfs, 2) <> 0 Then
        tekst = "LET OP: aansluiting sluit niet (restant overzicht " & Format$(somRest, "#,##0") & ", verzamellijst " & _
                Format$(somBron, "#,##0") & "). Overzicht NIET definitief."
        ws.Range("A4").Value = tekst
        ws.Range("A4").Font.Bold = True
        ws.Range("A4").Font.Color = RGB(255, 255, 255)
        ws.Range("A4").Interior.Color = RGB(192, 0, 0)
        ws.Rows(4).RowHeight = 15
        MsgBox tekst, vbCritical
    End If

    KnopTerug ws, ws.Range("A3")
    ' raad en jaarrekening: alleen t/m het jaar, dus geen budgetkolommen voor latere jaren
    If soort <> "naam" Then ws.Range(ws.Columns(nK + 1), ws.Columns(nK + 5)).Delete

    ' kolombreedte en rijhoogte op de definitieve kolommen (na verbergen/verwijderen)
    BladOpmaken ws, 5, eersteData, rij

    ' 8. Afdrukken en vastzetten
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("C6").Select
    ActiveWindow.FreezePanes = True
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .PrintTitleRows = "$5:$5"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Pagina &P van &N"
    End With
    On Error GoTo Fout
    BladBeveiligen ws
    TabbladenOrdenen
    ws.Activate
    Set OverzichtMaken = ws
    LogRegel "kredietoverzicht", , , soort, , , ws.Name & " | " & selTekst & " | " & nSel & " kredieten"
Einde:
    Application.ScreenUpdating = True
    Exit Function
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout bij het maken van het overzicht: " & Err.Description, vbCritical
End Function

' Tussentotaalregel met SUBTOTAL over de bedragkolommen. niveau 1 = programma, 2 = investeringsgroep, 3 = samengesteld krediet.
Private Sub Tussentotaal(ByVal ws As Worksheet, ByRef rij As Long, ByVal vanRij As Long, ByVal label As String, _
                         ByVal cEerste As Long, ByVal cLaatste As Long, ByVal niveau As Long)
    Dim c As Long
    If vanRij < 6 Or vanRij > rij - 1 Then Exit Sub
    ws.Cells(rij, 2).Value = label
    For c = cEerste To cLaatste
        ws.Cells(rij, c).Formula = "=SUBTOTAL(9," & ws.Range(ws.Cells(vanRij, c), ws.Cells(rij - 1, c)).Address(False, False) & ")"
    Next c
    With ws.Range(ws.Cells(rij, 2), ws.Cells(rij, cLaatste))
        .Font.Bold = (niveau < 3)
        .Font.Italic = (niveau = 3)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With
    rij = rij + 1
    If niveau < 3 Then rij = rij + 1             ' lege regel na een groep
End Sub

' True als de naam voorkomt in de cel (tijdschrijvers staan met komma's in een cel).
Private Function NaamKomtVoor(ByVal cel As String, ByVal naam As String) As Boolean
    Dim deel As Variant
    For Each deel In Split(Replace(cel, ";", ","), ",")
        If StrComp(Trim$(CStr(deel)), naam, vbTextCompare) = 0 Then NaamKomtVoor = True: Exit Function
    Next deel
End Function

Private Function SorteerDeel(ByVal x As Variant) As String
    If IsGetal(x) Then SorteerDeel = Format$(x, "000000000") Else SorteerDeel = "~" & LCase$(Trim$(CStr(x)))
End Function

' Kolom J: genummerde deelkredieten eerst (op nummer), daarna standaard/stelpost alfabetisch.
Private Function SorteerJ(ByVal x As Variant) As String
    If IsGetal(x) Then SorteerJ = "0" & Format$(x, "000000000") Else SorteerJ = "1" & LCase$(Trim$(CStr(x)))
End Function

' Eenvoudige sortering (insertion sort) op de sleutels; sel en toek schuiven mee.
Private Sub SorteerSleutels(ByRef sl() As String, ByRef sel() As Long, ByRef toek() As Boolean, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, ts As String, tl As Long, tb As Boolean
    For i = lo + 1 To hi
        ts = sl(i): tl = sel(i): tb = toek(i)
        j = i - 1
        Do While j >= lo
            If StrComp(sl(j), ts, vbBinaryCompare) <= 0 Then Exit Do
            sl(j + 1) = sl(j): sel(j + 1) = sel(j): toek(j + 1) = toek(j)
            j = j - 1
        Loop
        sl(j + 1) = ts: sel(j + 1) = tl: toek(j + 1) = tb
    Next i
End Sub

' Bladnaam zonder verboden tekens, maximaal 31 posities, uniek.
Private Function VrijeBladnaam(ByVal naam As String) As String
    Dim t As Variant, basis As String, n As Long, kandidaat As String
    For Each t In Array("\", "/", "?", "*", "[", "]", ":")
        naam = Replace(naam, CStr(t), "-")
    Next t
    basis = Left$(naam, 31)
    kandidaat = basis
    n = 1
    Do While BladBestaatStil(kandidaat)
        n = n + 1
        kandidaat = Left$(basis, 31 - Len(CStr(n)) - 3) & " (" & n & ")"
    Loop
    VrijeBladnaam = kandidaat
End Function

Private Function BladBestaatStil(ByVal naam As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, naam, vbTextCompare) = 0 Then BladBestaatStil = True: Exit Function
    Next ws
End Function

' Blok 3 op het Dashboard: keuze rol en naam en de knoppen voor de overzichten.
Public Sub UitdraaiBlokPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, naam As Variant, i As Long, titels As Variant, acties As Variant, namen As Variant
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("B57").Value = "3. Uitdraaien"
    ws.Range("B57").Font.Bold = True
    ws.Range("B58").Value = "Overzicht op naam: rol"
    ws.Range("B59").Value = "Overzicht op naam: naam"
    With ws.Range("C58")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Budgetbeheerder,Projectleider,Tijdschrijver"
        If Leeg(.Value) Then .Value = "Budgetbeheerder"
        .Interior.Color = RGB(255, 255, 204)
    End With
    With ws.Range("C59")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(""Lst_""&$C$58)"
        .Interior.Color = RGB(255, 255, 204)
    End With
    ws.Range("B64").Value = "4. Financieel adviseur"
    ws.Range("B64").Font.Bold = True
    On Error Resume Next
    For Each naam In Array("btnOverzichtRaad", "btnOverzichtNaam", "btnOverzichtJaarrekening", "btnAlleBudgethouders", "btnTijdschrijven", "btnUitvraag", "btnActualisatie")
        ws.Shapes(CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    ws.Rows(61).RowHeight = 26
    Set b = ws.Buttons.Add(ws.Range("B61").Left, ws.Range("B61").Top + 2, 165, 22)
    b.Name = "btnOverzichtNaam"
    b.Caption = "Kredietoverzicht op naam"
    b.OnAction = "KredietoverzichtOpNaam"
    ws.Rows(65).RowHeight = 26
    ws.Rows(66).RowHeight = 26
    namen = Array("btnOverzichtRaad", "btnOverzichtJaarrekening", "btnUitvraag", "btnAlleBudgethouders", "btnTijdschrijven", "btnActualisatie")
    titels = Array("Kredietoverzicht raad", "Kredietoverzicht jaarrekening", "Uitvraag kadernota/begroting", _
                   "Overzichten alle budgethouders", "Overzicht tijdschrijven", "Actualisatie verwerken (C59)")
    acties = Array("KredietoverzichtMaken", "KredietoverzichtJaarrekening", "UitvraagBegroting", _
                   "KredietoverzichtAlleBudgethouders", "OverzichtTijdschrijven", "ActualisatieVerwerken")
    For i = 0 To 5
        Set b = ws.Buttons.Add(ws.Range("B65").Left + (i Mod 3) * 170, ws.Range("B65").Offset(i \ 3, 0).Top + 2, 165, 22)
        b.Name = namen(i)
        b.Caption = titels(i)
        b.OnAction = acties(i)
    Next i
End Sub

' Koppen van M en N in de verzamellijst (korte, duidelijke namen) en kolom A (volgnummer) gecentreerd.
Private Sub KoppenStartEindjaar()
    With Vz()
        .Range("M1").Value = "Startjaar uitgaven"
        .Range("N1").Value = "Eindjaar uitgaven"
        .Range("A1:A" & CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)).HorizontalAlignment = xlCenter   ' volgnummer gecentreerd
    End With
End Sub

' Eenmalig: maakt de toelichting (AJ) leeg bij alle regels met investeringsgroep 'voertuigen'. Oude tekst gaat naar het Logboek.
Public Sub ToelichtingVoertuigenWissen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, n As Long
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "U heeft geen recht om dit te doen.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        If LCase$(Trim$(CStr(wv.Cells(r, 9).Value))) = "voertuigen" And Not Leeg(wv.Cells(r, 36).Value) Then n = n + 1
    Next r
    If n = 0 Then
        MsgBox "Er zijn geen voertuigen met een toelichting.", vbInformation
        Exit Sub
    End If
    If MsgBox("Bij " & n & " voertuigen de toelichting leegmaken? De oude tekst gaat naar het Logboek.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Toelichting voertuigen") <> vbYes Then Exit Sub
    For r = 2 To laatste
        If LCase$(Trim$(CStr(wv.Cells(r, 9).Value))) = "voertuigen" And Not Leeg(wv.Cells(r, 36).Value) Then
            LogRegel "toelichting voertuig gewist", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Toelichting", wv.Cells(r, 36).Value, "", "rij " & r
            wv.Cells(r, 36).ClearContents
        End If
    Next r
    MsgBox "Toelichting gewist bij " & n & " voertuigen.", vbInformation
End Sub

' ----------------------------------------------------------------------
' Opslaan in de SharePoint-mappen onder 'Stambestand kredieten'
'   01 Applicatiebeheer Tijdschrijven : overzicht voor de beheerder tijdschrijven
'   02 Budgetbeheerders en projectleiders : kredietoverzichten op naam
'   03 kredietoverzichten : kredietoverzicht per 31-12
' De hoofdmap is de map boven 'export coda' (uit Par_PadUitgavenbestand), anders eenmalig kiezen.
' ----------------------------------------------------------------------

' Geeft het pad van een submap van 'Stambestand kredieten' (met \ aan het eind); maakt de submap zo nodig aan. Leeg = afgebroken.
Private Function MapStambestand(ByVal submap As String) As String
    Dim basis As String, pad As String
    On Error Resume Next
    basis = Evaluate(ThisWorkbook.Names("Par_MapStambestand").RefersTo)
    On Error GoTo 0
    If Not MapBestaat(basis) Then
        pad = CStr(ThisWorkbook.Names("Par_PadUitgavenbestand").RefersToRange.Value)
        If InStrRev(pad, "\") > 0 Then
            pad = Left$(pad, InStrRev(pad, "\") - 1)          ' map export coda
            If InStrRev(pad, "\") > 0 Then basis = Left$(pad, InStrRev(pad, "\"))
        End If
    End If
    If Not MapBestaat(basis) Then
        With Application.FileDialog(4)
            .Title = "Kies de map 'Stambestand kredieten'"
            If .Show <> -1 Then Exit Function
            basis = .SelectedItems(1)
        End With
    End If
    If Right$(basis, 1) <> "\" Then basis = basis & "\"
    ThisWorkbook.Names.Add Name:="Par_MapStambestand", RefersTo:="=""" & basis & """"
    pad = basis & submap
    If Not MapBestaat(pad) Then MkDir pad
    MapStambestand = pad & "\"
End Function

' Bewaart een kopie van het blad als los Excelbestand. Bestaat het bestand al, dan wordt het overschreven.
Private Function BladOpslaan(ByVal ws As Worksheet, ByVal map As String, ByVal bestand As String) As Boolean
    Dim wb As Workbook, t As Variant
    For Each t In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
        bestand = Replace(bestand, CStr(t), "-")
    Next t
    ws.Copy
    Set wb = ActiveWorkbook
    ' knoppen (Terug naar dashboard) horen niet in het losse bestand
    On Error Resume Next
    wb.Worksheets(1).Unprotect WACHTWOORD
    wb.Worksheets(1).Buttons.Delete
    If ws.ProtectContents Then wb.Worksheets(1).Protect Password:=WACHTWOORD, AllowFiltering:=True
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=map & bestand & ".xlsx", FileFormat:=51
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    BladOpslaan = True
End Function

' Maakt per budgetbeheerder (kolom O) een kredietoverzicht en bewaart het in map 02, in een submap per persoon.
Public Sub KredietoverzichtAlleBudgethouders()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, namen As Object, nm As Variant, ws As Worksheet
    Dim map As String, lj As Long, n As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Dit wordt gedaan door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    Set namen = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(wv.Cells(r, 1).Value) And Not Leeg(wv.Cells(r, 15).Value) Then namen(Trim$(CStr(wv.Cells(r, 15).Value))) = True
    Next r
    If MsgBox("Voor " & namen.Count & " budgetbeheerders een kredietoverzicht maken en opslaan in map '02 Budgetbeheerders en projectleiders' (submap per persoon)?", _
              vbYesNo + vbQuestion, "Kredietoverzichten budgethouders") <> vbYes Then Exit Sub
    map = MapStambestand("02 Budgetbeheerders en projectleiders")
    If Len(map) = 0 Then Exit Sub
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    mStil = True
    For Each nm In namen.Keys
        Set ws = OverzichtMaken("naam", wv, lj, 15, CStr(nm))
        If Not ws Is Nothing Then
            BladOpslaan ws, SubMap(map, CStr(nm)), "Kredietoverzicht " & nm & " " & Format$(Date, "dd-mm-yyyy")
            n = n + 1
        End If
    Next nm
    mStil = False
    ThisWorkbook.Worksheets("Dashboard").Activate
    MsgBox n & " kredietoverzichten opgeslagen in:" & vbCrLf & map, vbInformation
End Sub

' Overzicht voor de beheerder tijdschrijven: alleen kredieten waarvoor actie nodig is (nieuw, gewijzigd of afgesloten),
' kolommen Kredietnummer, Omschrijving, Status, Tijdschrijvers, Actie, Toegevoegd op. De projectleider is tevens
' tijdschrijver en staat daarom bij de tijdschrijvers. Vergeleken wordt met de stand van de vorige keer, die op het
' verborgen blad 'Stand tijdschrijven' staat (alle lopende kredieten). Geen stand = alles nieuw (beginstand).
' Wordt als los bestand bewaard in map 01 en als blad 'Laatste overzicht tijdschrijven'.
Public Sub OverzichtTijdschrijven()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const BL_STAND As String = "Stand tijdschrijven"
    Dim wv As Worksheet, ws As Worksheet, wst As Worksheet, map As String
    Dim laatste As Long, r As Long, nr As String, i As Long, rij As Long, actie As String, st As String
    Dim huidig As Object, vorig As Object, rec As Variant, oudRec As Variant, sleutels As Variant, naamTs As String
    Dim nNieuw As Long, nGewijzigd As Long, nAfgesloten As Long, beginstand As Boolean
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Dit overzicht wordt gemaakt door de financieel adviseur.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Fout
    map = MapStambestand("01 Applicatiebeheer Tijdschrijven")
    If Len(map) = 0 Then Exit Sub
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)

    ' huidige stand per kredietnummer (alle lopende kredieten): omschrijving, status, tijdschrijvers incl. projectleider
    Set huidig = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        nr = NrTekst(wv.Cells(r, 3).Value)
        st = Trim$(CStr(wv.Cells(r, 8).Value))
        If Len(nr) > 0 And nr <> "7999998" And UCase$(Left$(st, 1)) = "L" Then
            naamTs = NamenSamenvoegen(CStr(wv.Cells(r, 16).Value), CStr(wv.Cells(r, 17).Value))
            If Not huidig.Exists(nr) Then
                huidig(nr) = Array(Trim$(CStr(wv.Cells(r, 4).Value)), st, naamTs)
            Else
                rec = huidig(nr)
                rec(2) = NamenSamenvoegen(CStr(rec(2)), naamTs)
                huidig(nr) = rec
            End If
        End If
    Next r

    ' stand van de vorige keer: kredietnummer, omschrijving, status, tijdschrijvers, toegevoegd op
    Set vorig = CreateObject("Scripting.Dictionary")
    beginstand = Not BladBestaatStil(BL_STAND)
    If Not beginstand Then
        Set wst = ThisWorkbook.Worksheets(BL_STAND)
        For r = 3 To wst.Cells(wst.Rows.Count, 1).End(xlUp).Row
            vorig(CStr(wst.Cells(r, 1).Value)) = Array(CStr(wst.Cells(r, 2).Value), CStr(wst.Cells(r, 3).Value), _
                                                       CStr(wst.Cells(r, 4).Value), wst.Cells(r, 5).Value)
        Next r
    End If

    ' blad 'Laatste overzicht tijdschrijven' (vorige opvraging wordt vervangen)
    If BladBestaatStil(BL_TIJD) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(BL_TIJD).Unprotect WACHTWOORD
        ThisWorkbook.Worksheets(BL_TIJD).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = NieuwBlad()
    ws.Name = BL_TIJD
    ws.Tab.Color = RGB(0, 73, 118)
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "Kredieten voor tijdschrijven"
    With ws.Range("A1").Font
        .Bold = True
        .Size = 16
        .Color = RGB(0, 73, 118)
    End With
    ws.Rows(1).RowHeight = 26
    ws.Range("A2").Font.Size = 11
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Rows(3).RowHeight = 22
    ws.Range("A4:F4").Value = Array("Kredietnummer", "Omschrijving", "Status", "Tijdschrijvers", "Actie", "Toegevoegd op")
    With ws.Range("A4:F4")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 73, 118)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Columns("A").NumberFormat = "@"
    rij = 5

    ' nieuw en gewijzigd
    sleutels = huidig.Keys
    For i = 0 To huidig.Count - 1
        nr = sleutels(i)
        rec = huidig(nr)
        actie = ""
        If Not vorig.Exists(nr) Then
            actie = "nieuw"
        Else
            oudRec = vorig(nr)
            If oudRec(0) <> rec(0) Or NamenSamenvoegen("", CStr(oudRec(2))) <> rec(2) Then actie = "gewijzigd"
        End If
        If Len(actie) > 0 Then
            ws.Cells(rij, 1).Value = nr
            ws.Cells(rij, 2).Value = rec(0)
            ws.Cells(rij, 3).Value = rec(1)
            ws.Cells(rij, 4).Value = rec(2)
            ws.Cells(rij, 5).Value = actie
            If actie = "nieuw" Then ws.Cells(rij, 6).Value = Date Else ws.Cells(rij, 6).Value = oudRec(3)
            If actie = "nieuw" Then nNieuw = nNieuw + 1 Else nGewijzigd = nGewijzigd + 1
            rij = rij + 1
        End If
    Next i
    ' afgesloten: stond de vorige keer als lopend krediet in de stand, nu niet meer
    sleutels = vorig.Keys
    For i = 0 To vorig.Count - 1
        nr = sleutels(i)
        If Not huidig.Exists(nr) Then
            oudRec = vorig(nr)
            ws.Cells(rij, 1).Value = nr
            ws.Cells(rij, 2).Value = oudRec(0)
            ws.Cells(rij, 3).Value = HuidigeStatus(wv, laatste, nr)
            ws.Cells(rij, 4).Value = oudRec(2)
            ws.Cells(rij, 5).Value = "afgesloten"
            ws.Cells(rij, 6).Value = oudRec(3)
            nAfgesloten = nAfgesloten + 1
            rij = rij + 1
        End If
    Next i
    If rij = 5 Then
        ws.Cells(5, 1).Value = "Geen wijzigingen sinds de vorige keer."
        rij = 6
    End If
    ws.Range("A2").Value = IIf(beginstand, "Beginstand: alle lopende kredieten als nieuw", "Alleen kredieten waarvoor actie nodig is") & _
                           " - stand " & Format$(Date, "dd-mm-yyyy") & ": " & nNieuw & " nieuw, " & nGewijzigd & " gewijzigd, " & _
                           nAfgesloten & " afgesloten"

    ' opmaak
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 45
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 50
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 15
    ws.Range("A5:F" & rij).VerticalAlignment = xlTop
    ws.Range("B5:B" & rij & ",D5:D" & rij).WrapText = True
    ws.Range("A5:B" & rij & ",D5:D" & rij).HorizontalAlignment = xlLeft
    ws.Range("C4:C" & rij & ",E4:F" & rij).HorizontalAlignment = xlCenter
    ws.Range("E5:E" & rij).Font.Bold = True
    ws.Range("F5:F" & rij).NumberFormat = "dd-mm-yyyy"
    BladOpmaken ws, 4, 5, rij
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
    KnopTerug ws, ws.Range("A3")
    BladOpslaan ws, map, "Tijdschrijven kredieten " & Format$(Date, "yyyy-mm-dd")

    ' nieuwe stand bewaren (alle lopende kredieten), met de datum 'toegevoegd op'
    If beginstand Then
        Set wst = NieuwBlad()
        wst.Name = BL_STAND
    Else
        Set wst = ThisWorkbook.Worksheets(BL_STAND)
        wst.Unprotect WACHTWOORD
    End If
    wst.Cells.Clear
    wst.Range("A1").Value = "Stand van de lopende kredieten bij het laatste overzicht tijdschrijven. Niet wijzigen."
    wst.Range("A2:E2").Value = Array("Kredietnummer", "Omschrijving", "Status", "Tijdschrijvers", "Toegevoegd op")
    wst.Columns("A").NumberFormat = "@"
    sleutels = huidig.Keys
    For i = 0 To huidig.Count - 1
        nr = sleutels(i)
        rec = huidig(nr)
        wst.Cells(3 + i, 1).Value = nr
        wst.Cells(3 + i, 2).Value = rec(0)
        wst.Cells(3 + i, 3).Value = rec(1)
        wst.Cells(3 + i, 4).Value = rec(2)
        If vorig.Exists(nr) Then
            oudRec = vorig(nr)
            wst.Cells(3 + i, 5).Value = oudRec(3)
        Else
            wst.Cells(3 + i, 5).Value = Date
        End If
    Next i
    wst.Visible = xlSheetHidden

    TabbladenOrdenen
    ws.Activate
    LogRegel "overzicht tijdschrijven", , , "", , , nNieuw & " nieuw, " & nGewijzigd & " gewijzigd, " & nAfgesloten & " afgesloten"
    MsgBox "Overzicht tijdschrijven opgeslagen in:" & vbCrLf & map & vbCrLf & vbCrLf & _
           nNieuw & " nieuw, " & nGewijzigd & " gewijzigd, " & nAfgesloten & " afgesloten.", vbInformation
    Exit Sub
Fout:
    Application.DisplayAlerts = True
    MsgBox "Fout bij het overzicht tijdschrijven: " & Err.Description, vbCritical
End Sub

' Voegt namen (gescheiden door komma's) samen, zonder dubbele, op alfabet.
Private Function NamenSamenvoegen(ByVal a As String, ByVal b As String) As String
    Dim lijst As Object, deel As Variant, k As Variant, i As Long, j As Long, arr() As String, t As String
    Set lijst = CreateObject("Scripting.Dictionary")
    lijst.CompareMode = vbTextCompare
    For Each deel In Split(Replace(a & "," & b, ";", ","), ",")
        If Len(Trim$(CStr(deel))) > 0 Then lijst(Trim$(CStr(deel))) = True
    Next deel
    If lijst.Count = 0 Then Exit Function
    ReDim arr(0 To lijst.Count - 1)
    For Each k In lijst.Keys
        arr(i) = CStr(k): i = i + 1
    Next k
    For i = 1 To UBound(arr)
        t = arr(i): j = i - 1
        Do While j >= 0
            If StrComp(arr(j), t, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = t
    Next i
    NamenSamenvoegen = Join(arr, ", ")
End Function

Private Function HuidigeStatus(ByVal wv As Worksheet, ByVal laatste As Long, ByVal nr As String) As String
    Dim r As Long
    For r = 2 To laatste
        If NrTekst(wv.Cells(r, 3).Value) = nr Then HuidigeStatus = CStr(wv.Cells(r, 8).Value): Exit Function
    Next r
    HuidigeStatus = "niet meer in lijst"
End Function

' Eenmalig: zet in "Nog uit te geven" (AD) overal de formule die bij een afgesloten krediet (A) 0 geeft.
' Vervangt ook vaste nullen die met de hand bij afgesloten kredieten zijn gezet. SUBTOTAL-regels blijven ongemoeid.
Public Sub RestantFormuleBijwerken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, maxRij As Long, r As Long, n As Long, voor As Variant, na As Variant, k As Long, calcOud As Long
    Dim f As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    If MsgBox("In kolom 'Nog uit te geven' overal de formule zetten: bij een afgesloten krediet (A) is het restant 0, " & _
              "anders budget min werkelijke uitgaven. Doorgaan?", vbYesNo + vbQuestion, "Formule restant") <> vbYes Then Exit Sub
    Set wv = Vz()
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    voor = MeetStap8(wv, CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value))
    calcOud = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For r = 2 To maxRij
        If InStr(1, wv.Cells(r, 30).Formula, "SUBTOTAL", vbTextCompare) = 0 Then
            f = "=IF($A" & r & "="""","""",IF(LEFT($H" & r & ",1)=""A"",0,Z" & r & "-AA" & r & "-AB" & r & "-AC" & r & "))"
            If wv.Cells(r, 30).Formula <> f Then wv.Cells(r, 30).Formula = f: n = n + 1
        End If
    Next r
    Application.Calculation = calcOud
    Application.Calculate
    na = MeetStap8(wv, CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value))
    Application.ScreenUpdating = True
    For k = 0 To 9
        If Round(voor(k) - na(k), 2) <> 0 Then
            MsgBox "Let op: de nameting wijkt af in meting " & k & ". Controleer de verzamellijst.", vbCritical
            Exit For
        End If
    Next k
    LogRegel "formule restant bijwerken", , , "Nog uit te geven", , , n & " cellen aangepast"
    MsgBox n & " cellen in 'Nog uit te geven' aangepast.", vbInformation
End Sub

' Submap per persoon (wordt aangemaakt als die nog niet bestaat).
Private Function SubMap(ByVal map As String, ByVal naam As String) As String
    Dim t As Variant
    For Each t In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
        naam = Replace(naam, CStr(t), "-")
    Next t
    Dim pad As String
    pad = map & Trim$(naam)
    If Not MapBestaat(pad) Then MkDir pad
    SubMap = pad & "\"
End Function

Private Function RolNaam(ByVal kol As Long) As String
    Select Case kol
        Case 15: RolNaam = "Budgetbeheerder"
        Case 16: RolNaam = "Projectleider"
        Case 17: RolNaam = "Tijdschrijver"
    End Select
End Function

' Zet het logo van de gemeente rechtsboven op het overzicht. Het logo is een bestand 'logo*.png' (of .jpg)
' in de map 'Stambestand kredieten'. Geen logo gevonden: dan zonder logo.
Private Sub LogoPlaatsen(ByVal ws As Worksheet, ByVal kolRechts As Long)
    Dim basis As String, f As String, shp As Object
    On Error GoTo Klaar
    basis = Evaluate(ThisWorkbook.Names("Par_MapStambestand").RefersTo)
    If Not MapBestaat(basis) Then Exit Sub
    f = Dir(basis & "logo*.png")
    If Len(f) = 0 Then f = Dir(basis & "logo*.jp*g")
    If Len(f) = 0 Then Exit Sub
    Set shp = ws.Shapes.AddPicture(basis & f, msoFalse, msoTrue, 0, 2, -1, -1)
    shp.LockAspectRatio = msoTrue
    shp.Height = 44
    shp.Left = ws.Cells(1, kolRechts).Left + ws.Cells(1, kolRechts).Width - shp.Width
    shp.Top = 2
    shp.Placement = xlMove
Klaar:
End Sub

' ======================================================================
' Bladbeveiliging
' ======================================================================
' Alle mutaties lopen via het Dashboard. De bladen staan op slot; macro's mogen wel schrijven (UserInterfaceOnly).
' Die instelling wordt niet in het bestand bewaard, daarom zet Auto_Open de beveiliging bij elk openen opnieuw aan.
' Open blijven: invoercellen op het Dashboard en de gele invoerkolommen van de tabellen op Parameters.
' Filteren mag op de verzamellijst, sorteren niet.

Public Sub BeveiligingAan()
    Dim naam As Variant
    mHerbeveiligingGepland = False
    On Error Resume Next
    For Each naam In Array(BL_VZ, "Logboek", "Vorige upload", "Stand tijdschrijven", BL_MUTREG, BL_UITVRAAG, "Parameters", "Dashboard", "Controle", _
                           BL_MELD, BL_OVERZICHT, BL_OVERZICHT_NAAM)
        If BladBestaatStil(CStr(naam)) Then BladBeveiligen ThisWorkbook.Worksheets(CStr(naam))
    Next naam
    TabbladenOrdenen
    ' nooit eindigen op het Controle-blad als de macro daar niet begon
    On Error Resume Next
    If Not mActiefVoorMacro Is Nothing Then
        If ActiveSheet.Name = "Controle" And mActiefVoorMacro.Name <> "Controle" Then mActiefVoorMacro.Activate
    End If
    Set mActiefVoorMacro = Nothing
End Sub

' Alleen voor de financieel adviseur: beveiliging uit om iets met de hand te herstellen. Gaat bij het volgende openen weer aan.
Public Sub BeveiligingUit()
    Dim ws As Worksheet
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen de beveiliging uitzetten.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Beveiliging van alle bladen uitzetten? Wijzigingen met de hand komen niet in het Logboek." & vbCrLf & _
              "Bij het volgende openen gaat de beveiliging vanzelf weer aan (of draai BeveiligingAan).", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Beveiliging uit") <> vbYes Then Exit Sub
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect WACHTWOORD
    Next ws
    On Error GoTo 0
    LogRegel "beveiliging uit", , , "", , , "met de hand uitgezet"
    MsgBox "Beveiliging staat uit.", vbInformation
End Sub

Private Sub BladBeveiligen(ByVal ws As Worksheet)
    Dim lo As ListObject, kol As Variant, j As Long, k As Long, laatste As Long
    On Error Resume Next
    ws.Unprotect WACHTWOORD
    Select Case ws.Name
        Case "Parameters"
            ws.Cells.Locked = True
            For Each kol In Array(Array("Tbl_Rollen", "Naam"), Array("Tbl_Rollen", "Windows-naam"), Array("Tbl_Rollen", "Rol"), _
                                  Array("Tbl_Rollen", "Toevoegen"), Array("Tbl_Rollen", "Wijzigen"), Array("Tbl_Rollen", "Bevestigen"), _
                                  Array("Tbl_Programma", "Code"), Array("Tbl_Programma", "Omschrijving"), _
                                  Array("Tbl_Taakveld", "Taakveld"), Array("Tbl_Taakveld", "Omschrijving"), _
                                  Array("Tbl_Balanscode", "Code"), Array("Tbl_Balanscode", "Omschrijving"), Array("Tbl_BijdrageBron", "Bron"))
                Set lo = Nothing
                Set lo = ws.ListObjects(CStr(kol(0)))
                If Not lo Is Nothing Then lo.ListColumns(CStr(kol(1))).DataBodyRange.Locked = False
            Next kol
        Case "Dashboard"
            ws.Cells.Locked = True
            ws.Range("Kaart_Invoer").Locked = False
            For j = 0 To 5
                For k = 1 To 2
                    ws.Cells(BijdrageRij(j, k), 2).Locked = False      ' bron van een bijdrage (kiezen of typen)
                Next k
            Next j
            ws.Range("C58:C59").Locked = False
            ws.Range("G2").Locked = False                         ' tabbladkiezer
        Case BL_VZ
            If Not ws.AutoFilterMode Then
                laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
                ws.Range("A1:AW" & laatste).AutoFilter
            End If
    End Select
    ws.Protect Password:=WACHTWOORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

' Aan het begin van elke macro die iets wijzigt: beveiliging uit en direct na afloop van de macro weer aan
' (Application.OnTime start BeveiligingAan zodra de macro klaar is, ook na een fout of een tussentijds stoppen).
Private Sub WerkStart()
    If mActiefVoorMacro Is Nothing Then Set mActiefVoorMacro = ActiveSheet
    BeveiligingTijdelijkUit
    If mHerbeveiligingGepland Then Exit Sub       ' maar een keer inplannen per macro
    mHerbeveiligingGepland = True
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!BeveiligingAan"
End Sub

' Zet de beveiliging van alle bladen tijdelijk uit (voor macro's die veel schrijven). Daarna altijd BeveiligingAan.
Private Sub BeveiligingTijdelijkUit()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect WACHTWOORD
    Next ws
End Sub

' Haalt bladen weg die eerdere versies van de overzichtmacro hebben gemaakt (een blad per opvraging).
' Alleen bladen met een naam volgens dat patroon EN de titel van het overzicht in A1; andere bladen blijven staan.
' Draait bij openen en bij elk nieuw overzicht; ook met de hand te starten.
Public Sub OudeOverzichtbladenOpruimen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim i As Long, ws As Worksheet, nm As String, titel As String
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        nm = ws.Name
        titel = CStr(ws.Range("A1").Value)
        ' naam begint met "Krediet" of "Jaarrekening" (afgekapt op 31 tekens kan de datum half zijn), maar nooit de vaste bladen
        If (nm Like "Krediet*" Or nm Like "Jaarrekening*") And nm <> "Kredietoverzicht" And nm <> BL_OVERZICHT _
           And nm <> BL_OVERZICHT_NAAM And (titel Like "Kredietoverzicht investeringskredieten*" Or titel Like "Kredietoverzicht jaarrekening*") Then
            ws.Unprotect WACHTWOORD
            ws.Delete
        End If
    Next i
    Application.DisplayAlerts = True
End Sub

' Kolombreedte en rijhoogte zo dat niets onlogisch wordt afgekapt (AutoFit is hiervoor niet betrouwbaar):
' - kopregel: kolom minstens zo breed als het langste woord, rijhoogte naar het aantal regels;
' - bedragkolommen: breder zolang er #### wordt getoond;
' - regels: rijhoogte naar het aantal tekstregels in de cellen met terugloop (alleen zichtbare kolommen).
Private Sub BladOpmaken(ByVal ws As Worksheet, ByVal kopRij As Long, ByVal van As Long, ByVal tot As Long)
    Dim c As Long, r As Long, lk As Long, w As Variant, nodig As Double, n As Long, regels As Long, cel As Range, stap As Long
    lk = ws.Cells(kopRij, ws.Columns.Count).End(xlToLeft).Column
    Application.Calculate
    ' 1. kopregel: langste woord moet passen
    For c = 1 To lk
        If Not ws.Columns(c).Hidden Then
            For Each w In Split(Replace(CStr(ws.Cells(kopRij, c).Value), "/", "/ "), " ")
                nodig = (Len(CStr(w)) * 7.4 + 9) / 6
                If ws.Columns(c).ColumnWidth < nodig Then ws.Columns(c).ColumnWidth = nodig
            Next w
        End If
    Next c
    ' 2. bedragen: geen ####
    For c = 1 To lk
        If Not ws.Columns(c).Hidden Then
            For stap = 1 To 6
                n = 0
                For r = van To tot
                    Set cel = ws.Cells(r, c)
                    If Not IsEmpty(cel.Value) Then
                        If IsNumeric(cel.Value) And Left$(cel.Text, 1) = "#" Then n = n + 1: Exit For
                    End If
                Next r
                If n = 0 Then Exit For
                ws.Columns(c).ColumnWidth = ws.Columns(c).ColumnWidth + 2
            Next stap
        End If
    Next c
    ' 3. rijhoogte kopregel en regels
    For r = kopRij To tot
        If r = kopRij Or r >= van Then
            regels = 1
            For c = 1 To lk
                If Not ws.Columns(c).Hidden Then
                    Set cel = ws.Cells(r, c)
                    If cel.WrapText And Len(CStr(cel.Value)) > 0 And Not IsNumeric(cel.Value) Then
                        n = AantalRegels(CStr(cel.Value), ws.Columns(c).ColumnWidth, cel.Font.Bold)
                        If n > regels Then regels = n
                    End If
                End If
            Next c
            ws.Rows(r).RowHeight = Application.Min(409, regels * 12.75 + IIf(r = kopRij, 8, 3))
        End If
    Next r
End Sub

' Aantal tekstregels als de tekst met terugloop in een kolom van deze breedte staat (Arial 10).
Private Function AantalRegels(ByVal tekst As String, ByVal breedte As Double, ByVal vet As Boolean) As Long
    Dim ruimte As Double, teken As Double, regel As Double, w As Variant, lw As Double, n As Long, deel As Variant
    ruimte = breedte * 6 + 5 - 6                 ' pixels in de kolom, min wat marge
    If ruimte < 30 Then ruimte = 30                ' smalle of verborgen kolom: geen eindeloze lus
    teken = IIf(vet, 7.4, 6.8)                   ' gemiddelde tekenbreedte Arial 10 in pixels
    n = 0
    For Each deel In Split(Replace(tekst, vbCr, ""), vbLf)
        n = n + 1: regel = 0
        For Each w In Split(CStr(deel), " ")
            lw = Len(CStr(w)) * teken
            If regel = 0 Then
                regel = lw
            ElseIf regel + teken + lw <= ruimte Then
                regel = regel + teken + lw
            Else
                n = n + 1: regel = lw
            End If
            Do While regel > ruimte                ' woord langer dan de kolom
                n = n + 1: regel = regel - ruimte
            Loop
        Next w
    Next deel
    If n < 1 Then n = 1
    AantalRegels = n
End Function

' ======================================================================
' Uitvraag kadernota / begroting (financieel adviseur)
' ======================================================================
' Overzicht waarmee teamleiders de investeringen actualiseren. Kolommen uit de verzamellijst:
' C, D, O, S, Y, AD, AE t/m AI en AJ (toelichting leeg, in te vullen). Alleen lopende kredieten (L).
' Een blad 'Laatste uitvraag begroting' met alle budgetbeheerders, plus per budgetbeheerder een bestand
' in map 02 (submap per persoon).
Public Sub UitvraagBegroting()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, ws As Worksheet, map As String, namen As Object, nm As Variant, laatste As Long, r As Long, n As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "De uitvraag wordt gemaakt door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    Set namen = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(wv.Cells(r, 1).Value) And UCase$(Left$(Trim$(CStr(wv.Cells(r, 8).Value)), 1)) = "L" _
           And Not Leeg(wv.Cells(r, 15).Value) Then namen(Trim$(CStr(wv.Cells(r, 15).Value))) = True
    Next r
    If MsgBox("Uitvraag kadernota/begroting maken voor " & namen.Count & " budgetbeheerders?" & vbCrLf & _
              "Per budgetbeheerder komt een bestand in map '02 Budgetbeheerders en projectleiders'.", _
              vbYesNo + vbQuestion, "Uitvraag kadernota/begroting") <> vbYes Then Exit Sub
    map = MapStambestand("02 Budgetbeheerders en projectleiders")
    Application.ScreenUpdating = False
    If Len(map) > 0 Then
        For Each nm In namen.Keys
            Set ws = UitvraagBlad(CStr(nm))
            If Not ws Is Nothing Then
                BladOpslaan ws, SubMap(map, CStr(nm)), "Uitvraag kadernota-begroting " & nm & " " & Format$(Date, "dd-mm-yyyy")
                n = n + 1
            End If
        Next nm
    End If
    Set ws = UitvraagBlad("")                         ' als laatste: het blad met iedereen blijft in de werkmap staan
    Application.ScreenUpdating = True
    If Not ws Is Nothing Then ws.Activate
    LogRegel "uitvraag kadernota/begroting", , , "", , , n & " bestanden"
    MsgBox "Uitvraag gemaakt. " & n & " bestanden opgeslagen in:" & vbCrLf & map, vbInformation
End Sub

' Bouwt het blad 'Laatste uitvraag begroting' voor een budgetbeheerder (of iedereen als naam leeg is).
Private Function UitvraagBlad(ByVal naam As String) As Worksheet
    Dim wv As Worksheet, ws As Worksheet, v As Variant, laatste As Long, n As Long, r As Long, i As Long, k As Long, c As Long
    Dim sel() As Long, sl() As String, tk() As Boolean, nSel As Long, lj As Long, bronKol As Variant, kop As Variant
    Dim rij As Long, eersteData As Long, vBh As String, bh As String, startBh As Long, nK As Long, tv As String

    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    v = wv.Range("A2:AW" & laatste).Value
    n = UBound(v, 1)
    ReDim sel(1 To n): ReDim sl(1 To n): ReDim tk(1 To n)
    For r = 1 To n
        If Not Leeg(v(r, 1)) And UCase$(Left$(Trim$(CStr(v(r, 8))), 1)) = "L" Then
            If Len(naam) = 0 Or StrComp(Trim$(CStr(v(r, 15))), naam, vbTextCompare) = 0 Then
                nSel = nSel + 1
                sel(nSel) = r
                sl(nSel) = LCase$(Trim$(CStr(v(r, 15)))) & Chr$(1) & SorteerDeel(v(r, 2)) & Chr$(1) & NrTekst(v(r, 3)) & Chr$(1) & _
                           Format$(v(r, 1), "000000000")
            End If
        End If
    Next r
    If nSel = 0 Then Exit Function
    SorteerSleutels sl, sel, tk, 1, nSel

    If BladBestaatStil(BL_UITVRAAG) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(BL_UITVRAAG).Unprotect WACHTWOORD
        ThisWorkbook.Worksheets(BL_UITVRAAG).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = NieuwBlad()
    ws.Name = BL_UITVRAAG
    ws.Tab.Color = RGB(0, 73, 118)
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10

    ' kolommen uit de verzamellijst: C, D, O, S, Y, AD, AE t/m AI, AJ
    bronKol = Array(3, 4, 15, 19, 25, 30, 31, 32, 33, 34, 35, 36)
    nK = UBound(bronKol) + 1
    ws.Range("A1").Value = "Uitvraag actualisatie investeringen kadernota / begroting " & (lj + 1)
    With ws.Range("A1").Font
        .Bold = True
        .Size = 16
        .Color = RGB(0, 73, 118)
    End With
    ws.Range("A2").Value = IIf(Len(naam) > 0, "Budgetbeheerder: " & naam & " - ", "") & "stand " & Format$(Date, "dd-mm-yyyy") & _
                           ". Geef per krediet aan wat er moet gebeuren (kolom Actualisatie), pas zo nodig de bedragen aan en licht toe." & _
                           " Terugschuiven kan niet naar " & lj & "."
    ws.Range("A2").Font.Size = 11
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Rows(1).RowHeight = 26
    ' rij 4: verwijzing naar de kolom in de verzamellijst
    For k = 0 To nK - 1
        ws.Cells(4, k + 1).Value = "kolom " & Split(wv.Cells(1, bronKol(k)).Address(True, False), "$")(0)
        kop = wv.Cells(1, bronKol(k)).Text
        If bronKol(k) = 36 Then kop = "Toelichting op de wijziging (in te vullen)"
        ws.Cells(5, k + 1).Value = kop
    Next k
    With ws.Range(ws.Cells(4, 1), ws.Cells(4, nK))
        .Font.Size = 8
        .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(5, 1), ws.Cells(5, nK))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 73, 118)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(5).RowHeight = 30
    ws.Cells(4, 13).Value = "niet wijzigen"
    ws.Cells(5, 13).Value = "Volgnr"                ' sleutel om terug te lezen (verborgen kolom)
    For k = 1 To 5
        ws.Cells(5, 13 + k).Value = "Was " & (lj + k)
    Next k

    rij = 6: eersteData = 6: vBh = Chr$(0)
    For i = 1 To nSel
        r = sel(i)
        bh = Trim$(CStr(v(r, 15)))
        If bh <> vBh Then
            If vBh <> Chr$(0) Then Tussentotaal ws, rij, startBh, "Totaal " & vBh, 6, 11, 1
            ws.Cells(rij, 1).Value = IIf(Len(bh) > 0, bh, "(geen budgetbeheerder)")
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK)).Font.Bold = True
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK)).Interior.Color = RGB(214, 228, 240)
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK)).Font.Color = RGB(0, 73, 118)
            rij = rij + 1: startBh = rij
            vBh = bh
        End If
        ws.Cells(rij, 1).NumberFormat = "@"
        ws.Cells(rij, 1).Value = NrTekst(v(r, 3))
        ws.Cells(rij, 2).Value = v(r, 4)
        ws.Cells(rij, 3).Value = v(r, 15)
        tv = CStr(CodeNaarWeergave("Tbl_Taakveld", v(r, 19)))
        If Len(tv) > 30 Then tv = Left$(tv, 29) & "."      ' taakveld moet binnen de kolom passen
        ws.Cells(rij, 4).Value = tv
        ws.Cells(rij, 5).Value = v(r, 25)
        ws.Cells(rij, 6).Value = Nul(v(r, 30))
        For k = 1 To 5
            ws.Cells(rij, 6 + k).Value = Nul(v(r, 30 + k))
        Next k
        ws.Cells(rij, 13).Value = v(r, 1)
        For k = 1 To 5
            ws.Cells(rij, 13 + k).Value = Nul(v(r, 30 + k))   ' "was"-bedragen (verborgen) voor de automatische wijzigingstekst
        Next k
        rij = rij + 1
    Next i
    Tussentotaal ws, rij, startBh, "Totaal " & vBh, 6, 11, 1
    ws.Cells(rij, 1).Value = "Eindtotaal"
    For c = 6 To 11
        ws.Cells(rij, c).Formula = "=SUBTOTAL(9," & ws.Range(ws.Cells(eersteData, c), ws.Cells(rij - 1, c)).Address(False, False) & ")"
    Next c
    With ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK))
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With

    ' voorstellen voor nieuwe kredieten: lege regels onderaan (gemarkeerd met NIEUW in de verborgen kolom)
    Dim nvStart As Long, nvEind As Long
    rij = rij + 2
    ws.Cells(rij, 1).Value = "Bestaand beleid: nieuwe jaarschijf of vervanging (vul omschrijving, taakveld, budgetten en toelichting in). " & _
                             "Geen nieuwe investeringen; de afschrijftermijn bepaalt de financieel adviseur."
    ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK)).Font.Bold = True
    ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK)).Interior.Color = RGB(214, 228, 240)
    ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK)).Font.Color = RGB(0, 73, 118)
    nvStart = rij + 1: nvEind = rij + 10
    For k = nvStart To nvEind
        ws.Cells(k, 3).Value = naam
        ws.Cells(k, 13).Value = "NIEUW"
    Next k
    rij = nvEind + 1

    ' kolom L: nieuwe omschrijving (invullen), kolom M: wijziging (automatisch), kolom N: onderbouwing (invullen)
    ws.Range(ws.Columns(12), ws.Columns(13)).Insert
    nK = nK + 2
    ws.Cells(4, 12).Value = "in te vullen"
    ws.Cells(4, 13).Value = "automatisch"
    ws.Cells(4, 14).Value = "in te vullen (kolom AJ)"
    ws.Cells(5, 12).Value = "Nieuwe omschrijving (alleen bij wijziging)"
    ws.Cells(5, 13).Value = "Wijziging (automatisch)"
    ws.Cells(5, 14).Value = "Onderbouwing / toelichting"
    ws.Range(ws.Cells(4, 12), ws.Cells(4, 14)).Font.Size = 8
    ws.Range(ws.Cells(4, 12), ws.Cells(4, 14)).Font.Color = RGB(128, 128, 128)
    ws.Range(ws.Cells(4, 12), ws.Cells(4, 14)).HorizontalAlignment = xlCenter
    With ws.Range(ws.Cells(5, 12), ws.Cells(5, 13))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 73, 118)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Columns(15), ws.Columns(20)).Hidden = True
    ws.Cells.Locked = True
    ' voorstellen nieuwe kredieten: omschrijving, taakveld, afschrijving, budgetten en toelichting invullen
    For k = nvStart To nvEind
        ws.Range(ws.Cells(k, 2), ws.Cells(k, 2)).Interior.Color = RGB(255, 255, 204)
        ws.Cells(k, 4).Interior.Color = RGB(255, 255, 204)          ' afschrijftermijn (E) niet invulbaar
        ws.Range(ws.Cells(k, 7), ws.Cells(k, 11)).Interior.Color = RGB(255, 255, 204)
        ws.Cells(k, 14).Interior.Color = RGB(255, 255, 204)
        ws.Range(ws.Cells(k, 2), ws.Cells(k, 2)).Locked = False
        ws.Cells(k, 4).Locked = False
        ws.Range(ws.Cells(k, 7), ws.Cells(k, 11)).Locked = False
        ws.Cells(k, 14).Locked = False
        ws.Cells(k, 13).Value = "Nieuw krediet"
    Next k
    ' bestaande kredieten: budgetten (G-K), nieuwe omschrijving (L) en onderbouwing (N) invullen; M rekent zelf
    Dim fWijz As String, jaren As String, eur As String
    eur = ChrW(8364)
    jaren = (lj + 1) & "," & (lj + 2) & "," & (lj + 3) & "," & (lj + 4) & "," & (lj + 5)
    For k = eersteData To rij - 1
        If Not Leeg(ws.Cells(k, 2).Value) And Not ws.Cells(k, 6).HasFormula And k < nvStart Then   ' kredietregel, geen totaalregel
            ws.Range(ws.Cells(k, 7), ws.Cells(k, 12)).Interior.Color = RGB(255, 255, 204)
            ws.Cells(k, 14).Interior.Color = RGB(255, 255, 204)
            ws.Range(ws.Cells(k, 7), ws.Cells(k, 12)).Locked = False
            ws.Cells(k, 14).Locked = False
            ' teksten: "Budget vervallen" (alles naar 0), "Budget 2028 vervallen" (alleen jaarschijf naar 0),
            ' "Budget 2028 verhoogd/verlaagd met EUR x" (met het jaar), of bij een verschuiving
            ' "Budget verlaagd met EUR x, van 2027 naar 2028" (oud en nieuw jaar)
            fWijz = "=LET(n,G#:K#,o,P#:T#,j,{" & jaren & "},d,SUM(n)-SUM(o)," & _
                    "van,TEXTJOIN("", "",TRUE,IF(n<o,j,"""")),naar,TEXTJOIN("", "",TRUE,IF(n>o,j,""""))," & _
                    "vv,TEXTJOIN("", "",TRUE,IF((n=0)*(o<>0),j,""""))," & _
                    "vs,IF(AND(van<>"""",naar<>""""),""van ""&van&"" naar ""&naar,"""")," & _
                    "jr,IF(vs="""",IF(naar<>"""",naar,van)&"" "","""")," & _
                    "bd,IF(d>0,""Budget ""&jr&""verhoogd met " & eur & " ""&FIXED(d,0),IF(d<0,""Budget ""&jr&""verlaagd met " & eur & " ""&FIXED(-d,0),""""))," & _
                    "IF(AND(SUM(o)<>0,SUM(n)=0),""Budget vervallen""," & _
                    "IF(AND(naar="""",van<>"""",van=vv),""Budget ""&vv&"" vervallen""," & _
                    "IF(bd="""",IF(vs="""","""",""Budget ""&vs),bd&IF(vs="""","""","", ""&vs)))))"
            ws.Cells(k, 13).Formula2 = Replace(fWijz, "#", CStr(k))
            With ws.Cells(k, 13).FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""""")
                .Interior.Color = RGB(255, 200, 120)          ' markering: er is iets gewijzigd
                .Font.Bold = True
            End With
        End If
    Next k
    ws.Range(ws.Cells(eersteData, 6), ws.Cells(rij, 11)).NumberFormat = "#,##0;-#,##0;-"
    ws.Range(ws.Cells(5, 6), ws.Cells(rij, 11)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(5, 5), ws.Cells(rij, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(eersteData, 1), ws.Cells(rij, 4)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(eersteData, 12), ws.Cells(rij, 14)).HorizontalAlignment = xlLeft
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 45
    ws.Columns(3).ColumnWidth = 22
    ws.Columns(4).ColumnWidth = 32
    ws.Columns(5).ColumnWidth = 8
    ws.Range(ws.Columns(6), ws.Columns(11)).ColumnWidth = 15
    ws.Columns(12).ColumnWidth = 30
    ws.Columns(13).ColumnWidth = 40
    ws.Columns(14).ColumnWidth = 60
    ws.Range(ws.Cells(eersteData, 2), ws.Cells(rij, 3)).WrapText = True
    ws.Range(ws.Cells(eersteData, 4), ws.Cells(rij, 4)).WrapText = False
    ws.Range(ws.Cells(eersteData, 12), ws.Cells(rij, 14)).WrapText = True
    ws.Range(ws.Cells(eersteData, 1), ws.Cells(rij, nK)).VerticalAlignment = xlTop
    BladOpmaken ws, 5, eersteData, rij
    LogoPlaatsen ws, 14
    KnopTerug ws, ws.Range("A3")
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("C6").Select
    ActiveWindow.FreezePanes = True
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .PrintTitleRows = "$4:$5"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Pagina &P van &N"
    End With
    On Error GoTo 0
    ws.Protect Password:=WACHTWOORD, DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    TabbladenOrdenen
    ws.Activate
    Set UitvraagBlad = ws
End Function

' Tabbladvolgorde: Dashboard, Verzamellijst, Laatste overzicht op naam, Laatste kredietoverzicht, Laatste uitvraag begroting,
' Laatste overzicht tijdschrijven, Meldingen uitgaven, daarna de overige bladen, en Voertuigen en Controle als laatste.
' Draait bij openen, na elk overzicht en bij DashboardKnoppenPlaatsen; ook met de hand te starten.
Public Sub TabbladenOrdenen()
    Dim naam As Variant, pos As Long, actief As Object, scherm As Boolean
    On Error Resume Next
    Set actief = ActiveSheet
    scherm = Application.ScreenUpdating
    Application.ScreenUpdating = False                 ' verplaatsen maakt een blad actief: niet laten zien
    pos = 0
    For Each naam In Array("Dashboard", BL_VZ, BL_OVERZICHT_NAAM, BL_OVERZICHT, BL_UITVRAAG, BL_TIJD, BL_MELD)
        If BladBestaatStil(CStr(naam)) Then
            pos = pos + 1
            If ThisWorkbook.Worksheets(CStr(naam)).Index <> pos Then
                ThisWorkbook.Worksheets(CStr(naam)).Move Before:=ThisWorkbook.Worksheets(pos)
            End If
        End If
    Next naam
    ' Voertuigen en daarna Controle helemaal achteraan; alleen verplaatsen als het nog niet zo staat
    If BladBestaatStil("Controle") And BladBestaatStil("Voertuigen") Then
        Dim w As Worksheet, moet As Boolean
        moet = (ThisWorkbook.Worksheets("Controle").Index < ThisWorkbook.Worksheets("Voertuigen").Index)
        For Each w In ThisWorkbook.Worksheets           ' een zichtbaar ander blad achter Voertuigen? (verborgen bladen tellen niet)
            If w.Visible = xlSheetVisible And w.Name <> "Voertuigen" And w.Name <> "Controle" _
               And w.Index > ThisWorkbook.Worksheets("Voertuigen").Index Then moet = True
        Next w
        If moet Then
            ThisWorkbook.Worksheets("Voertuigen").Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            ThisWorkbook.Worksheets("Controle").Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        End If
    End If
    If Not actief Is Nothing Then actief.Activate
    Application.ScreenUpdating = scherm
    TabbladkiezerVullen
End Sub

' ======================================================================
' Navigatie: knop "Terug naar dashboard" op elk blad en een tabbladkiezer op het Dashboard
' ======================================================================

' Eenmalig/bij openen: 'Verzamellijst alle kredieten' heet voortaan 'Verzamellijst'.
' Formules, namen en spiegelbladen passen zich in Excel vanzelf aan.
' Het blad Verzamellijst; hernoemt eerst een blad met de oude naam (zodat geen enkele macro op de oude naam vastloopt).
Private Function Vz() As Worksheet
    VerzamellijstHernoemen
    Set Vz = ThisWorkbook.Worksheets(BL_VZ)
End Function

Public Sub VerzamellijstHernoemen()
    On Error Resume Next
    If BladBestaatStil("Verzamellijst alle kredieten") And Not BladBestaatStil(BL_VZ) Then
        ThisWorkbook.Worksheets("Verzamellijst alle kredieten").Name = BL_VZ
    End If
End Sub

Public Sub NaarDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
    ThisWorkbook.Worksheets("Dashboard").Range("A1").Select
End Sub

Public Sub NaarTabblad()
    Dim naam As String
    naam = Trim$(CStr(ThisWorkbook.Worksheets("Dashboard").Range("G2").Value))
    If Len(naam) = 0 Or Not BladBestaatStil(naam) Then
        MsgBox "Kies eerst in cel G2 een tabblad.", vbExclamation
        Exit Sub
    End If
    If ThisWorkbook.Worksheets(naam).Visible <> xlSheetVisible Then
        MsgBox "Dit tabblad is verborgen.", vbExclamation
        Exit Sub
    End If
    ThisWorkbook.Worksheets(naam).Activate
End Sub

' Zet de knop "Terug naar dashboard" op een blad (een bestaande knop wordt vervangen).
Private Sub KnopTerug(ByVal ws As Worksheet, ByVal cel As Range)
    Dim b As Object, beveiligd As Boolean
    If ws.Name = "Dashboard" Then Exit Sub
    On Error Resume Next
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    ws.Buttons("btnTerug").Delete
    Set b = ws.Buttons.Add(cel.Left + 1, cel.Top + 1, 130, 18)
    b.Name = "btnTerug"
    b.Caption = "Terug naar dashboard"
    b.OnAction = "NaarDashboard"
    b.Placement = xlFreeFloating
    If beveiligd Then BladBeveiligen ws
End Sub

' Knop "Terug naar dashboard" op alle zichtbare bladen.
Public Sub KnoppenTerugPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, c As Long
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.Name <> "Dashboard" Then
            Select Case ws.Name
                Case BL_VZ
                    ws.Range("A1:B1").VerticalAlignment = xlBottom
                    If ws.Rows(1).RowHeight < 45 Then ws.Rows(1).RowHeight = 45
                    KnopTerug ws, ws.Range("A1")
                Case "Controle"
                    KnopTerug ws, ws.Range("A2").Offset(0, 0)
                    ws.Buttons("btnTerug").Left = ws.Range("A2").Left + 4 * 160
                    ws.Buttons("btnTerug").Top = ws.Range("A2").Top + 2
                Case BL_MELD, BL_TIJD
                    KnopTerug ws, ws.Range(IIf(ws.Name = BL_MELD, "D1", "A3"))
                Case BL_OVERZICHT, BL_OVERZICHT_NAAM, BL_UITVRAAG
                    KnopTerug ws, ws.Range("A3")
                Case Else
                    ' rechts naast het gebruikte deel van rij 1
                    c = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
                    If c < 3 Then c = 3
                    KnopTerug ws, ws.Cells(1, c)
            End Select
        End If
    Next ws
End Sub

' Tabbladkiezer op het Dashboard: F2 tekst, G2 keuzelijst, H2 knop.
Public Sub TabbladkiezerPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("F2").Value = "Naar tabblad:"
    ws.Range("F2").HorizontalAlignment = xlRight
    ws.Range("G2").Interior.Color = RGB(255, 255, 204)
    If ws.Columns("G").ColumnWidth < 30 Then ws.Columns("G").ColumnWidth = 30
    On Error Resume Next
    ws.Buttons("btnNaarTabblad").Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(ws.Range("H2").Left + 22, ws.Range("H2").Top + 1, 90, 18)   ' ruimte voor het keuzepijltje van G2
    b.Name = "btnNaarTabblad"
    b.Caption = "Ga naar tabblad"
    b.OnAction = "NaarTabblad"
    TabbladkiezerVullen
End Sub

' Keuzelijst in G2 met alle zichtbare tabbladen behalve het Dashboard.
Private Sub TabbladkiezerVullen()
    Dim ws As Worksheet, lijst As String, wd As Worksheet, beveiligd As Boolean
    On Error Resume Next
    ' keuzelijst aanpassen lukt niet op een beveiligd blad, ook niet voor macro's: even open
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    beveiligd = wd.ProtectContents
    If beveiligd Then wd.Unprotect WACHTWOORD
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.Name <> "Dashboard" And InStr(ws.Name, ",") = 0 Then
            lijst = lijst & IIf(Len(lijst) > 0, ",", "") & ws.Name
        End If
    Next ws
    If Len(lijst) > 255 Then lijst = Left$(lijst, InStrRev(Left$(lijst, 255), ",") - 1)
    With ThisWorkbook.Worksheets("Dashboard").Range("G2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=lijst
    End With
    ' voorselectie: het eerste tabblad uit de lijst (Verzamellijst) als er nog niets (geldigs) staat
    With ThisWorkbook.Worksheets("Dashboard").Range("G2")
        If Len(Trim$(CStr(.Value))) = 0 Or Not BladBestaatStil(CStr(.Value)) Then .Value = Split(lijst, ",")(0)
    End With
    If beveiligd Then BladBeveiligen wd
End Sub


' Kredietnummer 7999998 is een tijdelijk nummer voor kredieten die na het begrotingsjaar starten.
' Zodra een krediet met 7999998 in het begrotingsjaar (lopend jaar + 1) of eerder valt (jaar uit de status in kolom H),
' moet het een echt kredietnummer krijgen: ALARM op de meldingenlijst. Voorbeeld: lopend jaar 2027 -> alle 7999998 met L2028 of eerder.
Private Sub Alert7999998(ByVal vS As Variant, ByVal lj As Long)
    Dim i As Long, st As String, jaar As Long
    For i = 1 To UBound(vS, 1)
        If Not Leeg(vS(i, 1)) And NrTekst(vS(i, 3)) = "7999998" Then
            st = UCase$(Trim$(CStr(vS(i, 8))))
            jaar = Val(Mid$(st, 2))
            If Left$(st, 1) = "L" And jaar > 0 And jaar <= lj + 1 Then
                Meld 1, "ALARM: echt kredietnummer nodig (7999998)", jaar, "7999998", CStr(vS(i, 4)), i + 1, Empty, Empty, _
                     "Krediet loopt in " & jaar & " (begrotingsjaar " & (lj + 1) & "): vervang 7999998 door een echt kredietnummer."
            End If
        End If
    Next i
End Sub

' ======================================================================
' Actualisatie kadernota / begroting verwerken (financieel adviseur)
' ======================================================================
' De teamleider (budgetbeheerder) vult zijn uitvraagbestand in map 02\<naam> in. De financieel adviseur controleert
' het bestand en kiest daarna op het Dashboard (C58 = Budgetbeheerder, C59 = naam) deze knop. De macro leest het
' nieuwste uitvraagbestand van die persoon, toont een samenvatting en werkt na akkoord de verzamellijst bij:
' - budget per jaarschijf (AE t/m AI) zoals ingevuld;
' - omschrijving (D, en de korte omschrijving F) als een nieuwe omschrijving is ingevuld;
' - toelichting (AJ) wordt vervangen als de teamleider een toelichting heeft gegeven.
' Het lopend jaar wordt niet aangeraakt. Geen controle door de administratie (die speelt bij deze lijst geen rol).
' Voorstellen voor nieuwe kredieten komen op het blad 'Voorstellen nieuwe kredieten'; de financieel adviseur
' voert ze na beoordeling in via de invoerkaart.
Public Sub ActualisatieVerwerken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const BL_VOORSTEL As String = "Voorstellen nieuwe kredieten"
    Dim wd As Worksheet, wv As Worksheet, wb As Workbook, d As Variant, map As String, naam As String, f As String
    Dim bestand As String, dt As Date, laatste As Long, r As Long, i As Long, k As Long, rij As Long, lj As Long
    Dim volgRij As Object, vn As String, nieuwB(1 To 5) As Double, oudB As Double, wijzig As Boolean
    Dim nKred As Long, nNieuw As Long, somOud(1 To 5) As Double, somNieuw(1 To 5) As Double, tekst As String
    Dim plan() As Variant, nPlan As Long, voor As Variant, na As Variant, verschil As Double, wsVs As Worksheet
    Dim oms As String, toel As String, act As String, gebr As String, nAnders As Long, nieuweToel As String

    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur kan een actualisatie verwerken.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    If LCase$(Trim$(CStr(wd.Range("C58").Value))) <> "budgetbeheerder" Or Len(Trim$(CStr(wd.Range("C59").Value))) = 0 Then
        MsgBox "Kies op het Dashboard in C58 'Budgetbeheerder' en in C59 de teamleider van wie de actualisatie verwerkt moet worden.", vbExclamation
        Exit Sub
    End If
    naam = Trim$(CStr(wd.Range("C59").Value))
    map = MapStambestand("02 Budgetbeheerders en projectleiders")
    If Len(map) = 0 Then Exit Sub
    map = SubMap(map, naam)
    f = Dir(map & "Uitvraag kadernota-begroting*.xlsx")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then
            If Len(bestand) = 0 Or FileDateTime(map & f) > dt Then bestand = map & f: dt = FileDateTime(map & f)
        End If
        f = Dir()
    Loop
    If Len(bestand) = 0 Then
        MsgBox "Geen uitvraagbestand gevonden in:" & vbCrLf & map, vbExclamation
        Exit Sub
    End If
    If InStr(1, LogboekTekst(), Dir(bestand) & " (gewijzigd " & Format$(dt, "dd-mm-yyyy hh:mm") & ")", vbTextCompare) > 0 Then
        If MsgBox("Dit bestand is al eens verwerkt:" & vbCrLf & Dir(bestand) & vbCrLf & vbCrLf & "Toch opnieuw verwerken?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Actualisatie verwerken") <> vbYes Then Exit Sub
    End If

    On Error GoTo Fout
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(Filename:=bestand, ReadOnly:=True, UpdateLinks:=0)
    d = wb.Worksheets(1).Range("A1:T" & wb.Worksheets(1).UsedRange.Row + wb.Worksheets(1).UsedRange.Rows.Count).Value
    wb.Close SaveChanges:=False
    Set wb = Nothing
    ThisWorkbook.Activate
    If CStr(d(5, 15)) <> "Volgnr" Or CStr(d(5, 13)) <> "Wijziging (automatisch)" Then
        Application.ScreenUpdating = True
        MsgBox "Dit uitvraagbestand heeft nog de oude opzet (zonder volgnummer) en kan niet worden ingelezen." & vbCrLf & _
               "Maak de uitvraag opnieuw met de knop 'Uitvraag kadernota/begroting'.", vbExclamation
        Exit Sub
    End If

    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    Set volgRij = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(wv.Cells(r, 1).Value) Then volgRij(CStr(wv.Cells(r, 1).Value)) = r
    Next r

    ' 1. plan maken: per bestaand krediet de wijzigingen, en de voorstellen voor nieuwe kredieten
    ReDim plan(1 To UBound(d, 1), 1 To 9)            ' rij, oms, toel, b1..b5, act
    For i = 6 To UBound(d, 1)
        vn = Trim$(Tkst(d(i, 15)))
        If UCase$(vn) = "NIEUW" Then
            If Len(Trim$(Tkst(d(i, 2)))) > 0 Then nNieuw = nNieuw + 1
        ElseIf Len(vn) > 0 And volgRij.Exists(vn) Then
            r = volgRij(vn)
            wijzig = False
            For k = 1 To 5
                nieuwB(k) = Nul(d(i, 6 + k))
                oudB = Nul(wv.Cells(r, 30 + k).Value)
                If Round(nieuwB(k) - oudB, 2) <> 0 Then wijzig = True
            Next k
            oms = Trim$(Tkst(d(i, 12)))                 ' L: nieuwe omschrijving
            act = Trim$(Tkst(d(i, 13)))                 ' M: automatische wijzigingstekst
            toel = Trim$(Tkst(d(i, 14)))                ' N: onderbouwing van de teamleider
            For k = 1 To 5
                If Round(Nul(d(i, 15 + k)) - Nul(wv.Cells(r, 30 + k).Value), 2) <> 0 Then nAnders = nAnders + 1: Exit For
            Next k
            If wijzig Or Len(oms) > 0 Or Len(toel) > 0 Then
                nPlan = nPlan + 1
                plan(nPlan, 1) = r: plan(nPlan, 2) = oms: plan(nPlan, 3) = toel: plan(nPlan, 9) = act
                For k = 1 To 5
                    plan(nPlan, 3 + k) = nieuwB(k)
                    somOud(k) = somOud(k) + Nul(wv.Cells(r, 30 + k).Value)
                    somNieuw(k) = somNieuw(k) + nieuwB(k)
                Next k
                nKred = nKred + 1
            End If
        End If
    Next i
    Application.ScreenUpdating = True
    If nPlan = 0 And nNieuw = 0 Then
        MsgBox "In het bestand van " & naam & " staan geen wijzigingen en geen voorstellen.", vbInformation
        Exit Sub
    End If

    ' 2. samenvatting en akkoord
    tekst = "Actualisatie van " & naam & vbCrLf & Dir(bestand) & vbCrLf & vbCrLf & _
            nKred & " kredieten met wijzigingen, " & nNieuw & " voorstellen voor nieuwe kredieten." & vbCrLf & vbCrLf & _
            "Budget per jaarschijf (oud -> nieuw), alleen de gewijzigde kredieten:" & vbCrLf
    For k = 1 To 5
        tekst = tekst & (lj + k) & ":  " & Format$(somOud(k), "#,##0") & "  ->  " & Format$(somNieuw(k), "#,##0") & vbCrLf
    Next k
    tekst = tekst & vbCrLf & "Het lopend jaar " & lj & " wordt niet gewijzigd." & vbCrLf
    If nAnders > 0 Then tekst = tekst & vbCrLf & "LET OP: bij " & nAnders & " kredieten is het budget in de verzamellijst sinds de uitvraag gewijzigd." & _
                               vbCrLf & "Het bedrag uit de uitvraag gaat dan voor." & vbCrLf
    tekst = tekst & vbCrLf & "Verwerken in de verzamellijst?"
    If MsgBox(tekst, vbYesNo + vbQuestion + vbDefaultButton2, "Actualisatie kadernota/begroting verwerken") <> vbYes Then Exit Sub

    ' 3. verwerken met nulmeting
    Application.ScreenUpdating = False
    voor = MeetStap8(wv, laatste)
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    verschil = 0
    For i = 1 To nPlan
        r = plan(i, 1)
        For k = 1 To 5
            oudB = Nul(wv.Cells(r, 30 + k).Value)
            If Round(plan(i, 3 + k) - oudB, 2) <> 0 Then
                LogRegel "actualisatie kadernota", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, wv.Cells(1, 30 + k).Text, oudB, plan(i, 3 + k), _
                         naam & " | " & plan(i, 9)
                MutatieRegistreren naam, wv.Cells(r, 1).Value, NrTekst(wv.Cells(r, 3).Value), CStr(wv.Cells(r, 4).Value), _
                                   wv.Cells(r, 19).Value, wv.Cells(r, 20).Value, wv.Cells(r, 25).Value, lj + k, oudB, CDbl(plan(i, 3 + k)), _
                                   CStr(plan(i, 9)), CStr(plan(i, 3)), Dir(bestand)
                If plan(i, 3 + k) = 0 Then wv.Cells(r, 30 + k).ClearContents Else wv.Cells(r, 30 + k).Value = plan(i, 3 + k)
                verschil = verschil + plan(i, 3 + k) - oudB
            End If
        Next k
        If Len(plan(i, 2)) > 0 Then
            LogRegel "actualisatie kadernota", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Omschrijving", wv.Cells(r, 4).Value, plan(i, 2), naam
            wv.Cells(r, 4).Value = Left$(plan(i, 2), 36)
            wv.Cells(r, 6).Value = KortOmschrijving(CStr(wv.Cells(r, 4).Value))
        End If
        ' toelichting AJ = automatische wijzigingstekst plus de onderbouwing van de teamleider
        nieuweToel = CStr(plan(i, 9))
        If Len(plan(i, 3)) > 0 Then nieuweToel = nieuweToel & IIf(Len(nieuweToel) > 0, ". ", "") & plan(i, 3)
        If Len(nieuweToel) > 0 Then
            LogRegel "actualisatie kadernota", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Toelichting", wv.Cells(r, 36).Value, nieuweToel, naam
            wv.Cells(r, 36).Value = nieuweToel
        End If
        wv.Cells(r, 48).Value = gebr
        wv.Cells(r, 49).Value = Now
    Next i
    Application.Calculate
    na = MeetStap8(wv, laatste)
    ' alleen AE t/m AI mogen veranderen, met precies het verwerkte verschil
    If na(0) <> voor(0) Or Round(na(1) - voor(1), 2) <> 0 Or Round(na(2) - voor(2), 2) <> 0 Or Round(na(3) - voor(3), 2) <> 0 _
       Or Round(na(4) - voor(4), 2) <> 0 Or Round((na(5) + na(6) + na(7) + na(8) + na(9)) - (voor(5) + voor(6) + voor(7) + voor(8) + voor(9)) - verschil, 2) <> 0 Then
        Application.ScreenUpdating = True
        MsgBox "Let op: de nameting wijkt af. Controleer de verzamellijst en het Logboek.", vbCritical
    End If

    ' 4. voorstellen nieuwe kredieten naar een eigen blad
    If nNieuw > 0 Then
        If Not BladBestaatStil(BL_VOORSTEL) Then
            Set wsVs = NieuwBlad()
            wsVs.Name = BL_VOORSTEL
            wsVs.Cells.Font.Name = "Arial"
            wsVs.Cells.Font.Size = 10
            wsVs.Range("A1").Value = "Voorstellen bestaand beleid (nieuwe jaarschijf / vervanging): beoordelen, afschrijftermijn bepalen en invoeren via de invoerkaart"
            wsVs.Range("A1").Font.Bold = True
            wsVs.Range("A1").Font.Size = 14
            wsVs.Range("A1").Font.Color = RGB(0, 73, 118)
            wsVs.Range("A4:M4").Value = Array("Ingediend op", "Budgetbeheerder", "Omschrijving", "Taakveld", "Afschr", _
                                              "Budget " & (lj + 1), "Budget " & (lj + 2), "Budget " & (lj + 3), "Budget " & (lj + 4), _
                                              "Budget " & (lj + 5), "Toelichting", "Bestand", "Verwerkt (ja/nee)")
            With wsVs.Range("A4:M4")
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(0, 73, 118)
                .WrapText = True
            End With
        Else
            Set wsVs = ThisWorkbook.Worksheets(BL_VOORSTEL)
        End If
        rij = Application.Max(5, wsVs.Cells(wsVs.Rows.Count, 3).End(xlUp).Row + 1)
        For i = 6 To UBound(d, 1)
            If UCase$(Trim$(Tkst(d(i, 15)))) = "NIEUW" And Len(Trim$(Tkst(d(i, 2)))) > 0 Then
                wsVs.Cells(rij, 1).Value = Date
                wsVs.Cells(rij, 2).Value = naam
                wsVs.Cells(rij, 3).Value = Tkst(d(i, 2))
                wsVs.Cells(rij, 4).Value = Tkst(d(i, 4))
                wsVs.Cells(rij, 5).Value = ""                ' afschrijftermijn bepaalt de financieel adviseur
                For k = 1 To 5
                    wsVs.Cells(rij, 5 + k).Value = Nul(d(i, 6 + k))
                Next k
                wsVs.Cells(rij, 11).Value = Tkst(d(i, 14))
                wsVs.Cells(rij, 12).Value = Dir(bestand)
                wsVs.Cells(rij, 13).Value = "nee"
                For k = 1 To 5
                    If Nul(d(i, 6 + k)) <> 0 Then
                        MutatieRegistreren naam, "", "nieuw", Tkst(d(i, 2)), Tkst(d(i, 4)), "", "", lj + k, 0, Nul(d(i, 6 + k)), _
                                           "Nieuw krediet (voorstel)", Tkst(d(i, 14)), Dir(bestand)
                    End If
                Next k
                rij = rij + 1
            End If
        Next i
        wsVs.Range("A5:A" & rij).NumberFormat = "dd-mm-yyyy"
        wsVs.Range("F5:J" & rij).NumberFormat = "#,##0;-#,##0;-"
        wsVs.Columns("A").ColumnWidth = 12: wsVs.Columns("B").ColumnWidth = 20: wsVs.Columns("C").ColumnWidth = 40
        wsVs.Columns("D").ColumnWidth = 30: wsVs.Columns("E").ColumnWidth = 7: wsVs.Range("F:J").ColumnWidth = 14
        wsVs.Columns("K").ColumnWidth = 60: wsVs.Columns("L").ColumnWidth = 30: wsVs.Columns("M").ColumnWidth = 10
        wsVs.Range("C5:C" & rij & ",K5:K" & rij).WrapText = True
        wsVs.Range("A5:M" & rij).VerticalAlignment = xlTop
        BladOpmaken wsVs, 4, 5, rij
        KnopTerug wsVs, wsVs.Range("A2")
    End If

    LogRegel "actualisatie kadernota verwerkt", , , "", , , naam & " | " & Dir(bestand) & " (gewijzigd " & Format$(dt, "dd-mm-yyyy hh:mm") & ") | " & _
             nKred & " kredieten, " & nNieuw & " voorstellen"
    Application.ScreenUpdating = True
    MsgBox "Verwerkt: " & nKred & " kredieten bijgewerkt." & IIf(nNieuw > 0, vbCrLf & nNieuw & " voorstellen voor nieuwe kredieten staan op het blad '" & _
           BL_VOORSTEL & "'. Voer die na beoordeling in via de invoerkaart.", ""), vbInformation
    Exit Sub
Fout:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Fout bij het verwerken van de actualisatie: " & Err.Description, vbCritical
End Sub

' Alle tekst uit de toelichtingskolom van het Logboek (om dubbel verwerken te herkennen).
Private Function LogboekTekst() As String
    Dim lg As Worksheet, v As Variant, i As Long, t As String
    On Error Resume Next
    Set lg = ThisWorkbook.Worksheets("Logboek")
    v = lg.Range("K1:K" & lg.Cells(lg.Rows.Count, 1).End(xlUp).Row).Value
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            t = t & CStr(v(i, 1)) & vbLf
        Next i
    End If
    LogboekTekst = t
End Function

' Legt elke verwerkte wijziging per jaarschijf vast op het blad 'Mutaties kadernota-begroting'.
' Met taakveld, balanscode en afschrijftermijn is dit de basis voor de berekening van de lastenaanpassingen
' (kapitaallasten) in de kadernota/begroting. Het register wordt aangevuld, nooit leeggemaakt.
Private Sub MutatieRegistreren(ByVal naam As String, ByVal volgnr As Variant, ByVal nr As String, ByVal oms As String, _
                               ByVal taakveld As Variant, ByVal balans As Variant, ByVal afschr As Variant, ByVal jaar As Long, _
                               ByVal oud As Double, ByVal nieuw As Double, ByVal act As String, ByVal toel As String, ByVal bestand As String)
    Dim ws As Worksheet, rij As Long
    If Not BladBestaatStil(BL_MUTREG) Then
        Set ws = NieuwBlad()
        ws.Name = BL_MUTREG
        ws.Cells.Font.Name = "Arial"
        ws.Cells.Font.Size = 10
        ws.Range("A1").Value = "Mutaties kadernota / begroting (basis voor de lastenaanpassingen)"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Color = RGB(0, 73, 118)
        ws.Range("A4:O4").Value = Array("Verwerkt op", "Verwerkt door", "Budgetbeheerder", "Volgnr", "Kredietnummer", "Omschrijving", _
                                        "Taakveld", "Balanscode", "Afschr", "Jaar", "Oud budget", "Nieuw budget", "Mutatie", _
                                        "Actualisatie", "Toelichting")
        With ws.Range("A4:O4")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 73, 118)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(4).RowHeight = 30
        ws.Columns("A").ColumnWidth = 16: ws.Columns("B:C").ColumnWidth = 18: ws.Columns("D").ColumnWidth = 8
        ws.Columns("E").ColumnWidth = 14: ws.Columns("F").ColumnWidth = 40: ws.Columns("G:H").ColumnWidth = 11
        ws.Columns("I:J").ColumnWidth = 7: ws.Columns("K:M").ColumnWidth = 15: ws.Columns("N").ColumnWidth = 26
        ws.Columns("O").ColumnWidth = 50
        ws.Columns("E").NumberFormat = "@"
        ws.Range("A4:O4").AutoFilter
        KnopTerug ws, ws.Range("A2")
    Else
        Set ws = ThisWorkbook.Worksheets(BL_MUTREG)
        ws.Unprotect WACHTWOORD
    End If
    rij = Application.Max(5, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1)
    ws.Cells(rij, 1).Value = Now
    ws.Cells(rij, 1).NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Cells(rij, 2).Value = ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value
    ws.Cells(rij, 3).Value = naam
    ws.Cells(rij, 4).Value = volgnr
    ws.Cells(rij, 5).Value = nr
    ws.Cells(rij, 6).Value = oms
    ws.Cells(rij, 7).NumberFormat = "@"
    ws.Cells(rij, 7).Value = CStr(taakveld)
    ws.Cells(rij, 8).Value = balans
    ws.Cells(rij, 9).Value = afschr
    ws.Cells(rij, 10).Value = jaar
    ws.Cells(rij, 11).Value = oud
    ws.Cells(rij, 12).Value = nieuw
    ws.Cells(rij, 13).Value = nieuw - oud
    ws.Cells(rij, 14).Value = act
    ws.Cells(rij, 15).Value = toel
    ws.Range(ws.Cells(rij, 11), ws.Cells(rij, 13)).NumberFormat = "#,##0;-#,##0;-"
    ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 15)).VerticalAlignment = xlTop
    ws.Cells(rij, 6).WrapText = True
    ws.Cells(rij, 15).WrapText = True
End Sub

Private Function IsOpbrengstActiva(ByVal toel As Variant) As Boolean
    IsOpbrengstActiva = (InStr(1, CStr(toel), OPBRENGST, vbTextCompare) > 0)
End Function

' Eenmalig: markeert de bekende opbrengsten activa (verkoop vastgoed Sportlaan 10 en voertuigen V06 en V19) door
' "Opbrengst activa." voor de toelichting te zetten op de regels met een negatief budget. Oude tekst gaat naar het Logboek.
Public Sub OpbrengstenActivaMarkeren()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, n As Long, nr As Variant, c As Long, negatief As Boolean, toel As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        For Each nr In Array("7203107", "7206072", "7205045", "7205041")
            If NrTekst(wv.Cells(r, 3).Value) = nr Then
                negatief = (Nul(wv.Cells(r, 26).Value) < 0)
                For c = 31 To 35
                    If Nul(wv.Cells(r, c).Value) < 0 Then negatief = True
                Next c
                toel = Trim$(CStr(wv.Cells(r, 36).Value))
                If negatief And Not IsOpbrengstActiva(toel) Then
                    LogRegel "opbrengst activa markeren", wv.Cells(r, 1).Value, nr, "Toelichting", toel, OPBRENGST & ". " & toel, "rij " & r
                    wv.Cells(r, 36).Value = OPBRENGST & IIf(Len(toel) > 0, ". " & toel, "")
                    n = n + 1
                End If
            End If
        Next nr
    Next r
    MsgBox n & " regel(s) gemarkeerd als opbrengst activa.", vbInformation
End Sub

' Nieuw tabblad voor Voertuigen en Controle (die blijven achteraan). Zo hoeven Voertuigen en Controle niet steeds
' verplaatst te worden; verplaatsen maakt een blad actief, waardoor het Controle-blad even in beeld kwam.
Private Function NieuwBlad() As Worksheet
    If BladBestaatStil("Voertuigen") Then
        Set NieuwBlad = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets("Voertuigen"))
    ElseIf BladBestaatStil("Controle") Then
        Set NieuwBlad = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets("Controle"))
    Else
        Set NieuwBlad = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
End Function

' ======================================================================
' Controle en kengetallen (stap 12)
' ======================================================================
' Stap 12 - Blok 5 op het Dashboard (vanaf rij 69) en de statusregel in C4 (functioneel ontwerp par. 8).
' ControleblokPlaatsen : zet de controles en kengetallen op het Dashboard. Waar het kan als formule (altijd actueel).
' ControlesVernieuwen  : rekent de controles uit die niet met een formule kunnen (spiegelbladen, verdeelsleutel,
'                        dubbele nummers, handmatige formules, SUBTOTAL-regels). Draait bij openen, na een upload en via de knop.
' Statusregel C4: "OK - alles sluit aan" als alle controles 0 zijn, anders "LET OP".


' Bereik van een kolom van de verzamellijst over de kredietregels (rij 2 t/m Par_LaatsteRij), als formuletekst.
Private Function Rg(ByVal kol As String) As String
    Rg = VL & "$" & kol & "$2:INDEX(" & VL & "$" & kol & ":$" & kol & ",Par_LaatsteRij)"
End Function

Public Sub ControleblokPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, r As Long, k As Long, fc As Object, b As Object, deel As Variant, aantalAf As String
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("B69:D110").Clear
    ws.Range("B69").Value = "5. Controle en kengetallen"
    ws.Range("B69").Font.Bold = True
    ws.Range("B70").Value = "Controles laatst vernieuwd"
    ws.Range("C70").NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Range("C70").HorizontalAlignment = xlLeft
    On Error Resume Next
    ws.Buttons("btnControles").Delete
    On Error GoTo 0
    Set b = ws.Buttons.Add(ws.Range("D70").Left + 2, ws.Range("D70").Top + 1, 150, 18)
    b.Name = "btnControles"
    b.Caption = "Controles vernieuwen"
    b.OnAction = "ControlesVernieuwen"
    ws.Range("B71:D71").Value = Array("Controle", "Aantal", "Waar (rij in de Verzamellijst) / wat te doen")
    ws.Range("B71:D71").Font.Bold = True
    ws.Range("B71:D71").Borders(xlEdgeBottom).LineStyle = xlContinuous

    ' de aantallen zijn formules (altijd actueel); de vindplaatsen in kolom D vult ControlesVernieuwen
    r = CB_EERSTE
    Regel ws, r, "Omschrijving langer dan toegestaan (kolom E en G)", "=COUNTIF(" & Rg("E") & ",""fout"")+COUNTIF(" & Rg("G") & ",""fout"")"
    Regel ws, r, "Verdeelsleutel samengesteld krediet niet 100%", ""
    Regel ws, r, "Afgesloten (A) met restantkrediet", "=COUNTIFS(" & Rg("H") & ",""A*""," & Rg("AD") & ",""<>0"")"
    Regel ws, r, "Zonder budgetbeheerder (O)", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("O") & ","""")"
    Regel ws, r, "Zonder projectleider (P)", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("P") & ","""")"
    Regel ws, r, "Zonder activanummer (X) - tijdelijk toegestaan", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("X") & ","""")"
    Regel ws, r, "Zonder afschrijftermijn (Y)", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("Y") & ","""")"
    Regel ws, r, "Dubbel: kredietnummer + activanummer", ""
    Regel ws, r, "Te controleren of afgekeurd", "=COUNTIF(" & Rg("AP") & ",""" & CtlTeControleren() & """)+COUNTIF(" & Rg("AP") & ",""Afgekeurd"")"
    Regel ws, r, "Hoofdletters of spaties in kolom I of J", "=SUMPRODUCT((" & Rg("A") & "<>"""")*(" & Rg("I") & "<>"""")*(1-EXACT(" & Rg("I") & _
          "&"""",LOWER(TRIM(" & Rg("I") & "&"""")))))+SUMPRODUCT((" & Rg("A") & "<>"""")*(" & Rg("J") & "<>"""")*(1-EXACT(" & Rg("J") & _
          "&"""",LOWER(TRIM(" & Rg("J") & "&"""")))))"
    Regel ws, r, "Werkbereik bijna vol (binnen 100 rijen)", "=IF(Par_LaatsteRij>Par_MaxRij-100,1,0)"
    Regel ws, r, "Status (H) niet L of A met jaartal", "=SUMPRODUCT((" & Rg("A") & "<>"""")*(1-(LEN(" & Rg("H") & ")=5)*((LEFT(" & Rg("H") & _
          ",1)=""L"")+(LEFT(" & Rg("H") & ",1)=""A""))*ISNUMBER(--MID(" & Rg("H") & ",2,4))))"
    Regel ws, r, "Zonder kredietnummer (C) - tijdelijk toegestaan", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("C") & ","""")"
    Regel ws, r, "7999998 in het begrotingsjaar of eerder", "=SUMPRODUCT((" & Rg("C") & "&""""=""7999998"")*(LEFT(" & Rg("H") & ",1)=""L"")*(IFERROR(--MID(" & _
          Rg("H") & ",2,4),9999)<=Par_LopendJaar+1))"
    Regel ws, r, "Handmatige formules in bedragkolommen", ""
    Regel ws, r, "SUBTOTAL-regels onder de kredieten", ""

    ' kengetallen (altijd formules)
    r = CB_LAATSTE + 2
    ws.Cells(r, 2).Value = "Kengetallen": ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    ' lopend: alleen t/m het lopend jaar (geen kredieten die pas in een toekomstig jaar starten)
    Kengetal ws, r, "=""Aantal lopende kredieten (L t/m ""&Par_LopendJaar&"")""", "=SUMPRODUCT((LEFT(" & Rg("H") & ",1)=""L"")*(IFERROR(--MID(" & _
             Rg("H") & ",2,4),9999)<=Par_LopendJaar))", "0"
    ' afgesloten per jaar: lopend jaar, de twee jaren ervoor, en eerder
    For k = 0 To 2
        Kengetal ws, r, "=""Aantal afgesloten in ""&(Par_LopendJaar-" & k & ")", "=COUNTIF(" & Rg("H") & ",""A""&(Par_LopendJaar-" & k & "))", "0"
    Next k
    Kengetal ws, r, "=""Aantal afgesloten voor ""&(Par_LopendJaar-2)", "=SUMPRODUCT((LEFT(" & Rg("H") & ",1)=""A"")*(IFERROR(--MID(" & _
             Rg("H") & ",2,4),9999)<Par_LopendJaar-2))", "0"
    Kengetal ws, r, "=""Budget t/m ""&Par_LopendJaar", "=SUM(" & Rg("Z") & ")", ""
    Kengetal ws, r, "=""Werkelijke uitgaven t/m ""&Par_LopendJaar", "=SUM(" & Rg("AA") & ")+SUM(" & Rg("AB") & ")+SUM(" & Rg("AC") & ")", ""
    Kengetal ws, r, "Restantkrediet", "=SUM(" & Rg("AD") & ")", ""
    For k = 1 To 5
        Kengetal ws, r, "=""Budget ""&(Par_LopendJaar+" & k & ")", "=SUM(" & Rg(Choose(k, "AE", "AF", "AG", "AH", "AI")) & ")", ""
    Next k
    Kengetal ws, r, "Te controleren door de administratie", "=COUNTIF(" & Rg("AP") & ",""" & CtlTeControleren() & """)", "0"
    Kengetal ws, r, "Afgekeurd", "=COUNTIF(" & Rg("AP") & ",""Afgekeurd"")", "0"

    ' rood bij een aantal > 0; zonder activanummer en zonder kredietnummer geel (tijdelijk toegestaan, telt niet mee)
    For Each deel In Array(CB_EERSTE + 5, CB_EERSTE + 12)
        Set fc = VoegCF(ws.Range(ws.Cells(deel, 3), ws.Cells(deel, 3)), "=AND(ISNUMBER(C" & deel & "),C" & deel & ">0)")
        If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 235, 156)
    Next deel
    For Each deel In Array(Array(CB_EERSTE, CB_EERSTE + 4), Array(CB_EERSTE + 6, CB_EERSTE + 11), Array(CB_EERSTE + 13, CB_LAATSTE))
        Set fc = VoegCF(ws.Range(ws.Cells(deel(0), 3), ws.Cells(deel(1), 3)), "=AND(ISNUMBER(C" & deel(0) & "),C" & deel(0) & ">0)")
        If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 150)
    Next deel
    ws.Range(ws.Cells(CB_EERSTE, 3), ws.Cells(CB_LAATSTE, 3)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(CB_EERSTE, 4), ws.Cells(CB_LAATSTE, 4)).WrapText = True
    ws.Range(ws.Cells(CB_EERSTE, 2), ws.Cells(CB_LAATSTE, 4)).VerticalAlignment = xlTop

    ' statusregel
    aantalAf = "(COUNTIF(C" & CB_EERSTE & ":C" & CB_LAATSTE & ","">0"")-N(C" & (CB_EERSTE + 5) & ">0)-N(C" & (CB_EERSTE + 12) & ">0))"
    ws.Range("C4").Formula = "=IF(" & aantalAf & "=0,""OK - alles sluit aan"",""LET OP - ""&" & aantalAf & _
                             "&"" controle(s) wijken af, zie blok 5 (rij " & CB_EERSTE & ")"")"
    ws.Range("C4").Font.Bold = True
    ws.Range("C4").FormatConditions.Delete
    Set fc = ws.Range("C4").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK - alles sluit aan""")
    fc.Interior.Color = RGB(198, 239, 206)
    Set fc = ws.Range("C4").FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""OK - alles sluit aan""")
    fc.Interior.Color = RGB(255, 150, 150)
    ControlesVernieuwen
End Sub

Private Sub Regel(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal formule As String)
    ws.Cells(r, 2).Value = label
    If Len(formule) > 0 Then ws.Cells(r, 3).Formula2 = formule Else ws.Cells(r, 3).ClearContents   ' Formula2: rekent met hele bereiken
    ws.Cells(r, 3).NumberFormat = "0"
    ws.Cells(r, 4).Font.Color = RGB(89, 89, 89)
    r = r + 1
End Sub

Private Sub Kengetal(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal formule As String, ByVal notatie As String)
    If Left$(label, 1) = "=" Then ws.Cells(r, 2).Formula = label Else ws.Cells(r, 2).Value = label
    ws.Cells(r, 3).Formula2 = formule
    ws.Cells(r, 3).NumberFormat = IIf(Len(notatie) > 0, notatie, "#,##0;-#,##0;-")
    ws.Cells(r, 3).HorizontalAlignment = xlRight
    r = r + 1
End Sub

' Rekent de controles uit die niet met een formule kunnen en zet bij elke controle in kolom D waar het zit
' (rij en kredietnummer in de Verzamellijst) en wat te doen.
Public Sub ControlesVernieuwen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, wv As Worksheet, laatste As Long, maxRij As Long, lj As Long, v As Variant, r As Long, i As Long
    Dim waar(0 To 15) As String, aant(0 To 15) As Long, hint As Variant, c As Variant, t As String, st As String, jr As Long
    Dim grp As Object, grpRij As Object, k As Variant, dubbel As Object, sleutel As String, cel As Range, eerste As String
    On Error GoTo Fout
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    If CStr(wd.Cells(CB_EERSTE, 2).Value) = "" Then Exit Sub        ' blok 5 nog niet geplaatst
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    v = wv.Range("A1:AW" & laatste).Value
    hint = Array("macro OmschrijvingenInkorten (max 36 tekens)", "kolom K van de groep (moet samen 100% zijn)", _
                 "restant moet 0 zijn; macro RestantFormuleBijwerken", "invullen via Krediet wijzigen", "invullen via Krediet wijzigen", _
                 "invullen zodra bekend", "invullen via Krediet wijzigen", "kredietnummer en activanummer controleren", _
                 "zie tabblad Controle", "macro LijstwaardenNormaliseren", "macro WerkbereikVerruimen", "status moet L of A met jaartal zijn", _
                 "invullen zodra bekend", "echt kredietnummer invullen", "macro FormulesNaarWaarden", "macro SubtotaalRegelsVerwijderen")

    Set grp = CreateObject("Scripting.Dictionary")
    Set grpRij = CreateObject("Scripting.Dictionary")
    Set dubbel = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(v(r, 1)) Then
            t = "rij " & r & " (" & IIf(Leeg(v(r, 3)), "volgnr " & v(r, 1), NrTekst(v(r, 3))) & ")"
            st = UCase$(Trim$(CStr(v(r, 8))))
            If CStr(v(r, 5)) = "fout" Or CStr(v(r, 7)) = "fout" Then Voeg waar(0), aant(0), t
            If IsGetal(v(r, 10)) And IsGetal(v(r, 11)) And Not HeeftNegatiefBudget(v, r) Then   ' bijdrageregels tellen niet mee
                grp(CStr(v(r, 10))) = grp(CStr(v(r, 10))) + CDbl(v(r, 11))
                grpRij(CStr(v(r, 10))) = grpRij(CStr(v(r, 10))) & IIf(Len(grpRij(CStr(v(r, 10)))) > 0, ", ", "") & r
            End If
            If Left$(st, 1) = "A" And Nul(v(r, 30)) <> 0 Then Voeg waar(2), aant(2), t
            If Leeg(v(r, 15)) Then Voeg waar(3), aant(3), t
            If Leeg(v(r, 16)) Then Voeg waar(4), aant(4), t
            If Leeg(v(r, 24)) Then Voeg waar(5), aant(5), t
            If Leeg(v(r, 25)) Then Voeg waar(6), aant(6), t
            If Not Leeg(v(r, 3)) And IsGetal(v(r, 24)) And NrTekst(v(r, 3)) <> "7999998" And Not HeeftNegatiefBudget(v, r) Then
                sleutel = NrTekst(v(r, 3)) & "|" & CStr(v(r, 24))
                If dubbel.Exists(sleutel) Then Voeg waar(7), aant(7), t & " = rij " & dubbel(sleutel) Else dubbel.Add sleutel, r
            End If
            If CStr(v(r, 42)) = CtlTeControleren() Or CStr(v(r, 42)) = "Afgekeurd" Then Voeg waar(8), aant(8), t
            For Each c In Array(9, 10)
                If Not Leeg(v(r, c)) And Not IsGetal(v(r, c)) Then
                    If StrComp(CStr(v(r, c)), LCase$(Trim$(CStr(v(r, c)))), vbBinaryCompare) <> 0 Then Voeg waar(9), aant(9), t: Exit For
                End If
            Next c
            If Not (Len(st) = 5 And (Left$(st, 1) = "L" Or Left$(st, 1) = "A") And IsAlleenCijfers(Mid$(st, 2))) Then Voeg waar(11), aant(11), t
            If Leeg(v(r, 3)) Then Voeg waar(12), aant(12), "rij " & r & " (volgnr " & v(r, 1) & ")"
            If NrTekst(v(r, 3)) = "7999998" And Left$(st, 1) = "L" Then
                jr = Val(Mid$(st, 2))
                If jr > 0 And jr <= lj + 1 Then Voeg waar(13), aant(13), "rij " & r & " (volgnr " & v(r, 1) & ", " & st & ")"
            End If
            For Each c In Array(26, 27, 28, 29, 31, 32, 33, 34, 35)
                If wv.Cells(r, c).HasFormula Then Voeg waar(14), aant(14), wv.Cells(r, c).Address(False, False)
            Next c
        End If
    Next r
    For Each k In grp.Keys
        If Abs(grp(k) - 1) > 0.0005 Then Voeg waar(1), aant(1), "groep " & k & " = " & Format$(grp(k) * 100, "0.0") & "% (rijen " & grpRij(k) & ")"
    Next k
    If laatste > maxRij - 100 Then Voeg waar(10), aant(10), "laatste rij " & laatste & ", werkbereik t/m " & maxRij
    Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).Find(What:="SUBTOTAL", LookIn:=xlFormulas, LookAt:=xlPart)
    If Not cel Is Nothing Then
        eerste = cel.Address
        Do
            Voeg waar(15), aant(15), cel.Address(False, False)
            Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).FindNext(cel)
        Loop While Not cel Is Nothing And cel.Address <> eerste
    End If

    For i = 0 To 15
        If aant(i) > 0 Then
            wd.Cells(CB_EERSTE + i, 4).Value = waar(i) & IIf(aant(i) > 12, " en " & (aant(i) - 12) & " meer", "") & "  >  " & hint(i)
        Else
            wd.Cells(CB_EERSTE + i, 4).Value = ""
        End If
    Next i
    ' aantallen van de controles zonder formule
    wd.Cells(CB_EERSTE + 1, 3).Value = aant(1)
    wd.Cells(CB_EERSTE + 7, 3).Value = aant(7)
    wd.Cells(CB_EERSTE + 14, 3).Value = aant(14)
    wd.Cells(CB_EERSTE + 15, 3).Value = aant(15)
    wd.Range(wd.Cells(CB_EERSTE, 2), wd.Cells(CB_LAATSTE, 2)).Rows.AutoFit
    wd.Range("C70").Value = Now
    If ActiveSheet.Name = "Controle" Then wd.Activate
    Exit Sub
Fout:
    wd.Range("C70").Value = "fout: " & Err.Description
End Sub

' Voegt een vindplaats toe aan de lijst (maximaal 12 tonen) en telt mee.
Private Sub Voeg(ByRef lijst As String, ByRef n As Long, ByVal t As String)
    n = n + 1
    If n <= 12 Then lijst = lijst & IIf(Len(lijst) > 0, "; ", "") & t
End Sub

Private Function HeeftNegatiefBudget(ByVal v As Variant, ByVal r As Long) As Boolean
    Dim c As Variant
    For Each c In Array(26, 31, 32, 33, 34, 35)
        If Nul(v(r, c)) < 0 Then HeeftNegatiefBudget = True: Exit Function
    Next c
End Function

' Zet handmatige formules in de bedragkolommen (Z, AA t/m AC, AE t/m AI) om naar hun uitkomst.
' De uitkomst verandert niet; de oude formule gaat naar het Logboek. Let op: een formule die naar een andere cel
' verwijst (bijv. =-ROUND(Z109*20%;0)) beweegt daarna niet meer mee met die cel.
Public Sub FormulesNaarWaarden()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, c As Variant, n As Long, waarde As Variant, voor As Variant, na As Variant, k As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        For Each c In Array(26, 27, 28, 29, 31, 32, 33, 34, 35)
            If wv.Cells(r, c).HasFormula Then n = n + 1
        Next c
    Next r
    If n = 0 Then
        MsgBox "Er staan geen handmatige formules in de bedragkolommen.", vbInformation
        Exit Sub
    End If
    If MsgBox(n & " handmatige formules omzetten naar hun uitkomst? De bedragen veranderen niet; de formules gaan naar het Logboek.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Formules naar waarden") <> vbYes Then Exit Sub
    Application.Calculate
    voor = MeetStap8(wv, laatste)
    Application.ScreenUpdating = False
    n = 0
    For r = 2 To laatste
        For Each c In Array(26, 27, 28, 29, 31, 32, 33, 34, 35)
            If wv.Cells(r, c).HasFormula Then
                waarde = wv.Cells(r, c).Value
                LogRegel "formule naar waarde", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, wv.Cells(r, c).Address(False, False), _
                         "'" & wv.Cells(r, c).Formula, waarde, ""      ' apostrof: in het Logboek als tekst, niet als formule
                wv.Cells(r, c).Value = waarde
                n = n + 1
            End If
        Next c
    Next r
    Application.Calculate
    na = MeetStap8(wv, laatste)
    Application.ScreenUpdating = True
    For k = 0 To 9
        If Round(na(k) - voor(k), 2) <> 0 Then
            MsgBox "Let op: de totalen wijken af na het omzetten (meting " & k & "). Controleer het Logboek.", vbCritical
            Exit For
        End If
    Next k
    ControlesVernieuwen
    MsgBox n & " formules omgezet naar hun uitkomst. Totalen ongewijzigd.", vbInformation
End Sub

' Herstelt omschrijvingen (kolom D) die langer zijn dan 36 tekens. Per regel een voorstel (straat -> str, daarna afkappen)
' dat u kunt aanpassen. Oude tekst gaat naar het Logboek.
Public Sub OmschrijvingenInkorten()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, oud As String, nieuw As String, n As Long
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "U heeft geen recht om dit te doen.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        oud = CStr(wv.Cells(r, 4).Value)
        If Not Leeg(wv.Cells(r, 1).Value) And Len(oud) > 36 Then
            nieuw = Replace(Replace(oud, "straat", "str"), "  ", " ")
            If Len(nieuw) > 36 Then nieuw = Left$(nieuw, 36)
            Do
                nieuw = InputBox("Volgnr " & wv.Cells(r, 1).Value & " (" & wv.Cells(r, 3).Value & ")" & vbCrLf & _
                                 "Huidig (" & Len(oud) & " tekens): " & oud & vbCrLf & vbCrLf & _
                                 "Nieuwe omschrijving, maximaal 36 tekens:", "Omschrijving te lang", nieuw)
                If Len(nieuw) = 0 Then Exit Do               ' annuleren = deze regel overslaan
            Loop While Len(nieuw) > 36
            If Len(nieuw) > 0 And nieuw <> oud Then
                LogRegel "omschrijving ingekort", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Omschrijving", oud, nieuw, "rij " & r
                wv.Cells(r, 4).Value = nieuw
                n = n + 1
            End If
        End If
    Next r
    Application.Calculate
    MsgBox n & " omschrijving(en) ingekort.", vbInformation
End Sub

' Haalt de oude SUBTOTAL-regels onder de kredieten weg (bijv. rij 1474). De kengetallen staan nu in blok 5 op het Dashboard.
Public Sub SubtotaalRegelsVerwijderen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, laatste As Long, maxRij As Long, cel As Range, lijst As Collection, c As Variant, n As Long, eerste As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    Set lijst = New Collection
    Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).Find(What:="SUBTOTAL", LookIn:=xlFormulas, LookAt:=xlPart)
    If Not cel Is Nothing Then
        eerste = cel.Address
        Do
            lijst.Add cel.Address(False, False)
            Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).FindNext(cel)
        Loop While Not cel Is Nothing And cel.Address <> eerste
    End If
    If lijst.Count = 0 Then
        MsgBox "Er staan geen SUBTOTAL-regels onder de kredieten.", vbInformation
        Exit Sub
    End If
    If MsgBox(lijst.Count & " SUBTOTAL-cellen onder de kredieten leegmaken?", vbYesNo + vbQuestion, "SUBTOTAL-regels") <> vbYes Then Exit Sub
    For Each c In lijst
        LogRegel "SUBTOTAL-regel verwijderd", , , CStr(c), "'" & wv.Range(c).Formula, "", ""
        wv.Range(c).ClearContents
        n = n + 1
    Next c
    ControlesVernieuwen
    MsgBox n & " SUBTOTAL-cellen leeggemaakt.", vbInformation
End Sub
