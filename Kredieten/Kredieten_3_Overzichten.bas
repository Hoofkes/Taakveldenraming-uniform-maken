Attribute VB_Name = "Kredieten_3_Overzichten"
Option Explicit

' Vervolg van de macro's voor het Dashboard investeringskredieten (gesplitst vanwege de grootte; declaraties in Kredieten_1_Basis).

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
                UvStandVerstuurd CStr(nm)                ' stand per budgethouder: verstuurd op
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
            If w.Visible = xlSheetVisible And w.Name <> "Voertuigen" And w.Name <> "Controle" And w.Name <> BL_KENG _
               And w.Index > ThisWorkbook.Worksheets("Voertuigen").Index Then moet = True
        Next w
        If moet Then
            ThisWorkbook.Worksheets("Voertuigen").Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
            ThisWorkbook.Worksheets("Controle").Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        End If
    End If
    ' Kengetallen helemaal achteraan
    If BladBestaatStil(BL_KENG) Then
        If ThisWorkbook.Worksheets(BL_KENG).Index <> ThisWorkbook.Worksheets.Count Then
            ThisWorkbook.Worksheets(BL_KENG).Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
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
Public Function Vz() As Worksheet
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
Public Sub KnopTerug(ByVal ws As Worksheet, ByVal cel As Range)
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
    b.PrintObject = False
    ws.Shapes("uitleg_btnTerug").Delete
    KnopUitleg ws, "btnTerug", "Terug naar dashboard: gaat naar het tabblad Dashboard."
    ' blauwe overzichtbladen: ook knoppen voor pdf en mail
    If IsBlauwBlad(ws.Name) Then
        ws.Buttons("btnPdf").Delete
        ws.Buttons("btnMail").Delete
        ws.Shapes("uitleg_btnPdf").Delete
        ws.Shapes("uitleg_btnMail").Delete
        Set b = ws.Buttons.Add(cel.Left + 136, cel.Top + 1, 110, 18)
        b.Name = "btnPdf"
        b.Caption = "Opslaan als pdf"
        b.OnAction = "BladNaarPdf"
        b.Placement = xlFreeFloating
        b.PrintObject = False
        Set b = ws.Buttons.Add(cel.Left + 252, cel.Top + 1, 110, 18)
        b.Name = "btnMail"
        b.Caption = "Mailen als pdf"
        b.OnAction = "BladMailen"
        b.Placement = xlFreeFloating
        b.PrintObject = False
        KnopUitleg ws, "btnPdf", "Opslaan als pdf: maakt een pdf van dit tabblad (A4 liggend, elke pagina begint met de kopregel). " & _
                   "De kolom Melding staat er niet in. U kiest zelf waar de pdf komt."
        KnopUitleg ws, "btnMail", "Mailen als pdf: maakt de pdf en opent een nieuw Outlook-bericht met de pdf als bijlage. " & _
                   "Vul zelf de ontvanger in en klik op Verzenden."
    End If
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
    If ws.Rows(2).RowHeight < 22 Then ws.Rows(2).RowHeight = 22
    Set b = ws.Buttons.Add(ws.Range("H2").Left + 22, ws.Range("H2").Top + 2, 90, 18)   ' ruimte voor het keuzepijltje van G2
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
Public Sub Alert7999998(ByVal vS As Variant, ByVal lj As Long)
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
' het bestand en kiest daarna op het Dashboard (rol = Budgetbeheerder en naam in blok 2) deze knop. De macro leest het
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
    If LCase$(Trim$(CStr(wd.Range("C" & NM_ROL).Value))) <> "budgetbeheerder" Or Len(Trim$(CStr(wd.Range("C" & NM_NAAM).Value))) = 0 Then
        MsgBox "Kies in blok 2 bij 'rol' Budgetbeheerder en bij 'naam' de teamleider van wie de uitvraag verwerkt moet worden.", vbExclamation
        Exit Sub
    End If
    naam = Trim$(CStr(wd.Range("C" & NM_NAAM).Value))
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

' ======================================================================
' Uitvraag kadernota/begroting verwerken, per budgethouder
' ======================================================================
' 1. 'Uitvraag verwerken' (blok 2, rol Budgetbeheerder + naam): leest het nieuwste ingevulde uitvraagbestand van die
'    budgethouder en zet per investering met een wijziging een regel op het blad 'Verwerking uitvraag'.
' 2. De financieel adviseur kiest per regel een besluit (goedkeuren / afwijzen / nog niet) en kan de bedragen,
'    de nieuwe omschrijving en de onderbouwing aanpassen (gele cellen). Hij heeft het laatste woord.
' 3. 'Besluiten verwerken' (op dat blad): alleen goedgekeurde regels gaan naar de Verzamellijst (budget per jaarschijf,
'    omschrijving, toelichting) en naar het register 'Mutaties kadernota-begroting'; voorstellen voor nieuwe
'    kredieten naar 'Voorstellen nieuwe kredieten'. Nulmeting vooraf en achteraf.
' 4. 'Stand uitvraag' toont per budgethouder: verstuurd, ingeleverd, verwerkt of nog open.
' 5. 'Begroting vastgesteld': pas dan gaan de verwerkte wijzigingen ter controle naar de financiele administratie.

Private Function UvKol(ByVal naam As String) As Long
    Select Case naam
        Case "volgnr": UvKol = 1
        Case "nr": UvKol = 2
        Case "oms": UvKol = 3
        Case "nieuweoms": UvKol = 4
        Case "was": UvKol = 5            ' 5 t/m 9
        Case "wordt": UvKol = 10         ' 10 t/m 14
        Case "wijziging": UvKol = 15
        Case "onderbouwing": UvKol = 16
        Case "besluit": UvKol = 17
        Case "opmerking": UvKol = 18
        Case "status": UvKol = 19
    End Select
End Function

' Stand per budgethouder (blad 'Stand uitvraag'). Geeft de rij van de budgethouder; maakt die zo nodig aan.
Private Function UvStandRij(ByVal naam As String, ByVal maken As Boolean) As Long
    Dim ws As Worksheet, r As Long
    If Not BladBestaatStil(BL_UVSTAND) Then
        If Not maken Then Exit Function
        Set ws = NieuwBlad()
        ws.Name = BL_UVSTAND
        ws.Tab.Color = RGB(237, 125, 49)
        ws.Cells.Font.Name = "Arial"
        ws.Cells.Font.Size = 10
        ws.Range("A1").Value = "Stand uitvraag kadernota / begroting per budgethouder"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Color = RGB(0, 73, 118)
        ws.Range("A4:I4").Value = Array("Budgethouder", "Uitvraag verstuurd op", "Bestand ingeleverd op", "Ingelezen op", _
                                        "Goedgekeurd", "Afgewezen", "Nog open", "Afgerond op", "Status")
        With ws.Range("A4:I4")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 73, 118)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(4).RowHeight = 30
        ws.Columns("A").ColumnWidth = 28: ws.Range("B:D").ColumnWidth = 16: ws.Range("E:G").ColumnWidth = 11
        ws.Columns("H").ColumnWidth = 14: ws.Columns("I").ColumnWidth = 34
        ws.Range("B:D,H:H").NumberFormat = "dd-mm-yyyy hh:mm"
        KnopTerug ws, ws.Range("A2")
    Else
        Set ws = ThisWorkbook.Worksheets(BL_UVSTAND)
        ws.Unprotect WACHTWOORD
    End If
    For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), naam, vbTextCompare) = 0 Then UvStandRij = r: Exit Function
    Next r
    If maken Then
        r = Application.Max(5, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1)
        ws.Cells(r, 1).Value = naam
        UvStandRij = r
    End If
End Function

' Wordt aangeroepen door UitvraagBegroting: nieuwe ronde, per budgethouder 'verstuurd op' vastleggen.
Private Sub UvStandVerstuurd(ByVal naam As String)
    Dim ws As Worksheet, r As Long
    r = UvStandRij(naam, True)
    Set ws = ThisWorkbook.Worksheets(BL_UVSTAND)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 9)).ClearContents
    ws.Cells(r, 2).Value = Now
    ws.Cells(r, 9).Value = "verstuurd, nog niet ingeleverd"
End Sub

' Nieuwste uitvraagbestand van een budgethouder (leeg = geen).
Private Function UvBestand(ByVal naam As String, ByRef dt As Date) As String
    Dim map As String, f As String
    map = MapStambestand("02 Budgetbeheerders en projectleiders")
    If Len(map) = 0 Then Exit Function
    map = SubMap(map, naam)
    If Len(map) = 0 Then Exit Function
    f = Dir(map & "Uitvraag kadernota-begroting*.xlsx")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then
            If Len(UvBestand) = 0 Or FileDateTime(map & f) > dt Then UvBestand = map & f: dt = FileDateTime(map & f)
        End If
        f = Dir()
    Loop
End Function

' Knop Stand uitvraag (blok 2): werkt 'ingeleverd op' bij uit de bestanden en toont het blad.
Public Sub UitvraagStand()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, r As Long, naam As String, dt As Date, bestand As String, verstuurd As Variant
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    If Not BladBestaatStil(BL_UVSTAND) Then
        MsgBox "Er is nog geen uitvraag verstuurd (knop 'Uitvraag kadernota/begroting').", vbInformation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(BL_UVSTAND)
    ws.Unprotect WACHTWOORD
    For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        naam = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(naam) > 0 Then
            verstuurd = ws.Cells(r, 2).Value
            dt = 0
            bestand = UvBestand(naam, dt)
            ' ingeleverd = het bestand is na het versturen nog opgeslagen (ingevuld door de budgethouder)
            If Len(bestand) > 0 And IsDate(verstuurd) Then
                If dt > CDate(verstuurd) + TimeSerial(0, 5, 0) Then ws.Cells(r, 3).Value = dt
            End If
            ws.Cells(r, 9).Value = UvStatusTekst(ws, r)
        End If
    Next r
    UvStandOpmaak ws
    ws.Activate
End Sub

Private Function UvStatusTekst(ByVal ws As Worksheet, ByVal r As Long) As String
    If Not IsEmpty(ws.Cells(r, 8).Value) Then
        UvStatusTekst = "afgerond"
    ElseIf Not IsEmpty(ws.Cells(r, 4).Value) Then
        UvStatusTekst = "ingelezen, " & Nul(ws.Cells(r, 7).Value) & " regel(s) nog open"
    ElseIf Not IsEmpty(ws.Cells(r, 3).Value) Then
        UvStatusTekst = "ingeleverd, nog verwerken"
    ElseIf Not IsEmpty(ws.Cells(r, 2).Value) Then
        UvStatusTekst = "nog niet ingeleverd - budgethouder aanspreken"
    Else
        UvStatusTekst = ""
    End If
End Function

Private Sub UvStandOpmaak(ByVal ws As Worksheet)
    Dim r As Long, lr As Long, t As String
    lr = Application.Max(5, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
    ws.Range("A5:I" & lr).Interior.ColorIndex = xlNone
    For r = 5 To lr
        t = CStr(ws.Cells(r, 9).Value)
        If t Like "nog niet ingeleverd*" Then
            ws.Cells(r, 9).Interior.Color = RGB(255, 150, 150)
        ElseIf t Like "ingeleverd*" Or t Like "ingelezen*" Then
            ws.Cells(r, 9).Interior.Color = RGB(255, 200, 120)
        ElseIf t = "afgerond" Then
            ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
        End If
    Next r
    ws.Range("A5:I" & lr).VerticalAlignment = xlTop
    ws.Range("E5:G" & lr).HorizontalAlignment = xlCenter
    BladBeveiligen ws
End Sub

' Knop Uitvraag verwerken (blok 2): leest het uitvraagbestand van de gekozen budgethouder in op 'Verwerking uitvraag'.
Public Sub UitvraagInlezen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, wv As Worksheet, ws As Worksheet, wb As Workbook, d As Variant, naam As String, bestand As String, dt As Date
    Dim lj As Long, laatste As Long, r As Long, i As Long, k As Long, rij As Long, vn As String, volgRij As Object
    Dim wijzig As Boolean, oms As String, toel As String, n As Long, nNieuw As Long, sr As Long, c As Long, nTot As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur verwerkt de uitvraag.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    If LCase$(Trim$(CStr(wd.Range("C" & NM_ROL).Value))) <> "budgetbeheerder" Or Len(Trim$(CStr(wd.Range("C" & NM_NAAM).Value))) = 0 Then
        MsgBox "Kies in blok 2 bij 'Overzicht op naam: rol' Budgetbeheerder en bij 'naam' de budgethouder.", vbExclamation
        Exit Sub
    End If
    naam = Trim$(CStr(wd.Range("C" & NM_NAAM).Value))
    ' openstaande besluiten van een andere budgethouder?
    If BladBestaatStil(BL_UVVERW) Then
        Set ws = ThisWorkbook.Worksheets(BL_UVVERW)
        If StrComp(CStr(ws.Range("B3").Value), naam, vbTextCompare) <> 0 Then
            n = 0
            For r = 6 To ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
                If Len(CStr(ws.Cells(r, UvKol("status")).Value)) = 0 And Len(CStr(ws.Cells(r, 3).Value)) > 0 Then n = n + 1
            Next r
            If n > 0 Then
                If MsgBox("Op 'Verwerking uitvraag' staan nog " & n & " regel(s) van " & ws.Range("B3").Value & " zonder verwerking." & vbCrLf & _
                          "Die blijven open in de Stand uitvraag en kunnen later opnieuw worden ingelezen." & vbCrLf & vbCrLf & "Doorgaan met " & naam & "?", _
                          vbYesNo + vbExclamation + vbDefaultButton2, "Uitvraag verwerken") <> vbYes Then Exit Sub
            End If
        End If
    End If
    bestand = UvBestand(naam, dt)
    If Len(bestand) = 0 Then
        MsgBox "Geen uitvraagbestand gevonden voor " & naam & ". Is de uitvraag al ingeleverd?", vbExclamation
        Exit Sub
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
        MsgBox "Dit uitvraagbestand heeft een oude opzet en kan niet worden ingelezen. Maak de uitvraag opnieuw.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    Set volgRij = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(wv.Cells(r, 1).Value) Then volgRij(CStr(wv.Cells(r, 1).Value)) = r
    Next r

    ' blad opbouwen
    If BladBestaatStil(BL_UVVERW) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(BL_UVVERW).Unprotect WACHTWOORD
        ThisWorkbook.Worksheets(BL_UVVERW).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = NieuwBlad()
    ws.Name = BL_UVVERW
    ws.Tab.Color = RGB(237, 125, 49)
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "Verwerking uitvraag kadernota / begroting " & (lj + 1)
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Color = RGB(0, 73, 118)
    ws.Range("A2").Value = "Kies per regel een besluit (kolom Q). Gele cellen mag u aanpassen. Daarna: knop 'Besluiten verwerken'."
    ws.Range("A2").Font.Italic = True
    ws.Range("A3").Value = "Budgethouder:"
    ws.Range("B3").Value = naam
    ws.Range("B3").Font.Bold = True
    ws.Range("D3").Value = Dir(bestand) & " (opgeslagen " & Format$(dt, "dd-mm-yyyy hh:mm") & ")"
    ws.Range("D3").Font.Color = RGB(89, 89, 89)
    ws.Range("A5:S5").Value = Array("Volgnr", "Kredietnummer", "Omschrijving (nu)", "Nieuwe omschrijving", _
                                    "Was " & (lj + 1), "Was " & (lj + 2), "Was " & (lj + 3), "Was " & (lj + 4), "Was " & (lj + 5), _
                                    "Wordt " & (lj + 1), "Wordt " & (lj + 2), "Wordt " & (lj + 3), "Wordt " & (lj + 4), "Wordt " & (lj + 5), _
                                    "Wijziging", "Onderbouwing (wordt toelichting)", "Besluit", "Opmerking financieel adviseur", "Verwerkt")
    With ws.Range("A5:S5")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 73, 118)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(5).RowHeight = 30
    rij = 6
    For i = 6 To UBound(d, 1)
        vn = Trim$(Tkst(d(i, 15)))
        If UCase$(vn) = "NIEUW" And Len(Trim$(Tkst(d(i, 2)))) > 0 Then
            ws.Cells(rij, 1).Value = "NIEUW"
            ws.Cells(rij, 3).Value = Tkst(d(i, 2))
            ws.Cells(rij, 4).Value = "taakveld: " & Tkst(d(i, 4))
            For k = 1 To 5
                ws.Cells(rij, UvKol("was") + k - 1).Value = 0
                ws.Cells(rij, UvKol("wordt") + k - 1).Value = Nul(d(i, 6 + k))
            Next k
            ws.Cells(rij, UvKol("wijziging")).Value = "Voorstel nieuw krediet (bestaand beleid)"
            ws.Cells(rij, UvKol("onderbouwing")).Value = Tkst(d(i, 14))
            nNieuw = nNieuw + 1
            rij = rij + 1
        ElseIf Len(vn) > 0 And volgRij.Exists(vn) Then
            r = volgRij(vn)
            wijzig = False
            For k = 1 To 5
                If Round(Nul(d(i, 6 + k)) - Nul(wv.Cells(r, 30 + k).Value), 2) <> 0 Then wijzig = True
            Next k
            oms = Trim$(Tkst(d(i, 12)))
            toel = Trim$(Tkst(d(i, 14)))
            ' de hele was-lijst komt op het blad (alle lopende kredieten van de budgethouder), ook zonder wijziging
            If True Then
                ws.Cells(rij, 1).Value = wv.Cells(r, 1).Value
                ws.Cells(rij, 2).NumberFormat = "@"
                ws.Cells(rij, 2).Value = NrTekst(wv.Cells(r, 3).Value)
                ws.Cells(rij, 3).Value = wv.Cells(r, 4).Value
                ws.Cells(rij, 4).Value = oms
                For k = 1 To 5
                    ws.Cells(rij, UvKol("was") + k - 1).Value = Nul(wv.Cells(r, 30 + k).Value)
                    ws.Cells(rij, UvKol("wordt") + k - 1).Value = Nul(d(i, 6 + k))
                    ' budget in de verzamellijst sinds de uitvraag gewijzigd: oranje
                    If Round(Nul(d(i, 15 + k)) - Nul(wv.Cells(r, 30 + k).Value), 2) <> 0 Then ws.Cells(rij, UvKol("was") + k - 1).Interior.Color = RGB(255, 200, 120)
                Next k
                ws.Cells(rij, UvKol("wijziging")).Value = Tkst(d(i, 13))
                ws.Cells(rij, UvKol("onderbouwing")).Value = toel
                If wijzig Or Len(oms) > 0 Or Len(toel) > 0 Then
                    n = n + 1
                    ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 3)).Font.Bold = True          ' gewijzigd: vet
                Else
                    ws.Cells(rij, UvKol("wijziging")).Value = GEEN_WIJZ
                    ws.Cells(rij, UvKol("wijziging")).Font.Color = RGB(150, 150, 150)
                End If
                nTot = nTot + 1
                rij = rij + 1
            End If
        End If
    Next i
    If rij = 6 Then
        ws.Cells(6, 3).Value = "Geen lopende kredieten en geen voorstellen in het bestand van " & naam & "."
        rij = 7
    End If
    ' opmaak en invoercellen
    ws.Cells.Locked = True
    If rij > 6 And nTot + nNieuw > 0 Then
        With ws.Range(ws.Cells(6, UvKol("besluit")), ws.Cells(rij - 1, UvKol("besluit")))
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="goedkeuren,afwijzen,nog niet"
        End With
        ' gewijzigde regels en voorstellen: 'nog niet' (u beslist); ongewijzigde regels: 'goedkeuren' (er verandert niets)
        For r = 6 To rij - 1
            If CStr(ws.Cells(r, UvKol("wijziging")).Value) = GEEN_WIJZ Then
                ws.Cells(r, UvKol("besluit")).Value = "goedkeuren"
            Else
                ws.Cells(r, UvKol("besluit")).Value = "nog niet"
            End If
        Next r
        Dim c2 As Variant
        For Each c2 In Array(UvKol("nieuweoms"), UvKol("onderbouwing"), UvKol("besluit"), UvKol("opmerking"))
            ws.Range(ws.Cells(6, c2), ws.Cells(rij - 1, c2)).Interior.Color = RGB(255, 255, 204)
            ws.Range(ws.Cells(6, c2), ws.Cells(rij - 1, c2)).Locked = False
        Next c2
        ws.Range(ws.Cells(6, UvKol("wordt")), ws.Cells(rij - 1, UvKol("wordt") + 4)).Interior.Color = RGB(255, 255, 204)
        ws.Range(ws.Cells(6, UvKol("wordt")), ws.Cells(rij - 1, UvKol("wordt") + 4)).Locked = False
    End If
    ws.Range(ws.Cells(6, UvKol("was")), ws.Cells(rij, UvKol("wordt") + 4)).NumberFormat = "#,##0;-#,##0;-"
    ws.Columns(1).ColumnWidth = 8: ws.Columns(2).ColumnWidth = 12: ws.Columns(3).ColumnWidth = 34: ws.Columns(4).ColumnWidth = 26
    ws.Range(ws.Columns(UvKol("was")), ws.Columns(UvKol("wordt") + 4)).ColumnWidth = 12
    ws.Columns(UvKol("wijziging")).ColumnWidth = 30: ws.Columns(UvKol("onderbouwing")).ColumnWidth = 40
    ws.Columns(UvKol("besluit")).ColumnWidth = 12: ws.Columns(UvKol("opmerking")).ColumnWidth = 30: ws.Columns(UvKol("status")).ColumnWidth = 18
    ws.Range(ws.Cells(6, 3), ws.Cells(rij, 4)).WrapText = True
    ws.Range(ws.Cells(6, UvKol("wijziging")), ws.Cells(rij, UvKol("opmerking"))).WrapText = True
    ws.Range(ws.Cells(6, 1), ws.Cells(rij, UvKol("status"))).VerticalAlignment = xlTop
    ws.Range(ws.Cells(6, 1), ws.Cells(rij, 2)).HorizontalAlignment = xlLeft
    BladOpmaken ws, 5, 6, rij
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("D6").Select
    ActiveWindow.FreezePanes = True
    ' knoppen
    KnopTerug ws, ws.Range("A4")
    Dim b As Object
    Set b = ws.Buttons.Add(ws.Range("A4").Left + 140, ws.Range("A4").Top + 1, 160, 18)
    b.Name = "btnUvVerwerken"
    b.Caption = "Besluiten verwerken"
    b.OnAction = "UitvraagBesluitenVerwerken"
    b.Placement = xlFreeFloating
    b.PrintObject = False
    KnopUitleg ws, "btnUvVerwerken", "Besluiten verwerken: alle regels met besluit 'goedkeuren' gaan naar de Verzamellijst (budget per jaarschijf, " & _
               "omschrijving, toelichting) en naar het register Mutaties kadernota-begroting. 'Afwijzen' wordt vastgelegd. 'Nog niet' blijft open."
    ws.Rows(4).RowHeight = 22
    ws.Protect Password:=WACHTWOORD, DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFormattingRows:=True
    ' stand bijwerken
    sr = UvStandRij(naam, True)
    With ThisWorkbook.Worksheets(BL_UVSTAND)
        .Unprotect WACHTWOORD
        If IsEmpty(.Cells(sr, 3).Value) Then .Cells(sr, 3).Value = dt
        .Cells(sr, 4).Value = Now
        .Cells(sr, 5).Value = 0: .Cells(sr, 6).Value = 0: .Cells(sr, 7).Value = nTot + nNieuw
        .Cells(sr, 8).ClearContents
        .Cells(sr, 9).Value = UvStatusTekst(ThisWorkbook.Worksheets(BL_UVSTAND), sr)
        UvStandOpmaak ThisWorkbook.Worksheets(BL_UVSTAND)
    End With
    ws.Activate
    TabbladenOrdenen
    ws.Activate
    LogRegel "uitvraag ingelezen", , , "", , , naam & " | " & Dir(bestand) & " | " & n & " wijzigingen, " & nNieuw & " voorstellen"
    Application.ScreenUpdating = True
    MsgBox naam & ": " & nTot & " lopende kredieten, waarvan " & n & " met een wijziging (vet), en " & nNieuw & " voorstel(len) voor nieuwe kredieten." & vbCrLf & vbCrLf & _
           "Ongewijzigde regels staan al op 'goedkeuren'. Kies bij de gewijzigde regels en de voorstellen een besluit in kolom Q " & _
           "en klik daarna op 'Besluiten verwerken'.", vbInformation, "Uitvraag verwerken"
    Exit Sub
Fout:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Fout bij het inlezen van de uitvraag: " & Err.Description, vbCritical
End Sub

' Knop Besluiten verwerken (blad 'Verwerking uitvraag').
Public Sub UitvraagBesluitenVerwerken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const BL_VOORSTEL As String = "Voorstellen nieuwe kredieten"
    Dim ws As Worksheet, wv As Worksheet, wsVs As Worksheet, naam As String, lj As Long, laatste As Long, r As Long, i As Long, k As Long
    Dim lr As Long, besluit As String, volgRij As Object, voor As Variant, na As Variant, verschil As Double, oudB As Double, nieuwB As Double
    Dim nGoed As Long, nAf As Long, nOpen As Long, toel As String, gebr As String, bestand As String, sr As Long, vsRij As Long, nB As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur verwerkt de uitvraag.", vbExclamation
        Exit Sub
    End If
    If Not BladBestaatStil(BL_UVVERW) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_UVVERW)
    naam = CStr(ws.Range("B3").Value)
    bestand = CStr(ws.Range("D3").Value)
    lr = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    For i = 6 To lr
        If Len(CStr(ws.Cells(i, UvKol("status")).Value)) = 0 And Len(CStr(ws.Cells(i, 3).Value)) > 0 Then
            Select Case CStr(ws.Cells(i, UvKol("besluit")).Value)
                Case "goedkeuren": nGoed = nGoed + 1
                Case "afwijzen": nAf = nAf + 1
                Case Else: nOpen = nOpen + 1
            End Select
        End If
    Next i
    If nGoed + nAf = 0 Then
        MsgBox "Er is nog geen regel met besluit 'goedkeuren' of 'afwijzen'.", vbInformation
        Exit Sub
    End If
    If MsgBox(naam & ":" & vbCrLf & nGoed & " goedkeuren, " & nAf & " afwijzen, " & nOpen & " nog niet." & vbCrLf & vbCrLf & _
              "Goedgekeurde regels nu verwerken in de Verzamellijst?", vbYesNo + vbQuestion, "Besluiten verwerken") <> vbYes Then Exit Sub

    On Error GoTo Fout
    Application.ScreenUpdating = False
    ws.Unprotect WACHTWOORD
    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    Set volgRij = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(wv.Cells(r, 1).Value) Then volgRij(CStr(wv.Cells(r, 1).Value)) = r
    Next r
    voor = MeetStap8(wv, laatste)
    For i = 6 To lr
        If Len(CStr(ws.Cells(i, UvKol("status")).Value)) = 0 And Len(CStr(ws.Cells(i, 3).Value)) > 0 Then
            besluit = CStr(ws.Cells(i, UvKol("besluit")).Value)
            If besluit = "afwijzen" Then
                ws.Cells(i, UvKol("status")).Value = "afgewezen " & Format$(Date, "dd-mm-yyyy")
                LogRegel "uitvraag afgewezen", ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, "", , , naam & " | " & CStr(ws.Cells(i, UvKol("opmerking")).Value)
            ElseIf besluit = "goedkeuren" Then
                If CStr(ws.Cells(i, 1).Value) = "NIEUW" Then
                    ' voorstel nieuw krediet: naar het blad Voorstellen nieuwe kredieten en het register
                    If Not BladBestaatStil(BL_VOORSTEL) Then
                        Set wsVs = NieuwBlad()
                        wsVs.Name = BL_VOORSTEL
                        wsVs.Cells.Font.Name = "Arial"
                        wsVs.Cells.Font.Size = 10
                        wsVs.Range("A1").Value = "Voorstellen bestaand beleid: afschrijftermijn bepalen en invoeren via de invoerkaart"
                        wsVs.Range("A1").Font.Bold = True
                        wsVs.Range("A1").Font.Size = 14
                        wsVs.Range("A4:M4").Value = Array("Goedgekeurd op", "Budgetbeheerder", "Omschrijving", "Taakveld", "Afschr", _
                                                          "Budget " & (lj + 1), "Budget " & (lj + 2), "Budget " & (lj + 3), "Budget " & (lj + 4), _
                                                          "Budget " & (lj + 5), "Toelichting", "Bestand", "Ingevoerd (ja/nee)")
                        wsVs.Range("A4:M4").Font.Bold = True
                        KnopTerug wsVs, wsVs.Range("A2")
                    Else
                        Set wsVs = ThisWorkbook.Worksheets(BL_VOORSTEL)
                        wsVs.Unprotect WACHTWOORD
                    End If
                    vsRij = Application.Max(5, wsVs.Cells(wsVs.Rows.Count, 3).End(xlUp).Row + 1)
                    wsVs.Cells(vsRij, 1).Value = Date
                    wsVs.Cells(vsRij, 2).Value = naam
                    wsVs.Cells(vsRij, 3).Value = ws.Cells(i, 3).Value
                    wsVs.Cells(vsRij, 4).Value = Replace(CStr(ws.Cells(i, 4).Value), "taakveld: ", "")
                    For k = 1 To 5
                        wsVs.Cells(vsRij, 5 + k).Value = Nul(ws.Cells(i, UvKol("wordt") + k - 1).Value)
                        If Nul(ws.Cells(i, UvKol("wordt") + k - 1).Value) <> 0 Then
                            MutatieRegistreren naam, "", "nieuw", CStr(ws.Cells(i, 3).Value), wsVs.Cells(vsRij, 4).Value, "", "", lj + k, 0, _
                                               Nul(ws.Cells(i, UvKol("wordt") + k - 1).Value), "Nieuw krediet (voorstel)", CStr(ws.Cells(i, UvKol("onderbouwing")).Value), bestand
                        End If
                    Next k
                    wsVs.Cells(vsRij, 11).Value = ws.Cells(i, UvKol("onderbouwing")).Value
                    wsVs.Cells(vsRij, 12).Value = bestand
                    wsVs.Cells(vsRij, 13).Value = "nee"
                    ws.Cells(i, UvKol("status")).Value = "naar Voorstellen " & Format$(Date, "dd-mm-yyyy")
                ElseIf volgRij.Exists(CStr(ws.Cells(i, 1).Value)) Then
                    r = volgRij(CStr(ws.Cells(i, 1).Value))
                    nB = 0
                    For k = 1 To 5
                        oudB = Nul(wv.Cells(r, 30 + k).Value)
                        nieuwB = Nul(ws.Cells(i, UvKol("wordt") + k - 1).Value)
                        If Round(nieuwB - oudB, 2) <> 0 Then
                            nB = nB + 1
                            LogRegel "uitvraag verwerkt", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, wv.Cells(1, 30 + k).Text, oudB, nieuwB, naam
                            MutatieRegistreren naam, wv.Cells(r, 1).Value, NrTekst(wv.Cells(r, 3).Value), CStr(wv.Cells(r, 4).Value), _
                                               wv.Cells(r, 19).Value, wv.Cells(r, 20).Value, wv.Cells(r, 25).Value, lj + k, oudB, nieuwB, _
                                               CStr(ws.Cells(i, UvKol("wijziging")).Value), CStr(ws.Cells(i, UvKol("onderbouwing")).Value), bestand
                            If nieuwB = 0 Then wv.Cells(r, 30 + k).ClearContents Else wv.Cells(r, 30 + k).Value = nieuwB
                            verschil = verschil + nieuwB - oudB
                        End If
                    Next k
                    If Len(Trim$(CStr(ws.Cells(i, 4).Value))) > 0 Then
                        LogRegel "uitvraag verwerkt", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Omschrijving", wv.Cells(r, 4).Value, ws.Cells(i, 4).Value, naam
                        wv.Cells(r, 4).Value = Left$(Trim$(CStr(ws.Cells(i, 4).Value)), 36)
                        wv.Cells(r, 6).Value = KortOmschrijving(CStr(wv.Cells(r, 4).Value))
                    End If
                    toel = Trim$(CStr(ws.Cells(i, UvKol("wijziging")).Value))
                    If toel = GEEN_WIJZ Then toel = ""
                    If Len(Trim$(CStr(ws.Cells(i, UvKol("onderbouwing")).Value))) > 0 Then toel = toel & IIf(Len(toel) > 0, ". ", "") & Trim$(CStr(ws.Cells(i, UvKol("onderbouwing")).Value))
                    If Len(toel) > 0 Then
                        toel = Left$(toel, MAX_TOEL)
                        LogRegel "uitvraag verwerkt", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Toelichting", wv.Cells(r, 36).Value, toel, naam
                        wv.Cells(r, 36).Value = toel
                    End If
                    If nB = 0 And Len(toel) = 0 And Len(Trim$(CStr(ws.Cells(i, 4).Value))) = 0 Then
                        ws.Cells(i, UvKol("status")).Value = "akkoord, geen wijziging"
                    Else
                        wv.Cells(r, 48).Value = gebr
                        wv.Cells(r, 49).Value = Now
                        ws.Cells(i, UvKol("status")).Value = "verwerkt " & Format$(Date, "dd-mm-yyyy")
                    End If
                Else
                    ws.Cells(i, UvKol("status")).Value = "volgnr niet gevonden"
                End If
            End If
        End If
    Next i
    Application.Calculate
    na = MeetStap8(wv, laatste)
    If na(0) <> voor(0) Or Round(na(1) - voor(1), 2) <> 0 Or Round(na(2) - voor(2), 2) <> 0 Or Round(na(3) - voor(3), 2) <> 0 _
       Or Round(na(4) - voor(4), 2) <> 0 Or Round((na(5) + na(6) + na(7) + na(8) + na(9)) - (voor(5) + voor(6) + voor(7) + voor(8) + voor(9)) - verschil, 2) <> 0 Then
        Application.ScreenUpdating = True
        MsgBox "Let op: de nameting wijkt af. Controleer de Verzamellijst en het Logboek.", vbCritical
    End If
    ' stand per budgethouder
    nGoed = 0: nAf = 0: nOpen = 0
    For i = 6 To lr
        If Len(CStr(ws.Cells(i, 3).Value)) > 0 And Len(CStr(ws.Cells(i, UvKol("besluit")).Value)) > 0 Then
            If CStr(ws.Cells(i, UvKol("status")).Value) Like "afgewezen*" Then
                nAf = nAf + 1
            ElseIf Len(CStr(ws.Cells(i, UvKol("status")).Value)) > 0 Then
                nGoed = nGoed + 1
            Else
                nOpen = nOpen + 1
            End If
        End If
    Next i
    sr = UvStandRij(naam, True)
    With ThisWorkbook.Worksheets(BL_UVSTAND)
        .Unprotect WACHTWOORD
        .Cells(sr, 5).Value = nGoed: .Cells(sr, 6).Value = nAf: .Cells(sr, 7).Value = nOpen
        If nOpen = 0 Then .Cells(sr, 8).Value = Now Else .Cells(sr, 8).ClearContents
        .Cells(sr, 9).Value = UvStatusTekst(ThisWorkbook.Worksheets(BL_UVSTAND), sr)
        UvStandOpmaak ThisWorkbook.Worksheets(BL_UVSTAND)
    End With
    ws.Protect Password:=WACHTWOORD, DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFormattingRows:=True
    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "Verwerkt. " & nGoed & " goedgekeurd, " & nAf & " afgewezen, " & nOpen & " nog open." & vbCrLf & vbCrLf & _
           "De wijzigingen gaan pas ter controle naar de financi" & ChrW(235) & "le administratie na vaststelling van de begroting " & _
           "(knop 'Begroting vastgesteld' in blok 2).", vbInformation
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout bij het verwerken: " & Err.Description, vbCritical
End Sub

' Knop Begroting vastgesteld (blok 2): alle via de uitvraag verwerkte wijzigingen die nog niet ter controle zijn
' aangeboden, gaan naar de financiele administratie (status 'Te controleren'). Vastgelegd in het register (kolom P).
Public Sub UitvraagNaarControle()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wr As Worksheet, wv As Worksheet, r As Long, lr As Long, vn As Variant, v As Variant, kred As Object, n As Long, k As Variant
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    If Not BladBestaatStil(BL_MUTREG) Then
        MsgBox "Er zijn nog geen verwerkte wijzigingen uit de uitvraag.", vbInformation
        Exit Sub
    End If
    Set wr = ThisWorkbook.Worksheets(BL_MUTREG)
    Set wv = Vz()
    wr.Unprotect WACHTWOORD
    If CStr(wr.Range("P4").Value) = "" Then wr.Range("P4").Value = "Ter controle aangeboden"
    lr = wr.Cells(wr.Rows.Count, 1).End(xlUp).Row
    Set kred = CreateObject("Scripting.Dictionary")
    For r = 5 To lr
        vn = wr.Cells(r, 4).Value
        If Len(CStr(wr.Cells(r, 16).Value)) = 0 And IsNumeric(vn) And Not IsEmpty(vn) Then kred(CStr(vn)) = True
    Next r
    If kred.Count = 0 Then
        MsgBox "Alle verwerkte wijzigingen zijn al ter controle aangeboden.", vbInformation
        BladBeveiligen wr
        Exit Sub
    End If
    If MsgBox("Is de begroting vastgesteld?" & vbCrLf & vbCrLf & kred.Count & " krediet(en) met wijzigingen uit de uitvraag gaan nu ter controle " & _
              "naar de financi" & ChrW(235) & "le administratie.", vbYesNo + vbQuestion + vbDefaultButton2, "Begroting vastgesteld") <> vbYes Then
        BladBeveiligen wr
        Exit Sub
    End If
    For Each k In kred.Keys
        v = Application.Match(CDbl(k), wv.Columns(1), 0)
        If Not IsError(v) Then
            wv.Cells(CLng(v), 42).Value = CtlTeControleren()
            wv.Cells(CLng(v), 43).ClearContents
            wv.Cells(CLng(v), 44).ClearContents
            wv.Cells(CLng(v), 45).Value = "uitvraag kadernota/begroting na vaststelling begroting"
            LogRegel "ter controle na vaststelling begroting", wv.Cells(CLng(v), 1).Value, wv.Cells(CLng(v), 3).Value, "", , , ""
            n = n + 1
        End If
    Next k
    For r = 5 To lr
        vn = wr.Cells(r, 4).Value
        If Len(CStr(wr.Cells(r, 16).Value)) = 0 And IsNumeric(vn) And Not IsEmpty(vn) Then wr.Cells(r, 16).Value = "aangeboden " & Format$(Date, "dd-mm-yyyy")
    Next r
    BladBeveiligen wr
    MsgBox n & " krediet(en) staan nu bij de financi" & ChrW(235) & "le administratie ter controle.", vbInformation
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

Public Function IsOpbrengstActiva(ByVal toel As Variant) As Boolean
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
Public Function NieuwBlad() As Worksheet
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

' Tabblad Kengetallen (achteraan): de kengetallen als formules op de Verzamellijst (altijd actueel).
Private Sub KengetallenBladMaken()
    Dim ws As Worksheet, r As Long, k As Long
    If BladBestaatStil(BL_KENG) Then
        Set ws = ThisWorkbook.Worksheets(BL_KENG)
        ws.Unprotect WACHTWOORD
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = BL_KENG
    End If
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 18
    ws.Range("B1").Value = "Kengetallen investeringskredieten"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B2").Value = "Formules op de Verzamellijst: altijd actueel."
    ws.Range("B2").Font.Italic = True
    r = 4
    ' lopend: alleen t/m het lopend jaar (geen kredieten die pas in een toekomstig jaar starten)
    Kengetal ws, r, "=""Aantal lopende kredieten (L t/m ""&Par_LopendJaar&"")""", "=SUMPRODUCT((LEFT(" & Rg("H") & ",1)=""L"")*(IFERROR(--MID(" & _
             Rg("H") & ",2,4),9999)<=Par_LopendJaar))", "0"
    ' afgesloten per jaar: lopend jaar, de twee jaren ervoor, en eerder
    For k = 0 To 2
        Kengetal ws, r, "=""Aantal afgesloten in ""&(Par_LopendJaar-" & k & ")", "=COUNTIF(" & Rg("H") & ",""A""&(Par_LopendJaar-" & k & "))", "0"
    Next k
    Kengetal ws, r, "=""Aantal afgesloten voor ""&(Par_LopendJaar-2)", "=SUMPRODUCT((LEFT(" & Rg("H") & ",1)=""A"")*(IFERROR(--MID(" & _
             Rg("H") & ",2,4),9999)<Par_LopendJaar-2))", "0"
    r = r + 1
    Kengetal ws, r, "=""Budget t/m ""&Par_LopendJaar", "=SUM(" & Rg("Z") & ")", ""
    Kengetal ws, r, "=""Werkelijke uitgaven t/m ""&Par_LopendJaar", "=SUM(" & Rg("AA") & ")+SUM(" & Rg("AB") & ")+SUM(" & Rg("AC") & ")", ""
    Kengetal ws, r, "Restantkrediet", "=SUM(" & Rg("AD") & ")", ""
    For k = 1 To 5
        Kengetal ws, r, "=""Budget ""&(Par_LopendJaar+" & k & ")", "=SUM(" & Rg(Choose(k, "AE", "AF", "AG", "AH", "AI")) & ")", ""
    Next k
    r = r + 1
    Kengetal ws, r, "Te controleren door de administratie", "=COUNTIF(" & Rg("AP") & ",""" & CtlTeControleren() & """)", "0"
    Kengetal ws, r, "Afgekeurd", "=COUNTIF(" & Rg("AP") & ",""Afgekeurd"")", "0"
    ActiveWindow.DisplayGridlines = False
    KnopTerug ws, ws.Range("D1")
    BladBeveiligen ws
End Sub

' Kolom 'Beheer' in de rollentabel (applicatiebeheerder: ziet alle blokken op het Dashboard). Wordt eenmalig
' toegevoegd; Michiel en Monique krijgen 'ja'. Daarna zelf aan te passen in de tabel op het blad Parameters.
Private Sub BeheerKolomZorgen()
    Dim wp As Worksheet, lo As ListObject, lc As ListColumn, r As ListRow, nm As String, beveiligd As Boolean
    On Error GoTo Klaar
    Set wp = ThisWorkbook.Worksheets("Parameters")
    Set lo = wp.ListObjects("Tbl_Rollen")
    On Error Resume Next
    Set lc = lo.ListColumns("Beheer")
    On Error GoTo Klaar
    If Not lc Is Nothing Then Exit Sub
    beveiligd = wp.ProtectContents
    If beveiligd Then wp.Unprotect WACHTWOORD
    Set lc = lo.ListColumns.Add
    lc.Name = "Beheer"
    For Each r In lo.ListRows
        nm = LCase$(Trim$(CStr(r.Range.Cells(1, 1).Value)))
        r.Range.Cells(1, lc.Index).Value = IIf(nm Like "michiel*" Or nm Like "monique*", "ja", "nee")
    Next r
    lc.DataBodyRange.Interior.Color = RGB(255, 255, 204)
    If beveiligd Then BladBeveiligen wp
    LogRegel "rollentabel", , , "Beheer", "", "kolom toegevoegd", "applicatiebeheerders: Michiel en Monique"
Klaar:
End Sub

' Dashboard per rol: financieel adviseur ziet blok 1 (invoerkaart) en blok 3, niet blok 4; financiele administratie
' ziet blok 4, niet blok 1 en 3 (zij controleert en past aan via de Verzamellijst). Zonder rechten: alleen blok 2 en 5.
' Applicatiebeheerder (kolom Beheer = ja in de rollentabel, standaard Michiel en Monique): ziet alles.
' Draait bij openen (na het kiezen van de gebruiker), bij Wie werkt er en bij het opbouwen van het Dashboard.
Public Sub DashboardNaarRol()
    Dim ws As Worksheet, adv As Boolean, fa As Boolean, shp As Object, beveiligd As Boolean
    On Error Resume Next
    If Not BladBestaatStil("Dashboard") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    If CStr(ws.Range("B" & NM_KOP).Value) <> "2. Financieel adviseur" Then Exit Sub   ' oud Dashboard: eerst DashboardAanmaken
    BeheerKolomZorgen
    adv = HeeftRecht("Toevoegen")
    fa = HeeftRecht("Bevestigen")
    If HeeftRecht("Beheer") Then adv = True: fa = True             ' applicatiebeheerder ziet alle blokken
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    ws.Rows("7:" & (KAART_KNOP + 1)).Hidden = Not adv                       ' blok 1 invoerkaart en knoppen eronder
    ws.Rows(eerste + 3).Hidden = True                                         ' korte omschrijving blijft altijd verborgen
    ws.Rows(NM_KOP & ":" & (FA_KOP - 1)).Hidden = Not adv                    ' blok 2 financieel adviseur
    ws.Rows(FA_KOP & ":" & (CB_KOP - 1)).Hidden = Not fa                     ' blok 3 financiele administratie
    For Each shp In ws.Shapes                                                 ' knoppen op verborgen rijen ook verbergen
        shp.Visible = Not ws.Rows(shp.TopLeftCell.Row).Hidden
    Next shp
    If beveiligd Then BladBeveiligen ws
End Sub

' Knop Naam vervangen (bovenaan het Dashboard): vervangt in een keer een budgetbeheerder, projectleider of
' tijdschrijver door een opvolger, in alle regels van de Verzamellijst waar die naam staat. Bij tijdschrijvers
' (meerdere namen in een cel) wordt alleen die ene naam vervangen. Elke wijziging gaat naar het Logboek.
Public Sub BudgethouderVervangen()
    NaamVervangenSoort "1"
End Sub

Public Sub ProjectleiderVervangen()
    NaamVervangenSoort "2"
End Sub

Public Sub TijdschrijverVervangen()
    NaamVervangenSoort "3"
End Sub

Public Sub NaamVervangen()
    NaamVervangenSoort ""
End Sub

Private Sub NaamVervangenSoort(ByVal keuze As String)
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, kols As Variant, oud As String, nieuw As String, alleenL As VbMsgBoxResult, rol As String
    Dim r As Long, laatste As Long, c As Variant, n As Long, waarde As String, nw As String, deel As Variant, gevonden As Boolean
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "U mag geen kredieten wijzigen.", vbExclamation
        Exit Sub
    End If
    If Len(keuze) = 0 Then
        keuze = Trim$(InputBox("Welke naam vervangen?" & vbCrLf & "1 = budgetbeheerder (kolom O)" & vbCrLf & "2 = projectleider (kolom P)" & vbCrLf & _
                               "3 = tijdschrijver (kolom Q)" & vbCrLf & "4 = in alle drie", "Naam vervangen", "4"))
    End If
    Select Case keuze
        Case "1": kols = Array(15): rol = "budgethouder"
        Case "2": kols = Array(16): rol = "projectleider"
        Case "3": kols = Array(17): rol = "tijdschrijver"
        Case "4": kols = Array(15, 16, 17): rol = "naam"
        Case "": Exit Sub
        Case Else
            MsgBox "Typ 1, 2, 3 of 4.", vbExclamation
            Exit Sub
    End Select
    oud = Trim$(InputBox("Welke " & rol & " moet worden vervangen? (naam precies zoals in de lijst)", "Vervangen: " & rol))
    If Len(oud) = 0 Then Exit Sub
    nieuw = Trim$(InputBox("Door wie wordt " & oud & " opgevolgd? (naam van de nieuwe " & rol & ")", "Vervangen: " & rol))
    If Len(nieuw) = 0 Then Exit Sub
    alleenL = MsgBox("Alleen bij lopende kredieten (L) vervangen?" & vbCrLf & vbCrLf & "Ja = alleen lopende kredieten (aanbevolen; " & _
                     "bij afgesloten kredieten blijft de oude naam staan)" & vbCrLf & "Nee = ook bij afgesloten kredieten (A)", _
                     vbYesNoCancel + vbQuestion, "Naam vervangen")
    If alleenL = vbCancel Then Exit Sub
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    ' eerst tellen
    For r = 2 To laatste
        If Not IsEmpty(wv.Cells(r, 1).Value) And (alleenL = vbNo Or UCase$(Left$(Trim$(CStr(wv.Cells(r, 8).Value)), 1)) = "L") Then
            For Each c In kols
                If NaamKomtVoor(CStr(wv.Cells(r, CLng(c)).Value), oud) Then n = n + 1
            Next c
        End If
    Next r
    If n = 0 Then
        MsgBox "De naam '" & oud & "' komt niet voor in de gekozen kolom(men).", vbInformation
        Exit Sub
    End If
    If MsgBox("'" & oud & "' vervangen door '" & nieuw & "' in " & n & " cel(len)?", vbYesNo + vbQuestion, "Naam vervangen") <> vbYes Then Exit Sub
    n = 0
    For r = 2 To laatste
        If Not IsEmpty(wv.Cells(r, 1).Value) And (alleenL = vbNo Or UCase$(Left$(Trim$(CStr(wv.Cells(r, 8).Value)), 1)) = "L") Then
            For Each c In kols
                waarde = CStr(wv.Cells(r, CLng(c)).Value)
                If NaamKomtVoor(waarde, oud) Then
                    nw = "": gevonden = False
                    For Each deel In Split(Replace(waarde, ";", ","), ",")
                        If StrComp(Trim$(CStr(deel)), oud, vbTextCompare) = 0 Then deel = nieuw
                        If InStr(1, ", " & nw & ",", ", " & Trim$(CStr(deel)) & ",", vbTextCompare) = 0 Then
                            nw = nw & IIf(Len(nw) > 0, ", ", "") & Trim$(CStr(deel))
                        End If
                    Next deel
                    wv.Cells(r, CLng(c)).Value = nw
                    LogRegel "naam vervangen", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, CStr(wv.Cells(1, CLng(c)).Value), waarde, nw, "rij " & r
                    n = n + 1
                End If
            Next c
        End If
    Next r
    KeuzelijstenVerversen True
    MsgBox "'" & oud & "' is vervangen door '" & nieuw & "' in " & n & " cel(len). Alles staat in het Logboek.", vbInformation
End Sub

Public Sub ControleblokPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, r As Long, k As Long, fc As Object, b As Object, deel As Variant, aantalAf As String
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("A" & CB_KOP & ":H" & (CB_LAATSTE + 40)).Clear
    ws.Range("B" & CB_KOP).Value = "4. Controles"
    ws.Range("B" & CB_KOP).Font.Bold = True
    ws.Range("B" & CB_KOP).Font.Size = 12
    ws.Range("B" & (CB_KOP + 1)).Value = "Controles laatst vernieuwd"
    ws.Range("C" & (CB_KOP + 1)).NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Range("C" & (CB_KOP + 1)).HorizontalAlignment = xlLeft
    On Error Resume Next
    ws.Buttons("btnControles").Delete
    ws.Shapes("uitleg_btnControles").Delete
    On Error GoTo 0
    ws.Rows(CB_KOP + 1).RowHeight = 24               ' knop past binnen de rij en loopt niet door de kopregel eronder
    Set b = ws.Buttons.Add(ws.Range("E" & (CB_KOP + 1)).Left + 2, ws.Range("E" & (CB_KOP + 1)).Top + 3, 170, 18)
    b.Name = "btnControles"
    b.Caption = "Controles vernieuwen"
    b.OnAction = "ControlesVernieuwen"
    ws.Range("B" & (CB_KOP + 2) & ":C" & (CB_KOP + 2)).Value = Array("Controle (rood = nakijken)", "Aantal")
    ws.Range("E" & (CB_KOP + 2)).Value = "Waar (rij in de Verzamellijst) / wat te doen"
    ws.Range("B" & (CB_KOP + 2) & ":F" & (CB_KOP + 2)).Font.Bold = True
    ws.Range("B" & (CB_KOP + 2) & ":F" & (CB_KOP + 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous

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

    KengetallenBladMaken

    ' rood bij een aantal > 0; zonder activanummer en zonder kredietnummer geel (tijdelijk toegestaan, telt niet mee)
    For Each deel In Array(CB_EERSTE + 5, CB_EERSTE + 12)
        Set fc = VoegCF(ws.Range(ws.Cells(deel, 3), ws.Cells(deel, 3)), "=AND(ISNUMBER(C" & deel & "),C" & deel & ">0)")
        If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 235, 156)
    Next deel
    For Each deel In Array(Array(CB_EERSTE, CB_EERSTE + 4), Array(CB_EERSTE + 6, CB_EERSTE + 11), Array(CB_EERSTE + 13, CB_LAATSTE))
        Set fc = VoegCF(ws.Range(ws.Cells(deel(0), 3), ws.Cells(deel(1), 3)), "=AND(ISNUMBER(C" & deel(0) & "),C" & deel(0) & ">0)")
        If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 150)
    Next deel
    ws.Range(ws.Cells(CB_EERSTE, 3), ws.Cells(CB_LAATSTE, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(CB_EERSTE, 5), ws.Cells(CB_LAATSTE, 5)).WrapText = False      ' tekst loopt door naar rechts
    ws.Range(ws.Cells(CB_EERSTE, 2), ws.Cells(CB_LAATSTE, 5)).VerticalAlignment = xlTop
    ' omschrijving rood als de controle afwijkt (zodat direct te zien is wat nagekeken moet worden)
    For Each deel In Array(Array(CB_EERSTE, CB_EERSTE + 4), Array(CB_EERSTE + 6, CB_EERSTE + 11), Array(CB_EERSTE + 13, CB_LAATSTE))
        Set fc = VoegCF(ws.Range(ws.Cells(deel(0), 2), ws.Cells(deel(1), 2)), "=AND(ISNUMBER($C" & deel(0) & "),$C" & deel(0) & ">0)")
        If Not fc Is Nothing Then fc.Font.Color = RGB(192, 0, 0): fc.Font.Bold = True
    Next deel

    ' statusregel
    aantalAf = "(COUNTIF(C" & CB_EERSTE & ":C" & CB_LAATSTE & ","">0"")-N(C" & (CB_EERSTE + 5) & ">0)-N(C" & (CB_EERSTE + 12) & ">0))"
    ws.Range("C4").Formula = "=IF(" & aantalAf & "=0,""OK - alles sluit aan"",""LET OP - ""&" & aantalAf & _
                             "&"" controle(s) wijken af: zie blok 4 onderaan (vanaf rij " & CB_EERSTE & ", rood = nakijken)"")"
    ws.Range("C4").Font.Bold = True
    ' klik op 'Controlestatus' (B4) = naar blok 5
    ws.Range("B4").Hyperlinks.Delete
    ws.Hyperlinks.Add Anchor:=ws.Range("B4"), Address:="", SubAddress:="'Dashboard'!B" & CB_KOP, TextToDisplay:="Controlestatus", _
                      ScreenTip:="Naar blok 4: controles"
    ws.Range("C4").FormatConditions.Delete
    Set fc = ws.Range("C4").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK - alles sluit aan""")
    fc.Interior.Color = RGB(198, 239, 206)
    Set fc = ws.Range("C4").FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""OK - alles sluit aan""")
    fc.Interior.Color = RGB(255, 150, 150)
    ControlesVernieuwen
End Sub

Public Sub Regel(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal formule As String)
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
    If CStr(wd.Cells(CB_KOP, 2).Value) <> "4. Controles" Then Exit Sub   ' blok 4 nog niet op deze plek (eerst DashboardAanmaken)
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
            st = UCase$(Trim$(Tkst(v(r, 8))))
            If Tkst(v(r, 5)) = "fout" Or Tkst(v(r, 7)) = "fout" Then Voeg waar(0), aant(0), t
            If IsGetal(v(r, 10)) And IsGetal(v(r, 11)) And Not HeeftNegatiefBudget(v, r) Then   ' bijdrageregels tellen niet mee
                grp(Tkst(v(r, 10))) = grp(Tkst(v(r, 10))) + CDbl(v(r, 11))
                grpRij(Tkst(v(r, 10))) = grpRij(Tkst(v(r, 10))) & IIf(Len(grpRij(Tkst(v(r, 10)))) > 0, ", ", "") & r
            End If
            If Left$(st, 1) = "A" And Nul(v(r, 30)) <> 0 Then Voeg waar(2), aant(2), t
            If Leeg(v(r, 15)) Then Voeg waar(3), aant(3), t
            If Leeg(v(r, 16)) Then Voeg waar(4), aant(4), t
            If Leeg(v(r, 24)) Then Voeg waar(5), aant(5), t
            If Leeg(v(r, 25)) Then Voeg waar(6), aant(6), t
            If Not Leeg(v(r, 3)) And IsGetal(v(r, 24)) And NrTekst(v(r, 3)) <> "7999998" And Not HeeftNegatiefBudget(v, r) Then
                sleutel = NrTekst(v(r, 3)) & "|" & Tkst(v(r, 24))
                If dubbel.Exists(sleutel) Then Voeg waar(7), aant(7), t & " = rij " & dubbel(sleutel) Else dubbel.Add sleutel, r
            End If
            If Tkst(v(r, 42)) = CtlTeControleren() Or Tkst(v(r, 42)) = "Afgekeurd" Then Voeg waar(8), aant(8), t
            For Each c In Array(9, 10)
                If Not Leeg(v(r, c)) And Not IsGetal(v(r, c)) Then
                    If StrComp(Tkst(v(r, c)), LCase$(Trim$(Tkst(v(r, c)))), vbBinaryCompare) <> 0 Then Voeg waar(9), aant(9), t: Exit For
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
            wd.Cells(CB_EERSTE + i, 5).Value = waar(i) & IIf(aant(i) > 12, " en " & (aant(i) - 12) & " meer", "") & "  >  " & hint(i)
        Else
            wd.Cells(CB_EERSTE + i, 5).Value = ""
        End If
    Next i
    ' aantallen van de controles zonder formule
    wd.Cells(CB_EERSTE + 1, 3).Value = aant(1)
    wd.Cells(CB_EERSTE + 7, 3).Value = aant(7)
    wd.Cells(CB_EERSTE + 14, 3).Value = aant(14)
    wd.Cells(CB_EERSTE + 15, 3).Value = aant(15)
    wd.Range(wd.Cells(CB_EERSTE, 2), wd.Cells(CB_LAATSTE, 2)).Rows.AutoFit
    wd.Range("C" & (CB_KOP + 1)).Value = Now
    If ActiveSheet.Name = "Controle" Then wd.Activate
    Exit Sub
Fout:
    wd.Range("C" & (CB_KOP + 1)).Value = "fout bij vernieuwen: " & Err.Description & " (rij " & r & ")"
End Sub

' Voegt een vindplaats toe aan de lijst (maximaal 12 tonen) en telt mee.
Private Sub Voeg(ByRef lijst As String, ByRef n As Long, ByVal t As String)
    n = n + 1
    If n <= 12 Then lijst = lijst & IIf(Len(lijst) > 0, "; ", "") & t
End Sub

' Regel in de verzamellijst met een negatief budget (bijdrage van derden).
Public Function NegatiefBudgetRij(ByVal wv As Worksheet, ByVal r As Long) As Boolean
    Dim c As Variant
    For Each c In Array(26, 31, 32, 33, 34, 35)
        If Nul(wv.Cells(r, CLng(c)).Value) < 0 Then NegatiefBudgetRij = True: Exit Function
    Next c
End Function

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
    Dim wv As Worksheet, r As Long, laatste As Long, c As Variant, n As Long, Waarde As Variant, voor As Variant, na As Variant, k As Long
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
                Waarde = wv.Cells(r, c).Value
                LogRegel "formule naar waarde", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, wv.Cells(r, c).Address(False, False), _
                         "'" & wv.Cells(r, c).Formula, Waarde, ""      ' apostrof: in het Logboek als tekst, niet als formule
                wv.Cells(r, c).Value = Waarde
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

' ======================================================================
' Jaarovergang (stap 14)
' ======================================================================
' Stap 14 - Jaarovergang (functioneel ontwerp par. 2.7, aangevuld met latere afspraken)
' JaarovergangUitvoeren : knop in blok 4 (alleen financieel adviseur). Bij openen in een nieuw kalenderjaar wordt erom gevraagd.
' 1. Back-up 'Verzamellijst 31-12-<oud jaar>' als harde waarden (= stand per 31-12; bron voor het jaarrekeningoverzicht).
' 2. Per regel: Z = Z + AE; AA = AA + AB; AB = AC; AC leeg; AE..AH = AF..AI; AI leeg. AD blijft formule.
' 3. Status L<oud jaar of eerder> wordt L<nieuw jaar>; bij die kredieten Startjaar (M) en Eindjaar (N) +1 en minimaal het nieuwe jaar.
'    Status A en toekomstige L blijven ongewijzigd.
' 4. Par_LopendJaar + 1 (koppen en Dashboard schuiven vanzelf mee).
' 5. Controle op de totalen; bij een verschil alles terugzetten.

Public Sub JaarovergangUitvoeren()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, wb As Worksheet, oud As Long, nieuw As Long, laatste As Long, r As Long, k As Long, nm As String
    Dim z As Variant, aa As Variant, ab As Variant, ac As Variant, ae As Variant, h As Variant, m As Variant, n As Variant
    Dim oZ As Variant, oAA As Variant, oAB As Variant, oAC As Variant, oAE As Variant, oH As Variant, oStart As Variant, oEind As Variant
    Dim sZ As Double, sAE1 As Double, sW As Double, sAEtot As Double, sAFtot As Double, st As String, jr As Long, nRol As Long
    Dim cZ As Double, cW As Double, cAE As Double, afwijking As Boolean

    If Not HeeftRecht("Toevoegen") Then
        MsgBox "De jaarovergang wordt uitgevoerd door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    oud = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    nieuw = oud + 1
    nm = "Verzamellijst 31-12-" & oud              ' een / mag niet in een tabbladnaam
    If Year(Date) <= oud Then
        If MsgBox("LET OP: het jaar " & oud & " is nog niet verstreken." & vbCrLf & _
                  "De jaarovergang mag pas vanaf 1 januari " & nieuw & " worden uitgevoerd." & vbCrLf & vbCrLf & _
                  "Alleen doorgaan als u dit op een KOPIE van het bestand test. Doorgaan?", _
                  vbYesNo + vbCritical + vbDefaultButton2, "Jaarovergang te vroeg") <> vbYes Then Exit Sub
    End If
    If BladBestaatStil(nm) Then
        MsgBox "Het blad '" & nm & "' bestaat al: de jaarovergang van " & oud & " naar " & nieuw & " is al uitgevoerd?" & vbCrLf & _
               "Er is niets gewijzigd.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Jaarovergang van " & oud & " naar " & nieuw & " uitvoeren?" & vbCrLf & vbCrLf & _
              "- er komt eerst een back-up '" & nm & "' (stand per 31-12-" & oud & ");" & vbCrLf & _
              "- alle bedragkolommen schuiven een jaar op;" & vbCrLf & _
              "- lopende kredieten L" & oud & " en eerder worden L" & nieuw & ", startjaar en eindjaar + 1." & vbCrLf & vbCrLf & _
              "Doe dit eerst op een kopie van het bestand.", vbYesNo + vbExclamation + vbDefaultButton2, "Jaarovergang") <> vbYes Then Exit Sub

    On Error GoTo Fout
    Application.ScreenUpdating = False
    Application.Calculate
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)

    ' 1. back-up als harde waarden
    wv.Copy Before:=ThisWorkbook.Worksheets("Dashboard")
    Set wb = ActiveSheet
    wb.Unprotect WACHTWOORD
    wb.Name = nm
    On Error Resume Next
    wb.Buttons.Delete
    If wb.FilterMode Then wb.ShowAllData
    On Error GoTo Fout
    wb.UsedRange.Value = wb.UsedRange.Value
    wb.Range("A1").AddComment "Stand per 31-12-" & oud & ", gemaakt bij de jaarovergang op " & Format$(Now, "dd-mm-yyyy hh:mm") & ". Alleen lezen."
    wb.Tab.Color = RGB(128, 128, 128)
    wb.Protect Password:=WACHTWOORD, AllowFiltering:=True
    wb.Move Before:=ThisWorkbook.Worksheets("Voertuigen")

    ' 2. inlezen (oude waarden bewaren om terug te kunnen zetten)
    oZ = wv.Range("Z2:Z" & laatste).Value: oAA = wv.Range("AA2:AA" & laatste).Value
    oAB = wv.Range("AB2:AB" & laatste).Value: oAC = wv.Range("AC2:AC" & laatste).Value
    oAE = wv.Range("AE2:AI" & laatste).Value: oH = wv.Range("H2:H" & laatste).Value
    oStart = wv.Range("M2:M" & laatste).Value: oEind = wv.Range("N2:N" & laatste).Value
    z = oZ: aa = oAA: ab = oAB: ac = oAC: ae = oAE: h = oH: m = oStart: n = oEind

    For r = 1 To laatste - 1
        sZ = sZ + Nul(oZ(r, 1)): sAE1 = sAE1 + Nul(oAE(r, 1))
        sW = sW + Nul(oAA(r, 1)) + Nul(oAB(r, 1)) + Nul(oAC(r, 1))
        For k = 2 To 5
            sAFtot = sAFtot + Nul(oAE(r, k))
        Next k
        If Not Leeg(wv.Cells(r + 1, 1).Value) Then
            z(r, 1) = Waarde(Nul(oZ(r, 1)) + Nul(oAE(r, 1)), oZ(r, 1), oAE(r, 1))
            aa(r, 1) = Waarde(Nul(oAA(r, 1)) + Nul(oAB(r, 1)), oAA(r, 1), oAB(r, 1))
            ab(r, 1) = oAC(r, 1)
            ac(r, 1) = Empty
            For k = 1 To 4
                ae(r, k) = oAE(r, k + 1)
            Next k
            ae(r, 5) = Empty
            st = UCase$(Trim$(Tkst(oH(r, 1))))
            jr = Val(Mid$(st, 2))
            If Left$(st, 1) = "L" And jr > 0 And jr <= oud Then
                h(r, 1) = "L" & nieuw
                nRol = nRol + 1
                If IsGetal(oStart(r, 1)) Then m(r, 1) = Application.Max(CLng(oStart(r, 1)) + 1, nieuw)
                If IsGetal(oEind(r, 1)) Then n(r, 1) = Application.Max(CLng(oEind(r, 1)) + 1, nieuw)
            End If
        End If
    Next r

    ' 3. wegschrijven en lopend jaar ophogen
    wv.Range("Z2:Z" & laatste).Value = z: wv.Range("AA2:AA" & laatste).Value = aa
    wv.Range("AB2:AB" & laatste).Value = ab: wv.Range("AC2:AC" & laatste).Value = ac
    wv.Range("AE2:AI" & laatste).Value = ae
    wv.Range("H2:H" & laatste).Value = h
    wv.Range("M2:M" & laatste).Value = m: wv.Range("N2:N" & laatste).Value = n
    ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value = nieuw
    Application.Calculate

    ' 4. controle op de totalen
    For r = 2 To laatste
        cZ = cZ + Nul(wv.Cells(r, 26).Value)
        cW = cW + Nul(wv.Cells(r, 27).Value) + Nul(wv.Cells(r, 28).Value) + Nul(wv.Cells(r, 29).Value)
        For k = 31 To 35
            cAE = cAE + Nul(wv.Cells(r, k).Value)
        Next k
    Next r
    afwijking = (Round(cZ - (sZ + sAE1), 2) <> 0) Or (Round(cW - sW, 2) <> 0) Or (Round(cAE - sAFtot, 2) <> 0)
    If afwijking Then
        wv.Range("Z2:Z" & laatste).Value = oZ: wv.Range("AA2:AA" & laatste).Value = oAA
        wv.Range("AB2:AB" & laatste).Value = oAB: wv.Range("AC2:AC" & laatste).Value = oAC
        wv.Range("AE2:AI" & laatste).Value = oAE: wv.Range("H2:H" & laatste).Value = oH
        wv.Range("M2:M" & laatste).Value = oStart: wv.Range("N2:N" & laatste).Value = oEind
        ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value = oud
        Application.DisplayAlerts = False
        wb.Unprotect WACHTWOORD
        wb.Delete
        Application.DisplayAlerts = True
        Application.Calculate
        Application.ScreenUpdating = True
        MsgBox "De controle op de totalen klopt niet. Alles is teruggezet; er is niets gewijzigd.", vbCritical
        Exit Sub
    End If

    LogRegel "jaarovergang", , , "Par_LopendJaar", oud, nieuw, nRol & " lopende kredieten naar L" & nieuw & "; back-up '" & nm & "'"
    KeuzelijstenVerversen True
    ControlesVernieuwen
    ' overzichten van voor de jaarovergang tonen nog het oude jaar ('start na <oud jaar>', budgetkolommen): weghalen.
    ' Een nieuw overzicht rekent altijd met het nieuwe lopende jaar. Opgeslagen bestanden in de mappen blijven bewaard.
    Dim oudBlad As Variant
    Application.DisplayAlerts = False
    For Each oudBlad In Array(BL_OVERZICHT, BL_OVERZICHT_NAAM, BL_UITVRAAG)
        If BladBestaatStil(CStr(oudBlad)) Then
            ThisWorkbook.Worksheets(CStr(oudBlad)).Unprotect WACHTWOORD
            ThisWorkbook.Worksheets(CStr(oudBlad)).Delete
        End If
    Next oudBlad
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Dashboard").Activate
    MsgBox "Jaarovergang naar " & nieuw & " uitgevoerd." & vbCrLf & vbCrLf & _
           "- de oude bladen 'Laatste kredietoverzicht', 'Laatste overzicht op naam' en 'Laatste uitvraag begroting' zijn weggehaald; " & _
           "maak ze opnieuw (dan met " & nieuw & ")" & vbCrLf & _
           "- back-up: blad '" & nm & "' (bron voor het kredietoverzicht jaarrekening)" & vbCrLf & _
           "- " & nRol & " lopende kredieten naar L" & nieuw & vbCrLf & _
           "- totalen gecontroleerd: budget t/m " & nieuw & " = oud budget + budget " & nieuw & "; werkelijke uitgaven gelijk", vbInformation
    Exit Sub
Fout:
    Dim foutTekst As String
    foutTekst = Err.Description
    On Error Resume Next
    ' gaat het mis voordat er iets is verschoven, dan de (half gemaakte) back-up weer weghalen
    If Not wb Is Nothing And CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value) = oud And IsEmpty(oZ) Then
        Application.DisplayAlerts = False
        wb.Unprotect WACHTWOORD
        wb.Delete
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Fout bij de jaarovergang: " & foutTekst & vbCrLf & "Controleer het bestand; gebruik zo nodig de kopie.", vbCritical
End Sub

' Leeg laten als beide bronnen leeg waren, anders het bedrag.
Private Function Waarde(ByVal bedrag As Double, ByVal a As Variant, ByVal b As Variant) As Variant
    If Leeg(a) And Leeg(b) Then Waarde = Empty Else Waarde = bedrag
End Function

' Bij openen: is het kalenderjaar verder dan het lopend jaar, dan vragen om de jaarovergang.
Public Sub JaarovergangSignaleren()
    Dim lj As Long
    On Error Resume Next
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    If lj = 0 Or Year(Date) <= lj Then Exit Sub
    If HeeftRecht("Toevoegen") Then
        If MsgBox("Het is " & Year(Date) & ", maar het lopend jaar in dit bestand is nog " & lj & "." & vbCrLf & _
                  "Jaarovergang nu uitvoeren?", vbYesNo + vbQuestion, "Jaarovergang") = vbYes Then JaarovergangUitvoeren
    Else
        MsgBox "Het lopend jaar in dit bestand is nog " & lj & ". De financieel adviseur moet de jaarovergang nog uitvoeren.", vbInformation
    End If
End Sub

' ======================================================================
' Uitleg bij de knoppen
' ======================================================================
' Naast elke knop op het Dashboard een knopje "?"; klikken toont de uitleg van die knop.
Public Sub KnoppenUitlegPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("J3:K40").Clear                       ' eerder blok 'Uitleg knoppen' opruimen
    KnopUitleg ws, "btnWijzigen", "Krediet wijzigen: zoek een bestaand krediet (op kredietnummer, omschrijving, activanummer of volgnummer). " & _
               "Het komt op de invoerkaart; pas aan en klik op 'Ingevuld? Verwerken!'."
    KnopUitleg ws, "btnKaartWissen", "Scherm leegmaken: maakt de invoerkaart leeg, er wordt niets opgeslagen."
    KnopUitleg ws, "btnVerwijderen", "Krediet verwijderen: typ een kredietnummer, activanummer of een deel van de omschrijving en kies het krediet. " & _
               "Alleen bij echte fouten en zonder werkelijke uitgaven. Alleen financieel adviseur. De gegevens gaan naar het Logboek."
    KnopUitleg ws, "btnUitgaven", "Verwerken werkelijk: leest de nieuwste uitdraai uit de map 'export coda' en werkt de werkelijke " & _
               "uitgaven bij. Meldingen komen op het blad 'Meldingen uitgaven'; daarna wordt het kredietoverzicht raad gemaakt."
    KnopUitleg ws, "btnTijdschrijvers", "Tijdschrijvers kiezen: kies een of meer tijdschrijvers voor dit krediet uit de lijst."
    KnopUitleg ws, "btnVerwerken", "Ingevuld? Verwerken!: slaat de invoerkaart op. Nieuw krediet = toevoegen, geladen krediet = wijziging opslaan."
    KnopUitleg ws, "btnOverzichtNaam", "Kredietoverzicht op naam: kies in blok 2 de rol en de naam. Alleen lopende kredieten. " & _
               "Wordt ook bewaard in map 02 (submap per persoon)."
    KnopUitleg ws, "btnOverzichtRaad", "Kredietoverzicht raad, stand vandaag: lopende kredieten t/m dit jaar plus de in dit jaar afgesloten " & _
               "kredieten, zonder subsidies en bijdragen. Wordt bewaard in map 03."
    KnopUitleg ws, "btnOverzichtJaarrekening", "Kredietoverzicht jaarrekening: stand per 31-12 van vorig jaar (uit de back-up van de jaarovergang). " & _
               "Wordt bewaard in map 03."
    KnopUitleg ws, "btnUitvraag", "Uitvraag kadernota/begroting: maakt per budgetbeheerder een bestand in map 02 waarin de teamleider " & _
               "de budgetten per jaar actualiseert en toelicht."
    KnopUitleg ws, "btnAlleBudgethouders", "Overzichten alle budgethouders: maakt voor elke budgetbeheerder een kredietoverzicht en bewaart het in map 02."
    KnopUitleg ws, "btnTijdschrijven", "Overzicht tijdschrijven: voor de beheerder tijdschrijven, alleen nieuwe, gewijzigde en afgesloten " & _
               "kredieten. Wordt bewaard in map 01."
    KnopUitleg ws, "btnActualisatie", "Uitvraag verwerken (financieel adviseur): kies hierboven rol Budgetbeheerder en de naam. Zijn ingevulde " & _
               "uitvraag komt op het blad 'Verwerking uitvraag': per investering goedkeuren, afwijzen of nog niet (bedragen mag u aanpassen). " & _
               "Daarna op dat blad 'Besluiten verwerken'."
    KnopUitleg ws, "btnUvStand", "Stand uitvraag: per budgethouder verstuurd, ingeleverd, verwerkt of nog open. Rood = nog niet ingeleverd " & _
               "(budgethouder aanspreken)."
    KnopUitleg ws, "btnUvControle", "Begroting vastgesteld: pas na vaststelling van de begroting gaan alle via de uitvraag verwerkte wijzigingen " & _
               "ter controle naar de financi" & ChrW(235) & "le administratie."
    KnopUitleg ws, "btnJaarovergang", "Jaarovergang: pas uitvoeren als het jaar voorbij is (vanaf 1 januari). Back-up per 31-12, bedragen schuiven een jaar op, lopende kredieten naar het nieuwe jaar. " & _
               "Alleen financieel adviseur; eerst op een kopie testen."
    KnopUitleg ws, "btnControles", "Controles vernieuwen: rekent de controles in blok 4 opnieuw uit en toont per controle waar het zit."
    KnopUitleg ws, "btnNaarTabblad", "Ga naar tabblad: gaat naar het tabblad dat in G2 is gekozen."
    KnopUitleg ws, "btnAfgekeurd", "Afgekeurd krediet herstellen (financieel adviseur): zet een door de financi" & ChrW(235) & "le administratie " & _
               "afgekeurd krediet op de invoerkaart, met de reden. Pas aan en klik op 'Ingevuld? Verwerken!'."
    KnopUitleg ws, "btnHercontrole", "Aanbieden voor hercontrole (financieel adviseur): het herstelde krediet gaat weer naar de financi" & ChrW(235) & _
               "le administratie (te controleren). Zij hebben altijd het laatste woord."
    KnopUitleg ws, "btnBhVervangen", "Budgethouder vervangen: vervangt in een keer een budgetbeheerder (kolom O) door een opvolger, " & _
               "in alle kredieten waar die naam staat (standaard alleen lopende kredieten). Elke wijziging gaat naar het Logboek."
    KnopUitleg ws, "btnPlVervangen", "Projectleider vervangen: vervangt in een keer een projectleider (kolom P) door een opvolger, " & _
               "in alle kredieten waar die naam staat (standaard alleen lopende kredieten). Elke wijziging gaat naar het Logboek."
    KnopUitleg ws, "btnTsVervangen", "Tijdschrijver vervangen: vervangt in een keer een tijdschrijver (kolom Q) door een opvolger; " & _
               "andere tijdschrijvers in dezelfde cel blijven staan. Standaard alleen lopende kredieten. Elke wijziging gaat naar het Logboek."
End Sub

' Klein knopje "?" naast de knop. Klikken toont de uitleg in een venster. Dit is een gewone formulierknop, net als de
' andere knoppen, zodat het ook op het beveiligde Dashboard werkt (scherminfo bij zweven is in Excel niet betrouwbaar).
Public Sub KnopUitleg(ByVal ws As Worksheet, ByVal knop As String, ByVal tekst As String)
    Dim b As Object, q As Object, had As Boolean
    On Error Resume Next
    Set b = ws.Buttons(knop)
    If b Is Nothing Then Exit Sub
    Set q = ws.Shapes("uitleg_" & knop)
    had = Not q Is Nothing
    ' knop smaller maken voor het vraagteken, tenzij dat al gebeurd is (vraagteken staat dan al rechts naast de knop)
    If had Then
        If q.Left < b.Left + b.Width Then b.Width = b.Width - 20
        q.Delete
    Else
        b.Width = b.Width - 20
    End If
    Set q = ws.Buttons.Add(b.Left + b.Width + 2, b.Top, 18, b.Height)
    q.Name = "uitleg_" & knop
    q.Caption = "?"
    q.OnAction = "KnopUitlegTonen"
    UitlegBewaren ws, "uitleg_" & knop, tekst
End Sub

' De uitleg wordt bewaard in een verborgen naam per knop, zodat hij ook na sluiten en openen beschikbaar is.
' Een tekst in een formule mag maximaal 255 tekens zijn: langere uitleg wordt in stukken opgeslagen ("..."&"...").
Private Sub UitlegBewaren(ByVal ws As Worksheet, ByVal naam As String, ByVal tekst As String)
    Dim f As String
    tekst = Replace(tekst, """", "'")
    Do While Len(tekst) > 250
        f = f & """" & Left$(tekst, 250) & """&"
        tekst = Mid$(tekst, 251)
    Loop
    f = "=" & f & """" & tekst & """"
    ThisWorkbook.Names.Add Name:="Uitleg_" & Mid$(naam, Len("uitleg_") + 1), RefersTo:=f, Visible:=False
End Sub

' Klik op een vraagteken: toont de uitleg van de knop ernaast.
Public Sub KnopUitlegTonen()
    Dim naam As String, knop As String, tekst As String, titel As String
    On Error Resume Next
    naam = CStr(Application.Caller)
    knop = Mid$(naam, Len("uitleg_") + 1)
    tekst = ThisWorkbook.Names("Uitleg_" & knop).RefersTo          ' ="tekst" of ="deel 1"&"deel 2"
    tekst = Replace(Mid$(tekst, 3, Len(tekst) - 3), """&""", "")
    titel = ActiveSheet.Buttons(knop).Caption
    On Error GoTo 0
    If Len(tekst) = 0 Then Exit Sub
    MsgBox tekst, vbInformation, "Uitleg" & IIf(Len(titel) > 0, ": " & titel, "")
End Sub


' ======================================================================
' Blauwe overzichtbladen: pdf maken en mailen
' ======================================================================
' Blauwe tabbladen: Laatste kredietoverzicht, Laatste overzicht op naam, Laatste uitvraag begroting en
' Laatste overzicht tijdschrijven. Op elk blad staan de knoppen "Opslaan als pdf" en "Mailen als pdf".
' De kolom Melding (alleen voor financien) staat nooit in de pdf.

Public Function IsBlauwBlad(ByVal naam As String) As Boolean
    IsBlauwBlad = (naam = BL_OVERZICHT Or naam = BL_OVERZICHT_NAAM Or naam = BL_UITVRAAG Or naam = BL_TIJD)
End Function

' Kolomnummer van de kolom 'Melding' (kopregel 4 of 5); 0 als die er niet is.
Public Function KolomMelding(ByVal ws As Worksheet) As Long
    Dim c As Long, r As Long
    For r = 4 To 5
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If StrComp(Trim$(CStr(ws.Cells(r, c).Value)), "Melding", vbTextCompare) = 0 Then KolomMelding = c: Exit Function
        Next c
    Next r
End Function

' Kopregel van het blad (de regel met de kolomnamen).
Private Function KopRijBlad(ByVal ws As Worksheet) As Long
    KopRijBlad = IIf(ws.Name = BL_TIJD, 4, 5)
End Function

' Voorgestelde bestandsnaam (zonder .pdf) en submap van 'Stambestand kredieten'.
Private Sub PdfNaam(ByVal ws As Worksheet, ByRef naam As String, ByRef mapNaam As String)
    Dim t As Variant
    Select Case ws.Name
        Case BL_OVERZICHT
            naam = "Kredietoverzicht " & CStr(ws.Range("A2").Value)
            mapNaam = "03 kredietoverzichten"
        Case BL_OVERZICHT_NAAM
            naam = "Kredietoverzicht " & Replace(CStr(ws.Range("A2").Value), ": ", " ")
            mapNaam = "02 Budgetbeheerders en projectleiders"
        Case BL_UITVRAAG
            naam = "Uitvraag kadernota-begroting " & Format$(Date, "dd-mm-yyyy")
            mapNaam = "02 Budgetbeheerders en projectleiders"
        Case Else
            naam = "Tijdschrijven kredieten " & Format$(Date, "dd-mm-yyyy")
            mapNaam = "01 Applicatiebeheer Tijdschrijven"
    End Select
    For Each t In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
        naam = Replace(naam, CStr(t), "-")
    Next t
    naam = Left$(Trim$(naam), 120)
End Sub

' Maakt de pdf van het blad: afdrukbereik, kopregel op elke pagina, liggend, passend op de breedte, kolom Melding verborgen.
Private Function PdfMaken(ByVal ws As Worksheet, ByVal pad As String) As Boolean
    Dim beveiligd As Boolean, cM As Long, mVerborgen As Boolean, kop As Long, lc As Long, lr As Long, c As Long, b As Object
    On Error GoTo Fout
    Application.ScreenUpdating = False
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    kop = KopRijBlad(ws)
    cM = KolomMelding(ws)
    If cM > 0 Then
        mVerborgen = ws.Columns(cM).Hidden
        ws.Columns(cM).Hidden = True
    End If
    ' laatste kolom en regel van het overzicht
    lc = ws.Cells(kop, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lc
        If ws.Cells(ws.Rows.Count, c).End(xlUp).Row > lr Then lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    Next c
    ' vervolgregels van een vorige pdf weghalen
    VervolgRegelsWeg ws, kop, lr
    ' rijhoogte opnieuw bepalen (een toelichting kan met de hand zijn verbeterd)
    BladOpmaken ws, kop, kop + 1, lr
    For Each b In ws.Buttons
        b.PrintObject = False
    Next b
    On Error Resume Next                         ' paginainstellingen kunnen per printer verschillen
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Address
        .PrintTitleRows = IIf(ws.Name = BL_UITVRAAG, "$4:$5", "$" & kop & ":$" & kop)
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .HeaderMargin = Application.CentimetersToPoints(0.6)
        .FooterMargin = Application.CentimetersToPoints(0.6)
        .CenterHorizontally = True
        .PrintGridlines = False
        .CenterFooter = "Pagina &P van &N"
    End With
    Application.PrintCommunication = True
    On Error GoTo Fout
    ' pagina's netjes afbreken: elke pagina begint met de donkerblauwe kopregel en een lichtblauwe regel
    PaginasIndelen ws, kop, lr, lc
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pad, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    PdfMaken = True
Einde:
    On Error Resume Next
    If cM > 0 Then ws.Columns(cM).Hidden = mVerborgen
    If beveiligd Then BladBeveiligen ws
    Application.ScreenUpdating = True
    Exit Function
Fout:
    Application.PrintCommunication = True
    MsgBox "De pdf kon niet worden gemaakt: " & Err.Description, vbCritical
    Resume Einde
End Function

' Lichtblauwe regel (programma, budgetbeheerder, 'Nog niet gestart', blok samengestelde kredieten).
Private Function IsLichtblauw(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    IsLichtblauw = (ws.Cells(r, 1).Interior.Color = RGB(214, 228, 240)) And Len(CStr(ws.Cells(r, 1).Value)) > 0
End Function

' Kopje van een investeringsgroep of project (kolom A leeg, kolom B vet en schuin).
Private Function IsGroepKopje(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    IsGroepKopje = Leeg(ws.Cells(r, 1).Value) And Len(CStr(ws.Cells(r, 2).Value)) > 0 And _
                   ws.Cells(r, 2).Font.Bold = True And ws.Cells(r, 2).Font.Italic = True
End Function

' Haalt de regels '... (vervolg)' van een vorige pdf weg (alleen lichtblauwe regels met die tekst).
Private Sub VervolgRegelsWeg(ByVal ws As Worksheet, ByVal kop As Long, ByRef lr As Long)
    Dim r As Long
    ws.ResetAllPageBreaks
    For r = lr To kop + 1 Step -1
        If IsLichtblauw(ws, r) And CStr(ws.Cells(r, 1).Value) Like "* (vervolg)" Then
            ws.Rows(r).Delete
            lr = lr - 1
        End If
    Next r
End Sub

' Zet de pagina-einden zelf (A4 liggend, passend op de breedte). Elke pagina begint met de kopregel (herhaald) en
' daaronder een lichtblauwe regel. Afbreken gebeurt bij voorkeur voor een lichtblauwe regel. Past een programma niet
' op een pagina, dan wordt afgebroken voor een groepskopje (of anders een gewone regel) en komt bovenaan de nieuwe
' pagina een lichtblauwe regel '<programma> (vervolg)'.
Private Sub PaginasIndelen(ByVal ws As Worksheet, ByVal kop As Long, ByRef lr As Long, ByVal lc As Long)
    Dim c As Long, r As Long, breed As Double, schaal As Double, titelH As Double, ruimte As Double, gebruikt As Double
    Dim h As Double, start As Long, lb As Long, gk As Long, brk As Long, bron As Long, eersteTitel As Long, tekst As String
    Dim nPag As Long
    On Error GoTo Klaar
    ws.ResetAllPageBreaks
    For c = 1 To lc
        If Not ws.Columns(c).Hidden Then breed = breed + ws.Columns(c).Width
    Next c
    If breed <= 0 Then Exit Sub
    ' A4 liggend 841,9 x 595,3 punten; marges links/rechts 1 cm, boven/onder 1,5 cm
    schaal = (841.9 - 2 * 28.35) / breed
    If schaal > 1 Then schaal = 1
    schaal = Int(schaal * 100) / 100
    eersteTitel = IIf(ws.Name = BL_UITVRAAG, 4, kop)
    For r = eersteTitel To kop
        titelH = titelH + ws.Rows(r).Height
    Next r
    ruimte = (595.3 - 2 * 42.5) / schaal * 0.96 - titelH      ' 4% veiligheidsmarge
    If ruimte < 100 Then Exit Sub
    For r = 1 To eersteTitel - 1                              ' titel bovenaan de eerste pagina
        gebruikt = gebruikt + ws.Rows(r).Height
    Next r
    start = kop + 1: r = kop + 1
    Do While r <= lr
        h = ws.Rows(r).Height
        If r > start Then
            If IsLichtblauw(ws, r) Then lb = r
            If IsGroepKopje(ws, r) Then gk = r
        End If
        If gebruikt + h > ruimte And r > start Then
            If lb > start Then
                brk = lb
            Else
                If gk > start Then brk = gk Else brk = r
                ' lichtblauwe vervolgregel met het programma (of de budgetbeheerder) erboven
                bron = 0
                For c = brk - 1 To kop + 1 Step -1
                    If IsLichtblauw(ws, c) Then bron = c: Exit For
                Next c
                If bron > 0 Then
                    tekst = CStr(ws.Cells(bron, 1).Value)
                    If Not tekst Like "* (vervolg)" Then tekst = tekst & " (vervolg)"
                    ws.Rows(brk).Insert
                    ws.Rows(bron).Copy Destination:=ws.Rows(brk)
                    ws.Range(ws.Cells(brk, 2), ws.Cells(brk, lc)).ClearContents
                    ws.Cells(brk, 1).Value = tekst
                    ws.Rows(brk).RowHeight = ws.Rows(bron).Height
                    lr = lr + 1
                End If
            End If
            ws.HPageBreaks.Add Before:=ws.Rows(brk)
            nPag = nPag + 1
            If nPag > 500 Then Exit Do                           ' noodrem
            start = brk: r = brk: gebruikt = 0: lb = 0: gk = 0
        Else
            gebruikt = gebruikt + h
            r = r + 1
        End If
    Loop
Klaar:
End Sub

' Knop "Opslaan als pdf" op een blauw blad.
Public Sub BladNaarPdf()
    Dim ws As Worksheet, naam As String, sm As String, map As String, pad As Variant
    Set ws = ActiveSheet
    If Not IsBlauwBlad(ws.Name) Then Exit Sub
    PdfNaam ws, naam, sm
    map = MapStambestand(sm)
    pad = Application.GetSaveAsFilename(InitialFileName:=map & naam & ".pdf", _
                                        FileFilter:="PDF-bestand (*.pdf), *.pdf", Title:="Pdf opslaan")
    If VarType(pad) = vbBoolean Then Exit Sub    ' geannuleerd
    If LCase$(Right$(CStr(pad), 4)) <> ".pdf" Then pad = pad & ".pdf"
    If PdfMaken(ws, CStr(pad)) Then
        LogRegel "pdf gemaakt", , , "", , , ws.Name & " | " & pad
        MsgBox "Pdf opgeslagen:" & vbCrLf & pad, vbInformation
    End If
End Sub

' Knop "Mailen als pdf" op een blauw blad: maakt de pdf en opent een nieuw Outlook-bericht met de pdf als bijlage.
' Het bericht wordt niet vanzelf verstuurd: vul de ontvanger in en klik zelf op Verzenden.
Public Sub BladMailen()
    Dim ws As Worksheet, naam As String, sm As String, pad As String, ol As Object, m As Object
    Set ws = ActiveSheet
    If Not IsBlauwBlad(ws.Name) Then Exit Sub
    PdfNaam ws, naam, sm
    pad = Environ$("TEMP") & "\" & naam & ".pdf"
    If Not PdfMaken(ws, pad) Then Exit Sub
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        MsgBox "Outlook kon niet worden gestart (de nieuwe Outlook ondersteunt dit niet)." & vbCrLf & _
               "De pdf staat klaar in:" & vbCrLf & pad & vbCrLf & "Voeg die zelf toe als bijlage.", vbExclamation
        Exit Sub
    End If
    Set m = ol.CreateItem(0)
    m.Subject = naam
    m.Attachments.Add pad
    m.Display                                    ' eerst tonen, dan staat de handtekening er al in
    m.HTMLBody = "<p style='font-family:Arial;font-size:10pt'>Beste collega,<br><br>Bijgaand het " & _
                 LCase$(Left$(CStr(ws.Range("A1").Value), 1)) & Mid$(CStr(ws.Range("A1").Value), 2) & _
                 IIf(ws.Name = BL_UITVRAAG, "", " (" & CStr(ws.Range("A2").Value) & ")") & ".</p>" & m.HTMLBody
    LogRegel "pdf gemaild", , , "", , , ws.Name & " | " & naam
End Sub
