Attribute VB_Name = "Kredieten_4_Overzichten"
Option Explicit

' Koppen van M en N in de verzamellijst (korte, duidelijke namen) en kolom A (volgnummer) gecentreerd.
Public Sub KoppenStartEindjaar()
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
Public Function MapStambestand(ByVal SubMap As String) As String
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
    pad = basis & SubMap
    If Not MapBestaat(pad) Then
        On Error Resume Next
        MkDir pad
        On Error GoTo 0
        If Not MapBestaat(pad) Then
            If Not mStil Then MsgBox "De map kon niet worden aangemaakt:" & vbCrLf & pad & vbCrLf & vbCrLf & _
                                     "Er is niets opgeslagen. Controleer of u in die map mag schrijven.", vbExclamation
            Exit Function
        End If
    End If
    MapStambestand = pad & "\"
End Function

' Bewaart een kopie van het blad als los Excelbestand. Bestaat het bestand al, dan wordt het overschreven.
' Lukt opslaan niet (bestand staat open, map niet bereikbaar), dan wordt de losse werkmap gesloten en volgt een melding.
Public Function BladOpslaan(ByVal ws As Worksheet, ByVal map As String, ByVal bestand As String) As Boolean
    Dim wb As Workbook, t As Variant, Fout As String
    If Len(map) = 0 Then Exit Function           ' map kon niet worden aangemaakt (melding is al gegeven)
    For Each t In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
        bestand = Replace(bestand, CStr(t), "-")
    Next t
    On Error GoTo mislukt
    ws.Copy
    Set wb = ActiveWorkbook
    ' knoppen (Terug naar dashboard) horen niet in het losse bestand
    On Error Resume Next
    wb.Worksheets(1).Unprotect WACHTWOORD
    wb.Worksheets(1).Buttons.Delete
    ' kolom Melding is alleen voor financien: niet in het opgeslagen of verzonden bestand
    If KolomMelding(wb.Worksheets(1)) > 0 Then wb.Worksheets(1).Columns(KolomMelding(wb.Worksheets(1))).Delete
    If ws.ProtectContents Then wb.Worksheets(1).Protect Password:=WACHTWOORD, AllowFiltering:=True, AllowFormattingRows:=True
    On Error GoTo mislukt
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=map & bestand & ".xlsx", FileFormat:=51
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    BladOpslaan = True
    Exit Function

mislukt:
    Fout = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    If Not mStil Then MsgBox "Opslaan is niet gelukt:" & vbCrLf & map & bestand & ".xlsx" & vbCrLf & vbCrLf & Fout & vbCrLf & vbCrLf & _
                             "Staat het bestand misschien open? Sluit het en probeer het opnieuw. Het blad in deze werkmap is wel gemaakt.", vbExclamation
End Function

' Maakt per budgetbeheerder (kolom O) een kredietoverzicht en bewaart het in map 02, in een submap per persoon.
Public Sub KredietoverzichtAlleBudgethouders()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, namen As Object, nm As Variant, ws As Worksheet
    Dim map As String, lj As Long, n As Long, mislukt As String
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
    On Error GoTo Fout
    mStil = True
    For Each nm In namen.Keys
        Set ws = OverzichtMaken("naam", wv, lj, 15, CStr(nm))
        If Not ws Is Nothing Then
            If BladOpslaan(ws, SubMap(map, CStr(nm)), "Kredietoverzicht " & nm & " " & Format$(Date, "dd-mm-yyyy")) Then
                n = n + 1
            Else
                mislukt = mislukt & "- " & nm & vbCrLf
            End If
        End If
    Next nm
    mStil = False
    ThisWorkbook.Worksheets("Dashboard").Activate
    MsgBox n & " kredietoverzichten opgeslagen in:" & vbCrLf & map & _
           IIf(Len(mislukt) > 0, vbCrLf & vbCrLf & "Niet gelukt (bestand open of map niet bereikbaar):" & vbCrLf & mislukt, ""), _
           IIf(Len(mislukt) > 0, vbExclamation, vbInformation)
    Exit Sub

Fout:
    mStil = False
    MsgBox "Fout bij de overzichten voor alle budgethouders: " & Err.Description & vbCrLf & _
           n & " overzicht(en) waren al opgeslagen.", vbCritical
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

    ' samengestelde kredieten (getal in J): medewerkers schrijven op het hoofddeel (wegen bij wegen/riolering,
    ' ruwbouw bij scholen en gebouwen), zodat zij hun uren niet zelf hoeven te verdelen. De andere delen krijgen
    ' geen eigen regel; hun tijdschrijvers komen bij het hoofddeel. Geen hoofddeel in de groep: elk deel een eigen regel.
    Dim vT As Variant, gHoofd As Object, gOms As Object, deelNaar As Object, nrNaar As Object, g As String, hn As String
    vT = wv.Range("A2:AW" & laatste).Value
    JGetalNormaliseren vT
    Set gHoofd = CreateObject("Scripting.Dictionary")
    Set gOms = CreateObject("Scripting.Dictionary")
    Set deelNaar = CreateObject("Scripting.Dictionary")     ' rij in de verzamellijst -> kredietnummer hoofddeel
    Set nrNaar = CreateObject("Scripting.Dictionary")       ' kredietnummer deel -> kredietnummer hoofddeel
    For r = 1 To UBound(vT, 1)
        hn = NrTekst(vT(r, 3))
        If IsGetal(vT(r, 10)) And UCase$(Left$(Trim$(CStr(vT(r, 8))), 1)) = "L" And Len(hn) > 0 And hn <> "7999998" Then
            g = CStr(vT(r, 10))
            If IsHoofddeel(vT, r) And Not gHoofd.Exists(g) Then gHoofd(g) = hn: gOms(g) = Trim$(CStr(vT(r, 4)))
        End If
    Next r
    For r = 1 To UBound(vT, 1)
        hn = NrTekst(vT(r, 3))
        If IsGetal(vT(r, 10)) And Len(hn) > 0 Then
            g = CStr(vT(r, 10))
            If gHoofd.Exists(g) And Not IsHoofddeel(vT, r) And hn <> gHoofd(g) Then
                deelNaar(r + 1) = gHoofd(g)
                nrNaar(hn) = gHoofd(g)
            End If
        End If
    Next r

    ' huidige stand per kredietnummer (alle lopende kredieten): omschrijving, status, tijdschrijvers incl. projectleider
    Set huidig = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        nr = NrTekst(wv.Cells(r, 3).Value)
        st = Trim$(CStr(wv.Cells(r, 8).Value))
        If deelNaar.Exists(r) Then nr = deelNaar(r)  ' deel van een samengesteld krediet: uren op het hoofddeel
        If Len(nr) > 0 And nr <> "7999998" And UCase$(Left$(st, 1)) = "L" Then
            naamTs = NamenSamenvoegen(CStr(wv.Cells(r, 16).Value), CStr(wv.Cells(r, 17).Value))
            If Not huidig.Exists(nr) Then
                g = ""
                If deelNaar.Exists(r) Then g = CStr(wv.Cells(r, 10).Value)
                huidig(nr) = Array(IIf(Len(g) > 0, gOms(g), Trim$(CStr(wv.Cells(r, 4).Value))), st, naamTs)
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
            If nrNaar.Exists(nr) Then ws.Cells(rij, 5).Value = "afgesloten (uren schrijven op " & nrNaar(nr) & ")"
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
    ws.Range("E5:E" & rij).WrapText = True
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
    Dim f As String, oudFm As Variant, geschreven As Boolean, foutTekst As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    If MsgBox("In kolom 'Nog uit te geven' overal de formule zetten: bij een afgesloten krediet (A) is het restant 0, " & _
              "anders budget min werkelijke uitgaven. Doorgaan?", vbYesNo + vbQuestion, "Formule restant") <> vbYes Then Exit Sub
    Set wv = Vz()
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    On Error GoTo Fout
    voor = MeetStap8(wv, CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value))
    oudFm = wv.Range(wv.Cells(2, 30), wv.Cells(maxRij, 30)).Formula     ' oude formules bewaren
    calcOud = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    geschreven = True                            ' vanaf hier wordt bij een fout alles teruggezet
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
            wv.Range(wv.Cells(2, 30), wv.Cells(maxRij, 30)).Formula = oudFm
            Application.Calculate
            MsgBox "De nameting wijkt af in meting " & k & ". De oude formules zijn teruggezet. Er is niets gewijzigd.", vbCritical
            Exit Sub
        End If
    Next k
    LogRegel "formule restant bijwerken", , , "Nog uit te geven", , , n & " cellen aangepast"
    MsgBox n & " cellen in 'Nog uit te geven' aangepast.", vbInformation
    Exit Sub
Fout:
    foutTekst = Err.Description
    On Error Resume Next
    If geschreven Then wv.Range(wv.Cells(2, 30), wv.Cells(maxRij, 30)).Formula = oudFm
    If calcOud <> 0 Then Application.Calculation = calcOud
    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox "Fout: " & foutTekst & vbCrLf & IIf(geschreven, "De oude formules zijn teruggezet. ", "") & "Er is niets gewijzigd.", vbCritical
End Sub

' Submap per persoon (wordt aangemaakt als die nog niet bestaat).
Public Function SubMap(ByVal map As String, ByVal naam As String) As String
    Dim t As Variant
    For Each t In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
        naam = Replace(naam, CStr(t), "-")
    Next t
    Dim pad As String
    If Len(map) = 0 Then Exit Function
    pad = map & Trim$(naam)
    If Not MapBestaat(pad) Then
        On Error Resume Next
        MkDir pad
        On Error GoTo 0
        If Not MapBestaat(pad) Then
            If Not mStil Then MsgBox "De map kon niet worden aangemaakt:" & vbCrLf & pad & vbCrLf & vbCrLf & _
                                     "Er is niets opgeslagen. Controleer of u in die map mag schrijven.", vbExclamation
            Exit Function
        End If
    End If
    SubMap = pad & "\"
End Function

Public Function RolNaam(ByVal kol As Long) As String
    Select Case kol
        Case 15: RolNaam = "Budgetbeheerder"
        Case 16: RolNaam = "Projectleider"
        Case 17: RolNaam = "Tijdschrijver"
    End Select
End Function

' Zet het logo van de gemeente rechtsboven op het overzicht. Het logo is een bestand 'logo*.png' (of .jpg)
' in de map 'Stambestand kredieten'. Geen logo gevonden: dan zonder logo.
Public Sub LogoPlaatsen(ByVal ws As Worksheet, ByVal kolRechts As Long)
    Dim basis As String, f As String, shp As Object
    On Error GoTo klaar
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
klaar:
End Sub

' ======================================================================
' Bladbeveiliging
' ======================================================================
' Alle mutaties lopen via het Dashboard. De bladen staan op slot; macro's mogen wel schrijven (UserInterfaceOnly).
' Die instelling wordt niet in het bestand bewaard, daarom zet Auto_Open de beveiliging bij elk openen opnieuw aan.
' Open blijven: invoercellen op het Dashboard en de gele invoerkolommen van de tabellen op Parameters.
' Filteren mag op de verzamellijst, sorteren niet.

Public Sub BeveiligingAan()
    Dim naam As Variant, bw As Worksheet, actief0 As Object
    On Error Resume Next
    Set actief0 = ActiveSheet
    On Error GoTo 0
    mHerbeveiligingGepland = False
    Application.EnableEvents = True
    On Error Resume Next
    For Each naam In Array(BL_VZ, "Logboek", "Vorige upload", "Stand tijdschrijven", BL_MUTREG, BL_UITVRAAG, "Parameters", "Dashboard", "Controle", _
                           BL_MELD, BL_OVERZICHT, BL_OVERZICHT_NAAM, BL_KENG)
        If BladBestaatStil(CStr(naam)) Then BladBeveiligen ThisWorkbook.Worksheets(CStr(naam))
    Next naam
    ' back-up van 31-12 en voorstellen nieuwe kredieten: ook na elke macro weer op slot
    For Each bw In ThisWorkbook.Worksheets
        If bw.Name Like "Verzamellijst 31-12-####" Or bw.Name = "Voorstellen nieuwe kredieten" Then BladBeveiligen bw
    Next bw
    TabbladenOrdenen
    ' nooit eindigen op het Controle-blad als de macro daar niet begon
    On Error Resume Next
    If Not mActiefVoorMacro Is Nothing Then
        If ActiveSheet.Name = "Controle" And mActiefVoorMacro.Name <> "Controle" Then mActiefVoorMacro.Activate
    End If
    If mBlijfOpDashboard Then
        mBlijfOpDashboard = False
        ThisWorkbook.Worksheets("Dashboard").Activate
    End If
    ' nooit eindigen op het blad Kengetallen als je (of de macro) daar niet begon
    If Not mActiefVoorMacro Is Nothing Then
        If ActiveSheet.Name = BL_KENG And mActiefVoorMacro.Name <> BL_KENG Then mActiefVoorMacro.Activate
    ElseIf Not actief0 Is Nothing Then
        If ActiveSheet.Name = BL_KENG And actief0.Name <> BL_KENG Then actief0.Activate
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

' Naam samengesteld krediet (kaartveld 30): bij een nieuw krediet alleen invulbaar als bij Samengesteld/standaard/stelpost
' "nieuw samengesteld krediet" is gekozen. Bij een volgend deelkrediet van een bestaande groep staat de naam dus vast.
' Bij het laden van een bestaand krediet (wijzigen) blijft het veld gewoon bewerkbaar.
Public Function KaartNaamGeblokkeerd(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    If Len(CStr(ws.Range("Kaart_Rijnummer").Value)) > 0 Then Exit Function
    KaartNaamGeblokkeerd = (LCase$(Trim$(CStr(ws.Cells(KaartRij(8), KaartKol(8)).Value))) <> "nieuw samengesteld krediet")
End Function

Public Sub BladBeveiligen(ByVal ws As Worksheet)
    Dim lo As ListObject, kol As Variant, j As Long, k As Long, laatste As Long
    On Error Resume Next
    ws.Unprotect WACHTWOORD
    Select Case ws.Name
        Case "Parameters"
            ws.Cells.Locked = True
            For Each kol In Array(Array("Tbl_Rollen", "Naam"), Array("Tbl_Rollen", "Windows-naam"), Array("Tbl_Rollen", "Rol"), _
                                  Array("Tbl_Rollen", "Toevoegen"), Array("Tbl_Rollen", "Wijzigen"), Array("Tbl_Rollen", "Bevestigen"), _
                                  Array("Tbl_Rollen", "Beheer"), _
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
                    ws.Cells(BijdrageRij(j, k), 5).Locked = False      ' bron van een bijdrage (kiezen of typen)
                Next k
            Next j
            ws.Range("G2").Locked = False                         ' tabbladkiezer
            With ws.Cells(KaartRij(30), KaartKol(30))              ' naam samengesteld krediet: geblokkeerd tenzij nieuw samengesteld
                If KaartNaamGeblokkeerd(ws) Then
                    .Locked = True
                    .Interior.Color = RGB(217, 217, 217)
                Else
                    .Locked = False
                    .Interior.Color = RGB(221, 235, 247)
                End If
            End With
        Case BL_VZ
            If Not ws.AutoFilterMode Then
                laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
                ws.Range("A1:AW" & laatste).AutoFilter
            End If
    End Select
    ' blauwe overzichtbladen: rijhoogte mag met de hand worden aangepast (na het verbeteren van een toelichting)
    ws.Protect Password:=WACHTWOORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=IsBlauwBlad(ws.Name)
End Sub

' Aan het begin van elke macro die iets wijzigt: beveiliging uit en direct na afloop van de macro weer aan
' (Application.OnTime start BeveiligingAan zodra de macro klaar is, ook na een fout of een tussentijds stoppen).
Public Sub WerkStart()
    If mActiefVoorMacro Is Nothing Then Set mActiefVoorMacro = ActiveSheet
    Application.EnableEvents = False             ' macro's die de invoerkaart vullen geven geen melding (weer aan in BeveiligingAan)
    BeveiligingTijdelijkUit
    If mHerbeveiligingGepland Then Exit Sub       ' maar een keer inplannen per macro
    mHerbeveiligingGepland = True
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!BeveiligingAan"
End Sub

' Zet de beveiliging van alle bladen tijdelijk uit (voor macro's die veel schrijven). Daarna altijd BeveiligingAan.
Public Sub BeveiligingTijdelijkUit()
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
' - regels: rijhoogte 13,5. Heeft het blad een kolom Toelichting, dan wordt een regel alleen hoger bij een lange
'   toelichting (andere kolommen tellen dan niet mee). Zonder kolom Toelichting: naar de cellen met terugloop.
Public Sub BladOpmaken(ByVal ws As Worksheet, ByVal kopRij As Long, ByVal van As Long, ByVal tot As Long)
    Dim c As Long, r As Long, lk As Long, w As Variant, nodig As Double, n As Long, regels As Long, cel As Range, stap As Long
    Dim toel() As Boolean, heeftToel As Boolean
    lk = ws.Cells(kopRij, ws.Columns.Count).End(xlToLeft).Column
    ReDim toel(1 To lk)
    For c = 1 To lk
        If InStr(1, CStr(ws.Cells(kopRij, c).Value), "toelichting", vbTextCompare) > 0 Then toel(c) = True: heeftToel = True
    Next c
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
                If Not ws.Columns(c).Hidden And (r = kopRij Or toel(c) Or Not heeftToel) Then
                    Set cel = ws.Cells(r, c)
                    If cel.WrapText And Len(CStr(cel.Value)) > 0 And Not IsNumeric(cel.Value) Then
                        n = AantalRegels(CStr(cel.Value), ws.Columns(c).ColumnWidth, cel.Font.Bold)
                        If n > regels Then regels = n
                    End If
                End If
            Next c
            If r = kopRij Then
                ws.Rows(r).RowHeight = Application.Min(409, regels * 12.75 + 8)
            Else
                ws.Rows(r).RowHeight = Application.Min(409, 13.5 + (regels - 1) * 12.75)   ' standaard 13,5; langere tekst: per regel 12,75 erbij
            End If
        End If
    Next r
End Sub

' Aantal tekstregels als de tekst met terugloop in een kolom van deze breedte staat (Arial 10).
Private Function AantalRegels(ByVal tekst As String, ByVal breedte As Double, ByVal vet As Boolean) As Long
    Dim ruimte As Double, teken As Double, Regel As Double, w As Variant, lw As Double, n As Long, deel As Variant
    ruimte = breedte * 6 + 5 - 6                 ' pixels in de kolom, min wat marge
    If ruimte < 30 Then ruimte = 30                ' smalle of verborgen kolom: geen eindeloze lus
    teken = IIf(vet, 7.4, 6.8)                   ' gemiddelde tekenbreedte Arial 10 in pixels
    n = 0
    For Each deel In Split(Replace(tekst, vbCr, ""), vbLf)
        n = n + 1: Regel = 0
        For Each w In Split(CStr(deel), " ")
            lw = Len(CStr(w)) * teken
            If Regel = 0 Then
                Regel = lw
            ElseIf Regel + teken + lw <= ruimte Then
                Regel = Regel + teken + lw
            Else
                n = n + 1: Regel = lw
            End If
            Do While Regel > ruimte                ' woord langer dan de kolom
                n = n + 1: Regel = Regel - ruimte
            Loop
        Next w
    Next deel
    If n < 1 Then n = 1
    AantalRegels = n
End Function

' Vervolg van de macro's voor het Dashboard investeringskredieten (gesplitst vanwege de grootte; declaraties in Kredieten_1_Basis).

' Alleen een getal (bedrag in euro's) toestaan in de budgetcellen van het uitvraagbestand.
Private Sub UvBedragenControle(ByVal rng As Range)
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-1000000000", Formula2:="1000000000"
        .IgnoreBlank = True
        .ErrorTitle = "Bedrag"
        .ErrorMessage = "Typ een bedrag in euro's (alleen een getal)."
    End With
End Sub

' Bouwt het tijdelijke blad waarmee het uitvraagbestand van een budgetbeheerder wordt gemaakt (naam leeg = iedereen).
Public Function UitvraagBlad(ByVal naam As String) As Worksheet
    Dim wv As Worksheet, ws As Worksheet, v As Variant, laatste As Long, n As Long, r As Long, i As Long, k As Long, c As Long
    Dim sel() As Long, sl() As String, tk() As Boolean, nSel As Long, lj As Long, bronKol As Variant, kop As Variant
    Dim rij As Long, eersteData As Long, vBh As String, bh As String, startBh As Long, nK As Long, tv As String

    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    v = wv.Range("A2:AW" & laatste).Value
    JGetalNormaliseren v
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

    If BladBestaatStil(BL_UVTEMP) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(BL_UVTEMP).Unprotect WACHTWOORD
        ThisWorkbook.Worksheets(BL_UVTEMP).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = NieuwBlad()
    ws.Name = BL_UVTEMP
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
    ws.Range("A2").Value = IIf(Len(naam) > 0, "Budgetbeheerder: " & naam & " - ", "") & "stand " & Format$(Date, "dd-mm-yyyy") & ". " & _
                           "Hieronder staan al je lopende kredieten met het restantkrediet. Per krediet kun je de budgetten van de komende jaren " & _
                           "verhogen, verlagen, doorschuiven naar een later jaar of terugschuiven naar een eerder jaar. Terugschuiven naar " & lj & _
                           " (het lopende jaar) kan niet. Naast de automatische toelichting kun je een eigen toelichting geven. " & _
                           "Gouden cellen mag je aanpassen, grijze cellen zijn vast."
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
    Dim nvStart As Long, nvEind As Long, nvKop As Long, nvUitleg As Long
    rij = rij + 2
    nvKop = rij: nvUitleg = rij + 1
    ws.Cells(nvKop, 1).Value = "Nieuwe investeringskredieten bestaand beleid (nog niet geraamd)"
    ws.Range(ws.Cells(nvKop, 1), ws.Cells(nvKop, nK)).Font.Bold = True
    ws.Range(ws.Cells(nvKop, 1), ws.Cells(nvKop, nK)).Interior.Color = RGB(214, 228, 240)
    ws.Range(ws.Cells(nvKop, 1), ws.Cells(nvKop, nK)).Font.Color = RGB(0, 73, 118)
    ws.Cells(nvUitleg, 1).Value = "Voer hier nog niet geraamde investeringen voor bestaand beleid in, bijvoorbeeld een nieuwe jaarschijf of vervangingsinvesteringen. " & _
                                  "Ideaal is dat hier alleen investeringen staan die in het laatste jaar (" & (lj + 5) & ") aan de orde zijn. " & _
                                  "Eerdere vervangingsinvesteringen hadden al bij de vorige uitvraag aangeleverd moeten worden. " & _
                                  "Vul omschrijving, taakveld, budgetten en toelichting in. De afschrijftermijn bepaalt de financieel adviseur."
    nvStart = rij + 2: nvEind = rij + 11
    For k = nvStart To nvEind
        ws.Cells(k, 3).Value = naam
        ws.Cells(k, 13).Value = "NIEUW"
    Next k
    rij = nvEind + 1

    ' kolom L: nieuwe omschrijving (invullen), kolom M: wijziging (automatisch), kolom N: onderbouwing (invullen)
    ws.Range(ws.Columns(12), ws.Columns(13)).Insert
    nK = nK + 2
    ws.Range(ws.Cells(nvUitleg, 1), ws.Cells(nvUitleg, 14)).Merge          ' uitleg bij de nieuwe kredieten over de volle breedte
    ws.Cells(nvUitleg, 1).WrapText = True
    ws.Cells(nvUitleg, 1).VerticalAlignment = xlTop
    ws.Cells(nvUitleg, 1).Font.Italic = True
    ws.Cells(nvUitleg, 1).Font.Color = RGB(89, 89, 89)
    ws.Range("A2:N2").Merge                                  ' uitleg bovenaan over de volle breedte
    ws.Range("A2").WrapText = True
    ws.Range("A2").VerticalAlignment = xlTop
    ws.Rows(2).RowHeight = 64
    ws.Cells(4, 12).Value = "in te vullen"
    ws.Cells(4, 13).Value = "automatisch, vast"
    ws.Cells(4, 14).Value = "in te vullen (komt in de toelichting)"
    ws.Cells(5, 12).Value = "Nieuwe omschrijving (alleen bij wijziging)"
    ws.Cells(5, 13).Value = "Automatische toelichting (vast, niet aanpassen)"
    ws.Cells(5, 14).Value = "Eigen toelichting budgethouder (invullen)"
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
    ws.Cells(4, 21).Value = "automatisch, vast"
    ws.Cells(5, 21).Value = "Beschikbaar " & (lj + 1) & " (restant " & lj & " + budget " & (lj + 1) & ")"
    With ws.Range(ws.Cells(4, 21), ws.Cells(4, 21))
        .Font.Size = 8
        .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(5, 21)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 73, 118)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Columns(15), ws.Columns(21)).Hidden = True                ' hulpkolommen: volgnummer, was-bedragen en beschikbaar eerste jaar
    ws.Cells.Locked = True
    ' voorstellen nieuwe kredieten: omschrijving, taakveld, afschrijving, budgetten en toelichting invullen
    For k = nvStart To nvEind
        ws.Range(ws.Cells(k, 2), ws.Cells(k, 2)).Interior.Color = RGB(255, 217, 102)
        ws.Cells(k, 4).Interior.Color = RGB(255, 217, 102)          ' afschrijftermijn (E) niet invulbaar
        ws.Range(ws.Cells(k, 7), ws.Cells(k, 11)).Interior.Color = RGB(255, 217, 102)
        ws.Cells(k, 14).Interior.Color = RGB(255, 217, 102)
        ws.Range(ws.Cells(k, 2), ws.Cells(k, 2)).Locked = False
        ws.Cells(k, 4).Locked = False
        ws.Range(ws.Cells(k, 7), ws.Cells(k, 11)).Locked = False
        ws.Cells(k, 14).Locked = False
        ws.Cells(k, 13).Value = "Nieuw krediet"
        ws.Cells(k, 13).Interior.Color = RGB(217, 217, 217)
        UvBedragenControle ws.Range(ws.Cells(k, 7), ws.Cells(k, 11))
    Next k
    ' bestaande kredieten: budgetten (G-K), nieuwe omschrijving (L) en onderbouwing (N) invullen; M rekent zelf
    Dim fWijz As String, jaren As String, eur As String
    eur = ChrW(8364)
    jaren = (lj + 1) & "," & (lj + 2) & "," & (lj + 3) & "," & (lj + 4) & "," & (lj + 5)
    For k = eersteData To rij - 1
        If Not Leeg(ws.Cells(k, 2).Value) And Not ws.Cells(k, 6).HasFormula And k < nvStart Then   ' kredietregel, geen totaalregel
            ws.Range(ws.Cells(k, 7), ws.Cells(k, 12)).Interior.Color = RGB(255, 217, 102)
            ws.Cells(k, 14).Interior.Color = RGB(255, 217, 102)
            ws.Range(ws.Cells(k, 7), ws.Cells(k, 12)).Locked = False
            ws.Cells(k, 14).Locked = False
            ws.Cells(k, 13).Interior.Color = RGB(217, 217, 217)             ' automatische toelichting: vast
            ws.Cells(k, 21).Formula = "=F" & k & "+G" & k                    ' beschikbaar in het eerste jaar: restant + budget van dat jaar
            ws.Cells(k, 21).Interior.Color = RGB(217, 217, 217)
            ws.Cells(k, 21).NumberFormat = "#,##0;-#,##0;-"
            ws.Cells(k, 21).HorizontalAlignment = xlRight
            UvBedragenControle ws.Range(ws.Cells(k, 7), ws.Cells(k, 11))
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
                .Font.Color = RGB(192, 0, 0)                   ' markering: er is iets gewijzigd
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
    ws.Columns(14).ColumnWidth = 75                      ' eigen toelichting: ruim, met terugloop
    ws.Range(ws.Cells(eersteData, 2), ws.Cells(rij, 3)).WrapText = True
    ws.Range(ws.Cells(eersteData, 4), ws.Cells(rij, 4)).WrapText = False
    ws.Range(ws.Cells(eersteData, 12), ws.Cells(rij, 14)).WrapText = True
    ws.Range(ws.Cells(eersteData, 1), ws.Cells(rij, nK)).VerticalAlignment = xlTop
    BladOpmaken ws, 5, eersteData, rij              ' rijhoogte 13,5; alleen hoger bij een lange toelichting
    ws.Rows(nvKop).RowHeight = 16                              ' kopje en uitleg van het blok nieuwe kredieten
    ws.Rows(nvUitleg).RowHeight = 42
    ws.Cells(nvUitleg, 1).WrapText = True                    ' terugloop in de uitleg (de opmaak van kolom D hierboven kan die uitzetten)
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
                Case "Logboek"
                    ' de koprij is vastgezet: de knop staat linksboven en blijft altijd zichtbaar
                    ws.Range("A1:K1").VerticalAlignment = xlBottom
                    If ws.Rows(1).RowHeight < 45 Then ws.Rows(1).RowHeight = 45
                    KnopTerug ws, ws.Range("A1")
                Case "Controle"
                    KnopTerug ws, ws.Range("A2").Offset(0, 0)
                    ws.Buttons("btnTerug").Left = ws.Range("A2").Left + 4 * 160
                    ws.Buttons("btnTerug").Top = ws.Range("A2").Top + 2
                Case BL_MELD, BL_TIJD
                    KnopTerug ws, ws.Range(IIf(ws.Name = BL_MELD, "F1", "A3"))
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

' Knop Gebruiker verwijderen bovenaan het Dashboard, op de regel van de gebruiker (alleen zichtbaar voor de financieel adviseur).
Public Sub GebruikerKnopPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    ws.Shapes("uitleg_btnGebruikerVerwijderen").Delete
    ws.Shapes("btnGebruikerVerwijderen").Delete
    On Error GoTo 0
    If ws.Rows(2).RowHeight < 22 Then ws.Rows(2).RowHeight = 22
    Set b = ws.Buttons.Add(ws.Range("E2").Left + 2, ws.Range("E2").Top + 2, 170, 18)
    b.Name = "btnGebruikerVerwijderen"
    b.Caption = "Gebruiker verwijderen"
    b.OnAction = "GebruikerVerwijderen"
    b.Placement = xlFreeFloating
    b.PrintObject = False
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
