Attribute VB_Name = "Kredieten_7_Raadplegen"
Option Explicit

' Krediet raadplegen: zoeken op verschillende manieren en alles van een krediet bekijken, zonder iets te wijzigen.
' Het resultaat staat op het blauwe blad 'Krediet raadplegen' (een vorige raadpleging wordt vervangen).
Public Const BL_RAADPLEGEN As String = "Krediet raadplegen"
Private Const MAX_KEUZE As Long = 12           ' zoveel treffers in het keuzevenster; meer = lijst op het blad
Private Const LAATSTE_KOL As Long = 49         ' kolom AW van de verzamellijst

' Knop bovenaan het Dashboard (rechts naast 'Ga naar tabblad'), voor iedereen zichtbaar. Wordt ook aangeroepen
' vanuit DashboardKnoppenPlaatsen; kan los worden gedraaid om alleen deze knop te plaatsen.
Public Sub RaadplegenKnopPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    ws.Shapes("uitleg_btnRaadplegen").Delete
    ws.Shapes("btnRaadplegen").Delete
    On Error GoTo 0
    If ws.Rows(2).RowHeight < 22 Then ws.Rows(2).RowHeight = 22
    Set b = ws.Buttons.Add(ws.Range("H2").Left + 122, ws.Range("H2").Top + 2, 150, 18)
    b.Name = "btnRaadplegen"
    b.Caption = "Krediet raadplegen"
    b.OnAction = "KredietRaadplegen"
    b.Placement = xlFreeFloating
    b.PrintObject = False
    KnopUitleg ws, "btnRaadplegen", "Krediet raadplegen: alleen bekijken, er wordt niets gewijzigd. Je kiest waarop je zoekt: " & _
               "1 = volgnummer (kolom A), 2 = kredietnummer (C), 3 = omschrijving (D), 4 = activanummer (X), " & _
               "5 = budgetbeheerder, 6 = projectleider, 7 = tijdschrijver, 8 = taakveld, 9 = status (bijv. L2026 of A), " & _
               "0 = overal zoeken (alle kolommen). Een deel van de tekst is genoeg, behalve bij het volgnummer. " & _
               "Zijn er meer treffers, dan kies je er een uit de lijst. Bij meer dan " & MAX_KEUZE & " treffers komen ze " & _
               "allemaal op het blad 'Krediet raadplegen'; zoek dan opnieuw op volgnummer. Je ziet alle gegevens van " & _
               "het krediet, de andere delen als het een samengesteld krediet is, en de wijzigingen uit het Logboek."
End Sub

Public Sub KredietRaadplegen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, v As Variant, laatste As Long, r As Long, c As Long, n As Long, kk As Long
    Dim keuze As String, zoek As String, wat As String, lijst As String, antw As String, hit As Boolean
    Dim treffers() As Long

    On Error GoTo Fout
    keuze = Trim$(InputBox("Zoeken op:" & vbCrLf & _
                           "1 = volgnummer (kolom A)" & vbCrLf & _
                           "2 = kredietnummer (kolom C)" & vbCrLf & _
                           "3 = omschrijving (kolom D)" & vbCrLf & _
                           "4 = activanummer (kolom X)" & vbCrLf & _
                           "5 = budgetbeheerder" & vbCrLf & _
                           "6 = projectleider" & vbCrLf & _
                           "7 = tijdschrijver" & vbCrLf & _
                           "8 = taakveld" & vbCrLf & _
                           "9 = status (bijv. L2026 of A)" & vbCrLf & _
                           "0 = overal zoeken (alle kolommen)" & vbCrLf & vbCrLf & _
                           "Een deel van de tekst is genoeg (behalve bij het volgnummer).", "Krediet raadplegen", "2"))
    If Len(keuze) = 0 Then Exit Sub
    If Len(keuze) <> 1 Or InStr("0123456789", keuze) = 0 Then
        MsgBox "Typ een cijfer van 0 tot en met 9.", vbExclamation
        Exit Sub
    End If
    wat = Choose(Val(keuze) + 1, "overal", "volgnummer", "kredietnummer", "omschrijving", "activanummer", _
                 "budgetbeheerder", "projectleider", "tijdschrijver", "taakveld", "status")
    zoek = Trim$(InputBox("Wat zoek je? (" & wat & ")", "Krediet raadplegen"))
    If Len(zoek) = 0 Then Exit Sub

    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    v = wv.Range(wv.Cells(1, 1), wv.Cells(laatste, LAATSTE_KOL)).Value
    ReDim treffers(1 To laatste)
    For r = 2 To laatste
        If Not Leeg(v(r, 1)) Then
            hit = False
            Select Case keuze
                Case "1": hit = (Trim$(CStr(v(r, 1))) = zoek)
                Case "2": hit = Bevat(NrTekst(v(r, 3)), zoek)
                Case "3": hit = Bevat(v(r, 4), zoek)
                Case "4": hit = Bevat(v(r, 24), zoek)
                Case "5": hit = Bevat(v(r, 15), zoek)
                Case "6": hit = Bevat(v(r, 16), zoek)
                Case "7": hit = Bevat(v(r, 17), zoek)
                Case "8": hit = Bevat(v(r, 19), zoek) Or Bevat(CodeNaarWeergave("Tbl_Taakveld", v(r, 19)), zoek)
                Case "9": hit = (StrComp(Left$(Trim$(CStr(v(r, 8))), Len(zoek)), zoek, vbTextCompare) = 0)
                Case "0"
                    For c = 1 To LAATSTE_KOL
                        If Not IsError(v(r, c)) Then
                            If Bevat(v(r, c), zoek) Then hit = True: Exit For
                        End If
                    Next c
            End Select
            If hit Then
                n = n + 1
                treffers(n) = r
            End If
        End If
    Next r

    If n = 0 Then
        MsgBox "Niets gevonden op " & wat & " '" & zoek & "'.", vbInformation
        Exit Sub
    End If
    If n > MAX_KEUZE Then
        TreffersTonen wv, treffers, n, wat & " '" & zoek & "'"
        MsgBox n & " treffers. Ze staan op het blad '" & BL_RAADPLEGEN & "'." & vbCrLf & _
               "Zoek daarna opnieuw op volgnummer (1) om een krediet te bekijken.", vbInformation
        Exit Sub
    End If
    If n = 1 Then
        r = treffers(1)
    Else
        For kk = 1 To n
            r = treffers(kk)
            lijst = lijst & kk & " = " & NrTekst(v(r, 3)) & " | " & Left$(CStr(v(r, 4)), 30) & " | " & v(r, 8) & vbCrLf
        Next kk
        antw = Trim$(InputBox(n & " treffers:" & vbCrLf & vbCrLf & lijst & vbCrLf & "Typ het nummer:", "Krediet raadplegen"))
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
    KredietTonen wv, v, r, laatste
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout bij het raadplegen: " & Err.Description, vbCritical
End Sub

' Tekst bevat de zoekterm (hoofdletters maken niet uit).
Private Function Bevat(ByVal waarde As Variant, ByVal zoek As String) As Boolean
    If IsError(waarde) Then Exit Function
    Bevat = (InStr(1, CStr(waarde), zoek, vbTextCompare) > 0)
End Function

' Maakt het blad 'Krediet raadplegen' leeg en klaar (een vorige raadpleging wordt vervangen).
Private Function RaadpleegBlad(ByVal titel As String, ByVal ondertitel As String) As Worksheet
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    If BladBestaatStil(BL_RAADPLEGEN) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(BL_RAADPLEGEN).Unprotect WACHTWOORD
        ThisWorkbook.Worksheets(BL_RAADPLEGEN).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = NieuwBlad()
    ws.Name = BL_RAADPLEGEN
    ws.Tab.Color = RGB(0, 73, 118)
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = titel
    With ws.Range("A1").Font
        .Bold = True
        .Size = 16
        .Color = RGB(0, 73, 118)
    End With
    ws.Rows(1).RowHeight = 26
    ws.Range("A2").Value = ondertitel
    ws.Range("A2").Font.Size = 11
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Rows(3).RowHeight = 22
    Set RaadpleegBlad = ws
End Function

' Kopregel van een blok (donkerblauw, witte letters) en een sectiekop (lichtblauw).
Private Sub RpKop(ByVal ws As Worksheet, ByVal rij As Long, ByVal koppen As Variant)
    Dim k As Long
    For k = 0 To UBound(koppen)
        ws.Cells(rij, k + 1).Value = koppen(k)
    Next k
    With ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 8))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 73, 118)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub RpSectie(ByVal ws As Worksheet, ByVal rij As Long, ByVal tekst As String)
    ws.Cells(rij, 1).Value = tekst
    With ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 8))
        .Font.Bold = True
        .Font.Color = RGB(0, 73, 118)
        .Interior.Color = RGB(214, 228, 240)
    End With
End Sub

' Laatste stap: opmaak, afdrukken, vastzetten en tonen.
Private Sub BladAfronden(ByVal ws As Worksheet, ByVal lr As Long)
    ws.Range(ws.Cells(5, 1), ws.Cells(lr, 8)).VerticalAlignment = xlTop
    KnopTerug ws, ws.Range("A3")
    AfdrukA3 ws, "", 4, lr, 8
    ws.Protect Password:=WACHTWOORD, DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFormattingRows:=True
    TabbladenOrdenen
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
End Sub

' Veel treffers: een lijst op het blad, om daarna op volgnummer verder te zoeken.
Private Sub TreffersTonen(ByVal wv As Worksheet, ByRef treffers() As Long, ByVal n As Long, ByVal gezocht As String)
    Dim ws As Worksheet, i As Long, r As Long, rij As Long
    Set ws = RaadpleegBlad("Krediet raadplegen - zoekresultaat", "Gezocht op " & gezocht & ": " & n & " treffers, stand " & _
                           Format$(Date, "dd-mm-yyyy") & ". Zoek opnieuw op volgnummer om een krediet te bekijken.")
    RpKop ws, 4, Array("Volgnr", "Kredietnummer", "Omschrijving", "Status", "Budgetbeheerder", "Projectleider", "Activanummer", "Restantkrediet")
    rij = 5
    For i = 1 To n
        r = treffers(i)
        ws.Cells(rij, 1).Value = wv.Cells(r, 1).Value
        ws.Cells(rij, 2).NumberFormat = "@"
        ws.Cells(rij, 2).Value = NrTekst(wv.Cells(r, 3).Value)
        ws.Cells(rij, 3).Value = wv.Cells(r, 4).Value
        ws.Cells(rij, 4).Value = wv.Cells(r, 8).Value
        ws.Cells(rij, 5).Value = wv.Cells(r, 15).Value
        ws.Cells(rij, 6).Value = wv.Cells(r, 16).Value
        ws.Cells(rij, 7).NumberFormat = "@"
        ws.Cells(rij, 7).Value = CStr(wv.Cells(r, 24).Value)
        ws.Cells(rij, 8).Value = wv.Cells(r, 30).Value
        rij = rij + 1
    Next i
    ws.Range(ws.Cells(5, 8), ws.Cells(rij, 8)).NumberFormat = "#,##0;-#,##0;-"
    ws.Range(ws.Cells(4, 1), ws.Cells(rij, 7)).HorizontalAlignment = xlLeft
    ws.Columns(1).ColumnWidth = 9
    ws.Columns(2).ColumnWidth = 16
    ws.Columns(3).ColumnWidth = 50
    ws.Columns(4).ColumnWidth = 9
    ws.Range(ws.Columns(5), ws.Columns(6)).ColumnWidth = 24
    ws.Columns(7).ColumnWidth = 16
    ws.Columns(8).ColumnWidth = 15
    ws.Range(ws.Cells(5, 3), ws.Cells(rij, 3)).WrapText = True
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
    BladAfronden ws, rij
End Sub

' Alles van een krediet: alle velden van de verzamellijst, de andere delen van een samengesteld krediet en het Logboek.
Private Sub KredietTonen(ByVal wv As Worksheet, ByRef v As Variant, ByVal r As Long, ByVal laatste As Long)
    Dim ws As Worksheet, rij As Long, c As Long, x As Long, t As String, veld As String, i As Long
    Dim lg As Worksheet, lv As Variant, lLaatste As Long, nLog As Long, volgnr As String, nr As String

    volgnr = Trim$(CStr(v(r, 1)))
    nr = NrTekst(v(r, 3))
    Set ws = RaadpleegBlad("Krediet " & nr & " - " & CStr(v(r, 4)), _
                           "Volgnummer " & volgnr & " (rij " & r & " van de verzamellijst), stand " & Format$(Date, "dd-mm-yyyy") & _
                           ". Alleen bekijken: wijzigen gaat via 'Krediet wijzigen'.")
    ws.Columns(1).ColumnWidth = 15
    ws.Columns(2).ColumnWidth = 34
    ws.Columns(3).ColumnWidth = 22
    ws.Columns(4).ColumnWidth = 12
    ws.Range(ws.Columns(5), ws.Columns(7)).ColumnWidth = 22
    ws.Columns(8).ColumnWidth = 60

    ' 1. Alle gegevens: kolom, veldnaam, waarde (zoals in de verzamellijst getoond; codes met omschrijving)
    RpKop ws, 4, Array("Kolom", "Veld", "Waarde")
    ws.Range(ws.Cells(4, 3), ws.Cells(4, 8)).Merge
    rij = 5
    For c = 1 To LAATSTE_KOL
        veld = Trim$(CStr(wv.Cells(1, c).Text))
        If Len(veld) > 0 Then
            t = CStr(wv.Cells(r, c).Text)
            If c = 2 Then t = CStr(CodeNaarWeergave("Tbl_Programma", v(r, 2)))
            If c = 19 And Not Leeg(v(r, 19)) Then t = CStr(v(r, 19)) & " - " & CStr(CodeNaarWeergave("Tbl_Taakveld", v(r, 19)))
            If Len(Trim$(t)) = 0 Then t = "-"
            ws.Cells(rij, 1).Value = Split(wv.Cells(1, c).Address(True, False), "$")(0)
            ws.Cells(rij, 2).Value = veld
            ws.Range(ws.Cells(rij, 3), ws.Cells(rij, 8)).Merge
            ws.Cells(rij, 3).NumberFormat = "@"
            ws.Cells(rij, 3).Value = t
            ws.Cells(rij, 3).HorizontalAlignment = xlLeft
            ws.Cells(rij, 3).WrapText = True
            If t = "-" Then ws.Cells(rij, 3).Font.Color = RGB(166, 166, 166)
            If Len(t) > 110 Then ws.Rows(rij).RowHeight = Application.Min(409, 13.5 + Int(Len(t) / 110) * 12.75)
            If rij Mod 2 = 0 Then ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 8)).Interior.Color = RGB(242, 242, 242)
            rij = rij + 1
        End If
    Next c
    ws.Range(ws.Cells(5, 1), ws.Cells(rij, 1)).Font.Color = RGB(128, 128, 128)
    ws.Range(ws.Cells(5, 1), ws.Cells(rij, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(5, 2), ws.Cells(rij, 2)).Font.Bold = True

    ' 2. Samengesteld krediet: alle delen met hetzelfde getal in kolom J
    If IsGetal(v(r, 10)) Then
        rij = rij + 1
        RpSectie ws, rij, "Samengesteld krediet: alle delen (zelfde nummer in kolom J: " & v(r, 10) & ")"
        rij = rij + 1
        RpKop ws, rij, Array("Volgnr", "Omschrijving", "Kredietnummer", "Status", "Budget t/m lopend jaar", "Werkelijke uitgaven", "Restantkrediet", "Toelichting")
        rij = rij + 1
        For x = 2 To laatste
            If Not Leeg(v(x, 1)) And CStr(v(x, 10)) = CStr(v(r, 10)) Then
                ws.Cells(rij, 1).Value = v(x, 1)
                ws.Cells(rij, 2).Value = v(x, 4)
                ws.Cells(rij, 3).NumberFormat = "@"
                ws.Cells(rij, 3).Value = NrTekst(v(x, 3))
                ws.Cells(rij, 4).Value = v(x, 8)
                ws.Cells(rij, 5).Value = Nul(v(x, 26))
                ws.Cells(rij, 6).Value = Nul(v(x, 27)) + Nul(v(x, 28)) + Nul(v(x, 29))
                ws.Cells(rij, 7).Value = Nul(v(x, 30))
                ws.Cells(rij, 8).Value = v(x, 36)
                ws.Range(ws.Cells(rij, 5), ws.Cells(rij, 7)).NumberFormat = "#,##0;-#,##0;-"
                If x = r Then ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 8)).Font.Bold = True     ' het geraadpleegde deel
                rij = rij + 1
            End If
        Next x
        ws.Range(ws.Cells(rij - 1, 1), ws.Cells(rij - 1, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If

    ' 3. Historie uit het Logboek (op volgnummer of kredietnummer), nieuwste eerst, maximaal 100 regels
    rij = rij + 1
    RpSectie ws, rij, "Wijzigingen (Logboek), nieuwste eerst"
    rij = rij + 1
    RpKop ws, rij, Array("Datum", "Actie", "Gebruiker", "Veld", "Oud", "Nieuw", "", "Toelichting")
    rij = rij + 1
    If BladBestaatStil("Logboek") Then
        Set lg = ThisWorkbook.Worksheets("Logboek")
        lLaatste = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row
        If lLaatste >= 2 Then
            lv = lg.Range("A1:K" & lLaatste).Value
            For i = lLaatste To 2 Step -1
                If (Len(volgnr) > 0 And Trim$(CStr(lv(i, 6))) = volgnr) Or (Len(nr) > 0 And Trim$(CStr(lv(i, 7))) = nr) Then
                    ws.Cells(rij, 1).Value = lv(i, 1)
                    ws.Cells(rij, 1).NumberFormat = "dd-mm-yyyy hh:mm"
                    ws.Cells(rij, 2).Value = lv(i, 5)
                    ws.Cells(rij, 3).Value = lv(i, 2)
                    ws.Cells(rij, 4).Value = lv(i, 8)
                    ws.Cells(rij, 5).NumberFormat = "@"
                    ws.Cells(rij, 5).Value = CStr(lv(i, 9))
                    ws.Cells(rij, 6).NumberFormat = "@"
                    ws.Cells(rij, 6).Value = CStr(lv(i, 10))
                    ws.Cells(rij, 8).Value = lv(i, 11)
                    rij = rij + 1
                    nLog = nLog + 1
                    If nLog >= 100 Then Exit For
                End If
            Next i
        End If
    End If
    If nLog = 0 Then
        ws.Cells(rij, 1).Value = "Geen wijzigingen gevonden in het Logboek."
        ws.Cells(rij, 1).Font.Italic = True
        rij = rij + 1
    End If
    BladAfronden ws, rij
End Sub
