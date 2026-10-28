Attribute VB_Name = "Kredieten_6_Uitvraag"
Option Explicit

' ======================================================================
' Uitvraag kadernota/begroting (stap 7)
' ======================================================================
' Een blauw blad 'Uitvraag' (BL_UITVRAAG) met alle budgetbeheerders. Per budgetbeheerder staat eerst een kopregel met de
' data (verstuurd, ingeleverd, opgehaald, verwerkt) en daaronder zijn lopende kredieten. Drie knoppen op het Dashboard:
'   1. Uitvraag versturen      : UitvraagBegroting   (bestand per budgetbeheerder in map 02 + het blad Uitvraag)
'   2. Aanleveringen ophalen   : UitvraagOphalen     (een budgetbeheerder kiezen; alleen diens regels worden bijgewerkt)
'   3. Uitvraag verwerken      : UitvraagVerwerken   (goedgekeurde regels naar de Verzamellijst, direct Te controleren)
' Kolommen op het blad: 1 volgnr, 2 kredietnummer, 3 omschrijving, 4 nieuwe omschrijving, 5-9 was, 10-14 wordt, 15 wijziging,
' 16 onderbouwing, 17 besluit, 18 opmerking, 19 status, 20 signaal. Verborgen: 21 besluit door de macro, 22-26 aangeleverd,
' 27 aangeleverde omschrijving, 28 aangeleverde onderbouwing, 29 budgetbeheerder, 30 soort rij (KOP of REG), 31-34 datums.

Private Const UV_EERSTE As Long = 6
Private Const UV_SIGNAAL As String = "aanlevering gewijzigd na beoordeling"

Private mWv As Worksheet
Private mD As Variant
Private mVolgRij As Object
Private mNWijz As Long
Private mNSig As Long
Private mNNieuw As Long

Public Function UvKol(ByVal naam As String) As Long
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
        Case "signaal": UvKol = 20
        Case "besluit0": UvKol = 21      ' besluit zoals de macro het zette
        Case "aang": UvKol = 22          ' 22 t/m 26
        Case "aangoms": UvKol = 27
        Case "aangtoel": UvKol = 28
        Case "bh": UvKol = 29
        Case "type": UvKol = 30
        Case "verstuurd": UvKol = 31
        Case "ingeleverd": UvKol = 32
        Case "opgehaald": UvKol = 33
        Case "verwerkt": UvKol = 34
    End Select
End Function

' Nieuwste uitvraagbestand van een budgethouder (leeg = geen).
Public Function UvBestand(ByVal naam As String, ByRef dt As Date) As String
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

' Stond de Dashboard-indeling al op de nieuwe opzet (2a, 2b, 2c met de drie uitvraagknoppen)?
Public Function DashboardIndelingKlopt() As Boolean
    Dim ws As Worksheet, s As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    If CStr(ws.Range("B" & (NM_KOP + 1)).Value) <> "2a. Kredietoverzichten" Then Exit Function
    Set s = ws.Shapes("btnUvOphalen")
    If s Is Nothing Then Exit Function
    If CStr(ws.Cells(CB_LAATSTE + 1, 2).Value) Like "Werkelijke uitgaven hoger*" Then Exit Function     ' oude extra controleregel opruimen
    DashboardIndelingKlopt = True
End Function

' Plaatst de knoppen van het Dashboard opnieuw als de indeling nog de oude is. Een fout wordt gemeld in plaats van stilzwijgend overgeslagen.
Public Sub DashboardIndelingBijwerken()
    Dim foutTekst As String
    On Error GoTo Fout
    If DashboardIndelingKlopt() Then Exit Sub
    DashboardKnoppenPlaatsen
    Exit Sub
Fout:
    foutTekst = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    MsgBox "De nieuwe indeling van het Dashboard kon niet worden gemaakt." & vbCrLf & vbCrLf & "Fout: " & foutTekst & vbCrLf & vbCrLf & _
           "Stuur deze melding door. Het Dashboard blijft zoals het was.", vbExclamation, "Dashboard"
End Sub

' Bouwt het Dashboard altijd opnieuw op (ook als de indeling al nieuw lijkt) en meldt wat er gebeurt. Starten via Ontwikkelaars > Macro's.
Public Sub DashboardOpnieuwOpbouwen()
    Dim foutTekst As String
    On Error GoTo Fout
    DashboardKnoppenPlaatsen
    MsgBox "Het Dashboard is opnieuw opgebouwd." & vbCrLf & _
           "Onderdeel 2c moet nu drie knoppen hebben: '1. Uitvraag versturen', '2. Aanleveringen ophalen' en '3. Uitvraag verwerken'.", _
           vbInformation, "Dashboard"
    Exit Sub
Fout:
    foutTekst = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    MsgBox "Het Dashboard kon niet worden opgebouwd." & vbCrLf & vbCrLf & "Fout: " & foutTekst, vbExclamation, "Dashboard"
End Sub

' Is de uitvraag gemaakt in het lopende jaar? Anders een melding.
Public Function UvJaarKlopt() As Boolean
    Dim t As String, j As Long, lj As Long
    On Error Resume Next
    t = ThisWorkbook.Names("Uitvraag_Jaar").RefersTo
    j = CLng(Val(Mid$(t, 2)))
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    On Error GoTo 0
    If j = 0 Then
        MsgBox "Er is niet vastgelegd voor welk jaar deze uitvraag is gemaakt." & vbCrLf & _
               "Maak de uitvraag opnieuw met '1. Uitvraag versturen'.", vbExclamation, "Uitvraag"
    ElseIf j <> lj Then
        MsgBox "Deze uitvraag is gemaakt in " & j & ", maar het lopende jaar is nu " & lj & "." & vbCrLf & _
               "Ophalen en verwerken kan alleen in het jaar waarin de uitvraag is gemaakt.", vbExclamation, "Uitvraag"
    Else
        UvJaarKlopt = True
    End If
End Function

' Aantal regels op het blad Uitvraag dat nog niet is verwerkt: een besluit, geen status en wel een wijziging of voorstel.
' De jaarovergang wordt hierdoor geblokkeerd.
Public Function UvOnverwerkt() As Long
    Dim ws As Worksheet, r As Long, n As Long
    If Not BladBestaatStil(BL_UITVRAAG) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(BL_UITVRAAG)
    For r = UV_EERSTE To UvLaatsteRij(ws)
        If CStr(ws.Cells(r, 30).Value) = "REG" Then
            If Len(CStr(ws.Cells(r, 17).Value)) > 0 And Len(CStr(ws.Cells(r, 19).Value)) = 0 _
               And CStr(ws.Cells(r, 15).Value) <> GEEN_WIJZ Then n = n + 1
        End If
    Next r
    UvOnverwerkt = n
End Function

Private Function UvLaatsteRij(ByVal ws As Worksheet) As Long
    UvLaatsteRij = Application.Max(UV_EERSTE - 1, ws.Cells(ws.Rows.Count, 30).End(xlUp).Row)
End Function

Private Function UvKopRij(ByVal ws As Worksheet, ByVal naam As String) As Long
    Dim r As Long
    For r = UV_EERSTE To UvLaatsteRij(ws)
        If CStr(ws.Cells(r, 30).Value) = "KOP" Then
            If StrComp(Trim$(CStr(ws.Cells(r, 29).Value)), naam, vbTextCompare) = 0 Then UvKopRij = r: Exit Function
        End If
    Next r
End Function

Private Function UvSectieEinde(ByVal ws As Worksheet, ByVal kop As Long) As Long
    Dim r As Long, lr As Long
    lr = UvLaatsteRij(ws)
    r = kop + 1
    Do While r <= lr
        If CStr(ws.Cells(r, 30).Value) = "KOP" Then Exit Do
        r = r + 1
    Loop
    UvSectieEinde = r - 1
End Function

Private Function UvDatum(ByVal v As Variant) As String
    If IsDate(v) Then UvDatum = Format$(CDate(v), "dd-mm-yyyy") Else UvDatum = "-"
End Function

' Tekst op de kopregel van een budgetbeheerder: status, data en aantal open regels.
' Rood = nog niet ingeleverd, oranje = ingeleverd maar nog niet opgehaald, blauw = opgehaald met open regels, groen = afgerond.
Private Sub UvKopTekst(ByVal ws As Worksheet, ByVal kop As Long)
    Dim r As Long, eind As Long, nOpen As Long, nSig As Long, t As String, st As String, kleur As Long
    eind = UvSectieEinde(ws, kop)
    For r = kop + 1 To eind
        If CStr(ws.Cells(r, 30).Value) = "REG" Then
            If Len(CStr(ws.Cells(r, 17).Value)) > 0 And Len(CStr(ws.Cells(r, 19).Value)) = 0 _
               And CStr(ws.Cells(r, 15).Value) <> GEEN_WIJZ Then nOpen = nOpen + 1
            If Len(CStr(ws.Cells(r, 20).Value)) > 0 Then nSig = nSig + 1
        End If
    Next r
    If Not IsDate(ws.Cells(kop, 32).Value) Then
        st = "NOG NIET INGELEVERD": kleur = RGB(192, 0, 0)
    ElseIf Not IsDate(ws.Cells(kop, 33).Value) Then
        st = "ingeleverd, nog niet opgehaald": kleur = RGB(197, 90, 17)
    ElseIf nOpen > 0 Then
        st = "opgehaald, " & nOpen & " regel(s) open": kleur = RGB(0, 73, 118)
    Else
        st = "afgerond": kleur = RGB(0, 112, 60)
    End If
    t = st & "  |  verstuurd " & UvDatum(ws.Cells(kop, 31).Value) & "  |  ingeleverd " & UvDatum(ws.Cells(kop, 32).Value) & _
        "  |  opgehaald " & UvDatum(ws.Cells(kop, 33).Value) & "  |  verwerkt " & UvDatum(ws.Cells(kop, 34).Value) & _
        IIf(nSig > 0, "  |  " & nSig & " met signaal", "")
    ws.Cells(kop, 5).Value = t
    ws.Cells(kop, 5).Font.Color = kleur
    ws.Cells(kop, 5).Font.Bold = True
End Sub

' Rijhoogte van de regels: 13,5, alleen hoger bij een lange toelichting (dat regelt BladOpmaken).
Private Sub UvRijenPassend(ByVal ws As Worksheet)
    Dim r As Long
    For r = UV_EERSTE To UvLaatsteRij(ws)
        If CStr(ws.Cells(r, 30).Value) = "KOP" Then
            ws.Rows(r).RowHeight = 13.5                       ' kopregel van een budgetbeheerder: altijd de standaardhoogte
        End If
    Next r
End Sub

' Opmaak en invoercellen van een regel.
Private Sub UvRegelOpmaak(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Variant
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 20)).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 20)).Font.Bold = False
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 20)).Font.Color = RGB(0, 0, 0)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 20)).Locked = True
    For Each c In Array(4, 10, 11, 12, 13, 14, 16, 17, 18)
        ws.Cells(r, CLng(c)).Interior.Color = RGB(255, 217, 102)
        ws.Cells(r, CLng(c)).Locked = False
    Next c
    With ws.Cells(r, 17)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="goedkeuren,afwijzen,nog niet"
    End With
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 14)).NumberFormat = "#,##0;-#,##0;-"
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).WrapText = True
    ws.Range(ws.Cells(r, 15), ws.Cells(r, 18)).WrapText = True
    ws.Range(ws.Cells(r, 20), ws.Cells(r, 20)).WrapText = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 20)).VerticalAlignment = xlTop
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).HorizontalAlignment = xlLeft
End Sub

Private Function UvJaarWaarde() As Long
    Dim t As String
    On Error Resume Next
    t = ThisWorkbook.Names("Uitvraag_Jaar").RefersTo
    UvJaarWaarde = CLng(Val(Mid$(t, 2)))
End Function

' Bouwt het blad Uitvraag. Zonder naam: het blad wordt opnieuw gemaakt met alle budgetbeheerders met lopende kredieten.
' Met een naam: alleen de regels van die budgetbeheerder worden (opnieuw) op het bestaande blad gezet; de rest blijft staan.
Private Function UvBladOpbouwen(ByVal alleenNaam As String) As Worksheet
    Dim wv As Worksheet, ws As Worksheet, v As Variant, laatste As Long, n As Long, r As Long, i As Long, k As Long, lj As Long
    Dim sel() As Long, sl() As String, tk() As Boolean, nSel As Long, rij As Long, bh As String, vBh As String
    Dim bouwNieuw As Boolean, kop As Long, ins As Long, r2 As Long, kopStart As Long
    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    v = wv.Range("A2:AW" & laatste).Value
    JGetalNormaliseren v
    n = UBound(v, 1)
    ReDim sel(1 To n): ReDim sl(1 To n): ReDim tk(1 To n)
    For r = 1 To n
        If Not Leeg(v(r, 1)) And UCase$(Left$(Trim$(CStr(v(r, 8))), 1)) = "L" And Len(Trim$(CStr(v(r, 15)))) > 0 Then
            If Len(alleenNaam) = 0 Or StrComp(Trim$(CStr(v(r, 15))), alleenNaam, vbTextCompare) = 0 Then
                nSel = nSel + 1
                sel(nSel) = r
                sl(nSel) = LCase$(Trim$(CStr(v(r, 15)))) & Chr$(1) & SorteerDeel(v(r, 2)) & Chr$(1) & NrTekst(v(r, 3)) & Chr$(1) & _
                           Format$(v(r, 1), "000000000")
            End If
        End If
    Next r
    If nSel = 0 Then Exit Function
    SorteerSleutels sl, sel, tk, 1, nSel

    bouwNieuw = (Len(alleenNaam) = 0) Or Not BladBestaatStil(BL_UITVRAAG)
    If Not bouwNieuw Then
        If UvJaarWaarde() <> lj Then bouwNieuw = True
    End If
    If bouwNieuw Then
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
        ws.Range("A1").Value = "Uitvraag kadernota / begroting " & (lj + 1) & " (budgetten " & (lj + 1) & " t/m " & (lj + 5) & ")"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 16
        ws.Range("A1").Font.Color = RGB(0, 73, 118)
        ws.Rows(1).RowHeight = 26
        ws.Range("A2").Value = "Gemaakt in " & lj & ". Per budgetbeheerder: eerst '2. Aanleveringen ophalen', kies per regel een besluit (kolom Q) en klik daarna op '3. Uitvraag verwerken' op het Dashboard. Gouden cellen mag je aanpassen."
        ws.Range("A2").Value = ws.Range("A2").Value & vbLf & "Kopregel per budgetbeheerder: rood = nog niet ingeleverd, oranje = ingeleverd maar nog niet opgehaald, " & _
                               "blauw = opgehaald met open regels, groen = afgerond. De status wordt bijgewerkt als je '2. Aanleveringen ophalen' klikt: " & _
                               "de macro kijkt dan in map 02 of het bestand van de budgetbeheerder na het versturen is opgeslagen."
        ws.Range("A2:S2").Merge
        ws.Range("A2").WrapText = True
        ws.Range("A2").VerticalAlignment = xlTop
        ws.Rows(2).RowHeight = 58
        ws.Range("A2").Font.Italic = True
        ws.Range("A2").Font.Color = RGB(89, 89, 89)
        ws.Range("A5:T5").Value = Array("Volgnr", "Kredietnummer", "Omschrijving (nu)", "Nieuwe omschrijving", _
                                        "Was " & (lj + 1), "Was " & (lj + 2), "Was " & (lj + 3), "Was " & (lj + 4), "Was " & (lj + 5), _
                                        "Wordt " & (lj + 1), "Wordt " & (lj + 2), "Wordt " & (lj + 3), "Wordt " & (lj + 4), "Wordt " & (lj + 5), _
                                        "Wijziging", "Onderbouwing (wordt toelichting)", "Besluit", "Opmerking financieel adviseur", "Verwerkt", "Signaal")
        With ws.Range("A5:T5")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 73, 118)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(5).RowHeight = 30
        ws.Cells.Locked = True
        rij = UV_EERSTE
    Else
        Set ws = ThisWorkbook.Worksheets(BL_UITVRAAG)
        ws.Unprotect WACHTWOORD
        kop = UvKopRij(ws, alleenNaam)
        If kop > 0 Then ws.Rows(kop & ":" & UvSectieEinde(ws, kop)).Delete
        ins = UvLaatsteRij(ws) + 1
        For r2 = UV_EERSTE To UvLaatsteRij(ws)
            If CStr(ws.Cells(r2, 30).Value) = "KOP" Then
                If StrComp(LCase$(Trim$(CStr(ws.Cells(r2, 29).Value))), LCase$(alleenNaam), vbBinaryCompare) > 0 Then ins = r2: Exit For
            End If
        Next r2
        ws.Rows(ins & ":" & (ins + nSel)).Insert        ' een kopregel en nSel regels
        ws.Range(ws.Cells(ins, 1), ws.Cells(ins + nSel, 34)).ClearContents
        rij = ins
    End If
    kopStart = rij

    vBh = Chr$(0)
    For i = 1 To nSel
        r = sel(i)
        bh = Trim$(CStr(v(r, 15)))
        If bh <> vBh Then
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 20)).Locked = True
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 20)).WrapText = False
            ws.Rows(rij).RowHeight = 13.5
            ws.Cells(rij, 17).Validation.Delete
            ws.Range(ws.Cells(rij, 5), ws.Cells(rij, 14)).NumberFormat = "General"
            ws.Cells(rij, 1).Value = bh
            ws.Cells(rij, 29).Value = bh
            ws.Cells(rij, 30).Value = "KOP"
            ws.Cells(rij, 31).Value = Now
            With ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 20))
                .Font.Bold = True
                .Interior.Color = RGB(214, 228, 240)
                .Font.Color = RGB(0, 73, 118)
            End With
            UvKopTekst ws, rij
            rij = rij + 1
            vBh = bh
        End If
        ws.Cells(rij, 1).Value = v(r, 1)
        ws.Cells(rij, 2).NumberFormat = "@"
        ws.Cells(rij, 2).Value = NrTekst(v(r, 3))
        ws.Cells(rij, 3).Value = v(r, 4)
        For k = 1 To 5
            ws.Cells(rij, 4 + k).Value = Nul(v(r, 30 + k))
        Next k
        ws.Cells(rij, 29).Value = bh
        ws.Cells(rij, 30).Value = "REG"
        UvRegelOpmaak ws, rij
        rij = rij + 1
    Next i
    ' kopregels opnieuw tellen (regels staan er nu)
    For r = kopStart To rij - 1
        If CStr(ws.Cells(r, 30).Value) = "KOP" Then UvKopTekst ws, r
    Next r
    If bouwNieuw Then
        ws.Columns(1).ColumnWidth = 8: ws.Columns(2).ColumnWidth = 12: ws.Columns(3).ColumnWidth = 34: ws.Columns(4).ColumnWidth = 26
        ws.Range(ws.Columns(5), ws.Columns(14)).ColumnWidth = 11
        ws.Columns(15).ColumnWidth = 30: ws.Columns(16).ColumnWidth = 40: ws.Columns(17).ColumnWidth = 12
        ws.Columns(18).ColumnWidth = 30: ws.Columns(19).ColumnWidth = 18: ws.Columns(20).ColumnWidth = 30
        ws.Range(ws.Columns(21), ws.Columns(34)).Hidden = True
    End If
    BladOpmaken ws, 5, UV_EERSTE, UvLaatsteRij(ws)
    UvRijenPassend ws
    If bouwNieuw Then
        KnopTerug ws, ws.Range("A4")
        ws.Rows(4).RowHeight = 22
        ws.Activate
        ActiveWindow.FreezePanes = False
        ws.Range("D6").Select
        ActiveWindow.FreezePanes = True
    End If
    BladBeveiligen ws
    Set UvBladOpbouwen = ws
End Function

' Laat kiezen voor wie de uitvraag wordt gemaakt: alle budgetbeheerders (0) of een. Geeft "#ALLE", een naam, of "" als er niets is gekozen.
Private Function UvKiesBereik(ByRef nm() As String, ByVal n As Long) As String
    Dim lijst As String, i As Long, antw As String, res As String, nr As Long, hit As Long, getoond As Long
    Do
        lijst = "0 = alle budgetbeheerders" & vbCrLf
        getoond = 0
        For i = 1 To n
            If Len(lijst) < 800 Then
                lijst = lijst & i & " = " & nm(i) & vbCrLf
                getoond = i
            End If
        Next i
        If getoond < n Then lijst = lijst & "... en nog " & (n - getoond) & " namen: typ een deel van de naam." & vbCrLf
        antw = Trim$(InputBox("Voor wie maak je de uitvraag?" & vbCrLf & vbCrLf & lijst & vbCrLf & "Typ het nummer of een deel van de naam:", _
                              "Uitvraag versturen", "0"))
        If Len(antw) = 0 Then Exit Function
        If IsNumeric(antw) Then
            nr = CLng(antw)
            If nr = 0 Then
                UvKiesBereik = "#ALLE"
                Exit Function
            End If
            If nr >= 1 And nr <= n Then
                UvKiesBereik = nm(nr)
                Exit Function
            End If
            MsgBox "Dat nummer staat niet in de lijst.", vbExclamation, "Uitvraag versturen"
        Else
            hit = 0: res = ""
            For i = 1 To n
                If InStr(1, nm(i), antw, vbTextCompare) > 0 Then
                    hit = hit + 1
                    res = nm(i)
                End If
            Next i
            If hit = 1 Then
                UvKiesBereik = res
                Exit Function
            ElseIf hit = 0 Then
                MsgBox "Geen budgetbeheerder gevonden met '" & antw & "'.", vbInformation, "Uitvraag versturen"
            Else
                MsgBox "Meerdere namen passen bij '" & antw & "'. Typ meer van de naam, of het nummer.", vbInformation, "Uitvraag versturen"
            End If
        End If
    Loop
End Function

Private Sub UvTijdelijkWeg()
    If BladBestaatStil(BL_UVTEMP) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(BL_UVTEMP).Unprotect WACHTWOORD
        ThisWorkbook.Worksheets(BL_UVTEMP).Delete
        Application.DisplayAlerts = True
    End If
End Sub

' ----------------------------------------------------------------------
' 1. Uitvraag versturen
' ----------------------------------------------------------------------
Public Sub UitvraagBegroting()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, ws As Worksheet, map As String, namen As Object, nm As Variant, laatste As Long, r As Long, n As Long, lj As Long
    Dim foutTekst As String, lijst() As String, nL As Long, i As Long, j As Long, t As String, bereik As String, alle As Boolean, vraag As String
    Dim pad As String, bn As String, t2 As Variant
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "De uitvraag wordt gemaakt door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    Set namen = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(wv.Cells(r, 1).Value) And UCase$(Left$(Trim$(CStr(wv.Cells(r, 8).Value)), 1)) = "L" _
           And Not Leeg(wv.Cells(r, 15).Value) Then namen(Trim$(CStr(wv.Cells(r, 15).Value))) = True
    Next r
    If namen.Count = 0 Then
        MsgBox "Er zijn geen lopende kredieten met een budgetbeheerder.", vbInformation
        Exit Sub
    End If
    ' gesorteerde lijst met namen
    nL = namen.Count
    ReDim lijst(1 To nL)
    i = 0
    For Each nm In namen.Keys
        i = i + 1
        lijst(i) = CStr(nm)
    Next nm
    For i = 2 To nL
        t = lijst(i)
        j = i - 1
        Do While j >= 1
            If StrComp(lijst(j), t, vbTextCompare) <= 0 Then Exit Do
            lijst(j + 1) = lijst(j)
            j = j - 1
        Loop
        lijst(j + 1) = t
    Next i
    bereik = UvKiesBereik(lijst, nL)
    If Len(bereik) = 0 Then Exit Sub
    alle = (bereik = "#ALLE")

    If alle Then
        If BladBestaatStil(BL_UITVRAAG) Then
            If MsgBox("Er staat al een uitvraag (blad '" & BL_UITVRAAG & "')." & vbCrLf & vbCrLf & _
                      "Een nieuwe uitvraag voor iedereen vervangt dat blad, ook wat al is aangeleverd en wat nog niet is verwerkt." & vbCrLf & vbCrLf & _
                      "Toch een nieuwe uitvraag maken?", vbYesNo + vbExclamation + vbDefaultButton2, "Uitvraag versturen") <> vbYes Then Exit Sub
        End If
        vraag = "Uitvraag kadernota/begroting maken voor " & namen.Count & " budgetbeheerders?"
    Else
        If BladBestaatStil(BL_UITVRAAG) Then
            If UvJaarWaarde() <> lj Then
                If MsgBox("Het blad '" & BL_UITVRAAG & "' hoort bij een ander jaar en wordt vervangen door een blad met alleen " & bereik & "." & vbCrLf & vbCrLf & _
                          "Doorgaan?", vbYesNo + vbExclamation + vbDefaultButton2, "Uitvraag versturen") <> vbYes Then Exit Sub
            ElseIf UvKopRij(ThisWorkbook.Worksheets(BL_UITVRAAG), bereik) > 0 Then
                If MsgBox("De regels van " & bereik & " op het blad '" & BL_UITVRAAG & "' worden vervangen, ook wat al is aangeleverd en wat nog niet is verwerkt." & vbCrLf & _
                          "De andere budgetbeheerders blijven staan." & vbCrLf & vbCrLf & _
                          "Doorgaan?", vbYesNo + vbExclamation + vbDefaultButton2, "Uitvraag versturen") <> vbYes Then Exit Sub
            End If
        End If
        vraag = "Uitvraag kadernota/begroting maken voor " & bereik & "?"
    End If
    If MsgBox(vraag & vbCrLf & "Het bestand komt in map '02 Budgetbeheerders en projectleiders'." & vbCrLf & _
              "Er wordt niets gemaild.", vbYesNo + vbQuestion, "Uitvraag versturen") <> vbYes Then Exit Sub
    map = MapStambestand("02 Budgetbeheerders en projectleiders")
    If Len(map) = 0 Then Exit Sub

    On Error GoTo Fout
    Application.ScreenUpdating = False
    For Each nm In namen.Keys
        If alle Or StrComp(CStr(nm), bereik, vbTextCompare) = 0 Then
            Set ws = UitvraagBlad(CStr(nm))
            If Not ws Is Nothing Then
                If BladOpslaan(ws, SubMap(map, CStr(nm)), "Uitvraag kadernota-begroting " & nm & " " & Format$(Date, "dd-mm-yyyy")) Then n = n + 1
            End If
        End If
    Next nm
    UvTijdelijkWeg
    Set ws = UvBladOpbouwen(IIf(alle, "", bereik))
    If ws Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Het blad Uitvraag kon niet worden gemaakt.", vbExclamation
        Exit Sub
    End If
    ThisWorkbook.Names.Add Name:="Uitvraag_Jaar", RefersTo:="=" & lj, Visible:=False
    TabbladenOrdenen
    ws.Activate
    On Error Resume Next
    If alle Then
        Application.GoTo ws.Range("A" & UV_EERSTE), True
    ElseIf UvKopRij(ws, bereik) > 0 Then
        Application.GoTo ws.Cells(UvKopRij(ws, bereik), 1), True      ' naar de zojuist gemaakte budgetbeheerder
    End If
    On Error GoTo Fout
    LogRegel "uitvraag kadernota/begroting", , , "", , , IIf(alle, n & " bestanden, alle budgetbeheerders", "bestand voor " & bereik) & ", lopend jaar " & lj
    Application.ScreenUpdating = True
    If alle Then
        MsgBox "Uitvraag gemaakt. " & n & " bestand(en) opgeslagen in:" & vbCrLf & map & vbCrLf & vbCrLf & _
               "Geef de bestanden zelf aan de budgetbeheerders. Op het blad '" & BL_UITVRAAG & "' staat de datum van versturen.", vbInformation
    Else
        ' naam van het bestand zoals BladOpslaan die maakt
        bn = "Uitvraag kadernota-begroting " & bereik & " " & Format$(Date, "dd-mm-yyyy")
        For Each t2 In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
            bn = Replace(bn, CStr(t2), "-")
        Next t2
        pad = SubMap(map, bereik) & bn & ".xlsx"
        If n > 0 And Len(Dir(pad)) > 0 Then
            If MsgBox("Uitvraag gemaakt voor " & bereik & "." & vbCrLf & vbCrLf & "Het bestand staat in:" & vbCrLf & pad & vbCrLf & vbCrLf & _
                      "Wil je het uitvraagoverzicht van " & bereik & " nu openen om te controleren?" & vbCrLf & _
                      "Sluit het daarna zonder op te slaan, tenzij je zelf iets wilt aanpassen.", vbYesNo + vbQuestion, "Uitvraag versturen") = vbYes Then
                Workbooks.Open Filename:=pad
            End If
        Else
            MsgBox "Uitvraag gemaakt voor " & bereik & ". Het bestand staat in:" & vbCrLf & map, vbInformation
        End If
    End If
    Exit Sub
Fout:
    foutTekst = Err.Description
    On Error Resume Next
    UvTijdelijkWeg
    Application.ScreenUpdating = True
    MsgBox "Fout bij het maken van de uitvraag: " & foutTekst, vbCritical
End Sub

' ----------------------------------------------------------------------
' 2. Aanleveringen ophalen
' ----------------------------------------------------------------------
Public Sub UitvraagOphalen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, wb As Workbook, r As Long, i As Long, k As Long, n As Long, lr As Long, kop As Long, eind As Long
    Dim kNaam() As String, kDt() As Date, kBest() As String, kRij() As Long, nK As Long
    Dim naam As String, bestand As String, dt As Date, verstuurd As Variant, opgehaald As Variant, nieuw As Boolean
    Dim lijst As String, antw As String, sel As Long, vn As String, sleutel As String, d As Variant
    Dim secRij As Object, teVoegen As Collection, itm As Variant, laatsteVz As Long, rN As Long, foutTekst As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur haalt aanleveringen op.", vbExclamation
        Exit Sub
    End If
    If Not BladBestaatStil(BL_UITVRAAG) Then
        MsgBox "Er is nog geen uitvraag. Gebruik eerst '1. Uitvraag versturen'.", vbInformation
        Exit Sub
    End If
    If Not UvJaarKlopt() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_UITVRAAG)
    ws.Unprotect WACHTWOORD
    lr = UvLaatsteRij(ws)
    ' welke budgetbeheerders hebben iets nieuws aangeleverd?
    For r = UV_EERSTE To lr
        If CStr(ws.Cells(r, 30).Value) = "KOP" Then
            naam = Trim$(CStr(ws.Cells(r, 29).Value))
            dt = 0
            bestand = UvBestand(naam, dt)
            verstuurd = ws.Cells(r, 31).Value
            opgehaald = ws.Cells(r, 33).Value
            nieuw = False
            If Len(bestand) > 0 And IsDate(verstuurd) Then
                If dt > CDate(verstuurd) + TimeSerial(0, 5, 0) Then
                    ws.Cells(r, 32).Value = dt                 ' ingeleverd: het bestand is na het versturen opgeslagen
                    If Not IsDate(opgehaald) Then
                        nieuw = True
                    ElseIf dt > CDate(opgehaald) + TimeSerial(0, 0, 5) Then
                        nieuw = True
                    End If
                End If
            End If
            UvKopTekst ws, r
            If nieuw Then
                nK = nK + 1
                ReDim Preserve kNaam(1 To nK): ReDim Preserve kDt(1 To nK): ReDim Preserve kBest(1 To nK): ReDim Preserve kRij(1 To nK)
                kNaam(nK) = naam: kDt(nK) = dt: kBest(nK) = bestand: kRij(nK) = r
            End If
        End If
    Next r
    If nK = 0 Then
        BladBeveiligen ws
        MsgBox "Er zijn geen nieuwe of gewijzigde aanleveringen." & vbCrLf & _
               "Een aanlevering is nieuw als het bestand van de budgetbeheerder na het versturen is opgeslagen en sinds het vorige ophalen is gewijzigd.", vbInformation, "Aanleveringen ophalen"
        Exit Sub
    End If
    For i = 1 To nK
        If Len(lijst) < 800 Then
            lijst = lijst & i & " = " & kNaam(i) & " (opgeslagen " & Format$(kDt(i), "dd-mm-yyyy hh:mm") & _
                    IIf(IsDate(ws.Cells(kRij(i), 33).Value), ", eerder opgehaald", ", nog niet opgehaald") & ")" & vbCrLf
        End If
    Next i
    antw = Trim$(InputBox("Nieuwe of gewijzigde aanleveringen:" & vbCrLf & vbCrLf & lijst & vbCrLf & "Typ het nummer van de budgetbeheerder:", _
                          "Aanleveringen ophalen", "1"))
    If Len(antw) = 0 Or Not IsNumeric(antw) Then
        BladBeveiligen ws
        Exit Sub
    End If
    sel = CLng(antw)
    If sel < 1 Or sel > nK Then
        BladBeveiligen ws
        MsgBox "Dat nummer staat niet in de lijst.", vbExclamation
        Exit Sub
    End If
    naam = kNaam(sel): bestand = kBest(sel): dt = kDt(sel)

    On Error GoTo Fout
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(Filename:=bestand, ReadOnly:=True, UpdateLinks:=0)
    d = wb.Worksheets(1).Range("A1:T" & wb.Worksheets(1).UsedRange.Row + wb.Worksheets(1).UsedRange.Rows.Count).Value
    wb.Close SaveChanges:=False
    Set wb = Nothing
    ThisWorkbook.Activate
    If CStr(d(5, 15)) <> "Volgnr" Or Not (CStr(d(5, 13)) Like "*utomatisch*") Then
        Application.ScreenUpdating = True
        BladBeveiligen ws
        MsgBox "Dit uitvraagbestand heeft een oude opzet en kan niet worden ingelezen. Maak de uitvraag opnieuw.", vbExclamation
        Exit Sub
    End If
    mD = d
    Set mWv = Vz()
    laatsteVz = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    Set mVolgRij = CreateObject("Scripting.Dictionary")
    For r = 2 To laatsteVz
        If Not Leeg(mWv.Cells(r, 1).Value) Then mVolgRij(CStr(mWv.Cells(r, 1).Value)) = r
    Next r
    mNWijz = 0: mNSig = 0: mNNieuw = 0

    kop = UvKopRij(ws, naam)
    eind = UvSectieEinde(ws, kop)
    Set secRij = CreateObject("Scripting.Dictionary")
    For r = kop + 1 To eind
        If CStr(ws.Cells(r, 30).Value) = "REG" Then
            vn = CStr(ws.Cells(r, 1).Value)
            If vn = "NIEUW" Then
                secRij("NIEUW|" & LCase$(Trim$(CStr(ws.Cells(r, 3).Value)))) = r
            Else
                secRij(vn) = r
            End If
        End If
    Next r
    Set teVoegen = New Collection
    For i = 6 To UBound(d, 1)
        vn = Trim$(Tkst(d(i, 15)))
        sleutel = ""
        If UCase$(vn) = "NIEUW" Then
            If Len(Trim$(Tkst(d(i, 2)))) > 0 Then sleutel = "NIEUW|" & LCase$(Trim$(Tkst(d(i, 2))))
        ElseIf Len(vn) > 0 And mVolgRij.Exists(vn) Then
            sleutel = vn
        End If
        If Len(sleutel) > 0 Then
            If secRij.Exists(sleutel) Then
                UvMergeRegel ws, CLng(secRij(sleutel)), i, (Left$(sleutel, 6) = "NIEUW|"), False
            Else
                teVoegen.Add Array(i, sleutel)
            End If
        End If
    Next i
    ' regels die er nog niet stonden (nieuwe voorstellen, of kredieten die na het versturen zijn gestart) onder in de sectie
    If teVoegen.Count > 0 Then
        ws.Rows((eind + 1) & ":" & (eind + teVoegen.Count)).Insert
        rN = eind
        For Each itm In teVoegen
            rN = rN + 1
            ws.Range(ws.Cells(rN, 1), ws.Cells(rN, 34)).ClearContents
            ws.Cells(rN, 29).Value = naam
            ws.Cells(rN, 30).Value = "REG"
            If Left$(CStr(itm(1)), 6) = "NIEUW|" Then
                ws.Cells(rN, 1).Value = "NIEUW"
                ws.Cells(rN, 3).Value = Tkst(mD(CLng(itm(0)), 2))
            Else
                r = CLng(mVolgRij(CStr(itm(1))))
                ws.Cells(rN, 1).Value = mWv.Cells(r, 1).Value
                ws.Cells(rN, 2).NumberFormat = "@"
                ws.Cells(rN, 2).Value = NrTekst(mWv.Cells(r, 3).Value)
                ws.Cells(rN, 3).Value = mWv.Cells(r, 4).Value
            End If
            UvRegelOpmaak ws, rN
            UvMergeRegel ws, rN, CLng(itm(0)), (Left$(CStr(itm(1)), 6) = "NIEUW|"), True
        Next itm
    End If
    kop = UvKopRij(ws, naam)
    ws.Cells(kop, 32).Value = dt
    ws.Cells(kop, 33).Value = Now
    ws.Cells(kop, 34).ClearContents
    UvKopTekst ws, kop
    BladOpmaken ws, 5, UV_EERSTE, UvLaatsteRij(ws)
    UvRijenPassend ws
    BladBeveiligen ws
    ws.Activate
    LogRegel "uitvraag opgehaald", , , "", , , naam & " | " & Dir(bestand) & " | " & mNWijz & " wijzigingen, " & mNNieuw & " voorstellen, " & mNSig & " signalen"
    Application.ScreenUpdating = True
    MsgBox naam & ": aanlevering opgehaald." & vbCrLf & vbCrLf & mNWijz & " regel(s) met een wijziging (vet), " & mNNieuw & " voorstel(len) voor nieuwe kredieten" & _
           IIf(mNSig > 0, ", en " & mNSig & " regel(s) met het signaal '" & UV_SIGNAAL & "'", "") & "." & vbCrLf & vbCrLf & _
           "Ongewijzigde regels staan op 'goedkeuren'. Kies bij de overige regels een besluit in kolom Q.", vbInformation, "Aanleveringen ophalen"
    Exit Sub
Fout:
    foutTekst = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Fout bij het ophalen van de aanlevering: " & foutTekst & vbCrLf & _
           "Controleer het blad '" & BL_UITVRAAG & "' voor je verdergaat.", vbCritical
End Sub

' Zet een regel uit het uitvraagbestand (rij i van mD) op regel r van het blad Uitvraag. Regels die de adviseur al heeft
' aangepast of beoordeeld worden niet overschreven; wijkt de nieuwe aanlevering af, dan komt er een signaal.
Private Sub UvMergeRegel(ByVal ws As Worksheet, ByVal r As Long, ByVal i As Long, ByVal isNieuw As Boolean, ByVal nieuweRij As Boolean)
    Dim k As Long, vr As Long, w(1 To 5) As Double, oms As String, toel As String, dTekst As String, wijzTekst As String
    Dim wijzig As Boolean, gewijzigd As Boolean, eerder As Boolean, beoordeeld As Boolean, aangGewijzigd As Boolean, besl As String
    Dim vn As String
    For k = 1 To 5
        w(k) = Nul(mD(i, 6 + k))
    Next k
    toel = Trim$(Tkst(mD(i, 14)))
    If isNieuw Then
        dTekst = "taakveld: " & Tkst(mD(i, 4))
        wijzTekst = "Voorstel nieuw krediet (bestaand beleid)"
        gewijzigd = True
    Else
        vn = Trim$(Tkst(mD(i, 15)))
        vr = CLng(mVolgRij(vn))
        oms = Trim$(Tkst(mD(i, 12)))
        dTekst = oms
        wijzTekst = Tkst(mD(i, 13))
        For k = 1 To 5
            If Round(w(k) - Nul(mWv.Cells(vr, 30 + k).Value), 2) <> 0 Then wijzig = True
        Next k
        gewijzigd = (wijzig Or Len(oms) > 0 Or Len(toel) > 0)
    End If
    If Not nieuweRij Then eerder = (Len(CStr(ws.Cells(r, 21).Value)) > 0)
    If eerder Then
        For k = 1 To 5
            If Round(w(k) - Nul(ws.Cells(r, 21 + k).Value), 2) <> 0 Then aangGewijzigd = True
        Next k
        If dTekst <> CStr(ws.Cells(r, 27).Value) Or toel <> CStr(ws.Cells(r, 28).Value) Then aangGewijzigd = True
        beoordeeld = (Len(CStr(ws.Cells(r, 19).Value)) > 0)
        If CStr(ws.Cells(r, 17).Value) <> CStr(ws.Cells(r, 21).Value) Then beoordeeld = True
        For k = 1 To 5
            If Round(Nul(ws.Cells(r, 9 + k).Value) - Nul(ws.Cells(r, 21 + k).Value), 2) <> 0 Then beoordeeld = True
        Next k
        If Trim$(CStr(ws.Cells(r, 4).Value)) <> CStr(ws.Cells(r, 27).Value) Then beoordeeld = True
        If Trim$(CStr(ws.Cells(r, 16).Value)) <> CStr(ws.Cells(r, 28).Value) Then beoordeeld = True
    End If
    If eerder And beoordeeld Then
        ' niet overschrijven
        If aangGewijzigd Then
            ws.Cells(r, 20).Value = UV_SIGNAAL & " (" & Format$(Date, "dd-mm-yyyy") & ")"
            For k = 1 To 5
                ws.Cells(r, 21 + k).Value = w(k)
            Next k
            ws.Cells(r, 27).Value = dTekst
            ws.Cells(r, 28).Value = toel
            mNSig = mNSig + 1
        End If
        If Not isNieuw And Len(CStr(ws.Cells(r, 19).Value)) = 0 Then
            For k = 1 To 5
                ws.Cells(r, 4 + k).Value = Nul(mWv.Cells(vr, 30 + k).Value)
            Next k
        End If
    Else
        ws.Cells(r, 4).Value = dTekst
        For k = 1 To 5
            ws.Cells(r, 9 + k).Value = w(k)
            ws.Cells(r, 21 + k).Value = w(k)
            If isNieuw Then
                ws.Cells(r, 4 + k).Value = 0
            Else
                ws.Cells(r, 4 + k).Value = Nul(mWv.Cells(vr, 30 + k).Value)
                ' budget in de Verzamellijst sinds de uitvraag gewijzigd: oranje
                If Round(Nul(mD(i, 15 + k)) - Nul(mWv.Cells(vr, 30 + k).Value), 2) <> 0 Then
                    ws.Cells(r, 4 + k).Interior.Color = RGB(255, 200, 120)
                Else
                    ws.Cells(r, 4 + k).Interior.ColorIndex = xlNone
                End If
            End If
        Next k
        ws.Cells(r, 16).Value = toel
        ws.Cells(r, 27).Value = dTekst
        ws.Cells(r, 28).Value = toel
        ws.Cells(r, 20).ClearContents
        If gewijzigd Then
            ws.Cells(r, 15).Value = wijzTekst
            ws.Cells(r, 15).Font.Color = RGB(0, 0, 0)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
            besl = "nog niet"
            mNWijz = mNWijz + 1
            If isNieuw Then mNNieuw = mNNieuw + 1: mNWijz = mNWijz - 1
        Else
            ws.Cells(r, 15).Value = GEEN_WIJZ
            ws.Cells(r, 15).Font.Color = RGB(150, 150, 150)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = False
            besl = "goedkeuren"
        End If
        ws.Cells(r, 17).Value = besl
        ws.Cells(r, 21).Value = besl
    End If
End Sub

' ----------------------------------------------------------------------
' 3. Uitvraag verwerken
' ----------------------------------------------------------------------
Public Sub UitvraagVerwerken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const BL_VOORSTEL As String = "Voorstellen nieuwe kredieten"
    Dim ws As Worksheet, wv As Worksheet, wsVs As Worksheet, wr As Worksheet, naam As String, lj As Long, laatste As Long, r As Long, i As Long, k As Long
    Dim lr As Long, besluit As String, volgRij As Object, voor As Variant, na As Variant, verschil As Double, oudB As Double, nieuwB As Double
    Dim nGoed As Long, nGoedZonder As Long, nAf As Long, nOpen As Long, toel As String, gebr As String, vsRij As Long, nB As Long
    Dim bkD As Variant, bkF As Variant, bkAE As Variant, bkAP As Variant, bkStat As Variant
    Dim vsBestond As Boolean, vsVoor As Long, mutBestond As Boolean, mutVoor As Long
    Dim geschreven As Boolean, klaar As Boolean, foutTekst As String, wijzigVz As Boolean, nTer As Long, bestandTxt As String, kopR As Long
    Dim nVerwerkt As Long, nAfgewezen As Long, kopPerNaam As Object, itm As Variant, open1 As Long

    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur verwerkt de uitvraag.", vbExclamation
        Exit Sub
    End If
    If Not BladBestaatStil(BL_UITVRAAG) Then
        MsgBox "Er is nog geen uitvraag. Gebruik eerst '1. Uitvraag versturen'.", vbInformation
        Exit Sub
    End If
    If Not UvJaarKlopt() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_UITVRAAG)
    lr = UvLaatsteRij(ws)
    For i = UV_EERSTE To lr
        If CStr(ws.Cells(i, 30).Value) = "REG" And Len(CStr(ws.Cells(i, 19).Value)) = 0 Then
            Select Case CStr(ws.Cells(i, 17).Value)
                Case "goedkeuren"
                    nGoed = nGoed + 1
                    If CStr(ws.Cells(i, 15).Value) = GEEN_WIJZ Then nGoedZonder = nGoedZonder + 1
                Case "afwijzen": nAf = nAf + 1
                Case "nog niet": nOpen = nOpen + 1
            End Select
        End If
    Next i
    If nGoed + nAf = 0 Then
        MsgBox "Er is geen regel met besluit 'goedkeuren' of 'afwijzen' die nog verwerkt moet worden." & vbCrLf & _
               "Haal eerst een aanlevering op met '2. Aanleveringen ophalen' en kies per regel een besluit.", vbInformation, "Uitvraag verwerken"
        Exit Sub
    End If
    If MsgBox("Alle budgetbeheerders samen:" & vbCrLf & nGoed & " goedkeuren (waarvan " & nGoedZonder & " zonder wijziging), " & nAf & " afwijzen, " & nOpen & " nog niet." & vbCrLf & vbCrLf & _
              "Goedgekeurde regels nu verwerken in de Verzamellijst?" & vbCrLf & _
              "Een gewijzigd krediet krijgt direct de status 'Te controleren door financi" & ChrW(235) & "le administratie'.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Uitvraag verwerken") <> vbYes Then Exit Sub

    Set wv = Vz()
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    Set volgRij = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(wv.Cells(r, 1).Value) Then volgRij(CStr(wv.Cells(r, 1).Value)) = r
    Next r
    Set kopPerNaam = CreateObject("Scripting.Dictionary")
    For i = UV_EERSTE To lr
        If CStr(ws.Cells(i, 30).Value) = "KOP" Then kopPerNaam(Trim$(CStr(ws.Cells(i, 29).Value))) = i
    Next i

    On Error GoTo Fout
    Application.ScreenUpdating = False
    ws.Unprotect WACHTWOORD
    ' reservekopie van alles wat deze macro kan wijzigen, om bij een fout alles terug te zetten
    bkD = wv.Range("D2:D" & laatste).Value
    bkF = wv.Range("F2:F" & laatste).Value
    bkAE = wv.Range("AE2:AJ" & laatste).Value
    bkAP = wv.Range("AP2:AW" & laatste).Value
    bkStat = ws.Range(ws.Cells(UV_EERSTE, 19), ws.Cells(lr, 19)).Value
    vsBestond = BladBestaatStil(BL_VOORSTEL)
    If vsBestond Then vsVoor = ThisWorkbook.Worksheets(BL_VOORSTEL).Cells(ThisWorkbook.Worksheets(BL_VOORSTEL).Rows.Count, 3).End(xlUp).Row
    mutBestond = BladBestaatStil(BL_MUTREG)
    If mutBestond Then mutVoor = ThisWorkbook.Worksheets(BL_MUTREG).Cells(ThisWorkbook.Worksheets(BL_MUTREG).Rows.Count, 1).End(xlUp).Row
    voor = MeetStap8(wv, laatste)
    geschreven = True

    For i = UV_EERSTE To lr
        If CStr(ws.Cells(i, 30).Value) = "REG" And Len(CStr(ws.Cells(i, 19).Value)) = 0 Then
            naam = Trim$(CStr(ws.Cells(i, 29).Value))
            bestandTxt = "Uitvraag " & naam
            If kopPerNaam.Exists(naam) Then
                kopR = CLng(kopPerNaam(naam))
                If IsDate(ws.Cells(kopR, 33).Value) Then bestandTxt = "Uitvraag " & naam & " (opgehaald " & Format$(CDate(ws.Cells(kopR, 33).Value), "dd-mm-yyyy hh:mm") & ")"
            End If
            besluit = CStr(ws.Cells(i, 17).Value)
            If besluit = "afwijzen" Then
                ws.Cells(i, 19).Value = "afgewezen " & Format$(Date, "dd-mm-yyyy")
                LogRegel "uitvraag afgewezen", ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, "", , , naam & " | " & CStr(ws.Cells(i, 18).Value)
                nAfgewezen = nAfgewezen + 1
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
                                               Nul(ws.Cells(i, UvKol("wordt") + k - 1).Value), "Nieuw krediet (voorstel)", CStr(ws.Cells(i, UvKol("onderbouwing")).Value), bestandTxt
                        End If
                    Next k
                    wsVs.Cells(vsRij, 11).Value = ws.Cells(i, UvKol("onderbouwing")).Value
                    wsVs.Cells(vsRij, 12).Value = bestandTxt
                    wsVs.Cells(vsRij, 13).Value = "nee"
                    ws.Cells(i, 19).Value = "naar Voorstellen " & Format$(Date, "dd-mm-yyyy")
                    LogRegel "uitvraag voorstel nieuw krediet", "", "", "", , , naam & " | " & CStr(ws.Cells(i, 3).Value)
                    nVerwerkt = nVerwerkt + 1
                ElseIf volgRij.Exists(CStr(ws.Cells(i, 1).Value)) Then
                    r = volgRij(CStr(ws.Cells(i, 1).Value))
                    nB = 0
                    wijzigVz = False
                    For k = 1 To 5
                        oudB = Nul(wv.Cells(r, 30 + k).Value)
                        nieuwB = Nul(ws.Cells(i, UvKol("wordt") + k - 1).Value)
                        If Round(nieuwB - oudB, 2) <> 0 Then
                            nB = nB + 1
                            LogRegel "uitvraag verwerkt", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, wv.Cells(1, 30 + k).Text, oudB, nieuwB, naam
                            MutatieRegistreren naam, wv.Cells(r, 1).Value, NrTekst(wv.Cells(r, 3).Value), CStr(wv.Cells(r, 4).Value), _
                                               wv.Cells(r, 19).Value, wv.Cells(r, 20).Value, wv.Cells(r, 25).Value, lj + k, oudB, nieuwB, _
                                               CStr(ws.Cells(i, UvKol("wijziging")).Value), CStr(ws.Cells(i, UvKol("onderbouwing")).Value), bestandTxt
                            If nieuwB = 0 Then wv.Cells(r, 30 + k).ClearContents Else wv.Cells(r, 30 + k).Value = nieuwB
                            verschil = verschil + nieuwB - oudB
                            wijzigVz = True
                        End If
                    Next k
                    If Len(Trim$(CStr(ws.Cells(i, 4).Value))) > 0 Then
                        LogRegel "uitvraag verwerkt", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Omschrijving", wv.Cells(r, 4).Value, ws.Cells(i, 4).Value, naam
                        wv.Cells(r, 4).Value = Left$(Trim$(CStr(ws.Cells(i, 4).Value)), 36)
                        wv.Cells(r, 6).Value = KortOmschrijving(CStr(wv.Cells(r, 4).Value))
                        wijzigVz = True
                    End If
                    toel = Trim$(CStr(ws.Cells(i, UvKol("wijziging")).Value))
                    If toel = GEEN_WIJZ Then toel = ""
                    If Len(Trim$(CStr(ws.Cells(i, UvKol("onderbouwing")).Value))) > 0 Then toel = toel & IIf(Len(toel) > 0, ". ", "") & Trim$(CStr(ws.Cells(i, UvKol("onderbouwing")).Value))
                    If Len(toel) > 0 Then
                        toel = Left$(toel, MAX_TOEL)
                        LogRegel "uitvraag verwerkt", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Toelichting", wv.Cells(r, 36).Value, toel, naam
                        wv.Cells(r, 36).Value = toel
                        wijzigVz = True
                    End If
                    If Not wijzigVz Then
                        ws.Cells(i, 19).Value = "akkoord, geen wijziging"
                    Else
                        wv.Cells(r, 48).Value = gebr
                        wv.Cells(r, 49).Value = Now
                        ' direct ter controle bij de financiele administratie
                        wv.Cells(r, 42).Value = CtlTeControleren()
                        wv.Cells(r, 43).ClearContents
                        wv.Cells(r, 44).ClearContents
                        wv.Cells(r, 45).Value = "uitvraag kadernota/begroting"
                        LogRegel "ter controle na uitvraag", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "", , , naam
                        nTer = nTer + 1
                        ws.Cells(i, 19).Value = "verwerkt " & Format$(Date, "dd-mm-yyyy")
                    End If
                    nVerwerkt = nVerwerkt + 1
                Else
                    ws.Cells(i, 19).Value = "volgnr niet gevonden"
                End If
            End If
        End If
    Next i
    Application.Calculate
    na = MeetStap8(wv, laatste)
    If na(0) <> voor(0) Or Round(na(1) - voor(1), 2) <> 0 Or Round(na(2) - voor(2), 2) <> 0 Or Round(na(3) - voor(3), 2) <> 0 _
       Or Round(na(4) - voor(4), 2) <> 0 Or Round((na(5) + na(6) + na(7) + na(8) + na(9)) - (voor(5) + voor(6) + voor(7) + voor(8) + voor(9)) - verschil, 2) <> 0 Then
        UvTerugzetten wv, laatste, bkD, bkF, bkAE, bkAP, ws, lr, bkStat, vsBestond, vsVoor, mutBestond, mutVoor
        LogRegel "uitvraag verwerken teruggedraaid", , , "", , , "nameting wijkt af"
        Application.ScreenUpdating = True
        BladBeveiligen ws
        MsgBox "De nameting wijkt af. Alles is teruggezet. Er is niets verwerkt.", vbCritical, "Uitvraag verwerken"
        Exit Sub
    End If
    klaar = True
    ' kopregels bijwerken: verwerkt als er geen open regels meer zijn
    For Each itm In kopPerNaam.Keys
        kopR = CLng(kopPerNaam(itm))
        open1 = 0
        For r = kopR + 1 To UvSectieEinde(ws, kopR)
            If CStr(ws.Cells(r, 30).Value) = "REG" Then
                If Len(CStr(ws.Cells(r, 17).Value)) > 0 And Len(CStr(ws.Cells(r, 19).Value)) = 0 And CStr(ws.Cells(r, 15).Value) <> GEEN_WIJZ Then open1 = open1 + 1
            End If
        Next r
        If open1 = 0 And IsDate(ws.Cells(kopR, 33).Value) Then ws.Cells(kopR, 34).Value = Now
        UvKopTekst ws, kopR
    Next itm
    BladBeveiligen ws
    ws.Activate
    Application.ScreenUpdating = True
    MsgBox "Verwerkt." & vbCrLf & nVerwerkt & " regel(s) goedgekeurd en verwerkt, " & nAfgewezen & " afgewezen." & vbCrLf & _
           nTer & " krediet(en) staan nu ter controle bij de financi" & ChrW(235) & "le administratie." & vbCrLf & _
           "Regels met 'nog niet' blijven open op het blad '" & BL_UITVRAAG & "'." & vbCrLf & vbCrLf & _
           "Nieuwe kredieten staan op het blad 'Voorstellen nieuwe kredieten': voer die in via de invoerkaart.", vbInformation, "Uitvraag verwerken"
    Exit Sub
Fout:
    foutTekst = Err.Description
    On Error Resume Next
    If geschreven And Not klaar Then
        UvTerugzetten wv, laatste, bkD, bkF, bkAE, bkAP, ws, lr, bkStat, vsBestond, vsVoor, mutBestond, mutVoor
        LogRegel "uitvraag verwerken teruggedraaid", , , "", , , "fout: " & foutTekst
    End If
    Application.ScreenUpdating = True
    BladBeveiligen ws
    If klaar Then
        MsgBox "De uitvraag is verwerkt, maar daarna ging iets mis: " & foutTekst, vbExclamation, "Uitvraag verwerken"
    Else
        MsgBox "Fout: " & foutTekst & vbCrLf & IIf(geschreven, "Alles is teruggezet. ", "") & "Er is niets verwerkt.", vbCritical, "Uitvraag verwerken"
    End If
End Sub

' Zet de Verzamellijst, de status op het blad Uitvraag en de nieuwe registerregels terug zoals ze waren.
Private Sub UvTerugzetten(ByVal wv As Worksheet, ByVal laatste As Long, ByVal bkD As Variant, ByVal bkF As Variant, ByVal bkAE As Variant, _
                          ByVal bkAP As Variant, ByVal ws As Worksheet, ByVal lr As Long, ByVal bkStat As Variant, _
                          ByVal vsBestond As Boolean, ByVal vsVoor As Long, ByVal mutBestond As Boolean, ByVal mutVoor As Long)
    Const BL_VOORSTEL As String = "Voorstellen nieuwe kredieten"
    Dim wr As Worksheet
    On Error Resume Next
    wv.Range("D2:D" & laatste).Value = bkD
    wv.Range("F2:F" & laatste).Value = bkF
    wv.Range("AE2:AJ" & laatste).Value = bkAE
    wv.Range("AP2:AW" & laatste).Value = bkAP
    ws.Range(ws.Cells(UV_EERSTE, 19), ws.Cells(lr, 19)).Value = bkStat
    If BladBestaatStil(BL_VOORSTEL) Then
        If Not vsBestond Then
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets(BL_VOORSTEL).Unprotect WACHTWOORD
            ThisWorkbook.Worksheets(BL_VOORSTEL).Delete
            Application.DisplayAlerts = True
        Else
            Set wr = ThisWorkbook.Worksheets(BL_VOORSTEL)
            If wr.Cells(wr.Rows.Count, 3).End(xlUp).Row > vsVoor Then wr.Rows((vsVoor + 1) & ":" & wr.Cells(wr.Rows.Count, 3).End(xlUp).Row).Delete
        End If
    End If
    If BladBestaatStil(BL_MUTREG) Then
        If Not mutBestond Then
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets(BL_MUTREG).Unprotect WACHTWOORD
            ThisWorkbook.Worksheets(BL_MUTREG).Delete
            Application.DisplayAlerts = True
        Else
            Set wr = ThisWorkbook.Worksheets(BL_MUTREG)
            If wr.Cells(wr.Rows.Count, 1).End(xlUp).Row > mutVoor Then wr.Rows((mutVoor + 1) & ":" & wr.Cells(wr.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End If
    Application.Calculate
End Sub

