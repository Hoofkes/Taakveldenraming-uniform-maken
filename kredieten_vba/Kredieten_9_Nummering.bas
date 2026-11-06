Attribute VB_Name = "Kredieten_9_Nummering"
Option Explicit

' Kredietnummers (kolom C) per balansrubriek. Elke rubriek heeft een eigen reeks:
'   72 = economisch nut (EN), 73 = maatschappelijk nut (MN), dan 0 en het cijfer van de rubriek (A129 = 7).
'   A124EN en A124ENH delen de reeks 7204.
' Het eerstvolgende vrije nummer per reeks wordt bewaard in een verborgen naam (KrNr_7201 enzovoort). De beginstand
' komt uit de lijst van de financiele administratie (november 2026); daarna houdt de macro de nummers zelf bij.
' Een nummer dat al in de Verzamellijst staat, wordt altijd overgeslagen.
' Balanscodes die nog niet in de Verzamellijst voorkomen (A122EN, A125MN, A126MN, A129MN en nieuwe codes in de
' balanscodetabel) hebben geen beginstand: de financiele administratie vult daar de eerste keer zelf een nummer in
' (in kolom C van de Verzamellijst of met de knop Kredietnummer aanpassen). Daarna werkt het net als bij de andere.

Private Function Reeksen() As Variant
    ' reeks, balanscode(s), beginstand (eerstvolgende vrije nummer volgens de lijst van de financiele administratie;
    ' 0 = komt nog niet voor in de Verzamellijst, eerste nummer Invullen)
    Reeksen = Array(Array("7201", "A121EN", 7201026), Array("7202", "A122EN", 0), Array("7203", "A123EN", 7203115), _
                    Array("7204", "A124EN / A124ENH", 7204188), Array("7205", "A125EN", 7205060), Array("7206", "A126EN", 7206078), _
                    Array("7207", "A129EN", 7207038), Array("7304", "A124MN1", 7304180), Array("7305", "A125MN", 0), _
                    Array("7306", "A126MN", 0), Array("7307", "A129MN", 0))
End Function

' De vaste reeksen plus de balanscodes uit de balanscodetabel (Parameters) die nog geen reeks in de lijst hebben.
Private Function AlleReeksen() As Variant
    Dim c As New Collection, x As Variant, lo As ListObject, cel As Range, bc As String, rk As String, i As Long, gevonden As Boolean
    Dim uit() As Variant
    For Each x In Reeksen()
        c.Add Array(CStr(x(0)), CStr(x(1)), x(2))
    Next x
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects("Tbl_Balanscode")
    On Error GoTo 0
    If Not lo Is Nothing Then
        For Each cel In lo.ListColumns(1).DataBodyRange.Cells
            bc = UCase$(Trim$(CStr(cel.Value)))
            rk = NummerReeks(bc)
            If Len(rk) > 0 Then
                gevonden = False
                For i = 1 To c.Count
                    If c(i)(0) = rk Then
                        gevonden = True
                        If InStr(" / " & c(i)(1) & " / ", " / " & bc & " / ") = 0 Then
                            x = c(i): x(1) = x(1) & " / " & bc
                            c.Remove i
                            If i > c.Count Then c.Add x Else c.Add x, Before:=i
                        End If
                        Exit For
                    End If
                Next i
                If Not gevonden Then c.Add Array(rk, bc, 0)
            End If
        Next cel
    End If
    ReDim uit(0 To c.Count - 1)
    For i = 1 To c.Count
        uit(i - 1) = c(i)
    Next i
    AlleReeksen = uit
End Function

Private Function Beginstand(ByVal reeks As String) As Double
    Dim x As Variant
    For Each x In Reeksen()
        If x(0) = reeks Then Beginstand = x(2)
    Next x
End Function

' Naam van een balanscode die niet in de balanscodetabel (Parameters) staat.
Private Function VasteNaam(ByVal bc As String) As String
    Select Case bc
        Case "A125MN": VasteNaam = "Vervoermiddelen (maatschappelijk nut)"
        Case "A126MN": VasteNaam = "Machines, apparaten en installaties (maatschappelijk nut)"
        Case "A129MN": VasteNaam = "Overige materi" & ChrW(235) & "le vaste activa (maatschappelijk nut)"
    End Select
End Function

' Reeks bij een balanscode, bijvoorbeeld A121EN -> 7201, A124ENH -> 7204, A124MN1 -> 7304. Leeg als er geen reeks is.
Public Function NummerReeks(ByVal balanscode As String) As String
    Dim bc As String, cijfer As String
    bc = UCase$(Trim$(VoorCode(balanscode)))
    If Not (bc Like "A12#EN*" Or bc Like "A12#MN*") Then Exit Function
    cijfer = Mid$(bc, 4, 1)
    If cijfer = "9" Then cijfer = "7"
    If cijfer < "1" Or cijfer > "7" Then Exit Function
    NummerReeks = IIf(Mid$(bc, 5, 2) = "EN", "72", "73") & "0" & cijfer
End Function

' Bewaard eerstvolgende nummer van een reeks (zet de beginstand als die er nog niet is). 0 = nog geen nummer:
' de financiele administratie vult het eerste nummer zelf in.
Private Function Opgeslagen(ByVal reeks As String) As Double
    On Error Resume Next
    Opgeslagen = CDbl(Evaluate(ThisWorkbook.Names("KrNr_" & reeks).RefersTo))
    On Error GoTo 0
    If Opgeslagen > 0 Then Exit Function
    Opgeslagen = Beginstand(reeks)
    If Opgeslagen > 0 Then Bewaar reeks, Opgeslagen
End Function

Private Sub Bewaar(ByVal reeks As String, ByVal nr As Double)
    ThisWorkbook.Names.Add Name:="KrNr_" & reeks, RefersTo:="=" & Format$(nr, "0"), Visible:=False
End Sub

' Alle kredietnummers in de Verzamellijst (kolom C), als tekst.
Private Function NummersInLijst() As Object
    Dim d As Object, v As Variant, i As Long, laatste As Long
    Set d = CreateObject("Scripting.Dictionary")
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    If laatste < 2 Then Set NummersInLijst = d: Exit Function
    v = Vz().Range("C2:C" & laatste).Value
    For i = 1 To UBound(v, 1)
        If Not IsError(v(i, 1)) Then d(NrTekst(v(i, 1))) = True
    Next i
    Set NummersInLijst = d
End Function

' Zet per reeks het bewaarde nummer voorbij nummers die al in de Verzamellijst staan (bijvoorbeeld door de
' financiele administratie met de hand ingevuld). Draait bij openen, na toevoegen en bij Controles vernieuwen.
Public Sub KredietnummersBijwerken()
    Dim x As Variant, d As Object, nr As Double, oud As Double, hoogste As Double, rk As Variant
    On Error Resume Next
    Set d = NummersInLijst()
    ' eenmalig (versie 2): beginstanden van reeksen die nog niet in de Verzamellijst voorkomen, weer weghalen
    If Len(ThisWorkbook.Names("KrNr_Versie").Name) = 0 Then
        For Each rk In Array("7202", "7305", "7306", "7307")
            hoogste = HoogsteInLijst(d, CStr(rk))
            If hoogste = 0 Then ThisWorkbook.Names("KrNr_" & rk).Delete
        Next rk
        ThisWorkbook.Names.Add Name:="KrNr_Versie", RefersTo:="=2", Visible:=False
    End If
    For Each x In AlleReeksen()
        oud = Opgeslagen(CStr(x(0)))
        nr = oud
        If nr = 0 Then
            ' nog geen nummer: heeft de financiele administratie het eerste nummer in kolom C ingevuld?
            hoogste = HoogsteInLijst(d, CStr(x(0)))
            If hoogste > 0 Then nr = hoogste + 1
        End If
        If nr > 0 Then
            Do While d.Exists(Format$(nr, "0"))
                nr = nr + 1
            Loop
        End If
        If nr <> oud Then Bewaar CStr(x(0)), nr
    Next x
    TabelNummersVullen
End Sub

' Wat in de tabel op Dashboard Adm staat: het eerstvolgende vrije nummer, of 'Invullen' als er nog geen nummer is.
Private Function NummerTekst(ByVal reeks As String) As Variant
    If Opgeslagen(reeks) > 0 Then NummerTekst = Opgeslagen(reeks) Else NummerTekst = "Invullen"
End Function

' Zet de nummers in de tabel van blok 3 op Dashboard Adm (geen formules: die gaven soms #NAAM?).
Private Sub TabelNummersVullen()
    Dim ws As Worksheet, r As Long, rk As String, beveiligd As Boolean
    On Error Resume Next
    If Not BladBestaatStil(BL_ADM) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_ADM)
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    For r = ADM_NR_KOP + 3 To ADM_NR_KOP + 22
        rk = Trim$(CStr(ws.Cells(r, 4).Value))
        If Len(rk) = 4 And IsNumeric(rk) Then
            ws.Cells(r, 5).Value = NummerTekst(rk)
            ws.Cells(r, 5).Font.Color = IIf(VarType(ws.Cells(r, 5).Value) = vbString, RGB(192, 0, 0), RGB(0, 0, 0))
        End If
    Next r
    If beveiligd Then BladBeveiligen ws
End Sub

' Nieuwe regel in de Verzamellijst: opmaak (uitlijning, lettertype, randen, notatie) exact gelijk aan voorbeeldregel 272.
Public Sub RegelOpmaakAlsVoorbeeld(ByVal wv As Worksheet, ByVal r As Long)
    Const VOORBEELD As Long = 272
    On Error Resume Next
    If r = VOORBEELD Then Exit Sub
    wv.Range("A" & VOORBEELD & ":AW" & VOORBEELD).Copy
    wv.Range("A" & r & ":AW" & r).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    wv.Rows(r).RowHeight = wv.Rows(VOORBEELD).RowHeight
End Sub

' Hoogste nummer van een reeks in de Verzamellijst (alleen nummers van 7 cijfers met een einde onder 500), anders 0.
Private Function HoogsteInLijst(ByVal d As Object, ByVal reeks As String) As Double
    Dim k As Variant
    For Each k In d.Keys
        If Len(k) = 7 And Left$(k, 4) = reeks And IsNumeric(k) Then
            If CLng(Right$(k, 3)) < 500 And CDbl(k) > HoogsteInLijst Then HoogsteInLijst = CDbl(k)
        End If
    Next k
End Function

' Eerstvolgende vrije kredietnummer voor een balanscode (leeg als de balanscode geen reeks heeft).
Public Function VolgendKredietnummer(ByVal balanscode As String) As Variant
    Dim reeks As String, nr As Double, d As Object
    VolgendKredietnummer = ""
    reeks = NummerReeks(balanscode)
    If Len(reeks) = 0 Then Exit Function
    Set d = NummersInLijst()
    nr = Opgeslagen(reeks)
    If nr = 0 Then Exit Function                                  ' eerste nummer van deze reeks: Invullen
    Do While d.Exists(Format$(nr, "0"))
        nr = nr + 1
    Loop
    VolgendKredietnummer = nr
End Function

' Na het toevoegen van een krediet: nummer gebruikt, het bewaarde nummer van die reeks schuift door.
Public Sub KredietnummerGebruikt(ByVal balanscode As String, ByVal nr As Variant)
    Dim reeks As String
    On Error Resume Next
    reeks = NummerReeks(balanscode)
    If Len(reeks) = 0 Or Not IsNumeric(nr) Or Len(CStr(nr)) = 0 Then Exit Sub
    If Left$(Format$(CDbl(nr), "0"), 4) <> reeks Then Exit Sub
    If CDbl(nr) >= Opgeslagen(reeks) Or Opgeslagen(reeks) = 0 Then
        Bewaar reeks, CDbl(nr) + 1
        LogRegel "kredietnummer gebruikt", , nr, "Reeks " & reeks, Format$(CDbl(nr), "0"), Format$(CDbl(nr) + 1, "0"), "eerstvolgende vrije nummer"
    End If
    KredietnummersBijwerken
End Sub

' Invoerkaart: na het kiezen van de balanscode het eerstvolgende vrije kredietnummer invullen (alleen bij een nieuw
' krediet en als het nummer leeg is of nog een eerder voorgesteld nummer bevat).
Public Sub KaartKredietnummerVoorstellen(ByVal wd As Worksheet)
    Dim celNr As Range, bc As String, voorstel As Variant, huidig As String, x As Variant, wasVoorstel As Boolean
    On Error GoTo klaar
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then Exit Sub         ' bestaand krediet: nummer niet aanraken
    Set celNr = wd.Cells(KaartRij(2), KaartKol(2))
    bc = CStr(wd.Cells(KaartRij(17), KaartKol(17)).Value)
    voorstel = VolgendKredietnummer(bc)
    If Len(CStr(voorstel)) = 0 Then Exit Sub
    huidig = Trim$(CStr(celNr.Value))
    If Len(huidig) > 0 Then
        For Each x In AlleReeksen()                                              ' staat er een eerder voorstel (andere rubriek)?
            If Opgeslagen(CStr(x(0))) > 0 And huidig = Format$(Opgeslagen(CStr(x(0))), "0") Then wasVoorstel = True
        Next x
        If huidig = "7999998" Then wasVoorstel = False
        If Not wasVoorstel Then Exit Sub                                         ' zelf ingevuld nummer: laten staan
    End If
    Application.EnableEvents = False
    celNr.Value = voorstel
klaar:
    Application.EnableEvents = True
End Sub

' Dashboard Adm, blok 3: eerstvolgende vrije kredietnummer per balansrubriek, met een knop om een nummer aan te passen.
Public Sub KredietnummersBlokPlaatsen()
    Dim ws As Worksheet, x As Variant, r As Long, b As Object, oms As String
    If Not BladBestaatStil(BL_ADM) Then Exit Sub
    KredietnummersBijwerken
    Set ws = ThisWorkbook.Worksheets(BL_ADM)
    ws.Unprotect WACHTWOORD
    ws.Range("A" & ADM_NR_KOP & ":H" & (ADM_NR_KOP + 22)).Clear
    KopOpmaak ws, ADM_NR_KOP, "3. Eerstvolgende vrije kredietnummers"
    ws.Range("B" & (ADM_NR_KOP + 1)).Value = "Wordt vanzelf bijgehouden; de invoerkaart vult dit nummer in na het kiezen van de balanscode. " & _
                                             "'Zelf invullen': de eerste keer zelf een nummer geven (kolom C of knop hiernaast)."
    ws.Range("B" & (ADM_NR_KOP + 1)).Font.Italic = True
    ws.Range("B" & (ADM_NR_KOP + 1)).Font.Color = RGB(89, 89, 89)
    ws.Range("B" & (ADM_NR_KOP + 2) & ":E" & (ADM_NR_KOP + 2)).Value = Array("Balanscode", "Omschrijving", "Reeks", "Eerstvolgende vrije nummer")
    With ws.Range("B" & (ADM_NR_KOP + 2) & ":E" & (ADM_NR_KOP + 2))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    r = ADM_NR_KOP + 3
    For Each x In AlleReeksen()
        oms = ""
        On Error Resume Next
        oms = CStr(CodeNaarWeergave("Tbl_Balanscode", Split(CStr(x(1)), " / ")(0)))
        On Error GoTo 0
        If InStr(oms, " - ") > 0 Then oms = Mid$(oms, InStr(oms, " - ") + 3) Else oms = VasteNaam(CStr(x(1)))
        ws.Cells(r, 2).Value = x(1)
        ws.Cells(r, 3).Value = oms
        ws.Cells(r, 4).NumberFormat = "@"
        ws.Cells(r, 4).Value = x(0)
        ws.Cells(r, 5).Value = NummerTekst(CStr(x(0)))
        If VarType(ws.Cells(r, 5).Value) = vbString Then ws.Cells(r, 5).Font.Color = RGB(192, 0, 0)
        ws.Cells(r, 5).NumberFormat = "0"
        ws.Cells(r, 5).HorizontalAlignment = xlLeft
        ws.Cells(r, 5).Font.Bold = True
        r = r + 1
    Next x
    ws.Range("B" & (ADM_NR_KOP + 3) & ":E" & (r - 1)).Borders(xlInsideHorizontal).Color = RGB(217, 217, 217)
    If ws.Columns(4).ColumnWidth < 9 Then ws.Columns(4).ColumnWidth = 9        ' kolom Reeks: tekst moet passen
    ws.Range("C" & (ADM_NR_KOP + 3) & ":C" & (r - 1)).Font.Color = RGB(89, 89, 89)
    On Error Resume Next
    ws.Shapes("uitleg_btnAdmNummerAanpassen").Delete
    ws.Shapes("btnAdmNummerAanpassen").Delete
    On Error GoTo 0
    ws.Rows(r).RowHeight = 26                                              ' direct onder de tabel, links uitgelijnd zoals de andere knoppen
    Set b = ws.Buttons.Add(ws.Range("B" & r).Left, ws.Range("B" & r).Top + 2, KNOP_B, 22)
    b.Name = "btnAdmNummerAanpassen"
    b.Caption = "Kredietnummer aanpassen"
    b.OnAction = "KredietnummerAanpassen"
    b.Placement = xlMove
    b.PrintObject = False
    KnopUitleg ws, "btnAdmNummerAanpassen", "Kredietnummer aanpassen (financi" & ChrW(235) & "le administratie): kies de reeks en typ het " & _
               "eerstvolgende vrije nummer, bijvoorbeeld als in Coda al een nummer is gebruikt dat niet in de Verzamellijst staat. " & _
               "Komt in het Logboek. Normaal is dit niet nodig: de nummers worden vanzelf bijgehouden."
End Sub

' Knop op Dashboard Adm: het eerstvolgende vrije nummer van een reeks met de hand aanpassen.
Public Sub KredietnummerAanpassen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim x As Variant, lijst As String, n As Long, k As Long, antw As String, reeks As String, oud As Double, nieuw As Double
    Dim rs() As String
    If Not HeeftRecht("Bevestigen") Then
        MsgBox "Dit doet de financi" & ChrW(235) & "le administratie.", vbExclamation
        Exit Sub
    End If
    For Each x In AlleReeksen()
        n = n + 1
        ReDim Preserve rs(1 To n)
        rs(n) = x(0)
        lijst = lijst & n & " = " & x(1) & " (" & x(0) & "): nu " & IIf(Opgeslagen(CStr(x(0))) > 0, Format$(Opgeslagen(CStr(x(0))), "0"), "Invullen") & vbCrLf
    Next x
    antw = Trim$(InputBox("Welke reeks? Typ het nummer." & vbCrLf & vbCrLf & lijst, "Kredietnummer aanpassen"))
    If Len(antw) = 0 Then Exit Sub
    If Not IsNumeric(antw) Then MsgBox "Dat is geen nummer.", vbExclamation: Exit Sub
    k = CLng(antw)
    If k < 1 Or k > n Then MsgBox "Dat nummer staat niet in de lijst.", vbExclamation: Exit Sub
    reeks = rs(k)
    oud = Opgeslagen(reeks)
    antw = Trim$(InputBox("Eerstvolgende vrije nummer voor reeks " & reeks & IIf(oud > 0, " (nu " & Format$(oud, "0") & ")", " (eerste nummer)") & ":", _
                          "Kredietnummer aanpassen", IIf(oud > 0, Format$(oud, "0"), reeks & "001")))
    If Len(antw) = 0 Then Exit Sub
    If Not IsNumeric(antw) Or Left$(antw, 4) <> reeks Or Len(antw) <> 7 Then
        MsgBox "Typ een nummer van 7 cijfers dat begint met " & reeks & ". Er is niets veranderd.", vbExclamation
        Exit Sub
    End If
    nieuw = CDbl(antw)
    Bewaar reeks, nieuw
    KredietnummersBijwerken
    LogRegel "kredietnummer aangepast", , , "Reeks " & reeks, Format$(oud, "0"), Format$(nieuw, "0"), "met de hand"
    MsgBox "Reeks " & reeks & ": eerstvolgende vrije nummer is nu " & Format$(Opgeslagen(reeks), "0") & ".", vbInformation
End Sub
