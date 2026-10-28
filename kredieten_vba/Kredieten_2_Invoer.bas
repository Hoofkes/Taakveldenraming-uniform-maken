Attribute VB_Name = "Kredieten_2_Invoer"
Option Explicit

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
    GebruikerKnopPlaatsen
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

Public Function VoorCode(ByVal t As String) As String
    Dim p As Long
    p = InStr(t, " - ")
    If p > 0 Then VoorCode = Trim$(Left$(t, p - 1)) Else VoorCode = Trim$(t)
End Function

Public Function GetalOfTekst(ByVal v As Variant) As Variant
    If Not Leeg(v) And IsNumeric(v) Then
        GetalOfTekst = CDbl(v)
    Else
        GetalOfTekst = Trim$(CStr(v))
    End If
End Function

Public Function CodeInTabel(ByVal tabel As String, ByVal code As String) As Boolean
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

Public Sub RuimRij(ByVal wv As Worksheet, ByVal r As Long)
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
Public Function StatusBudgetVeld(ByVal stLetter As String, ByVal stJaar As String) As Long
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
                    If JGroep(cel.Value) = k.grp Then
                        If Len(Trim$(CStr(wv.Cells(cel.Row, 46).Value))) > 0 Then
                            k.naam = Trim$(CStr(wv.Cells(cel.Row, 46).Value))
                            Exit For
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
    BlijfOpDashboard                             ' ook bij een foutmelding terug op het Dashboard
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
    ZetKaartVolgnr wv.Cells(r, 1).Value
    ZetKaartVeld wd, 1, CodeNaarWeergave("Tbl_Programma", wv.Cells(r, 2).Value)
    ZetKaartVeld wd, 2, wv.Cells(r, 3).Value
    ZetKaartVeld wd, 3, wv.Cells(r, 4).Value
    h = CStr(wv.Cells(r, 8).Value)
    ZetKaartVeld wd, 5, UCase$(Left$(h, 1))
    ZetKaartVeld wd, 6, Mid$(h, 2)
    ZetKaartVeld wd, 7, wv.Cells(r, 9).Value
    v = wv.Cells(r, 10).Value
    If JGroep(v) > 0 Then
        ZetKaartVeld wd, 8, GroepWeergave(JGroep(v))
    ElseIf VarType(v) = vbString Then
        ZetKaartVeld wd, 8, v
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
    KaartRijenPassendMaken
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
    BlijfOpDashboard                             ' ook bij een foutmelding terug op het Dashboard
    Dim wd As Worksheet, wv As Worksheet
    Dim k As TKaart
    Dim r As Long, i As Long, n As Long, nw As Long, x As Long
    Dim cols() As Long, vals() As Variant
    Dim wijzig(1 To 45) As Boolean, oudW(1 To 45) As Variant
    Dim voor As Variant, na As Variant, verschil As Boolean
    Dim dlt(0 To 9) As Double, som As Double, waarsch As String
    Dim keuze As VbMsgBoxResult, bevestig As Boolean, gebr As String, herAanbieden As Boolean
    Dim nB As Long, kk As Long, jj As Long, fb As String, laatste As Long, rN As Long, volgnr0 As Long, maxRij As Long
    Dim bron(1 To 12) As String, bed(1 To 12, 0 To 5) As Variant, sumB(0 To 5) As Double, ext(0 To 9) As Double
    Dim geschreven As Boolean, klaar As Boolean, ctlGezet As Boolean, foutTekst As String, c As Long, c2 As Variant
    Dim terug(1 To 45) As Boolean, ctlOud(42 To 49) As Variant, nCtl As Long

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
    If KaartVolgnr() > 0 Then
        If CLng(wv.Cells(r, 1).Value) <> KaartVolgnr() Then
            MsgBox "Op deze rij staat nu een ander krediet dan dat u hebt geladen (het volgnummer klopt niet meer)." & vbCrLf & _
                   "Er is niets opgeslagen. Gebruik Scherm leegmaken en laad het krediet opnieuw.", vbCritical
            Exit Sub
        End If
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
            If Not ZonderControle(cols(i)) Then nCtl = nCtl + 1
        End If
    Next i

    ' bijdragen van derden (nieuwe regels)
    nB = LeesBijdragen(bron, bed, fb)
    If Len(fb) > 0 Then
        MsgBox "Niet opgeslagen. Dit moet nog worden aangepast:" & vbCrLf & vbCrLf & fb, vbExclamation
        Exit Sub
    End If
    If nw = 0 And nB = 0 Then
        If CStr(wv.Cells(r, 42).Value) = "Afgekeurd" And Not HeeftRecht("Bevestigen") Then
            AanbiedenTerControle r                ' niets gewijzigd: eventueel toch opnieuw aanbieden
            KaartWissen
        Else
            MsgBox "Er is niets gewijzigd.", vbInformation
        End If
        Exit Sub
    End If
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    rN = CLng(ThisWorkbook.Names("Par_EersteVrijeRij").RefersToRange.Value)
    volgnr0 = CLng(ThisWorkbook.Names("Par_VolgendVolgnr").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    If nB > 0 Then
        If rN + nB - 1 > maxRij Or rN <> laatste + 1 Then
            MsgBox "Er is geen ruimte voor de bijdragen (formules lopen tot rij " & maxRij & "). Voer eerst WerkbereikVerruimen uit.", vbExclamation
            Exit Sub
        End If
    End If
    For kk = 1 To nB
        For jj = 0 To 5
            sumB(jj) = sumB(jj) + Nul(bed(kk, jj))
        Next jj
    Next kk
    ext(1) = sumB(0)
    For jj = 1 To 5
        ext(4 + jj) = sumB(jj)
    Next jj

    If HeeftRecht("Bevestigen") Then
        keuze = MsgBox(nw & " wijziging(en)" & IIf(nB > 0, " en " & nB & " bijdrage(n)", "") & " opslaan." & vbCrLf & vbCrLf & "Ja = opslaan en meteen als gecontroleerd bevestigen" & vbCrLf & _
                       "Nee = opslaan, nog te controleren" & vbCrLf & "Annuleren = niet opslaan", vbYesNoCancel + vbQuestion, "Wijziging opslaan")
        If keuze = vbCancel Then Exit Sub
        bevestig = (keuze = vbYes)
    End If

    Application.ScreenUpdating = False
    voor = MeetStap8(wv, laatste + nB)
    geschreven = True                            ' vanaf hier wordt bij een fout alles teruggezet
    For i = 1 To n
        If wijzig(i) Then
            oudW(i) = wv.Cells(r, cols(i)).Value
            terug(i) = True
            ZetCel wv, r, cols(i), vals(i)
            If cols(i) = 26 Then dlt(1) = dlt(1) + Nul(vals(i)) - Nul(oudW(i))
            If cols(i) >= 31 And cols(i) <= 35 Then dlt(5 + cols(i) - 31) = dlt(5 + cols(i) - 31) + Nul(vals(i)) - Nul(oudW(i))
        End If
    Next i
    For kk = 1 To nB
        SchrijfBijdrageRegel wv, r, rN + kk - 1, volgnr0 + kk - 1, bron(kk), bed, kk
        ControleerFormules wv, rN + kk - 1
    Next kk
    Application.Calculate

    na = MeetStap8(wv, laatste + nB)
    If na(0) <> voor(0) + nB Then verschil = True
    For i = 1 To 9
        If Abs(na(i) - (voor(i) + dlt(i) + ext(i))) > 0.005 Then verschil = True
    Next i
    If verschil Then
        For i = 1 To n
            If wijzig(i) Then ZetCel wv, r, cols(i), oudW(i)
        Next i
        For kk = 1 To nB
            RuimRij wv, rN + kk - 1
        Next kk
        Application.Calculate
        Application.ScreenUpdating = True
        MsgBox "Verschil bij de nameting. De wijzigingen zijn teruggezet. Er is niets gewijzigd.", vbCritical
        Exit Sub
    End If

    gebr = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    If nw > 0 Then
        For c = 42 To 49
            ctlOud(c) = wv.Cells(r, c).Value         ' controlekolommen bewaren om terug te kunnen zetten
        Next c
        ctlGezet = True
    End If
    If nw > 0 Then
        If bevestig Then
            wv.Cells(r, 42).Value = CtlGecontroleerd()
            wv.Cells(r, 43).Value = gebr
            wv.Cells(r, 44).NumberFormat = "dd-mm-yyyy"
            wv.Cells(r, 44).Value = Date
        ElseIf CStr(wv.Cells(r, 42).Value) = "Afgekeurd" And Not HeeftRecht("Bevestigen") Then
            ' afgekeurd krediet: blijft afgekeurd tot de adviseur het opnieuw aanbiedt (hieronder of later met de knop)
            herAanbieden = True
        ElseIf nCtl = 0 Then
            ' alleen namen, budgetten, uitgaven, afschrijftermijn, omschrijving of toelichting gewijzigd: geen controle nodig,
            ' de controlestatus blijft zoals hij was
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
    For kk = 1 To nB
        LogRegel "toevoegen bijdrage", volgnr0 + kk - 1, wv.Cells(r, 3).Value, "", , , "rij " & (rN + kk - 1) & ", bron " & bron(kk) & ", bij volgnummer " & wv.Cells(r, 1).Value
    Next kk

    If k.grp > 0 Then
        For x = 2 To CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
            If JGroep(wv.Cells(x, 10).Value) = k.grp Then
                If VarType(wv.Cells(x, 11).Value) <> vbString And Not IsError(wv.Cells(x, 11).Value) Then som = som + wv.Cells(x, 11).Value
            End If
        Next x
        If Abs(som - 1) > 0.002 Then waarsch = vbCrLf & vbCrLf & "Let op: de verdeelsleutel van groep " & k.grp & " komt uit op " & Format(som * 100, "0.0") & "% in plaats van 100%."
    End If

    klaar = True                                 ' de wijziging is gelukt; fouten hierna draaien die niet meer terug
    KeuzelijstenVerversen True
    KaartWissen
    BlijfOpDashboard
    Application.ScreenUpdating = True
    MsgBox "Opgeslagen (" & nw & " veld(en) gewijzigd" & IIf(nB > 0, ", " & nB & " bijdrage(n) toegevoegd", "") & ", rij " & r & ")." & waarsch, vbInformation
    If herAanbieden Then AanbiedenTerControle r
    Exit Sub

Fout:
    foutTekst = Err.Description
    Application.CutCopyMode = False
    If geschreven And Not klaar Then
        On Error Resume Next
        For i = 1 To n
            If terug(i) Then ZetCel wv, r, cols(i), oudW(i)
        Next i
        For kk = 1 To nB
            RuimRij wv, rN + kk - 1
        Next kk
        If ctlGezet Then
            For Each c2 In Array(42, 43, 44, 48, 49)
                wv.Cells(r, CLng(c2)).Value = ctlOud(CLng(c2))
            Next c2
        End If
        Application.Calculate
        Application.ScreenUpdating = True
        MsgBox "Fout: " & foutTekst & vbCrLf & "De wijzigingen zijn teruggezet. Er is niets gewijzigd." & vbCrLf & _
               "Controleer voor de zekerheid het Logboek.", vbCritical
    ElseIf klaar Then
        Application.ScreenUpdating = True
        MsgBox "De wijziging is opgeslagen (rij " & r & "), maar daarna ging iets mis: " & foutTekst, vbExclamation
    Else
        Application.ScreenUpdating = True
        MsgBox "Fout: " & foutTekst & vbCrLf & "Er is niets gewijzigd (rij " & r & ").", vbCritical
    End If
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
    ' teller te controleren / afgekeurd: rechts naast de knoppen (stond eerst bovenaan in B5)
    ws.Range("B5").FormatConditions.Delete
    ws.Range("B5").ClearContents
    With ws.Range("E" & (FA_KOP + 2))
        .Formula = "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)&"" regels te controleren, ""&COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")&"" afgekeurd"""
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .FormatConditions.Delete
    End With
    Dim fc As Object
    Set fc = VoegCF(ws.Range("E" & (FA_KOP + 2)), "=COUNTIF(" & VL & "$AP:$AP,""Afgekeurd"")>0")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 150)
    Set fc = VoegCF(ws.Range("E" & (FA_KOP + 2)), "=COUNTIF(" & VL & "$AP:$AP,""" & CtlTeControleren() & """)>0")
    If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 200, 120)
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
    KnopUitleg ws, "btnFaAanpassen", "Bestaand krediet aanpassen (incidenteel, bijvoorbeeld bij andere wetgeving of uitgangspunten): je kiest eerst " & _
               "waarop je zoekt (kredietnummer, omschrijving, activanummer of volgnummer) en typt wat je zoekt. Je gaat naar die regel " & _
               "in de Verzamellijst; pas aan en klik bovenaan op 'Klaar: goedkeuren'. Wijzigingen gaan naar het Logboek."
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
        Application.GoTo wv.Cells(open0, 2), True
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
    Application.GoTo wv.Cells(r, 2), True
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
    On Error GoTo klaar
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
klaar:
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
    Dim r As Long, n As Long, Fout As String
    If Not FaRecht() Then Exit Sub
    r = WerkRij()
    If r = 0 Then
        FaKnoppenVerzamellijst False
        MsgBox "Er staat geen regel open.", vbInformation
        Exit Sub
    End If
    Fout = FaRegelFouten(Vz(), r)
    If Len(Fout) > 0 Then
        If MsgBox("Let op, in de regel klopt nog iets niet:" & vbCrLf & vbCrLf & Fout & vbCrLf & "Toch sluiten?", _
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
        Application.GoTo .Range("A" & FA_KOP), True
    End With
End Sub

' Knop 'Klaar: goedkeuren' (Verzamellijst): keurt de regel goed die open staat.
Public Sub RegelBevestigen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim nog As Long
    Dim wv As Worksheet, r As Long, st As String, gebr As String, n As Long, Fout As String, isOpen As Boolean
    On Error GoTo Fout
    If Not FaRecht() Then Exit Sub
    Set wv = Vz()
    r = WerkRij()
    isOpen = (r > 0)
    If Not isOpen Then r = ControleGekozenRij()
    If r = 0 Then Exit Sub
    Fout = FaRegelFouten(wv, r)
    If Len(Fout) > 0 Then
        MsgBox "Niet goedgekeurd. Dit moet nog worden aangepast in de Verzamellijst (rij " & r & "):" & vbCrLf & vbCrLf & Fout, vbExclamation
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
    ' nog meer te controleren: direct door naar het volgende krediet (zonder eerst terug naar het Dashboard)
    nog = Application.WorksheetFunction.CountIf(wv.Range("AP:AP"), CtlTeControleren())
    If nog > 0 Then
        If MsgBox("Goedgekeurd" & IIf(n > 0, " (" & n & " veld(en) aangepast, staan in het Logboek)", "") & "." & vbCrLf & _
                  "Nog te controleren: " & nog & "." & vbCrLf & vbCrLf & "Door naar het volgende krediet?", _
                  vbYesNo + vbQuestion, "Goedgekeurd") = vbYes Then
            FaVolgende
            Exit Sub
        End If
        NaarBlok6
        Exit Sub
    End If
    NaarBlok6
    MsgBox "Goedgekeurd" & IIf(n > 0, " (" & n & " veld(en) aangepast, staan in het Logboek)", "") & "." & vbCrLf & _
           "Er zijn geen kredieten meer te controleren.", vbInformation
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

' Vervolg van de macro's voor het Dashboard investeringskredieten (gesplitst vanwege de grootte; declaraties in Kredieten_1_Basis).

' Kolommen op de Verzamellijst waarvan een wijziging GEEN controle door de financiele administratie nodig heeft:
' omschrijving (D, F), namen (O, P, Q), afschrijftermijn (Y), budgetten (Z, AE t/m AI), werkelijke uitgaven (AA t/m AC)
' en toelichting (AJ). Alle andere wijzigingen gaan ter controle naar de financiele administratie.
Public Function ZonderControle(ByVal kol As Long) As Boolean
    Select Case kol
        Case 4, 6, 15, 16, 17, 25, 26, 27, 28, 29, 31, 32, 33, 34, 35, 36: ZonderControle = True
    End Select
End Function

' ======================================================================
' Tijdschrijvers toevoegen
' ======================================================================
' Toont de lijst met namen (20 per pagina). U typt de nummers van de tijdschrijvers die erbij moeten, gescheiden door komma's.
' De namen die al in het veld staan blijven altijd staan (in de lijst gemarkeerd met *). De knop voegt alleen toe.
' Iemand weghalen doet u met de hand in het veld.
Public Sub TijdschrijversKiezen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Const PER As Long = 20
    Dim wd As Worksheet, cel As Range, lst As Range
    Dim namen() As String, al() As Boolean, erbij() As Boolean
    Dim n As Long, i As Long, pagina As Long, paginas As Long, van As Long, tot As Long
    Dim deel As Variant, resultaat As String, huidig As String
    Dim prompt As String, antw As String, tok() As String, t As Variant
    Dim Fout As Boolean, nr As Long

    On Error GoTo Fout
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set cel = wd.Cells(KaartRij(14), 3)
    Set lst = ThisWorkbook.Names("Lst_Tijdschrijver").RefersToRange
    For i = 1 To lst.Cells.Count
        If Len(Trim$(CStr(lst.Cells(i, 1).Value))) > 0 Then
            n = n + 1
            ReDim Preserve namen(1 To n)
            ReDim Preserve al(1 To n)
            ReDim Preserve erbij(1 To n)
            namen(n) = Trim$(CStr(lst.Cells(i, 1).Value))
        End If
    Next i
    ' huidige namen: blijven staan, in dezelfde volgorde
    For Each deel In Split(CStr(cel.Value), ",")
        If Len(Trim$(CStr(deel))) > 0 Then
            If Len(resultaat) > 0 Then resultaat = resultaat & ", "
            resultaat = resultaat & Trim$(CStr(deel))
            For i = 1 To n
                If StrComp(namen(i), Trim$(CStr(deel)), vbTextCompare) = 0 Then al(i) = True: Exit For
            Next i
        End If
    Next deel
    huidig = resultaat

    paginas = 0
    If n > 0 Then paginas = (n - 1) \ PER + 1
    pagina = 1
    Do While pagina <= paginas
        van = (pagina - 1) * PER + 1
        tot = pagina * PER
        If tot > n Then tot = n
        prompt = "Pagina " & pagina & " van " & paginas & ". Typ de nummers van de tijdschrijvers die erbij moeten, gescheiden door komma's. " & _
                 "Leeg laten = niemand van deze pagina." & vbCrLf & "* = staat er al en blijft staan." & vbCrLf & vbCrLf
        For i = van To tot
            prompt = prompt & IIf(al(i), "* ", "   ") & i & " = " & namen(i) & vbCrLf
        Next i
        antw = InputBox(prompt, "Tijdschrijver toevoegen", "")
        If StrPtr(antw) = 0 Then Exit Sub
        tok = Split(Replace(Replace(antw, ";", ","), " ", ""), ",")
        Fout = False
        For Each t In tok
            If Len(CStr(t)) > 0 Then
                If Not IsNumeric(t) Then
                    Fout = True
                Else
                    nr = CLng(t)
                    If nr < van Or nr > tot Then Fout = True
                End If
            End If
        Next t
        If Fout Then
            MsgBox "Typ alleen nummers van deze pagina (" & van & " t/m " & tot & "), gescheiden door komma's.", vbExclamation
        Else
            For Each t In tok
                If Len(CStr(t)) > 0 Then erbij(CLng(t)) = True
            Next t
            pagina = pagina + 1
        End If
    Loop

    ' staat iemand niet in de lijst, dan typt u de naam hier
    antw = InputBox("Staat een tijdschrijver die erbij moet niet in de lijst? Typ de naam hier, meerdere namen gescheiden door komma's." & vbCrLf & _
                    "Laat leeg als dat niet nodig is.", "Tijdschrijver toevoegen", "")
    If StrPtr(antw) = 0 Then Exit Sub

    For i = 1 To n
        If erbij(i) And Not al(i) Then
            If Len(resultaat) > 0 Then resultaat = resultaat & ", "
            resultaat = resultaat & namen(i)
        End If
    Next i
    For Each deel In Split(antw, ",")
        If Len(Trim$(CStr(deel))) > 0 Then
            If Not NaamKomtVoor(resultaat, Trim$(CStr(deel))) Then
                If Len(resultaat) > 0 Then resultaat = resultaat & ", "
                resultaat = resultaat & Trim$(CStr(deel))
            End If
        End If
    Next deel
    If resultaat = huidig Then
        MsgBox "Er is niemand toegevoegd.", vbInformation
        Exit Sub
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
Public Function LeesBijdragen(ByRef bron() As String, ByRef bed() As Variant, ByRef fouten As String) As Long
    Dim wd As Worksheet, j As Long, k As Long, n As Long, g As Long, p As Long, jj As Long
    Dim v As Variant, tekst As String, standaard As String, jaar As String
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    For j = 0 To 5
        jaar = CStr(CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value) + j)
        For k = 1 To 2
            v = wd.Cells(BijdrageRij(j, k), 6).Value
            If Not Leeg(v) Then
                If Not IsNumeric(v) Then
                    fouten = fouten & "- bijdrage " & k & " in " & jaar & ": een bedrag moet een getal zijn" & vbCrLf
                Else
                    tekst = Trim$(CStr(wd.Cells(BijdrageRij(j, k), 5).Value))
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
Public Sub SchrijfBijdrageRegel(ByVal wv As Worksheet, ByVal rBron As Long, ByVal rNieuw As Long, ByVal volgnr As Long, _
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
Public Sub ControleerFormules(ByVal wv As Worksheet, ByVal r As Long)
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
' Zoekt een krediet: eerst kiezen waarop gezocht wordt (zelfde keuze als bij Krediet wijzigen), daarna de zoekterm typen.
' Meerdere treffers: keuzelijst. Geeft de rij in de verzamellijst, 0 = niets gekozen.
Public Function ZoekKredietRij(ByVal titel As String) As Long
    Dim wv As Worksheet, zoek As String, r As Long, laatste As Long, n As Long, kk As Long, antw As String, lijst As String
    Dim treffers(1 To 15) As Long, meer As Boolean, hit As Boolean, keuze As String
    Set wv = Vz()
    keuze = Trim$(InputBox("Zoeken op:" & vbCrLf & "1 = kredietnummer" & vbCrLf & "2 = omschrijving (een deel van de tekst)" & vbCrLf & _
                           "3 = activanummer" & vbCrLf & "4 = volgnummer", titel, "1"))
    If Len(keuze) = 0 Then Exit Function
    If keuze <> "1" And keuze <> "2" And keuze <> "3" And keuze <> "4" Then
        MsgBox "Typ 1, 2, 3 of 4.", vbExclamation
        Exit Function
    End If
    zoek = Trim$(InputBox("Wat zoekt u?", titel))
    If Len(zoek) = 0 Then Exit Function
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        If Not IsEmpty(wv.Cells(r, 1).Value) Then
            Select Case keuze
                Case "1": hit = (NrTekst(wv.Cells(r, 3).Value) = zoek)
                Case "2": hit = (InStr(1, CStr(wv.Cells(r, 4).Value), zoek, vbTextCompare) > 0)
                Case "3": hit = (Trim$(CStr(wv.Cells(r, 24).Value)) = zoek)
                Case "4": hit = (CStr(wv.Cells(r, 1).Value) = zoek)
            End Select
            If hit Then
                If n < 15 Then n = n + 1: treffers(n) = r Else meer = True
            End If
        End If
    Next r
    If n = 0 Then
        MsgBox "Niets gevonden voor '" & zoek & "'.", vbInformation
        Exit Function
    End If
    If n = 1 Then ZoekKredietRij = treffers(1): Exit Function
    For kk = 1 To n
        r = treffers(kk)
        lijst = lijst & kk & " = vnr " & wv.Cells(r, 1).Value & " | " & NrTekst(wv.Cells(r, 3).Value) & " | " & Left$(CStr(wv.Cells(r, 4).Value), 30) & " | " & _
                wv.Cells(r, 8).Value & " | budget " & Format$(Nul(wv.Cells(r, 26).Value), "#,##0;-#,##0;0") & vbCrLf
    Next kk
    antw = Trim$(InputBox("Meerdere treffers" & IIf(meer, " (eerste 15, maak de zoekterm preciezer)", "") & ":" & vbCrLf & vbCrLf & _
                          lijst & vbCrLf & "Typ het nummer van de regel:", titel))
    If Len(antw) = 0 Then Exit Function
    If Not IsNumeric(antw) Then MsgBox "Dat is geen nummer.", vbExclamation: Exit Function
    kk = CLng(antw)
    If kk < 1 Or kk > n Then MsgBox "Dat nummer staat niet in de lijst.", vbExclamation: Exit Function
    ZoekKredietRij = treffers(kk)
End Function

' Zet de code op het blad Dashboard die bij elke invoer op de kaart KaartGewijzigd aanroept (Worksheet_Change).
' Gaat vanzelf bij het opbouwen van het Dashboard en bij het openen. Vereist in Excel: Bestand > Opties >
' Vertrouwenscentrum > Instellingen voor het Vertrouwenscentrum > Macro-instellingen >
' 'Toegang tot het objectmodel van het VBA-project vertrouwen'. Lukt het niet, dan (alleen als stil = False) een uitleg.
Public Sub DashboardEventInstalleren(Optional ByVal stil As Boolean = False)
    Dim cM As Object, i As Long, tekst As String, heeftChange As Boolean, heeftLink As Boolean, anders As String
    On Error GoTo niet
    Set cM = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.Worksheets("Dashboard").CodeName).CodeModule
    For i = 1 To cM.CountOfLines
        tekst = cM.Lines(i, 1)
        If InStr(1, tekst, "KaartGewijzigd", vbTextCompare) > 0 Then heeftChange = True
        If InStr(1, tekst, "FaHyperlink", vbTextCompare) > 0 Then heeftLink = True
    Next i
    For i = 1 To cM.CountOfLines
        tekst = cM.Lines(i, 1)
        If Not heeftChange And InStr(1, tekst, "Sub Worksheet_Change", vbTextCompare) > 0 Then anders = anders & vbCrLf & "Worksheet_Change: KaartGewijzigd Target": heeftChange = True
        If Not heeftLink And InStr(1, tekst, "Sub Worksheet_FollowHyperlink", vbTextCompare) > 0 Then anders = anders & vbCrLf & "Worksheet_FollowHyperlink: FaHyperlink Target": heeftLink = True
    Next i
    If Not heeftChange Then cM.AddFromString "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                                             "    KaartGewijzigd Target" & vbCrLf & "End Sub"
    If Not heeftLink Then cM.AddFromString "Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)" & vbCrLf & _
                                           "    FaHyperlink Target" & vbCrLf & "End Sub"
    If Len(anders) > 0 And Not stil Then MsgBox "Op het blad Dashboard staat al eigen code. Zet daarin zelf de regel(s):" & anders, vbExclamation
    Exit Sub
niet:
    If stil Then Exit Sub
    MsgBox "De koppeling met het blad Dashboard (melding bij invoer op een lege kaart) kon niet automatisch worden aangezet." & vbCrLf & vbCrLf & _
           "Oplossing (eenmalig):" & vbCrLf & _
           "1. Bestand > Opties > Vertrouwenscentrum > Instellingen voor het Vertrouwenscentrum > Macro-instellingen." & vbCrLf & _
           "2. Vink aan: 'Toegang tot het objectmodel van het VBA-project vertrouwen'. OK." & vbCrLf & _
           "3. Is het VBA-project met een wachtwoord vergrendeld? Open het dan eerst in de VBA-editor (Alt+F11)." & vbCrLf & _
           "4. Draai daarna de macro DashboardEventInstalleren.", vbInformation, "Koppeling Dashboard"
End Sub

' ----------------------------------------------------------------------
' Afgekeurde kredieten: signaal aan de financieel adviseur, herstellen en opnieuw aanbieden ter controle.
' De financiele administratie heeft altijd het laatste woord: na herstel keurt de adviseur het krediet niet zelf
' goed, maar biedt hij het opnieuw aan (status 'Te controleren'); daarna controleert de administratie opnieuw.
' ----------------------------------------------------------------------

' Rijen in de verzamellijst met status Afgekeurd.
Private Function AfgekeurdeRijen() As Collection
    Dim wv As Worksheet, r As Long, laatste As Long, c As New Collection
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        If CStr(wv.Cells(r, 42).Value) = "Afgekeurd" Then c.Add r
    Next r
    Set AfgekeurdeRijen = c
End Function

Private Function AfgekeurdTekst(ByVal r As Long) As String
    With Vz()
        AfgekeurdTekst = "volgnr " & .Cells(r, 1).Value & " | " & NrTekst(.Cells(r, 3).Value) & " | " & Left$(CStr(.Cells(r, 4).Value), 30) & _
                         vbCrLf & "      reden: " & CStr(.Cells(r, 45).Value) & " (" & CStr(.Cells(r, 43).Value) & ")"
    End With
End Function

' Bij openen (alleen financieel adviseur): melding als er afgekeurde kredieten zijn.
Public Sub AfgekeurdSignaal()
    Dim c As Collection, i As Long, lijst As String
    On Error Resume Next
    If Not HeeftRecht("Toevoegen") Then Exit Sub
    Set c = AfgekeurdeRijen()
    If c.Count = 0 Then Exit Sub
    For i = 1 To Application.Min(c.Count, 8)
        lijst = lijst & "- " & AfgekeurdTekst(c(i)) & vbCrLf
    Next i
    If MsgBox("De financi" & ChrW(235) & "le administratie heeft " & c.Count & " krediet(en) afgekeurd:" & vbCrLf & vbCrLf & lijst & vbCrLf & _
              "Nu herstellen? (Dashboard blok 2: 'Afgekeurd krediet herstellen')", vbYesNo + vbExclamation, "Afgekeurde kredieten") = vbYes Then
        AfgekeurdHerstellen
    End If
End Sub

' Kiest een afgekeurd krediet (bij meer dan een: lijst). 0 = geen.
Private Function KiesAfgekeurd(ByVal titel As String) As Long
    Dim c As Collection, i As Long, lijst As String, antw As String
    Set c = AfgekeurdeRijen()
    If c.Count = 0 Then
        MsgBox "Er zijn geen afgekeurde kredieten.", vbInformation, titel
        Exit Function
    End If
    If c.Count = 1 Then KiesAfgekeurd = c(1): Exit Function
    For i = 1 To Application.Min(c.Count, 12)
        lijst = lijst & i & " = " & AfgekeurdTekst(c(i)) & vbCrLf
    Next i
    antw = Trim$(InputBox("Afgekeurde kredieten:" & vbCrLf & vbCrLf & lijst & vbCrLf & "Typ het nummer:", titel, "1"))
    If Len(antw) = 0 Or Not IsNumeric(antw) Then Exit Function
    If CLng(antw) < 1 Or CLng(antw) > Application.Min(c.Count, 12) Then Exit Function
    KiesAfgekeurd = c(CLng(antw))
End Function

' Knop Afgekeurd krediet herstellen (blok 3): zet het afgekeurde krediet op de invoerkaart met de reden.
Public Sub AfgekeurdHerstellen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim r As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Afgekeurde kredieten worden hersteld door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    r = KiesAfgekeurd("Afgekeurd krediet herstellen")
    If r = 0 Then Exit Sub
    ToonOpKaart r
    MsgBox "Afgekeurd door " & Vz().Cells(r, 43).Value & ":" & vbCrLf & Vz().Cells(r, 45).Value & vbCrLf & vbCrLf & _
           "Pas het krediet aan op de invoerkaart en klik op 'Ingevuld? Verwerken!'." & vbCrLf & _
           "Daarna vraagt de macro of u het opnieuw aanbiedt ter controle (kan ook later met 'Aanbieden voor hercontrole').", _
           vbInformation, "Afgekeurd krediet herstellen"
End Sub

' Knop Aanbieden voor hercontrole (blok 3): het herstelde krediet gaat weer naar de financiele administratie.
Public Sub HercontroleAanbieden()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, r As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen de financieel adviseur biedt een krediet opnieuw aan.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then
        r = CLng(wd.Range("Kaart_Rijnummer").Value)
        If CStr(Vz().Cells(r, 42).Value) <> "Afgekeurd" Then r = 0
    End If
    If r = 0 Then r = KiesAfgekeurd("Aanbieden voor hercontrole")
    If r = 0 Then Exit Sub
    AanbiedenTerControle r
End Sub

Public Sub AanbiedenTerControle(ByVal r As Long)
    Dim wv As Worksheet
    Set wv = Vz()
    If MsgBox("Volgnummer " & wv.Cells(r, 1).Value & ": " & wv.Cells(r, 4).Value & vbCrLf & vbCrLf & _
              "Hersteld en opnieuw aanbieden ter controle aan de financi" & ChrW(235) & "le administratie?", _
              vbYesNo + vbQuestion, "Aanbieden voor hercontrole") <> vbYes Then Exit Sub
    wv.Cells(r, 42).Value = CtlTeControleren()
    wv.Cells(r, 43).ClearContents
    wv.Cells(r, 44).ClearContents
    wv.Cells(r, 45).Value = "hersteld na afkeuring: " & CStr(wv.Cells(r, 45).Value)
    wv.Cells(r, 48).Value = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    wv.Cells(r, 49).Value = Now
    LogRegel "aangeboden voor hercontrole", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "", , , "rij " & r
    MsgBox "Aangeboden. Het krediet staat weer bij de financi" & ChrW(235) & "le administratie (te controleren).", vbInformation
End Sub

' Wordt aangeroepen vanuit het blad Dashboard (Worksheet_Change) zodra iemand iets op de invoerkaart typt.
' Is er geen krediet geladen (dus: nieuw krediet) en mag de gebruiker niet toevoegen (financiele administratie),
' dan wordt de invoer meteen ongedaan gemaakt met een melding. Macro's geven geen melding (events staan dan uit).
Public Sub KaartGewijzigd(ByVal Target As Range)
    Dim wd As Worksheet
    On Error GoTo klaar
    Set wd = Target.Worksheet
    If Intersect(Target, wd.Range("Kaart_Invoer")) Is Nothing Then Exit Sub
    KaartRijenPassendMaken Intersect(Target, wd.Range("Kaart_Invoer"))
    If Not Intersect(Target, wd.Cells(KaartRij(8), KaartKol(8))) Is Nothing Then BladBeveiligen wd     ' naamveld vergrendelen of vrijgeven
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then Exit Sub        ' wijzigen van een bestaand krediet: mag
    If HeeftRecht("Toevoegen") Then Exit Sub
    Application.EnableEvents = False
    Application.Undo
    Application.EnableEvents = True
    MsgBox "U mag geen nieuwe kredieten toevoegen; dat doet de financieel adviseur." & vbCrLf & vbCrLf & _
           "Wilt u een bestaand krediet aanpassen? Klik dan eerst op 'Krediet wijzigen' en zoek het krediet op.", _
           vbExclamation, "Nieuw krediet"
klaar:
    Application.EnableEvents = True
End Sub

' Zorgt dat tekst in de invoercellen (kolom C van de kaart) altijd past: de tekst loopt door op een tweede regel
' en de rij wordt precies zo hoog als nodig (minimaal de gewone hoogte). Zonder rng: de hele kaart.
Public Sub KaartRijenPassendMaken(Optional ByVal rng As Range)
    Dim wd As Worksheet, k As Range
    On Error Resume Next
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    If rng Is Nothing Then Set rng = wd.Range("Kaart_Invoer")
    For Each k In rng.Cells
        If k.Column = 3 Then
            If Not wd.Rows(k.Row).Hidden Then
                k.WrapText = True
                k.VerticalAlignment = xlCenter
                wd.Rows(k.Row).AutoFit
                If wd.Rows(k.Row).RowHeight < 17.1 Then wd.Rows(k.Row).RowHeight = 17.1
            End If
        End If
    Next k
End Sub

Public Sub KredietVerwijderen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    BlijfOpDashboard                             ' ook bij een foutmelding terug op het Dashboard
    Dim wd As Worksheet, wv As Worksheet, r As Long, i As Long, laatste As Long, volgnr As Variant
    Dim rijen As New Collection, uitg As Double, reden As String, grpTxt As String, rv As Variant
    Dim weg As Boolean, foutTekst As String

    On Error GoTo Fout
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen kredieten verwijderen.", vbExclamation
        Exit Sub
    End If
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then
        r = CLng(wd.Range("Kaart_Rijnummer").Value)          ' krediet staat al op de kaart
        If KaartVolgnr() > 0 And r >= 2 Then
            If CLng(wv.Cells(r, 1).Value) <> KaartVolgnr() Then
                MsgBox "Op deze rij staat nu een ander krediet dan dat u hebt geladen (het volgnummer klopt niet meer)." & vbCrLf & _
                       "Er is niets verwijderd. Gebruik Scherm leegmaken en laad het krediet opnieuw.", vbCritical
                Exit Sub
            End If
        End If
    Else
        r = ZoekKredietRij("Krediet verwijderen")             ' eerst kiezen waarop gezocht wordt
        If r = 0 Then Exit Sub
    End If
    If r < 2 Or IsEmpty(wv.Cells(r, 1).Value) Then
        MsgBox "De geladen rij bestaat niet meer. Gebruik Scherm leegmaken en laad het krediet opnieuw.", vbCritical
        Exit Sub
    End If
    ' het krediet = de regel met dit volgnummer, plus de bijdrageregels (negatief budget) met hetzelfde kredietnummer
    ' als er na het verwijderen geen andere regel met dat kredietnummer overblijft
    volgnr = wv.Cells(r, 1).Value
    Dim nrK As String, anderen As Boolean
    nrK = NrTekst(wv.Cells(r, 3).Value)
    For i = 2 To laatste
        If wv.Cells(i, 1).Value = volgnr Then
            rijen.Add i
        ElseIf Len(nrK) > 0 And nrK <> "7999998" And NrTekst(wv.Cells(i, 3).Value) = nrK And Not IsEmpty(wv.Cells(i, 1).Value) Then
            If Not NegatiefBudgetRij(wv, i) Then anderen = True
        End If
    Next i
    If Not anderen And Len(nrK) > 0 And nrK <> "7999998" Then
        For i = 2 To laatste
            If wv.Cells(i, 1).Value <> volgnr And NrTekst(wv.Cells(i, 3).Value) = nrK And Not IsEmpty(wv.Cells(i, 1).Value) Then
                If NegatiefBudgetRij(wv, i) Then rijen.Add i
            End If
        Next i
    End If
    For Each rv In rijen
        uitg = uitg + Abs(Nul(wv.Cells(rv, 27).Value)) + Abs(Nul(wv.Cells(rv, 28).Value)) + Abs(Nul(wv.Cells(rv, 29).Value))
    Next rv
    If uitg <> 0 Then
        MsgBox "Er zijn werkelijke uitgaven of inkomsten geboekt. Een krediet met werkelijke bedragen mag nooit weg." & vbCrLf & _
               "Sluit het af: wijzig de status in A.", vbExclamation
        Exit Sub
    End If
    If JGroep(wv.Cells(r, 10).Value) > 0 Then
        grpTxt = vbCrLf & "Let op: dit is een deelkrediet van groep " & JGroep(wv.Cells(r, 10).Value) & ". De verdeelsleutel van de andere deelkredieten verandert."
    End If
    If MsgBox("Volgnummer " & volgnr & ": " & wv.Cells(r, 4).Value & vbCrLf & rijen.Count & " regel(s), inclusief bijdrageregels." & grpTxt & vbCrLf & vbCrLf & _
              "Helemaal verwijderen? Alleen het Logboek houdt de gegevens vast. Dit is alleen bedoeld voor echte fouten.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Krediet verwijderen") <> vbYes Then Exit Sub
    reden = "verwijderd via het Dashboard door " & CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    If Not RegelsWeghalen(wv, rijen, reden) Then Exit Sub
    weg = True                                   ' het krediet is weg; fouten hierna draaien dat niet terug
    KeuzelijstenVerversen True
    KaartWissen
    ControleLijstVernieuwen
    BlijfOpDashboard
    MsgBox "Krediet verwijderd (" & rijen.Count & " regel(s)). De gegevens staan in het Logboek.", vbInformation
    Exit Sub
Fout:
    foutTekst = Err.Description
    Application.ScreenUpdating = True
    If weg Then
        MsgBox "Het krediet is verwijderd, maar daarna ging iets mis: " & foutTekst & vbCrLf & _
               "Ververs de keuzelijsten met de knop op het Dashboard.", vbExclamation
    Else
        MsgBox "Fout: " & foutTekst & vbCrLf & "Er is niets verwijderd.", vbCritical
    End If
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
        ControleLijstVernieuwen
        MsgBox rijen.Count & " regel(s) weggehaald." & IIf(Len(overslaan) > 0, vbCrLf & "Niet weggehaald (werkelijke bedragen):" & overslaan, ""), vbInformation
    End If
End Sub
