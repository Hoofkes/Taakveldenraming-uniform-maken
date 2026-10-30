Attribute VB_Name = "Kredieten_3_Verwerking"
Option Explicit

' Haalt regels helemaal weg zonder Excel-rijen te verwijderen: de regels eronder schuiven een plek omhoog
' (alleen invoerkolommen, formulekolommen blijven staan) en de onderste regels worden leeggemaakt.
' Alle gegevens van de weggehaalde regels gaan eerst naar het Logboek. Nulmeting vooraf en achteraf; bij verschil terugzetten.
Public Function RegelsWeghalen(ByVal wv As Worksheet, ByVal rijen As Collection, ByVal reden As String) As Boolean
    Dim laatste As Long, eerste As Long, n As Long, i As Long, c As Long, k As Long, doel As Long
    Dim kols As Variant, weg() As Boolean, oudW() As Variant, oudF() As Variant, rv As Variant
    Dim voor As Variant, na As Variant, verwacht(0 To 9) As Double, tekst As String, kolMeet As Variant, tmp() As String
    Dim logs As New Collection, lg As Variant, geschoven As Boolean, foutTekst As String

    On Error GoTo Fout
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

    ' logboektekst per regel klaarzetten (wordt pas na een geslaagde nameting weggeschreven)
    For Each rv In rijen
        tekst = ""
        For k = 0 To UBound(kols)
            If Not Leeg(wv.Cells(rv, kols(k)).Value) Then
                tekst = tekst & Split(wv.Cells(1, kols(k)).Address(True, False), "$")(0) & "=" & CStr(wv.Cells(rv, kols(k)).Value) & "; "
            End If
        Next k
        logs.Add Array(wv.Cells(rv, 1).Value, wv.Cells(rv, 3).Value, "rij " & rv, reden & " | " & tekst)
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
    geschoven = True                             ' vanaf hier wordt bij een fout alles teruggezet
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
            HerstelRegels wv, eerste, laatste, kols, oudW, oudF
            Application.Calculate
            Application.ScreenUpdating = True
            LogRegel "verwijderen teruggedraaid", , , "", , , "nulmeting wijkt af (meting " & k & ")"
            MsgBox "Verschil bij de nameting. Alles is teruggezet. Er is niets verwijderd.", vbCritical
            Exit Function
        End If
    Next k

    ' nameting geslaagd: nu pas de logregels
    For Each lg In logs
        LogRegel "verwijderen (helemaal weg)", lg(0), lg(1), lg(2), , , lg(3)
    Next lg
    Application.ScreenUpdating = True
    RegelsWeghalen = True
    Exit Function
Fout:
    foutTekst = Err.Description
    On Error Resume Next
    If geschoven Then
        HerstelRegels wv, eerste, laatste, kols, oudW, oudF
        Application.Calculate
        LogRegel "verwijderen teruggedraaid", , , "", , , "fout: " & foutTekst
    End If
    Application.ScreenUpdating = True
    MsgBox "Fout: " & foutTekst & vbCrLf & IIf(geschoven, "Alles is teruggezet. ", "") & "Er is niets verwijderd." & _
           IIf(geschoven, vbCrLf & "Controleer voor de zekerheid het Logboek.", ""), vbCritical
    RegelsWeghalen = False
End Function

' Zet de bewaarde waarden en notaties terug (gebruikt door RegelsWeghalen).
Private Sub HerstelRegels(ByVal wv As Worksheet, ByVal eerste As Long, ByVal laatste As Long, ByVal kols As Variant, _
                          ByRef oudW() As Variant, ByRef oudF() As Variant)
    Dim c As Long, i As Long
    For c = 0 To UBound(kols)
        wv.Range(wv.Cells(eerste, kols(c)), wv.Cells(laatste, kols(c))).Value = oudW(c)
        For i = eerste To laatste
            wv.Cells(i, kols(c)).NumberFormat = oudF(c)(i)
        Next i
    Next c
End Sub

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
Public Function KaartRij(ByVal i As Long) As Long
    If i <= 8 Then
        KaartRij = eerste + i - 1                ' rij 10 t/m 17 (rij 13 = korte omschrijving, verborgen; rij 17 = samengesteld)
    ElseIf i <= 22 Then
        KaartRij = eerste + i                    ' rij 19 t/m 32 (rij 18 = naam samengesteld krediet)
    ElseIf i <= 28 Then
        KaartRij = eerste + 4 + 3 * (i - 23)     ' budgetten rechts: rij 14, 17, 20, 23, 26, 29
    ElseIf i = 29 Then
        KaartRij = eerste + 23                   ' toelichting (rij 33)
    Else
        KaartRij = eerste + 8                    ' naam samengesteld krediet: direct onder 'samengesteld' (rij 18)
    End If
End Function

' Kolom van de invoercel van veld i: C voor de gewone velden, F voor de budgetten (veldnaam staat er links naast).
Public Function KaartKol(ByVal i As Long) As Long
    If i >= 23 And i <= 28 Then KaartKol = 6 Else KaartKol = 3
End Function

' Rij van bijdrage k (1 of 2) in jaar j (0 = lopend jaar t/m 5 = lopend jaar + 5).
Public Function BijdrageRij(ByVal j As Long, ByVal k As Long) As Long
    BijdrageRij = eerste + 4 + 3 * j + k         ' twee regels onder elk budget (kolom E bron, kolom F bedrag)
End Function

' Bronnen van bijdragen van derden (tabel op blad Parameters, zelf aan te vullen).
Public Sub BijdrageBronTabelAanmaken()
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
Public Sub HerstelBijdrageLabels(ByVal wd As Worksheet)
    Dim j As Long, k As Long
    For j = 0 To 5
        For k = 1 To 2
            wd.Cells(BijdrageRij(j, k), 5).Formula = "=""Bijdrage " & k & " - ""&(Par_LopendJaar+" & j & ")"
        Next k
    Next j
End Sub

' Geeft de zelf aan te passen cellen op blad Parameters een gele kleur en de automatische kolommen een grijze kleur.
Public Sub KleurParameters()
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
' - Uitgaven zonder krediet: alarm, telt mee in de aansluiting. Krediet niet in de uitdraai: bedrag blijft staan, met notitie;
'   stond het krediet in de vorige upload wel met een bedrag, dan wordt het bedrag 0 (melding en notitie).
' - Afgesloten krediet (A) waarvan het bedrag wijzigt: wel verwerkt, melding en notitie.
' - Harde poort: aansluiting per jaar exact 0,00 en geen foutwaarde in de verdeelsleutel K. Klopt een jaar niet, dan wordt
'   NIETS weggeschreven (ook het andere jaar niet) en blijven peildatum en vorige upload ongewijzigd.
' - Volgorde: berekenen, poort, wegschrijven, nameting (bij verschil terugzetten), vastleggen, notities, nabewerking.


Public Sub UitgavenBijwerken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, wsP As Worksheet, pad As String, peil As Date, lj As Long, laatste As Long
    Dim wbU As Workbook, d As Variant, nRij As Long, nKol As Long, kop As Long
    Dim cNr As Long, cKs As Long, cBud As Long, cJ(0 To 1) As Long, j As Long, y As Long
    Dim r As Long, i As Long, k As Long, nr As String, ks As String, v As Double, t As String
    Dim nrIdx As Object, nrs() As String, bU() As Double, bI() As Double, bO() As Double, bTot() As Double, heeftTot() As Boolean
    Dim bud() As Double, skip() As Boolean, nN As Long, det(0 To 1) As Double, eind(0 To 1) As Double, heeftEind As Boolean
    Dim geenNr(0 To 1) As Double, vS As Variant, ns As Long, stamRij As Object, rijen As Collection
    Dim grpRij As Object, oud() As Double, nieuw() As Double, gezet() As Boolean, schrijven As Boolean
    Dim zk As Double, ni As Double, onb As Double, verwacht As Double, som As Double, kol As Long
    Dim m0 As Variant, m1 As Variant, arr() As Variant, rng As Range, tekst As String, alarm As Long
    Dim geschreven(0 To 1) As Boolean, oudKol(0 To 1) As Variant, sOud(0 To 1) As Double, sNieuw(0 To 1) As Double
    Dim rv As Variant, grpBedrag As Object, g As Variant, vAD As Variant, rij0 As Long, bew() As Variant, nB As Long
    Dim sch(0 To 1) As Boolean, klopt(0 To 1) As Boolean, poortDicht As Boolean, klaar As Boolean, vervallen() As Boolean
    Dim nw(0 To 1) As Variant, gz(0 To 1) As Variant, od(0 To 1) As Variant, vv(0 To 1) As Variant

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
    JGetalNormaliseren vS
    ns = UBound(vS, 1)
    Set stamRij = CreateObject("Scripting.Dictionary")
    Set grpRij = CreateObject("Scripting.Dictionary")
    For i = 1 To ns
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

    ' 6. Per jaar berekenen. Er wordt nog niets weggeschreven: eerst moeten alle jaren aansluiten (harde poort).
    m0 = MeetStap8(wv, laatste)
    mBlok = False
    For y = 0 To 1
        If cJ(y) = 0 Then GoTo VolgendJaar
        kol = IIf(y = 0, 29, 28)                       ' AC = lopend jaar, AB = vorig jaar
        sch(y) = (y = 0) Or (peil <= DateSerial(lj, 3, 31))
        ' formules in de kolom (bijv. handmatige verdeling =bedrag*40%) worden vervangen door de waarde uit de uitdraai
        ReDim oud(1 To ns): ReDim nieuw(1 To ns): ReDim gezet(1 To ns): ReDim vervallen(1 To ns)
        For i = 1 To ns
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

        ' Regels die niet in de uitdraai staan. Stond het krediet in de vorige upload wel met een bedrag,
        ' dan is dat bedrag vervallen en wordt het 0. Anders blijft het bedrag staan.
        For i = 1 To ns
            If Not gezet(i) And oud(i) <> 0 And Not Leeg(vS(i, 1)) Then
                nr = NrTekst(vS(i, 3))
                If sch(y) And mHeeftVorige And Len(nr) > 0 Then
                    If mVorige.Exists(nr & "|" & (lj - y)) Then
                        If Round(mVorige(nr & "|" & (lj - y)), 2) <> 0 Then vervallen(i) = True
                    End If
                End If
                If vervallen(i) Then
                    nieuw(i) = 0
                    Meld 4, "Bedrag vervallen: krediet niet meer in de uitdraai", lj - y, nr, CStr(vS(i, 4)), i + 1, oud(i), 0, _
                         "Stond in de vorige upload wel met een bedrag. Het bedrag wordt 0."
                Else
                    ni = ni + oud(i)
                End If
            End If
        Next i

        If sch(y) Then
            ' Harde poort: aansluiting
            som = 0
            For i = 1 To ns
                nieuw(i) = Round(nieuw(i), 2)
                som = som + nieuw(i)
            Next i
            verwacht = det(y) - onb - geenNr(y) - zk + ni
            If Round(som - verwacht, 2) <> 0 Then
                Meld 1, "ALARM: aansluiting sluit niet", lj - y, "", "", 0, verwacht, som, _
                     "Verschil " & Format$(som - verwacht, "#,##0.00") & ". Er is NIETS bijgewerkt, ook niet in het andere jaar."
                GoTo VolgendJaar
            End If
            klopt(y) = True
            mAansl = mAansl & "Aansluiting " & (lj - y) & ": uitdraai " & Format$(det(y), "#,##0.00") & ", verzamellijst " & _
                     Format$(som, "#,##0.00") & IIf(Round(som - det(y), 2) <> 0, " (verschil verklaard door uitgaven zonder krediet / bedragen die blijven staan)", "") & ". Sluit.  "
            sOud(y) = 0: sNieuw(y) = som
            For i = 1 To ns
                sOud(y) = sOud(y) + oud(i)
            Next i
            nw(y) = nieuw: gz(y) = gezet: od(y) = oud: vv(y) = vervallen
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

    ' 7. Harde poort over alle jaren samen: sluit een jaar niet aan, of is een verdeelsleutel fout, dan wordt niets weggeschreven.
    poortDicht = mBlok
    For y = 0 To 1
        If cJ(y) > 0 And sch(y) And Not klopt(y) Then poortDicht = True
    Next y
    If poortDicht Then
        alarm = MeldingenSchrijven(Dir(pad), peil)
        BeveiligingAan
        Application.ScreenUpdating = True
        ThisWorkbook.Worksheets(BL_MELD).Activate
        MsgBox "Er is NIETS bijgewerkt: " & Dir(pad) & vbCrLf & vbCrLf & _
               "Een jaar sluit niet aan of een verdeelsleutel (kolom K) is fout. Zie de ALARM-regels op blad '" & BL_MELD & "'." & vbCrLf & _
               "Peildatum en vorige upload zijn ook niet aangepast. Los het op en verwerk het bestand opnieuw.", vbCritical, "Niets bijgewerkt"
        Exit Sub
    End If

    ' 8. Wegschrijven (alle jaren in een keer)
    For y = 0 To 1
        If cJ(y) > 0 And sch(y) Then
            kol = IIf(y = 0, 29, 28)
            ReDim arr(1 To ns, 1 To 1)
            For i = 1 To ns
                If Leeg(vS(i, 1)) Then arr(i, 1) = vS(i, kol) Else arr(i, 1) = nw(y)(i)
            Next i
            Set rng = wv.Range(wv.Cells(2, kol), wv.Cells(laatste, kol))
            oudKol(y) = rng.Formula                     ' formules bewaren voor terugzetten
            geschreven(y) = True
            rng.Value = arr
        End If
    Next y

    ' 9. Nameting: alleen AB/AC mogen veranderd zijn. Anders terugzetten (via Fout).
    Application.Calculate
    m1 = MeetStap8(wv, laatste)
    For j = 0 To 9
        If j <> 3 And j <> 4 Then
            If Round(m0(j) - m1(j), 2) <> 0 Then Err.Raise vbObjectError + 11, , "Nulmeting wijkt af in meting " & j & ". Alles is teruggedraaid."
        End If
    Next j
    klaar = True                                 ' vanaf hier zijn de bedragen definitief: bij een fout niet meer terugdraaien

    ' 10. Vastleggen: peildatum, notitie in de kop, logboek en de bedragen van deze upload (vergelijking voor de volgende keer)
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
    ReDim bew(1 To 2 * nN + 1, 1 To 3)
    For i = 1 To nN
        For y = 0 To 1
            If cJ(y) > 0 Then
                If Round(bU(i, y) + bI(i, y) + bO(i, y), 2) <> 0 Then
                    nB = nB + 1
                    bew(nB, 1) = nrs(i): bew(nB, 2) = lj - y: bew(nB, 3) = Round(bU(i, y) + bI(i, y) + bO(i, y), 2)
                End If
            End If
        Next y
    Next i
    VorigeUploadSchrijven bew, nB, Dir(pad)

    ' 11. Notities per regel (pas na de controle): oude weg, nieuwe bij regels die niet in de uitdraai staan
    For y = 0 To 1
        If geschreven(y) Then
            kol = IIf(y = 0, 29, 28)
            wv.Range(wv.Cells(2, kol), wv.Cells(laatste, kol)).ClearComments
            For i = 1 To ns
                If vv(y)(i) Then
                    wv.Cells(i + 1, kol).AddComment "Niet meer in uitdraai " & Dir(pad) & " (peildatum " & Format$(peil, "dd-mm-yyyy") & "). Bedrag " & _
                                                    Format$(od(y)(i), "#,##0.00") & " is vervallen en op 0 gezet."
                ElseIf Not gz(y)(i) And od(y)(i) <> 0 Then
                    wv.Cells(i + 1, kol).AddComment "Niet in uitdraai " & Dir(pad) & " (peildatum " & Format$(peil, "dd-mm-yyyy") & "). Bedrag blijft staan."
                End If
                If gz(y)(i) And Round(nw(y)(i) - od(y)(i), 2) <> 0 And UCase$(Left$(CStr(vS(i, 8)), 1)) = "A" Then
                    Meld 2, "Afgesloten krediet, toch uitgaven/inkomsten verwerkt", lj - y, NrTekst(vS(i, 3)), CStr(vS(i, 4)), i + 1, od(y)(i), nw(y)(i), _
                         "Status " & vS(i, 8) & ". Bedrag gewijzigd en wel verwerkt. Gemarkeerd met een notitie in kolom '" & wv.Cells(1, kol).Text & "'."
                    wv.Cells(i + 1, kol).AddComment "Afgesloten krediet (" & vS(i, 8) & "): bedrag gewijzigd van " & Format$(od(y)(i), "#,##0.00") & _
                                                    " naar " & Format$(nw(y)(i), "#,##0.00") & " (uitdraai " & Format$(peil, "dd-mm-yyyy") & ")."
                End If
            Next i
        End If
    Next y

    ' 12. Nabewerking
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
    If Not klaar Then
        For y = 0 To 1
            If geschreven(y) Then wv.Range(wv.Cells(2, IIf(y = 0, 29, 28)), wv.Cells(laatste, IIf(y = 0, 29, 28))).Formula = oudKol(y)
        Next y
    End If
    BeveiligingAan
    Application.ScreenUpdating = True
    If klaar Then
        MsgBox "De werkelijke uitgaven zijn bijgewerkt en vastgelegd, maar daarna ging iets mis:" & vbCrLf & foutTekst & vbCrLf & vbCrLf & _
               "Kijk op blad '" & BL_MELD & "' en maak het kredietoverzicht zo nodig opnieuw.", vbExclamation
    Else
        MsgBox "Fout: " & foutTekst & vbCrLf & "Er is niets gewijzigd.", vbCritical
    End If
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
        ' een foutwaarde in K (bijv. #DEEL/0!) blokkeert de hele verwerking; tekst (nvt, administratief) wordt overgeslagen
        If IsError(vS(r, 11)) And bedrag <> 0 Then
            mBlok = True
            Meld 1, "ALARM: verdeelsleutel (kolom K) is een foutwaarde", jaar, NrTekst(vS(r, 3)), CStr(vS(r, 4)), r + 1, Empty, bedrag, _
                 "Samengesteld krediet groep " & grp & ". Herstel kolom K. Er is NIETS bijgewerkt."
        End If
        ' alleen uitgavenregels; een bijdrageregel (negatief budget) telt niet mee in de verdeelsleutel
        If Nul(vS(r, 11)) > 0 And BudgetRegel(vS, r) >= 0 Then somK = somK + Nul(vS(r, 11)): laatsteR = r
    Next r
    If somK <= 0 Then
        nieuw(rijen(1)) = bedrag
        If bedrag <> 0 Then Meld 4, "Geen geldige verdeelsleutel in groep " & grp, jaar, NrTekst(vS(rijen(1), 3)), CStr(vS(rijen(1), 4)), rijen(1) + 1, Empty, bedrag, _
                                "Geen enkele regel van de groep heeft een percentage in kolom K. Het hele bedrag staat op deze regel. Controleer de verdeling."
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

Public Function MapBestaat(ByVal map As String) As Boolean
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

Public Function Tkst(ByVal v As Variant) As String
    If Not IsError(v) Then Tkst = CStr(v)
End Function

' True als minstens een cel in het bereik een formule heeft.
Private Function HeeftFormules(ByVal rng As Range) As Boolean
    Dim hf As Variant
    hf = rng.HasFormula
    If IsNull(hf) Then HeeftFormules = True Else HeeftFormules = CBool(hf)
End Function

Public Function IsAlleenCijfers(ByVal t As String) As Boolean
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

Public Function IsGetal(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: IsGetal = True
    End Select
End Function

' Groepsnummer uit kolom J: een getal, of een nummer dat als tekst is opgeslagen ("7"). 0 = geen groepsnummer
' (leeg, "standaard", "stelpost" of een andere tekst).
Public Function JGroep(ByVal x As Variant) As Long
    Dim s As String
    On Error Resume Next
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then Exit Function
    If IsGetal(x) Then
        JGroep = CLng(x)
    ElseIf VarType(x) = vbString Then
        s = Trim$(CStr(x))
        If Len(s) > 0 Then
            If Not s Like "*[!0-9]*" Then JGroep = CLng(s)
        End If
    End If
End Function

' Zet in een ingelezen tabel (kolom 10 = J) groepsnummers die als tekst zijn opgeslagen om naar een getal.
' Alleen in het geheugen; de cellen in de werkmap blijven zoals ze zijn.
Public Sub JGetalNormaliseren(ByRef v As Variant)
    Dim r As Long, s As String
    On Error Resume Next
    For r = LBound(v, 1) To UBound(v, 1)
        If VarType(v(r, 10)) = vbString Then
            s = Trim$(CStr(v(r, 10)))
            If Len(s) > 0 Then
                If Not s Like "*[!0-9]*" Then v(r, 10) = CDbl(s)
            End If
        End If
    Next r
End Sub

' Kredietnummer uit kolom C als tekst (getal zonder decimalen).
Public Function NrTekst(ByVal v As Variant) As String
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

Public Function Letter(ByVal k As Long) As String
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

Private Sub VorigeUploadSchrijven(ByRef bew() As Variant, ByVal nB As Long, ByVal bestand As String)
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
    If nB > 0 Then ws.Range("A3").Resize(nB, 3).Value = bew
    ws.Visible = xlSheetHidden
End Sub

' Meldingen: 1 = ALARM, 2 = niet toegestaan, 3 = toelichting nodig, 4 = let op, 5 = info.
Public Sub Meld(ByVal prio As Long, ByVal soort As String, ByVal jaar As Long, ByVal nr As String, ByVal oms As String, _
                 ByVal rij As Long, ByVal oud As Variant, ByVal nieuw As Variant, ByVal toel As String)
    If mMeld Is Nothing Then Set mMeld = New Collection
    mMeld.Add Array(prio, soort, jaar, nr, oms, rij, oud, nieuw, toel)
End Sub

' Schrijft blad 'Meldingen uitgaven' opnieuw, gesorteerd op ernst. Geeft het aantal ALARM-meldingen terug.
Private Function MeldingenSchrijven(ByVal bestand As String, ByVal peil As Date) As Long
    Dim ws As Worksheet, rij As Long, p As Long, m As Variant, kleur As Long, alarm As Long
    Dim wv As Worksheet, stat As Object, r As Long, laatste As Long, nrT As String, rijNr As Object, omsNr As Object, volgnr As Object
    ' status (kolom H), rijnummer(s) en omschrijving per kredietnummer, voor de kolommen Status, Rij en Omschrijving
    Set wv = Vz()
    Set stat = CreateObject("Scripting.Dictionary")
    Set rijNr = CreateObject("Scripting.Dictionary")
    Set omsNr = CreateObject("Scripting.Dictionary")
    Set volgnr = CreateObject("Scripting.Dictionary")
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        nrT = NrTekst(wv.Cells(r, 3).Value)
        If Len(nrT) > 0 Then
            If Not stat.Exists(nrT) Then
                stat.Add nrT, CStr(wv.Cells(r, 8).Value)
                omsNr.Add nrT, CStr(wv.Cells(r, 4).Value)
                rijNr.Add nrT, CStr(r)
                volgnr.Add nrT, CStr(wv.Cells(r, 1).Value)
            Else
                rijNr(nrT) = rijNr(nrT) & ", " & r
                volgnr(nrT) = volgnr(nrT) & ", " & CStr(wv.Cells(r, 1).Value)
            End If
        End If
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
    ws.Range("A4:K4").Value = Array("Volgnr", "Rij", "Soort", "Jaar", "Kredietnummer", "Omschrijving", "Status", "Oud / verwacht", "Nieuw / uitdraai", "Verschil", "Toelichting")
    With ws.Range("A4:K4")
        .Font.Bold = True
        .Interior.Color = RGB(208, 208, 208)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    rij = 5
    For p = 1 To 5
        For Each m In mMeld
            If m(0) = p Then
                ws.Cells(rij, 3).Value = m(1)
                ws.Cells(rij, 4).Value = m(2)
                ws.Cells(rij, 5).NumberFormat = "@"
                ws.Cells(rij, 5).Value = m(3)
                ws.Cells(rij, 6).Value = m(4)
                If Len(CStr(m(4))) = 0 And omsNr.Exists(CStr(m(3))) Then ws.Cells(rij, 6).Value = omsNr(CStr(m(3)))
                If m(5) > 1 Then
                    ws.Cells(rij, 7).Value = wv.Cells(m(5), 8).Value
                ElseIf stat.Exists(CStr(m(3))) Then
                    ws.Cells(rij, 7).Value = stat(CStr(m(3)))
                ElseIf Len(CStr(m(3))) > 0 Then
                    ws.Cells(rij, 7).Value = "geen krediet"
                End If
                If m(5) > 0 Then
                    ws.Cells(rij, 2).Value = m(5)
                    If m(5) > 1 Then ws.Cells(rij, 1).Value = wv.Cells(m(5), 1).Value
                ElseIf rijNr.Exists(CStr(m(3))) Then
                    ws.Cells(rij, 1).NumberFormat = "@"
                    ws.Cells(rij, 1).Value = volgnr(CStr(m(3)))         ' alle volgnummers van dit kredietnummer
                    ws.Cells(rij, 2).NumberFormat = "@"
                    ws.Cells(rij, 2).Value = rijNr(CStr(m(3)))          ' alle rijen van dit kredietnummer in de Verzamellijst
                End If
                If Not IsEmpty(m(6)) Then ws.Cells(rij, 8).Value = m(6)
                If Not IsEmpty(m(7)) Then ws.Cells(rij, 9).Value = m(7)
                If Not IsEmpty(m(6)) And Not IsEmpty(m(7)) Then ws.Cells(rij, 10).Value = m(7) - m(6)
                ws.Cells(rij, 11).Value = m(8)
                Select Case p
                    Case 1: kleur = RGB(255, 150, 150): alarm = alarm + 1
                        ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 11)).Font.Bold = True
                    Case 2: kleur = RGB(255, 190, 190)
                    Case 3: kleur = RGB(255, 200, 120)
                    Case 4: kleur = RGB(255, 255, 204)
                    Case Else: kleur = xlNone
                End Select
                If kleur <> xlNone Then ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 11)).Interior.Color = kleur
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
    ws.Range("H5:J" & Application.Max(5, rij)).NumberFormat = "#,##0.00;-#,##0.00;-"
    ws.Range("A:B,D:E,G:G").HorizontalAlignment = xlLeft
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 44
    ws.Columns("D").ColumnWidth = 7
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 36
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H:J").ColumnWidth = 16
    ws.Columns("K").ColumnWidth = 80
    ws.Range("A4:K" & rij).VerticalAlignment = xlTop
    ws.Range("A4:K4").WrapText = True
    ws.Range("A5:B" & rij & ",C5:C" & rij & ",F5:F" & rij & ",K5:K" & rij).WrapText = True
    BladOpmaken ws, 4, 5, rij
    KnopTerug ws, ws.Range("F1")
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
' KredietoverzichtJaarrekening  : stand per 31-12 vorig jaar uit het blad 'Verzamellijst 31-12-<jaar>' (ontstaat bij de jaarovergang):
'                                 alle L en A van dat jaar.
' UitdraaiBlokPlaatsen          : blok 2 op het Dashboard (2a, 2b en 2c met alle knoppen van de financieel adviseur).
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
    Dim keuze As String, naam As String, kol As Long, lj As Long, ws As Worksheet, map As String, lijstNaam As String
    keuze = Trim$(InputBox("Kredietoverzicht op naam voor:" & vbCrLf & "1 = budgetbeheerder" & vbCrLf & "2 = projectleider" & vbCrLf & _
                           "3 = tijdschrijver", "Kredietoverzicht op naam", "1"))
    Select Case keuze
        Case "1": kol = 15: lijstNaam = "Lst_Budgetbeheerder"
        Case "2": kol = 16: lijstNaam = "Lst_Projectleider"
        Case "3": kol = 17: lijstNaam = "Lst_Tijdschrijver"
        Case "": Exit Sub
        Case Else
            MsgBox "Typ 1, 2 of 3.", vbExclamation
            Exit Sub
    End Select
    naam = KiesNaamUitLijst(lijstNaam, "Kredietoverzicht op naam")
    If Len(naam) = 0 Then Exit Sub
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    Set ws = OverzichtMaken("naam", Vz(), lj, kol, naam)
    If ws Is Nothing Then Exit Sub
    map = MapStambestand("02 Budgetbeheerders en projectleiders")
    If Len(map) > 0 Then
        map = SubMap(map, naam)
        If BladOpslaan(ws, map, "Kredietoverzicht " & naam & " " & Format$(Date, "dd-mm-yyyy")) Then
            ws.Activate
            MsgBox "Opgeslagen in:" & vbCrLf & map, vbInformation
        Else
            ws.Activate
        End If
    End If
End Sub

' Laat de gebruiker een naam kiezen uit een lijst op Parameters (genummerd). Typ het nummer, of een deel van de naam
' om de lijst te verkleinen. Geeft de naam terug, of "" als er niets is gekozen.
Public Function KiesNaamUitLijst(ByVal lijstNaam As String, ByVal titel As String) As String
    Dim namen() As String, res() As String, n As Long, nr As Long, cel As Range, i As Long, t As String
    Dim lijst As String, antw As String, getoond As Long
    On Error GoTo Fout
    For Each cel In ThisWorkbook.Names(lijstNaam).RefersToRange.Cells
        t = Trim$(CStr(cel.Value))
        If Len(t) > 0 Then
            n = n + 1
            ReDim Preserve namen(1 To n)
            namen(n) = t
        End If
    Next cel
    If n = 0 Then
        MsgBox "De lijst met namen is leeg. Ververs de keuzeopties en probeer het opnieuw.", vbExclamation, titel
        Exit Function
    End If
    Do
        lijst = "": getoond = 0
        For i = 1 To n
            If Len(lijst) < 700 Then
                lijst = lijst & i & " = " & namen(i) & vbCrLf
                getoond = i
            End If
        Next i
        If getoond < n Then lijst = lijst & "... en nog " & (n - getoond) & " namen: typ een deel van de naam." & vbCrLf
        antw = Trim$(InputBox("Kies een naam:" & vbCrLf & vbCrLf & lijst & vbCrLf & "Typ het nummer of een deel van de naam:", titel))
        If Len(antw) = 0 Then Exit Function
        If IsNumeric(antw) Then
            i = CLng(antw)
            If i >= 1 And i <= n Then
                KiesNaamUitLijst = namen(i)
                Exit Function
            End If
            MsgBox "Dat nummer staat niet in de lijst.", vbExclamation, titel
        Else
            nr = 0
            For i = 1 To n
                If InStr(1, namen(i), antw, vbTextCompare) > 0 Then
                    nr = nr + 1
                    ReDim Preserve res(1 To nr)
                    res(nr) = namen(i)
                End If
            Next i
            If nr = 0 Then
                MsgBox "Geen naam gevonden met '" & antw & "'.", vbInformation, titel
            ElseIf nr = 1 Then
                KiesNaamUitLijst = res(1)
                Exit Function
            Else
                ReDim namen(1 To nr)
                For i = 1 To nr
                    namen(i) = res(i)
                Next i
                n = nr
            End If
        End If
    Loop
    Exit Function
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical, titel
End Function

Public Sub KredietoverzichtJaarrekening()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, jaar As Long, best As Worksheet, j As Long, map As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Dit overzicht wordt gemaakt door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Verzamellijst 31-12-####" Then
            j = CLng(Right$(ws.Name, 4))
            If j > jaar Then jaar = j: Set best = ws
        End If
    Next ws
    If best Is Nothing Then
        If MsgBox("Er is nog geen blad 'Verzamellijst 31-12-<jaar>' (ontstaat bij de jaarovergang, stand per 31-12)." & vbCrLf & vbCrLf & _
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
Public Function OverzichtMaken(ByVal soort As String, ByVal bron As Worksheet, ByVal jaar As Long, ByVal naamKol As Long, ByVal naam As String) As Worksheet
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
    JGetalNormaliseren v
    n = UBound(v, 1)

    ' samengestelde kredieten: per groep (J) de kredietnummers, voor de verwijzing in de toelichting
    Dim grpNrs As Object, gk As String
    Set grpNrs = CreateObject("Scripting.Dictionary")
    For r = 1 To n
        If Not Leeg(v(r, 1)) And IsGetal(v(r, 10)) And Len(NrTekst(v(r, 3))) > 0 Then
            gk = CStr(v(r, 10))
            If InStr(", " & grpNrs(gk) & ",", ", " & NrTekst(v(r, 3)) & ",") = 0 Then
                grpNrs(gk) = grpNrs(gk) & IIf(Len(grpNrs(gk)) > 0, ", ", "") & NrTekst(v(r, 3))
            End If
        End If
    Next r

    ' kopje per investeringsgroep: deelkredieten met een eigen groepsnaam (zoals 'school d1') krijgen geen eigen kopje
    Dim grpKop() As String
    grpKop = GroepKoppen(v)

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
                    If stLetter = "A" And sj = jaar Then neem = True        ' afgesloten in het lopend jaar: net als in het raadsoverzicht
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
                sleutel(nSel) = IIf(toekomst, "1", "0") & Chr$(1) & SorteerDeel(v(r, 2)) & Chr$(1) & LCase$(grpKop(r)) & Chr$(1) & _
                                SorteerJ(v(r, 10)) & Chr$(1) & Format$(v(r, 1), "000000000")
            End If
        End If
    Next r
    If nSel = 0 Then
        If Not mStil Then MsgBox "Geen kredieten gevonden voor deze selectie.", vbInformation
        GoTo Einde
    End If
    ' samengestelde kredieten (zelfde getal in J): aantal delen in deze selectie, voor het extra overzicht onderaan
    Dim grpAantal As Object, nHoofd As Long
    Set grpAantal = CreateObject("Scripting.Dictionary")
    For i = 1 To nSel
        If IsGetal(v(sel(i), 10)) Then grpAantal(CStr(v(sel(i), 10))) = grpAantal(CStr(v(sel(i), 10))) + 1
    Next i
    SorteerSleutels sleutel, sel, toek, 1, nSel
    nHoofd = nSel                                ' alle kredieten staan in de lijst per programma

    ' 2. Meldingen van de laatste upload (alleen voor de actuele verzamellijst)
    Set meldRij = CreateObject("Scripting.Dictionary")
    Set meldNr = CreateObject("Scripting.Dictionary")
    If soort <> "jaarrekening" And BladBestaat(BL_MELD) Then
        With ThisWorkbook.Worksheets(BL_MELD)
            For r = 5 To .Cells(.Rows.Count, 3).End(xlUp).Row
                If IsNumeric(.Cells(r, 2).Value) And Not Leeg(.Cells(r, 2).Value) Then
                    meldRij(CLng(.Cells(r, 2).Value)) = CStr(.Cells(r, 3).Value)
                ElseIf Not Leeg(.Cells(r, 5).Value) Then
                    meldNr(CStr(.Cells(r, 5).Value)) = CStr(.Cells(r, 3).Value)
                End If
            Next r
        End With
    End If

    ' 3. Kolommen
    If soort = "naam" Then
        kols = Array("Kredietnummer", "Omschrijving investeringskrediet", "Projectleider", "Status", "Afschr", "Budget t/m " & jaar, "Totaal werkelijke uitgaven", "Restantkrediet", _
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
    cStat = IIf(soort = "naam", 4, 3)            ' op naam: kolom C is de projectleider
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
    If soort = "naam" Then
        ' kop budgetten komende jaren lichtblauw, kop toelichting lichtgroen
        With ws.Range(ws.Cells(5, nK + 1), ws.Cells(5, nK + 5))
            .Interior.Color = RGB(189, 215, 238)
            .Font.Color = RGB(0, 73, 118)
        End With
        With ws.Cells(5, cToel)
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 73, 118)
        End With
    End If

    ' 5. Regels per programma en investeringsgroep (kolom I), met tussentotalen
    rij = 6: eersteData = 6
    vProg = Chr$(0): vGrp = Chr$(0): vSectie = -1
    For i = 1 To nHoofd
        r = sel(i)
        sectie = IIf(toek(i), 1, 0)
        prog = CStr(CodeNaarWeergave("Tbl_Programma", v(r, 2)))
        grp = grpKop(r)
        ' groep sluiten bij wissel
        If vGrp <> Chr$(0) And (grp <> vGrp Or prog <> vProg Or sectie <> vSectie) Then
            Tussentotaal ws, rij, startGrp, "Totaal " & vGrp, cEerste, cLaatste, 2
        End If
        If vProg <> Chr$(0) And (prog <> vProg Or sectie <> vSectie) Then
            Tussentotaal ws, rij, startProg, "Totaal " & vProg, cEerste, cLaatste, 1
        End If
        If sectie <> vSectie And sectie = 1 Then
            ws.Cells(rij, 1).Value = "Nog niet gestart (start na " & jaar & ")"
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Font.Bold = True
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Interior.Color = RGB(214, 228, 240)
            ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Font.Color = RGB(0, 73, 118)
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
        End If
        vProg = prog: vGrp = grp: vSectie = sectie

        toel = Trim$(CStr(v(r, 36)))
        ' los deelkrediet van een samengesteld krediet (de andere delen vallen buiten deze selectie): verwijzen naar de andere delen
        If IsGetal(v(r, 10)) Then
            gk = Trim$(Replace(Replace(", " & grpNrs(CStr(v(r, 10))) & ",", ", " & NrTekst(v(r, 3)) & ",", ","), ",", " "))
            gk = Replace(Trim$(gk), "  ", ", ")
            If Len(gk) > 0 Then toel = "Samengesteld krediet, ook " & gk & "." & IIf(Len(toel) > 0, " " & toel, "")
        End If
        OverzichtRegel ws, rij, v, r, soort, nK, cBud, cWerk, cRest, cAfs, cToel, cMeld, toel, meldRij, meldNr, _
                       somBud, somWerk, somAfs, somRest, somBron
        If soort = "naam" Then
            ' projectleider mag de budgethouder aanpassen
            ws.Cells(rij - 1, 3).Locked = False
            ws.Cells(rij - 1, 3).Interior.Color = RGB(255, 242, 204)
            ' extra regel onder het krediet: actuele aanvulling op de toelichting door budgethouder of projectleider
            ws.Cells(rij, 2).Value = "Actuele aanvulling toelichting:"
            ws.Cells(rij, 2).Font.Italic = True
            ws.Cells(rij, 2).Font.Color = RGB(89, 89, 89)
            ws.Cells(rij, cToel).Interior.Color = RGB(226, 239, 218)
            rij = rij + 1
        End If
    Next i
    If nHoofd > 0 Then
        Tussentotaal ws, rij, startGrp, "Totaal " & vGrp, cEerste, cLaatste, 2
        Tussentotaal ws, rij, startProg, "Totaal " & vProg, cEerste, cLaatste, 1
    End If

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

    ' 5b. Extra overzicht: samengestelde investeringskredieten (verzamelkredieten). Deze kredieten staan hierboven al
    '     gewoon bij hun programma; hier nog een keer per project bij elkaar (alle delen onder elkaar, zonder kopjes per deel),
    '     met een totaal per project, zodat zichtbaar is hoe het budget van het hele project loopt (bijv. een school of
    '     wegen met riolering). Telt NIET mee in het eindtotaal hierboven.
    Dim sl2() As String, sel2() As Long, tk2() As Boolean, n2 As Long
    Dim xBud As Double, xWerk As Double, xAfs As Double, xRest As Double, xBron As Double
    ReDim sl2(1 To nSel): ReDim sel2(1 To nSel): ReDim tk2(1 To nSel)
    For i = 1 To nSel
        r = sel(i)
        If IsGetal(v(r, 10)) Then
            If grpAantal(CStr(v(r, 10))) >= 2 Then
                n2 = n2 + 1
                sel2(n2) = r: tk2(n2) = toek(i)
                sl2(n2) = Format$(v(r, 10), "000000000") & Chr$(1) & Format$(v(r, 1), "000000000")
            End If
        End If
    Next i
    If n2 > 0 Then
        Dim startSec As Long, startPrj As Long, prjNaam As String, vPrj As String, prjRijen As Collection, ia As Long
        SorteerSleutels sl2, sel2, tk2, 1, n2
        rij = rij + 3
        ws.Cells(rij, 1).Value = "Samengestelde investeringskredieten (verzamelkredieten)"
        ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Font.Bold = True
        ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Interior.Color = RGB(214, 228, 240)
        ws.Range(ws.Cells(rij, 1), ws.Cells(rij, nK + 7)).Font.Color = RGB(0, 73, 118)
        rij = rij + 1: startSec = rij
        vPrj = Chr$(0)
        For i = 1 To n2
            r = sel2(i)
            If CStr(v(r, 10)) <> vPrj Then
                If vPrj <> Chr$(0) Then Tussentotaal ws, rij, startPrj, "Totaal " & prjNaam, cEerste, cLaatste, 2
                Set prjRijen = New Collection
                For ia = i To n2
                    If CStr(v(sel2(ia), 10)) <> CStr(v(r, 10)) Then Exit For
                    prjRijen.Add sel2(ia)
                Next ia
                prjNaam = ProjectNaam(v, prjRijen)
                ws.Cells(rij, 2).Value = prjNaam
                ws.Cells(rij, 2).Font.Bold = True
                ws.Cells(rij, 2).Font.Italic = True
                rij = rij + 1: startPrj = rij
                vPrj = CStr(v(r, 10))
            End If
            ' toelichting: waar het deel in het overzicht staat (programma en investeringsgroep)
            toel = "Staat bij: " & CodeNaarWeergave("Tbl_Programma", v(r, 2)) & " - " & grpKop(r) & "."
            If tk2(i) Then toel = toel & " Nog niet gestart (start " & Mid$(Trim$(CStr(v(r, 8))), 2) & ")."
            ' eigen tellers: dit blok telt niet mee in de aansluiting
            OverzichtRegel ws, rij, v, r, soort, nK, cBud, cWerk, cRest, cAfs, cToel, cMeld, toel, meldRij, meldNr, _
                           xBud, xWerk, xAfs, xRest, xBron
        Next i
        Tussentotaal ws, rij, startPrj, "Totaal " & prjNaam, cEerste, cLaatste, 2
        ws.Cells(rij, 1).Value = "Deze kredieten staan hierboven ook al bij hun programma; dit is alleen een extra indeling per project " & _
                                 "(telt niet mee in het eindtotaal)."
        ws.Cells(rij, 1).Font.Italic = True
        ws.Cells(rij, 1).Font.Color = RGB(89, 89, 89)
    End If


    ' 6. Opmaak
    With ws.Range(ws.Cells(eersteData, cEerste), ws.Cells(rij, cLaatste))
        .NumberFormat = "#,##0;-#,##0;-"
        .Font.Color = RGB(0, 0, 0)
    End With
    ws.Range(ws.Cells(5, cEerste), ws.Cells(rij, cLaatste)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(5, 1), ws.Cells(rij, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(5, cStat), ws.Cells(rij, cStat)).HorizontalAlignment = xlCenter
    If soort = "naam" Then ws.Range(ws.Cells(5, 5), ws.Cells(rij, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(5, cToel), ws.Cells(rij, cMeld)).HorizontalAlignment = xlLeft
    ws.Columns(1).ColumnWidth = IIf(soort = "naam", 16, 9)
    ws.Columns(2).ColumnWidth = 45
    ws.Columns(cStat).ColumnWidth = 9
    If soort = "naam" Then
        ws.Columns(3).ColumnWidth = 22
        ws.Columns(5).ColumnWidth = 8
        ws.Range(ws.Cells(eersteData, 3), ws.Cells(rij, 3)).HorizontalAlignment = xlLeft
    End If
    ws.Range(ws.Columns(cEerste), ws.Columns(cLaatste)).ColumnWidth = 15
    ws.Columns(cToel).ColumnWidth = 80
    ws.Columns(cMeld).ColumnWidth = 28
    ws.Range(ws.Cells(eersteData, 2), ws.Cells(rij, 2)).WrapText = True
    ws.Range(ws.Cells(eersteData, cToel), ws.Cells(rij, cMeld)).WrapText = True
    ' toelichting mag op dit blad met de hand worden verbeterd (raakt geen formules en niet de verzamellijst)
    ws.Range(ws.Cells(eersteData, cToel), ws.Cells(rij, cToel)).Locked = False
    With ws.Range(ws.Cells(eersteData, cToel), ws.Cells(rij, cToel)).Validation       ' ook bij verbeteren op dit blad: blijft passen
        .Delete
        .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlLessEqual, Formula1:=CStr(MAX_TOEL + 60)
        .IgnoreBlank = True
        .ErrorTitle = "Toelichting te lang"
        .ErrorMessage = "Deze toelichting past dan niet meer in het overzicht. Maak hem korter."
    End With
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
    AfdrukA3 ws, "$5:$5", 5, rij
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
Public Sub Tussentotaal(ByVal ws As Worksheet, ByRef rij As Long, ByVal vanRij As Long, ByVal label As String, _
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

' Een kredietregel op het overzicht (blok per programma en blok samengestelde kredieten). Telt de bedragen op voor de aansluiting.
Private Sub OverzichtRegel(ByVal ws As Worksheet, ByRef rij As Long, ByRef v As Variant, ByVal r As Long, ByVal soort As String, _
                           ByVal nK As Long, ByVal cBud As Long, ByVal cWerk As Long, ByVal cRest As Long, ByVal cAfs As Long, _
                           ByVal cToel As Long, ByVal cMeld As Long, ByVal toel As String, ByVal meldRij As Object, ByVal meldNr As Object, _
                           ByRef somBud As Double, ByRef somWerk As Double, ByRef somAfs As Double, ByRef somRest As Double, ByRef somBron As Double)
    Dim k As Long
    If soort = "naam" Then
        ws.Cells(rij, 1).NumberFormat = "@"
        ws.Cells(rij, 1).Value = NrTekst(v(r, 3))
    Else
        ws.Cells(rij, 1).Value = v(r, 1)
    End If
    ws.Cells(rij, 2).Value = v(r, 4)
    If soort = "naam" Then
        ws.Cells(rij, 3).Value = v(r, 16)            ' projectleider
        ws.Cells(rij, 4).Value = v(r, 8)
        ws.Cells(rij, 5).Value = v(r, 25)
    Else
        ws.Cells(rij, 3).Value = v(r, 8)
    End If
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
End Sub

' Hoofddeel van een samengesteld krediet: het deel waarop medewerkers hun uren schrijven.
' Combinatie wegen/riolering: het deel wegen (kolom I). Scholen en gebouwen: het deel ruwbouw (omschrijving).
Public Function IsHoofddeel(ByRef v As Variant, ByVal r As Long) As Boolean
    IsHoofddeel = (LCase$(Trim$(CStr(v(r, 9)))) = "wegen") Or (InStr(1, CStr(v(r, 4)), "ruwbouw", vbTextCompare) > 0)
End Function

' Naam van een samengesteld krediet (project): kolom AT als die is ingevuld; anders het gemeenschappelijke begin van de
' omschrijvingen (bijv. 'ikc bricks'); anders de omschrijving van het hoofddeel (wegen/ruwbouw); anders van het eerste deel.
Private Function ProjectNaam(ByRef v As Variant, ByVal rijen As Collection) As String
    Dim x As Variant, p As String, d As String, i As Long, n As Long
    For Each x In rijen
        If Len(Trim$(CStr(v(x, 46)))) > 0 Then ProjectNaam = Trim$(CStr(v(x, 46))): Exit Function
    Next x
    If rijen.Count >= 2 Then
        p = LCase$(CStr(v(rijen(1), 4)))
        For Each x In rijen
            d = LCase$(CStr(v(x, 4)))
            n = 0
            For i = 1 To Application.Min(Len(p), Len(d))
                If Mid$(p, i, 1) <> Mid$(d, i, 1) Then Exit For
                n = i
            Next i
            ' alleen hele woorden: terug naar de laatste spatie als het woord niet in alle omschrijvingen af is
            If n < Len(d) Or n < Len(p) Then
                If Mid$(d & " ", n + 1, 1) <> " " Or Mid$(p & " ", n + 1, 1) <> " " Then n = InStrRev(Left$(p, n), " ")
            End If
            p = Left$(p, n)
        Next x
        p = Trim$(p)
        Do While Len(p) > 0 And InStr(",-(/", Right$(p, 1)) > 0
            p = Trim$(Left$(p, Len(p) - 1))
        Loop
        If Len(p) >= 6 Then ProjectNaam = Trim$(Left$(CStr(v(rijen(1), 4)), Len(p))): Exit Function
    End If
    For Each x In rijen
        If IsHoofddeel(v, CLng(x)) Then ProjectNaam = Trim$(CStr(v(x, 4))): Exit Function
    Next x
    ProjectNaam = Trim$(CStr(v(rijen(1), 4)))
End Function

' Kopje (investeringsgroep) per regel van de verzamellijst. Normaal kolom I. Uitzondering: een groepsnaam die alleen
' bij de delen van een samengesteld krediet hoort en op een cijfer eindigt (zoals 'school d1', 'school c3') is geen
' echte groep. Die delen komen onder een echte groep van hetzelfde programma met samengestelde kredieten (bijv.
' 'scholen'), anders onder 'samengestelde kredieten'. Door de sortering op kolom J staan de delen onder elkaar.
Private Function GroepKoppen(ByRef v As Variant) As String()
    Dim n As Long, r As Long, g As String, uit() As String, iJ As Object, tel As Object, best As Object, k As Variant
    Dim isDeel As Boolean, p As String
    n = UBound(v, 1)
    ReDim uit(1 To n)
    Set iJ = CreateObject("Scripting.Dictionary")
    For r = 1 To n
        If Not Leeg(v(r, 1)) Then
            g = LCase$(Trim$(CStr(v(r, 9))))
            If IsGetal(v(r, 10)) Then k = "J" & CStr(v(r, 10)) Else k = "-"
            If Not iJ.Exists(g) Then
                iJ(g) = k
            ElseIf iJ(g) <> k Then
                iJ(g) = "-"
            End If
        End If
    Next r
    ' echte groepen met samengestelde kredieten per programma (meest gebruikte)
    Set tel = CreateObject("Scripting.Dictionary")
    For r = 1 To n
        g = LCase$(Trim$(CStr(v(r, 9))))
        If Not Leeg(v(r, 1)) And IsGetal(v(r, 10)) And Not GroepIsDeel(g, iJ) Then
            k = CStr(v(r, 2)) & Chr$(1) & Trim$(CStr(v(r, 9)))
            tel(k) = tel(k) + 1
        End If
    Next r
    Set best = CreateObject("Scripting.Dictionary")
    For Each k In tel.Keys
        p = Split(k, Chr$(1))(0)
        If Not best.Exists(p) Then
            best(p) = k
        ElseIf tel(k) > tel(best(p)) Then
            best(p) = k
        End If
    Next k
    For r = 1 To n
        g = LCase$(Trim$(CStr(v(r, 9))))
        If GroepIsDeel(g, iJ) Then
            If best.Exists(CStr(v(r, 2))) Then uit(r) = Split(best(CStr(v(r, 2))), Chr$(1))(1) Else uit(r) = "samengestelde kredieten"
        Else
            uit(r) = Trim$(CStr(v(r, 9)))
        End If
    Next r
    GroepKoppen = uit
End Function

Private Function GroepIsDeel(ByVal g As String, ByVal iJ As Object) As Boolean
    If Len(g) = 0 Then Exit Function
    ' eindigt op letter + cijfer als los woord (zoals 'school c3'), of hoort bij precies een samengesteld krediet en eindigt op een cijfer
    If g Like "* [a-z]#" Then GroepIsDeel = True: Exit Function
    If iJ.Exists(g) Then GroepIsDeel = (Left$(iJ(g), 1) = "J") And (Right$(g, 1) Like "#")
End Function

' True als de naam voorkomt in de cel (tijdschrijvers staan met komma's in een cel).
Public Function NaamKomtVoor(ByVal cel As String, ByVal naam As String) As Boolean
    Dim deel As Variant
    For Each deel In Split(Replace(cel, ";", ","), ",")
        If StrComp(Trim$(CStr(deel)), naam, vbTextCompare) = 0 Then NaamKomtVoor = True: Exit Function
    Next deel
End Function

Public Function SorteerDeel(ByVal x As Variant) As String
    If IsGetal(x) Then SorteerDeel = Format$(x, "000000000") Else SorteerDeel = "~" & LCase$(Trim$(CStr(x)))
End Function

' Kolom J: genummerde deelkredieten eerst (op nummer), daarna standaard/stelpost alfabetisch.
Private Function SorteerJ(ByVal x As Variant) As String
    If IsGetal(x) Then SorteerJ = "0" & Format$(x, "000000000") Else SorteerJ = "1" & LCase$(Trim$(CStr(x)))
End Function

' Eenvoudige sortering (insertion sort) op de sleutels; sel en toek schuiven mee.
Public Sub SorteerSleutels(ByRef sl() As String, ByRef sel() As Long, ByRef toek() As Boolean, ByVal lo As Long, ByVal hi As Long)
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

Public Function BladBestaatStil(ByVal naam As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, naam, vbTextCompare) = 0 Then BladBestaatStil = True: Exit Function
    Next ws
End Function

' Blok 2 op het Dashboard (financieel adviseur) in drie subonderdelen met elk een eigen kopregel:
' 2a Kredietoverzichten, 2b Omzetting en controle, 2c Uitvraag en verwerking kadernota/begroting (voorlopig de oude knoppen).
Public Sub UitdraaiBlokPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, naam As Variant, i As Long, titels As Variant, acties As Variant, namen As Variant, rijOff As Variant, kolOff As Variant
    Dim koppen As Variant, kopRij As Variant
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    ws.Range("C" & (NM_KOP + 1) & ":C" & (FA_KOP - 1)).Validation.Delete       ' oude keuzecellen rol en naam
    On Error GoTo 0
    ws.Range("A" & NM_KOP & ":H" & (FA_KOP - 1)).Clear
    ws.Range("B" & NM_KOP).Value = "2. Financieel adviseur"
    ws.Range("B" & NM_KOP).Font.Bold = True
    ws.Range("B" & NM_KOP).Font.Size = 12
    kopRij = Array(NM_KOP + 1, NM_KOP + 4, NM_KOP + 8)
    koppen = Array("2a. Kredietoverzichten", "2b. Omzetting en controle", "2c. Uitvraag en verwerking kadernota/begroting")
    For i = 0 To 2
        ws.Range("B" & kopRij(i)).Value = koppen(i)
        ws.Range("B" & kopRij(i)).Font.Bold = True
        ws.Range("B" & kopRij(i)).Font.Color = RGB(31, 78, 121)
        ws.Range("B" & kopRij(i) & ":F" & kopRij(i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Rows(kopRij(i)).RowHeight = 20
    Next i
    On Error Resume Next
    For Each naam In Array("btnOverzichtRaad", "btnOverzichtNaam", "btnOverzichtJaarrekening", "btnAlleBudgethouders", "btnTijdschrijven", _
                           "btnUitvraag", "btnActualisatie", "btnJaarovergang", "btnUitgaven", "btnAfgekeurd", "btnHercontrole", _
                           "btnBhVervangen", "btnPlVervangen", "btnTsVervangen", "btnExport", "btnUvStand", "btnUvControle", "btnUvOphalen", "btnUvVerwerkAlle")
        ws.Shapes(CStr(naam)).Delete
        ws.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    For Each naam In Array(NM_KOP + 2, NM_KOP + 3, NM_KOP + 5, NM_KOP + 6, NM_KOP + 7, NM_KOP + 9)
        ws.Rows(CLng(naam)).RowHeight = 26
    Next naam
    ' 2b: rij 1 jaarovergang en afgekeurd, rij 2 namen vervangen, rij 3 import en export naast elkaar
    namen = Array("btnOverzichtNaam", "btnOverzichtRaad", "btnOverzichtJaarrekening", "btnAlleBudgethouders", "btnTijdschrijven", _
                  "btnJaarovergang", "btnAfgekeurd", "btnHercontrole", _
                  "btnBhVervangen", "btnPlVervangen", "btnTsVervangen", "btnUitgaven", "btnExport", _
                  "btnUitvraag", "btnUvOphalen", "btnUvVerwerkAlle")
    titels = Array("Kredietoverzicht op naam", "Kredietoverzicht raad", "Kredietoverzicht jaarrekening", "Overzichten alle budgethouders", "Overzicht tijdschrijven", _
                   "Jaarovergang uitvoeren", "Afgekeurd krediet herstellen", "Aanbieden voor hercontrole", _
                   "Budgethouder vervangen", "Projectleider vervangen", "Tijdschrijver vervangen", "Import werkelijke uitgaven", "Export naar financieel pakket", _
                   "1. Uitvraag versturen", "2. Aanleveringen ophalen", "3. Uitvraag verwerken")
    acties = Array("KredietoverzichtOpNaam", "KredietoverzichtMaken", "KredietoverzichtJaarrekening", "KredietoverzichtAlleBudgethouders", "OverzichtTijdschrijven", _
                   "JaarovergangUitvoeren", "AfgekeurdHerstellen", "HercontroleAanbieden", _
                   "BudgethouderVervangen", "ProjectleiderVervangen", "TijdschrijverVervangen", "UitgavenBijwerken", "ExportNaarFinancieelPakket", _
                   "UitvraagBegroting", "UitvraagOphalen", "UitvraagVerwerken")
    rijOff = Array(NM_KOP + 2, NM_KOP + 2, NM_KOP + 2, NM_KOP + 2, NM_KOP + 3, _
                   NM_KOP + 5, NM_KOP + 5, NM_KOP + 5, _
                   NM_KOP + 6, NM_KOP + 6, NM_KOP + 6, NM_KOP + 7, NM_KOP + 7, _
                   NM_KOP + 9, NM_KOP + 9, NM_KOP + 9)
    kolOff = Array(0, 1, 2, 3, 0, 0, 1, 2, 0, 1, 2, 0, 1, 0, 1, 2)
    For i = 0 To UBound(namen)
        Set b = ws.Buttons.Add(ws.Range("B1").Left + kolOff(i) * 205, ws.Range("B" & rijOff(i)).Top + 2, 195, 22)
        b.Name = namen(i)
        b.Caption = titels(i)
        b.OnAction = acties(i)
    Next i
End Sub
