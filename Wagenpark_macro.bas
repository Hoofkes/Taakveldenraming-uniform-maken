Attribute VB_Name = "Wagenpark"
Option Explicit

' Wagenparkbeheer: jaarovergang en aanschafbevestiging.
' Bladen: Dashboard, Voertuigen <jaar>, Vervangingsinvesteringen, Buiten gebruik.
' Data begint op rij 5, kopregel op rij 4. Onder elke lijst staat eerst een lege regel.

Private Const DASH As String = "Dashboard"
Private Const VI As String = "Vervangingsinvesteringen"
Private Const BG As String = "Buiten gebruik"
Private Const RIJ1 As Long = 5

' Kolommen Voertuigen <jaar>
Private Const V_NR As Long = 1, V_KRED As Long = 2, V_KENT As Long = 3, V_DOEL As Long = 4
Private Const V_WIJZE As Long = 5, V_BER As Long = 6, V_JAAR As Long = 7, V_LEV As Long = 8
Private Const V_KOST As Long = 9, V_ACTJR As Long = 10, V_ACTIVA As Long = 11, V_AANB As Long = 12
Private Const V_BIJK As Long = 13, V_SLOT As Long = 14, V_EXPL As Long = 15, V_LEASE As Long = 16
Private Const V_STATUS As Long = 17

' Kolommen Vervangingsinvesteringen
Private Const I_NR As Long = 1, I_DOEL As Long = 2, I_REG As Long = 3, I_JAAR As Long = 4
Private Const I_EINDE As Long = 5, I_AFSCHR As Long = 10, I_ACTIVA As Long = 11, I_TEACT As Long = 12
Private Const I_ACTJR As Long = 13, I_JLEASE As Long = 14, I_EXPL As Long = 15, I_OPM As Long = 16

' ================= bij openen =================
Public Sub Auto_Open()
    On Error GoTo Fout
    MaakKnoppen
    If Year(Date) > LopendJaar() Then
        Application.ScreenUpdating = False
        Do While Year(Date) > LopendJaar()
            VoerJaarovergangUit
        Loop
        Application.ScreenUpdating = True
        MsgBox "Jaarovergang uitgevoerd. Lopend jaar is nu " & LopendJaar() & "." & vbCrLf & _
               "Controleer de lijsten en sla het bestand op.", vbInformation, "Wagenpark"
    End If
    VernieuwKeuzelijst
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Fout bij openen: " & Err.Description, vbExclamation, "Wagenpark"
End Sub

Public Sub JaarovergangHandmatig()
    If MsgBox("Jaarovergang van " & LopendJaar() & " naar " & LopendJaar() + 1 & " nu uitvoeren?" & vbCrLf & _
              "Terugdraaien kan alleen door te sluiten zonder op te slaan.", vbYesNo + vbQuestion, "Wagenpark") = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    VoerJaarovergangUit
    Application.ScreenUpdating = True
    MsgBox "Lopend jaar is nu " & LopendJaar() & ".", vbInformation, "Wagenpark"
End Sub

' ================= jaarovergang =================
Public Sub VoerJaarovergangUit()
    Dim wsV As Worksheet, wsI As Worksheet
    Dim nieuwJaar As Long, r As Long, laatste As Long
    Set wsV = VoertuigenBlad()
    Set wsI = ThisWorkbook.Worksheets(VI)
    nieuwJaar = LopendJaar() + 1
    laatste = LaatsteRij(wsI)
    For r = RIJ1 To laatste
        If Val(wsI.Cells(r, I_JAAR).Value) = nieuwJaar And wsI.Cells(r, I_REG).Value <> "Doorschuiven" Then
            ZetOverNaarVoertuigen wsI, r, wsV, nieuwJaar
            VolgendVoertuigInVI wsI, r
        End If
    Next r
    Naam("LopendJaar").Value = nieuwJaar
    wsV.Name = "Voertuigen " & nieuwJaar
    wsV.Range("A1").Value = "Voertuigen " & nieuwJaar
    SorteerVI
    SorteerVoertuigen
    VernieuwKeuzelijst
End Sub

Private Sub ZetOverNaarVoertuigen(wsI As Worksheet, r As Long, wsV As Worksheet, jaar As Long)
    Dim lease As Boolean, ro As Long, nr As Long, w As Variant
    lease = (wsI.Cells(r, I_REG).Value = "Lease-koop")
    ro = ZoekRij(wsV, Basis(CStr(wsI.Cells(r, I_NR).Value)), "In gebruik")
    nr = NieuweRij(wsV)

    wsV.Cells(nr, V_NR).Value = wsI.Cells(r, I_NR).Value
    wsV.Cells(nr, V_KRED).Value = "n.n.b."
    wsV.Cells(nr, V_KENT).Value = "n.n.b."
    wsV.Cells(nr, V_DOEL).Value = wsI.Cells(r, I_DOEL).Value
    wsV.Cells(nr, V_WIJZE).Value = wsI.Cells(r, I_REG).Value
    w = ""
    If ro > 0 Then w = wsV.Cells(ro, V_BER).Value
    wsV.Cells(nr, V_BER).Value = Leeg(w, "n.n.b.")
    wsV.Cells(nr, V_JAAR).Value = jaar
    wsV.Cells(nr, V_ACTJR).Value = wsI.Cells(r, I_ACTJR).Value
    wsV.Cells(nr, V_ACTIVA).Value = Leeg(wsI.Cells(r, I_ACTIVA).Value, "n.n.b.")
    w = wsI.Cells(r, I_EXPL).Value
    If Len(CStr(w)) = 0 And ro > 0 Then w = wsV.Cells(ro, V_EXPL).Value
    If lease Then
        Dim d As Variant
        d = LeaseDelen(wsI, r)
        wsV.Cells(nr, V_LEV).Value = Naam("p_Lev").Value
        wsV.Cells(nr, V_AANB).Value = d(0)
        wsV.Cells(nr, V_BIJK).Value = d(1) + d(2)
        wsV.Cells(nr, V_SLOT).Value = d(3)
        wsV.Cells(nr, V_LEASE).Value = wsI.Cells(r, I_JLEASE).Value
        wsV.Cells(nr, V_EXPL).Value = Leeg(w, "n.n.b.")
    Else
        wsV.Cells(nr, V_LEV).Value = wsI.Cells(r, I_AFSCHR).Value
        wsV.Cells(nr, V_KOST).Value = Leeg(wsI.Cells(r, I_TEACT).Value, "n.n.b.")
        If Len(CStr(w)) > 0 Then wsV.Cells(nr, V_EXPL).Value = w
    End If
    wsV.Cells(nr, V_STATUS).Value = "Te vervangen " & jaar
End Sub

' Zet de VI-regel door naar het voertuig daarna (volgnummer +1, aanschafjaar = einde levensduur).
Private Sub VolgendVoertuigInVI(wsI As Worksheet, r As Long)
    Dim v As String, b As String, nieuwJaar As Long, rd As Long
    v = CStr(wsI.Cells(r, I_NR).Value)
    b = Basis(v)
    nieuwJaar = wsI.Cells(r, I_EINDE).Value
    wsI.Cells(r, I_NR).Value = Nummer(b, Volg(v) + 1)
    wsI.Cells(r, I_JAAR).Value = nieuwJaar
    wsI.Cells(r, I_ACTIVA).ClearContents
    wsI.Cells(r, I_OPM).ClearContents
    If wsI.Cells(r, I_REG).Value = "Lease-koop" Then SchrijfVIRij wsI, r, Empty   ' volgend voertuig weer op standaardbedragen
    If b = DoorschuifVan() Then
        rd = ZoekVIRij(wsI, DoorschuifNaar())
        If rd > 0 Then
            wsI.Cells(rd, I_NR).Value = Nummer(DoorschuifNaar(), Volg(CStr(wsI.Cells(rd, I_NR).Value)) + 1)
            wsI.Cells(rd, I_JAAR).Value = nieuwJaar
        End If
    End If
End Sub

' ================= aanschaf bevestigen =================
Public Sub AanschafBevestigen()
    Dim wsV As Worksheet, v As String, b As String
    Dim rn As Long, ro As Long, rd As Long, datum As Date, datumUit As Date
    Dim naarNr As String, naarDoel As Variant

    Set wsV = VoertuigenBlad()
    v = Trim(CStr(Naam("AB_Voertuig").Value))
    If v = "" Then MsgBox "Kies eerst een voertuig.", vbExclamation, "Wagenpark": Exit Sub
    rn = ZoekNummer(wsV, v)
    If rn = 0 Then MsgBox v & " staat niet op '" & wsV.Name & "'.", vbExclamation, "Wagenpark": Exit Sub
    If Left(CStr(wsV.Cells(rn, V_STATUS).Value), 12) <> "Te vervangen" Then
        MsgBox v & " is al in gebruik.", vbExclamation, "Wagenpark": Exit Sub
    End If
    If Trim(CStr(Naam("AB_Kenteken").Value)) = "" Or Not IsDate(Naam("AB_Datum").Value) Then
        MsgBox "Vul kenteken en datum ingebruikname in.", vbExclamation, "Wagenpark": Exit Sub
    End If
    datum = CDate(Naam("AB_Datum").Value)
    If IsDate(Naam("AB_DatumUit").Value) Then datumUit = CDate(Naam("AB_DatumUit").Value) Else datumUit = datum
    b = Basis(v)

    If MsgBox("Aanschaf van " & v & " bevestigen?" & vbCrLf & _
              "Het huidige voertuig " & b & " gaat dan buiten gebruik" & _
              IIf(b = DoorschuifVan(), " (schuift door naar " & DoorschuifNaar() & ")", "") & ".", _
              vbYesNo + vbQuestion, "Wagenpark") = vbNo Then Exit Sub
    Application.ScreenUpdating = False

    ' 1. oud voertuig
    If b = DoorschuifVan() Then
        naarNr = Nummer(DoorschuifNaar(), 1)
        naarDoel = Empty
        rd = ZoekRij(wsV, DoorschuifNaar(), "In gebruik")
        If rd > 0 Then
            naarNr = Nummer(DoorschuifNaar(), Volg(CStr(wsV.Cells(rd, V_NR).Value)) + 1)
            naarDoel = wsV.Cells(rd, V_DOEL).Value
            NaarBuitenGebruik wsV, rd, datumUit
        End If
        ro = ZoekRij(wsV, b, "In gebruik")
        If ro > 0 Then
            wsV.Cells(ro, V_NR).Value = naarNr
            If Not IsEmpty(naarDoel) Then wsV.Cells(ro, V_DOEL).Value = naarDoel
        End If
    Else
        ro = ZoekRij(wsV, b, "In gebruik")
        If ro > 0 Then NaarBuitenGebruik wsV, ro, datumUit
    End If

    ' 2. nieuw voertuig in gebruik
    rn = ZoekNummer(wsV, v)
    wsV.Cells(rn, V_KENT).Value = Trim(CStr(Naam("AB_Kenteken").Value))
    wsV.Cells(rn, V_JAAR).Value = Year(datum)
    If wsV.Cells(rn, V_WIJZE).Value = "Lease-koop" Then
        wsV.Cells(rn, V_ACTJR).Value = Year(datum) + Naam("p_SlotJr").Value
    Else
        wsV.Cells(rn, V_ACTJR).Value = Year(datum)
        If Len(CStr(Naam("AB_Kosten").Value)) > 0 Then wsV.Cells(rn, V_KOST).Value = Naam("AB_Kosten").Value
    End If
    If Len(CStr(Naam("AB_Krediet").Value)) > 0 Then wsV.Cells(rn, V_KRED).Value = Naam("AB_Krediet").Value
    If Len(CStr(Naam("AB_Activa").Value)) > 0 Then wsV.Cells(rn, V_ACTIVA).Value = Naam("AB_Activa").Value
    wsV.Cells(rn, V_STATUS).Value = "In gebruik"
    Application.ScreenUpdating = True
    If MsgBox("Wil je de bedragen van " & v & " aanpassen aan de werkelijkheid?", vbYesNo + vbQuestion, "Wagenpark") = vbYes Then
        BedragenVoertuigWijzigen wsV, rn
    End If
    Application.ScreenUpdating = False

    ' 3. opruimen
    Dim n As Variant
    For Each n In Array("AB_Voertuig", "AB_Kenteken", "AB_Datum", "AB_Krediet", "AB_Activa", "AB_Kosten", "AB_DatumUit")
        Naam(CStr(n)).ClearContents
    Next n
    SorteerVoertuigen
    VernieuwKeuzelijst
    Application.ScreenUpdating = True
    MsgBox v & " is in gebruik genomen.", vbInformation, "Wagenpark"
End Sub

Private Sub NaarBuitenGebruik(wsV As Worksheet, r As Long, datum As Date)
    Dim wsB As Worksheet, nb As Long
    Set wsB = ThisWorkbook.Worksheets(BG)
    nb = NieuweRij(wsB)
    wsB.Range(wsB.Cells(nb, 1), wsB.Cells(nb, 7)).Interior.Pattern = xlNone
    wsB.Cells(nb, 1).Value = wsV.Cells(r, V_NR).Value
    wsB.Cells(nb, 2).Value = wsV.Cells(r, V_KRED).Value
    wsB.Cells(nb, 3).Value = wsV.Cells(r, V_KENT).Value
    wsB.Cells(nb, 4).Value = wsV.Cells(r, V_DOEL).Value
    wsB.Cells(nb, 5).Value = wsV.Cells(r, V_JAAR).Value
    wsB.Cells(nb, 6).Value = datum
    wsB.Cells(nb, 6).NumberFormat = "dd-mm-yyyy"
    wsB.Cells(nb, 7).Value = wsV.Cells(r, V_ACTIVA).Value
    wsV.Rows(r).Delete
End Sub

' ================= wijzigen achteraf: voertuigen =================
Public Sub VoertuigWijzigen()
    Dim wsV As Worksheet, v As Variant, r As Long, keuze As Variant, c As Long, w As Variant, lijst As String
    Set wsV = VoertuigenBlad()
    v = InvoerTekst("Welk voertuig wil je wijzigen? (bijv. V05-01)", "")
    If IsEmpty(v) Then Exit Sub
    r = ZoekNummer(wsV, CStr(v))
    If r = 0 Then MsgBox v & " staat niet op '" & wsV.Name & "'.", vbExclamation, "Wagenpark": Exit Sub
    For c = V_KRED To V_LEASE
        lijst = lijst & c & ". " & wsV.Cells(4, c).Value & ":  " & wsV.Cells(r, c).Text & vbCrLf
    Next c
    Do
        keuze = InvoerGetal(v & vbCrLf & vbCrLf & lijst & vbCrLf & "99. Alle bedragen in één keer" & vbCrLf & vbCrLf & _
                            "Kies een nummer (Annuleren = klaar):", "")
        If IsEmpty(keuze) Then Exit Do
        c = CLng(keuze)
        If c = 99 Then
            BedragenVoertuigWijzigen wsV, r
        ElseIf c >= V_KRED And c <= V_LEASE Then
            w = InvoerTekst(wsV.Cells(4, c).Value & " (nu: " & wsV.Cells(r, c).Text & ")", CStr(wsV.Cells(r, c).Value))
            If Not IsEmpty(w) Then
                If IsNumeric(w) And c <> V_KRED And c <> V_ACTIVA And c <> V_EXPL Then wsV.Cells(r, c).Value = CDbl(w) Else wsV.Cells(r, c).Value = w
            End If
        End If
        lijst = ""
        For c = V_KRED To V_LEASE
            lijst = lijst & c & ". " & wsV.Cells(4, c).Value & ":  " & wsV.Cells(r, c).Text & vbCrLf
        Next c
    Loop
    VernieuwKeuzelijst
End Sub

Private Sub BedragenVoertuigWijzigen(wsV As Worksheet, r As Long)
    Dim a As Variant, b As Variant, o As Variant, s As Variant, l As Variant, k As Variant
    If wsV.Cells(r, V_WIJZE).Value = "Lease-koop" Then
        a = InvoerGetal("Aanbetaling", wsV.Cells(r, V_AANB).Value): If IsEmpty(a) Then Exit Sub
        b = InvoerGetal("Belettering", Naam("p_Belet").Value): If IsEmpty(b) Then Exit Sub
        o = InvoerGetal("Opbouw", Naam("p_Opbouw").Value): If IsEmpty(o) Then Exit Sub
        s = InvoerGetal("Slotbetaling", wsV.Cells(r, V_SLOT).Value): If IsEmpty(s) Then Exit Sub
        l = InvoerGetal("Leasesom per jaar", wsV.Cells(r, V_LEASE).Value): If IsEmpty(l) Then Exit Sub
        wsV.Cells(r, V_AANB).Value = a
        wsV.Cells(r, V_BIJK).Value = b + o
        wsV.Cells(r, V_SLOT).Value = s
        wsV.Cells(r, V_LEASE).Value = l
    Else
        k = InvoerGetal("Aanschafkosten (werkelijk)", wsV.Cells(r, V_KOST).Value): If IsEmpty(k) Then Exit Sub
        wsV.Cells(r, V_KOST).Value = k
    End If
    SorteerVoertuigen
End Sub

Public Sub VoertuigAfvoeren()
    Dim wsV As Worksheet, v As Variant, r As Long, d As Variant, ri As Long
    Set wsV = VoertuigenBlad()
    v = InvoerTekst("Welk voertuig gaat buiten gebruik zonder vervanging? (bijv. V25-01)", "")
    If IsEmpty(v) Then Exit Sub
    r = ZoekNummer(wsV, CStr(v))
    If r = 0 Then MsgBox v & " staat niet op '" & wsV.Name & "'.", vbExclamation, "Wagenpark": Exit Sub
    d = InvoerTekst("Datum buiten gebruik (dd-mm-jjjj)", Format(Date, "dd-mm-yyyy"))
    If IsEmpty(d) Then Exit Sub
    If Not IsDate(d) Then MsgBox "Geen geldige datum.", vbExclamation, "Wagenpark": Exit Sub
    If MsgBox(v & " naar 'Buiten gebruik' verplaatsen?", vbYesNo + vbQuestion, "Wagenpark") = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    NaarBuitenGebruik wsV, r, CDate(d)
    SorteerVoertuigen
    ri = ZoekVIRij(ThisWorkbook.Worksheets(VI), Basis(CStr(v)))
    Application.ScreenUpdating = True
    If ri > 0 Then
        If MsgBox("Ook de geplande vervanging " & ThisWorkbook.Worksheets(VI).Cells(ri, I_NR).Value & _
                  " schrappen uit Vervangingsinvesteringen?", vbYesNo + vbQuestion, "Wagenpark") = vbYes Then
            ThisWorkbook.Worksheets(VI).Rows(ri).Delete
            SorteerVI
        End If
    End If
    VernieuwKeuzelijst
    MsgBox v & " staat nu op 'Buiten gebruik'.", vbInformation, "Wagenpark"
End Sub

' ================= toekomstige vervangingsinvesteringen =================
Public Sub InvesteringWijzigen()
    Dim ws As Worksheet, v As Variant, r As Long, keuze As Variant, w As Variant, rd As Long
    Set ws = ThisWorkbook.Worksheets(VI)
    v = InvoerTekst("Welke investering wil je wijzigen? (V-nummer, bijv. V05 of V05-02)", "")
    If IsEmpty(v) Then Exit Sub
    r = ZoekVIRij(ws, Basis(CStr(v)))
    If r = 0 Then MsgBox v & " staat niet op 'Vervangingsinvesteringen'.", vbExclamation, "Wagenpark": Exit Sub
    Do
        keuze = InvoerGetal(ws.Cells(r, I_NR).Value & "  (" & ws.Cells(r, I_REG).Value & ", " & ws.Cells(r, I_JAAR).Value & _
                ", te activeren " & Format(ws.Cells(r, I_TEACT).Value, "€ #,##0") & ")" & vbCrLf & vbCrLf & _
                "1. Bedragen" & vbCrLf & "2. Aanschafjaar" & vbCrLf & "3. Regime (koop / lease-koop)" & vbCrLf & _
                "4. Gebruiksdoel" & vbCrLf & "5. Afschrijftermijn (koop)" & vbCrLf & "6. Activanummer" & vbCrLf & _
                "7. Jaarlease" & vbCrLf & "8. Exploitatienummer" & vbCrLf & "9. Opmerking" & vbCrLf & _
                "10. Terug naar standaardbedragen (lease-koop)" & vbCrLf & vbCrLf & "Kies een nummer (Annuleren = klaar):", "")
        If IsEmpty(keuze) Then Exit Do
        Select Case CLng(keuze)
            Case 1: BedragenInvesteringWijzigen ws, r
            Case 2
                w = InvoerGetal("Aanschafjaar", ws.Cells(r, I_JAAR).Value)
                If Not IsEmpty(w) Then
                    ws.Cells(r, I_JAAR).Value = CLng(w)
                    If Basis(CStr(ws.Cells(r, I_NR).Value)) = DoorschuifVan() Then
                        rd = ZoekVIRij(ws, DoorschuifNaar())
                        If rd > 0 Then ws.Cells(rd, I_JAAR).Value = CLng(w)
                    End If
                End If
            Case 3
                w = InvoerGetal("Regime: 1 = Koop, 2 = Lease-koop", "")
                If Not IsEmpty(w) Then
                    ws.Cells(r, I_REG).Value = IIf(CLng(w) = 2, "Lease-koop", "Koop")
                    SchrijfVIRij ws, r, Empty
                    If CLng(w) <> 2 Then BedragenInvesteringWijzigen ws, r
                End If
            Case 4: ZetTekst ws.Cells(r, I_DOEL), "Gebruiksdoel"
            Case 5
                w = InvoerGetal("Afschrijftermijn (jaar)", ws.Cells(r, I_AFSCHR).Value)
                If Not IsEmpty(w) And ws.Cells(r, I_REG).Value = "Koop" Then ws.Cells(r, I_AFSCHR).Value = CLng(w)
            Case 6: ZetTekst ws.Cells(r, I_ACTIVA), "Activanummer"
            Case 7
                w = InvoerGetal("Jaarlease", ws.Cells(r, I_JLEASE).Value)
                If Not IsEmpty(w) Then ws.Cells(r, I_JLEASE).Value = w
            Case 8: ZetTekst ws.Cells(r, I_EXPL), "Exploitatienummer"
            Case 9: ZetTekst ws.Cells(r, I_OPM), "Opmerking"
            Case 10: SchrijfVIRij ws, r, Empty
        End Select
    Loop
    SorteerVI
End Sub

Private Sub BedragenInvesteringWijzigen(ws As Worksheet, r As Long)
    Dim d As Variant, a As Variant, b As Variant, o As Variant, s As Variant, k As Variant
    If ws.Cells(r, I_REG).Value = "Lease-koop" Then
        d = LeaseDelen(ws, r)
        a = InvoerGetal("Aanbetaling", d(0)): If IsEmpty(a) Then Exit Sub
        b = InvoerGetal("Belettering", d(1)): If IsEmpty(b) Then Exit Sub
        o = InvoerGetal("Opbouw", d(2)): If IsEmpty(o) Then Exit Sub
        s = InvoerGetal("Slotbetaling", d(3)): If IsEmpty(s) Then Exit Sub
        SchrijfVIRij ws, r, Array(a, b, o, s)
    ElseIf ws.Cells(r, I_REG).Value = "Koop" Then
        k = InvoerGetal("Koopprijs", ws.Cells(r, I_TEACT).Value): If IsEmpty(k) Then Exit Sub
        ws.Cells(r, I_TEACT).Value = k
        If Len(CStr(ws.Cells(r, I_AFSCHR).Value)) = 0 Or ws.Cells(r, I_AFSCHR).HasFormula Then
            k = InvoerGetal("Afschrijftermijn (jaar)", 8): If Not IsEmpty(k) Then ws.Cells(r, I_AFSCHR).Value = CLng(k)
        End If
    End If
End Sub

Public Sub InvesteringToevoegen()
    Dim ws As Worksheet, b As Variant, nr As Long, volgnr As Long, reg As Variant, j As Variant, doel As Variant
    Dim wsV As Worksheet, ro As Long
    Set ws = ThisWorkbook.Worksheets(VI)
    Set wsV = VoertuigenBlad()
    b = InvoerTekst("V-nummer van het nieuwe of uit te breiden voertuig (zonder volgnummer, bijv. V30)", "")
    If IsEmpty(b) Then Exit Sub
    b = UCase(Trim(Basis(CStr(b))))
    If ZoekVIRij(ws, CStr(b)) > 0 Then MsgBox b & " staat al op 'Vervangingsinvesteringen'. Gebruik 'Investering wijzigen'.", vbExclamation, "Wagenpark": Exit Sub
    ro = ZoekRij(wsV, CStr(b), "In gebruik")
    volgnr = 1
    If ro > 0 Then volgnr = Volg(CStr(wsV.Cells(ro, V_NR).Value)) + 1
    Dim doelNu As String
    If ro > 0 Then doelNu = CStr(wsV.Cells(ro, V_DOEL).Value)
    doel = InvoerTekst("Gebruiksdoel", doelNu)
    If IsEmpty(doel) Then Exit Sub
    reg = InvoerGetal("Regime: 1 = Koop, 2 = Lease-koop", 2): If IsEmpty(reg) Then Exit Sub
    j = InvoerGetal("Aanschafjaar", LopendJaar() + 1): If IsEmpty(j) Then Exit Sub

    Application.ScreenUpdating = False
    nr = NieuweRij(ws)
    ws.Range(ws.Cells(nr, 1), ws.Cells(nr, I_OPM)).ClearContents
    ws.Cells(nr, I_NR).Value = Nummer(CStr(b), volgnr)
    ws.Cells(nr, I_DOEL).Value = doel
    ws.Cells(nr, I_REG).Value = IIf(CLng(reg) = 2, "Lease-koop", "Koop")
    ws.Cells(nr, I_JAAR).Value = CLng(j)
    If ro > 0 Then ws.Cells(nr, I_EXPL).Value = wsV.Cells(ro, V_EXPL).Value
    SchrijfVIRij ws, nr, Empty
    Application.ScreenUpdating = True
    If CLng(reg) <> 2 Then
        BedragenInvesteringWijzigen ws, nr
    End If
    SorteerVI
    MsgBox Nummer(CStr(b), volgnr) & " is toegevoegd.", vbInformation, "Wagenpark"
End Sub

Public Sub InvesteringSchrappen()
    Dim ws As Worksheet, v As Variant, r As Long
    Set ws = ThisWorkbook.Worksheets(VI)
    v = InvoerTekst("Welke investering wil je schrappen? (V-nummer)", "")
    If IsEmpty(v) Then Exit Sub
    r = ZoekVIRij(ws, Basis(CStr(v)))
    If r = 0 Then MsgBox v & " staat niet op 'Vervangingsinvesteringen'.", vbExclamation, "Wagenpark": Exit Sub
    If MsgBox(ws.Cells(r, I_NR).Value & " (" & ws.Cells(r, I_JAAR).Value & ") schrappen?" & vbCrLf & _
              "Het huidige voertuig blijft in gebruik.", vbYesNo + vbQuestion, "Wagenpark") = vbNo Then Exit Sub
    ws.Rows(r).Delete
    SorteerVI
End Sub

' Schrijft de formules van een regel op Vervangingsinvesteringen.
' delen = Empty: standaardbedragen; anders Array(aanbetaling, belettering, opbouw, slotbetaling).
Private Sub SchrijfVIRij(ws As Worksheet, r As Long, delen As Variant)
    Dim reg As String, d1 As String, sl As String, c As Long, kop As String
    reg = ws.Cells(r, I_REG).Value
    ws.Cells(r, I_EINDE).Formula = "=IF($C" & r & "=""Lease-koop"",$D" & r & "+p_Lev,IF($C" & r & "=""Koop"",$D" & r & "+$J" & r & ",0))"
    ws.Cells(r, I_ACTJR).Formula = "=IF($C" & r & "=""Lease-koop"",$D" & r & "+p_SlotJr,IF($C" & r & "=""Koop"",$D" & r & ",0))"
    If reg = "Lease-koop" Then
        If IsEmpty(delen) Then
            ws.Cells(r, I_TEACT).Formula = "=p_Aanbet+p_Belet+p_Opbouw+p_Slot"
            d1 = "p_Aanbet+p_Belet+p_Opbouw": sl = "p_Slot"
        Else
            ws.Cells(r, I_TEACT).Formula = "=" & Getal(delen(0)) & "+" & Getal(delen(1)) & "+" & Getal(delen(2)) & "+" & Getal(delen(3))
            d1 = "(" & Getal(delen(0)) & "+" & Getal(delen(1)) & "+" & Getal(delen(2)) & ")": sl = Getal(delen(3))
        End If
        ws.Cells(r, I_AFSCHR).Formula = "=p_Afschr"
        If Len(CStr(ws.Cells(r, I_JLEASE).Value)) = 0 Or ws.Cells(r, I_JLEASE).Value = 0 Then ws.Cells(r, I_JLEASE).Formula = "=p_Lease"
    Else
        If ws.Cells(r, I_TEACT).HasFormula Then ws.Cells(r, I_TEACT).ClearContents
        If ws.Cells(r, I_AFSCHR).HasFormula Then ws.Cells(r, I_AFSCHR).ClearContents
        ws.Cells(r, I_JLEASE).Value = 0
    End If
    For c = 6 To 9
        kop = ws.Cells(4, c).Address(True, False)          ' bijv. F$4
        If reg = "Lease-koop" Then
            ws.Cells(r, c).Formula = "=IF(" & kop & "=$D" & r & "," & d1 & ",0)+IF(" & kop & "=$M" & r & "," & sl & ",0)"
        Else
            ws.Cells(r, c).Formula = "=IF(" & kop & "=$D" & r & ",N($L" & r & "),0)"
        End If
    Next c
End Sub

' Leest aanbetaling, belettering, opbouw en slotbetaling uit de formule 'Te activeren' (=a+b+o+s).
Private Function LeaseDelen(ws As Worksheet, r As Long) As Variant
    Dim f As String, p As Variant, i As Long, uit(3) As Double
    f = ws.Cells(r, I_TEACT).Formula
    p = Split(Mid(f, 2), "+")
    If Left(f, 1) = "=" And UBound(p) = 3 Then
        For i = 0 To 3
            p(i) = Trim(p(i))
            If IsNumeric(p(i)) Then uit(i) = Val(p(i)) Else uit(i) = Naam(CStr(p(i))).Value
        Next i
    Else
        uit(0) = Naam("p_Aanbet").Value: uit(1) = Naam("p_Belet").Value
        uit(2) = Naam("p_Opbouw").Value: uit(3) = Naam("p_Slot").Value
    End If
    LeaseDelen = uit
End Function

Private Function Getal(x As Variant) As String
    Getal = Replace(CStr(CDbl(x)), ",", ".")
End Function

Private Sub ZetTekst(cel As Range, label As String)
    Dim w As Variant
    w = InvoerTekst(label & " (nu: " & cel.Text & ")", CStr(cel.Value))
    If Not IsEmpty(w) Then cel.Value = w
End Sub

' Application.InputBox geeft False bij Annuleren; dan Empty teruggeven.
Private Function InvoerGetal(vraag As String, standaard As Variant) As Variant
    Dim w As Variant
    w = Application.InputBox(vraag, "Wagenpark", standaard, Type:=1)
    If VarType(w) = vbBoolean Then InvoerGetal = Empty Else InvoerGetal = w
End Function

Private Function InvoerTekst(vraag As String, standaard As String) As Variant
    Dim w As Variant
    w = Application.InputBox(vraag, "Wagenpark", standaard, Type:=2)
    If VarType(w) = vbBoolean Then
        InvoerTekst = Empty
    ElseIf Len(Trim(CStr(w))) = 0 Then
        InvoerTekst = Empty
    Else
        InvoerTekst = Trim(CStr(w))
    End If
End Function

' ================= sorteren, keuzelijst, knoppen =================
Private Sub SorteerVoertuigen()
    Dim ws As Worksheet, laatste As Long
    Set ws = VoertuigenBlad()
    laatste = LaatsteRij(ws)
    If laatste <= RIJ1 Then GoTo Totalen
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(RIJ1, V_STATUS), ws.Cells(laatste, V_STATUS)), Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(RIJ1, V_NR), ws.Cells(laatste, V_NR)), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(RIJ1, 1), ws.Cells(laatste, V_STATUS))
        .Header = xlNo
        .Apply
    End With
Totalen:
    ' totaalregels staan op laatste+2 en laatste+3
    Dim c As Variant, rng As String
    For Each c In Array(V_KOST, V_AANB, V_BIJK, V_SLOT)
        rng = ws.Range(ws.Cells(RIJ1, c), ws.Cells(laatste, c)).Address
        ws.Cells(laatste + 2, c).Formula = "=SUMIF($Q$" & RIJ1 & ":$Q$" & laatste & ",""In gebruik""," & rng & ")"
        ws.Cells(laatste + 3, c).Formula = "=SUMIF($Q$" & RIJ1 & ":$Q$" & laatste & ",""Te vervangen*""," & rng & ")"
    Next c
    ws.AutoFilterMode = False
    ws.Range(ws.Cells(4, 1), ws.Cells(laatste, V_STATUS)).AutoFilter
End Sub

Private Sub SorteerVI()
    Dim ws As Worksheet, laatste As Long, c As Long
    Set ws = ThisWorkbook.Worksheets(VI)
    laatste = LaatsteRij(ws)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(RIJ1, I_JAAR), ws.Cells(laatste, I_JAAR)), Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(RIJ1, I_NR), ws.Cells(laatste, I_NR)), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(RIJ1, 1), ws.Cells(laatste, I_OPM))
        .Header = xlNo
        .Apply
    End With
    For c = 6 To 9
        ws.Cells(laatste + 2, c).Formula = "=SUM(" & ws.Range(ws.Cells(RIJ1, c), ws.Cells(laatste, c)).Address(False, False) & ")"
    Next c
    ws.Cells(laatste + 2, I_TEACT).Formula = "=SUM(" & ws.Range(ws.Cells(RIJ1, I_TEACT), ws.Cells(laatste, I_TEACT)).Address(False, False) & ")"
End Sub

Private Sub VernieuwKeuzelijst()
    Dim wsD As Worksheet, ws As Worksheet, r As Long, k As Long
    Set wsD = ThisWorkbook.Worksheets(DASH)
    Set ws = VoertuigenBlad()
    wsD.Range("H2:H60").ClearContents
    k = 2
    For r = RIJ1 To LaatsteRij(ws)
        If Left(CStr(ws.Cells(r, V_STATUS).Value), 12) = "Te vervangen" Then
            wsD.Cells(k, 8).Value = ws.Cells(r, V_NR).Value
            k = k + 1
        End If
    Next r
End Sub

Private Sub MaakKnoppen()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DASH)
    Knop ws, "btnJaar", "E3", "Jaarovergang nu", "JaarovergangHandmatig"
    Knop ws, "btnAanschaf", "E7", "Aanschaf bevestigen", "AanschafBevestigen"
    Knop ws, "btnWijzig", "E10", "Voertuig wijzigen", "VoertuigWijzigen"
    Knop ws, "btnAfvoer", "E12", "Voertuig afvoeren", "VoertuigAfvoeren"
    Knop ws, "btnInvWijzig", "G7", "Investering wijzigen", "InvesteringWijzigen"
    Knop ws, "btnInvToev", "G10", "Investering toevoegen", "InvesteringToevoegen"
    Knop ws, "btnInvSchrap", "G12", "Investering schrappen", "InvesteringSchrappen"
End Sub

Private Sub Knop(ws As Worksheet, n As String, cel As String, tekst As String, actie As String)
    Dim b As Object
    If KnopBestaat(ws, n) Then Exit Sub
    Set b = ws.Buttons.Add(ws.Range(cel).Left, ws.Range(cel).Top, 170, 26)
    b.Name = n: b.Caption = tekst: b.OnAction = actie
End Sub

Private Function KnopBestaat(ws As Worksheet, n As String) As Boolean
    Dim b As Object
    On Error Resume Next
    Set b = ws.Buttons(n)
    KnopBestaat = Not b Is Nothing
    On Error GoTo 0
End Function

' ================= hulpfuncties =================
Private Function VoertuigenBlad() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 11) = "Voertuigen " Then Set VoertuigenBlad = ws: Exit Function
    Next ws
    Err.Raise vbObjectError + 1, , "Blad 'Voertuigen <jaar>' niet gevonden."
End Function

Private Function Naam(n As String) As Range
    Set Naam = ThisWorkbook.Names(n).RefersToRange
End Function

Private Function LopendJaar() As Long
    LopendJaar = CLng(Naam("LopendJaar").Value)
End Function

Private Function DoorschuifVan() As String
    DoorschuifVan = Trim(CStr(Naam("DoorschuifVan").Value))
End Function

Private Function DoorschuifNaar() As String
    DoorschuifNaar = Trim(CStr(Naam("DoorschuifNaar").Value))
End Function

' Laatste gevulde rij in kolom A vanaf rij 5 (stopt bij de eerste lege cel).
Private Function LaatsteRij(ws As Worksheet) As Long
    Dim r As Long
    r = RIJ1
    Do While Len(CStr(ws.Cells(r, 1).Value)) > 0
        r = r + 1
    Loop
    LaatsteRij = r - 1
End Function

' Voegt een hele rij in direct onder de lijst (opmaak van de rij erboven) en geeft het rijnummer terug.
Private Function NieuweRij(ws As Worksheet) As Long
    Dim r As Long
    r = LaatsteRij(ws) + 1
    ws.Rows(r).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    NieuweRij = r
End Function

Private Function ZoekRij(ws As Worksheet, basisNr As String, status As String) As Long
    Dim r As Long
    For r = RIJ1 To LaatsteRij(ws)
        If Basis(CStr(ws.Cells(r, V_NR).Value)) = basisNr And ws.Cells(r, V_STATUS).Value = status Then
            ZoekRij = r: Exit Function
        End If
    Next r
End Function

Private Function ZoekNummer(ws As Worksheet, v As String) As Long
    Dim r As Long
    For r = RIJ1 To LaatsteRij(ws)
        If CStr(ws.Cells(r, V_NR).Value) = v Then ZoekNummer = r: Exit Function
    Next r
End Function

Private Function ZoekVIRij(ws As Worksheet, basisNr As String) As Long
    Dim r As Long
    For r = RIJ1 To LaatsteRij(ws)
        If Basis(CStr(ws.Cells(r, I_NR).Value)) = basisNr Then ZoekVIRij = r: Exit Function
    Next r
End Function

Private Function Basis(v As String) As String
    If InStr(v, "-") > 0 Then Basis = Left(v, InStr(v, "-") - 1) Else Basis = v
End Function

Private Function Volg(v As String) As Long
    If InStr(v, "-") > 0 Then Volg = Val(Mid(v, InStr(v, "-") + 1)) Else Volg = 0
End Function

Private Function Nummer(basisNr As String, n As Long) As String
    Nummer = basisNr & "-" & Format(n, "00")
End Function

Private Function Leeg(w As Variant, vervang As String) As Variant
    If IsError(w) Then
        Leeg = vervang
    ElseIf Len(CStr(w)) = 0 Then
        Leeg = vervang
    Else
        Leeg = w
    End If
End Function
