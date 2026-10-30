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
Private Const I_ACTJR As Long = 13, I_EXPL As Long = 15, I_OPM As Long = 16

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
        wsV.Cells(nr, V_LEV).Value = Naam("p_Lev").Value
        wsV.Cells(nr, V_AANB).Value = Naam("p_Aanbet").Value
        wsV.Cells(nr, V_BIJK).Value = Naam("p_Belet").Value + Naam("p_Opbouw").Value
        wsV.Cells(nr, V_SLOT).Value = Naam("p_Slot").Value
        wsV.Cells(nr, V_LEASE).Value = Naam("p_Lease").Value
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
    Dim ws As Worksheet, b As Object
    Set ws = ThisWorkbook.Worksheets(DASH)
    If Not KnopBestaat(ws, "btnAanschaf") Then
        Set b = ws.Buttons.Add(ws.Range("E7").Left, ws.Range("E7").Top, 170, 26)
        b.Name = "btnAanschaf": b.Caption = "Aanschaf bevestigen": b.OnAction = "AanschafBevestigen"
    End If
    If Not KnopBestaat(ws, "btnJaar") Then
        Set b = ws.Buttons.Add(ws.Range("E3").Left, ws.Range("E3").Top, 170, 26)
        b.Name = "btnJaar": b.Caption = "Jaarovergang nu": b.OnAction = "JaarovergangHandmatig"
    End If
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
