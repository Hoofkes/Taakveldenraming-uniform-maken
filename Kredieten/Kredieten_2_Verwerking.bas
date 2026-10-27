Attribute VB_Name = "Kredieten_2_Verwerking"
Option Explicit

Private mBlok As Boolean          ' foutwaarde in de verdeelsleutel: werkelijke uitgaven niet wegschrijven

' Vervolg van de macro's voor het Dashboard investeringskredieten (gesplitst vanwege de grootte; declaraties in Kredieten_1_Basis).

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
    Dim Fout As Boolean, nr As Long

    On Error GoTo Fout
    Set wd = ThisWorkbook.Worksheets("Dashboard")
    Set cel = wd.Cells(KaartRij(14), 3)
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
' Zoekt een krediet met een zoekterm: kredietnummer, activanummer of een deel van de omschrijving (alles in een vak).
' Meerdere treffers: keuzelijst. Geeft de rij in de verzamellijst, 0 = niets gekozen.
Public Function ZoekKredietRij(ByVal titel As String) As Long
    Dim wv As Worksheet, zoek As String, r As Long, laatste As Long, n As Long, kk As Long, antw As String, lijst As String
    Dim treffers(1 To 15) As Long, meer As Boolean, hit As Boolean
    Set wv = Vz()
    zoek = Trim$(InputBox("Typ een kredietnummer, een activanummer of een deel van de omschrijving:", titel))
    If Len(zoek) = 0 Then Exit Function
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        If Not IsEmpty(wv.Cells(r, 1).Value) Then
            hit = (NrTekst(wv.Cells(r, 3).Value) = zoek) Or (Trim$(CStr(wv.Cells(r, 24).Value)) = zoek) Or _
                  (InStr(1, CStr(wv.Cells(r, 4).Value), zoek, vbTextCompare) > 0)
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
        lijst = lijst & kk & " = " & NrTekst(wv.Cells(r, 3).Value) & " | " & Left$(CStr(wv.Cells(r, 4).Value), 30) & " | " & _
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
    Dim cm As Object, i As Long, tekst As String, heeftChange As Boolean, heeftLink As Boolean, anders As String
    On Error GoTo Niet
    Set cm = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.Worksheets("Dashboard").CodeName).CodeModule
    For i = 1 To cm.CountOfLines
        tekst = cm.Lines(i, 1)
        If InStr(1, tekst, "KaartGewijzigd", vbTextCompare) > 0 Then heeftChange = True
        If InStr(1, tekst, "FaHyperlink", vbTextCompare) > 0 Then heeftLink = True
    Next i
    For i = 1 To cm.CountOfLines
        tekst = cm.Lines(i, 1)
        If Not heeftChange And InStr(1, tekst, "Sub Worksheet_Change", vbTextCompare) > 0 Then anders = anders & vbCrLf & "Worksheet_Change: KaartGewijzigd Target": heeftChange = True
        If Not heeftLink And InStr(1, tekst, "Sub Worksheet_FollowHyperlink", vbTextCompare) > 0 Then anders = anders & vbCrLf & "Worksheet_FollowHyperlink: FaHyperlink Target": heeftLink = True
    Next i
    If Not heeftChange Then cm.AddFromString "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                                             "    KaartGewijzigd Target" & vbCrLf & "End Sub"
    If Not heeftLink Then cm.AddFromString "Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)" & vbCrLf & _
                                           "    FaHyperlink Target" & vbCrLf & "End Sub"
    If Len(anders) > 0 And Not stil Then MsgBox "Op het blad Dashboard staat al eigen code. Zet daarin zelf de regel(s):" & anders, vbExclamation
    Exit Sub
Niet:
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
    On Error GoTo Klaar
    Set wd = Target.Worksheet
    If Intersect(Target, wd.Range("Kaart_Invoer")) Is Nothing Then Exit Sub
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then Exit Sub        ' wijzigen van een bestaand krediet: mag
    If HeeftRecht("Toevoegen") Then Exit Sub
    Application.EnableEvents = False
    Application.Undo
    Application.EnableEvents = True
    MsgBox "U mag geen nieuwe kredieten toevoegen; dat doet de financieel adviseur." & vbCrLf & vbCrLf & _
           "Wilt u een bestaand krediet aanpassen? Klik dan eerst op 'Krediet wijzigen' en zoek het krediet op.", _
           vbExclamation, "Nieuw krediet"
Klaar:
    Application.EnableEvents = True
End Sub

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
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    If Len(CStr(wd.Range("Kaart_Rijnummer").Value)) > 0 Then
        r = CLng(wd.Range("Kaart_Rijnummer").Value)          ' krediet staat al op de kaart
    Else
        r = ZoekKredietRij("Krediet verwijderen")             ' in een keer: zoeken op kredietnummer, activanummer of omschrijving
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
    If IsGetal(wv.Cells(r, 10).Value) Then
        grpTxt = vbCrLf & "Let op: dit is een deelkrediet van groep " & CLng(wv.Cells(r, 10).Value) & ". De verdeelsleutel van de andere deelkredieten verandert."
    End If
    If MsgBox("Volgnummer " & volgnr & ": " & wv.Cells(r, 4).Value & vbCrLf & rijen.Count & " regel(s), inclusief bijdrageregels." & grpTxt & vbCrLf & vbCrLf & _
              "Helemaal verwijderen? Alleen het Logboek houdt de gegevens vast. Dit is alleen bedoeld voor echte fouten.", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Krediet verwijderen") <> vbYes Then Exit Sub
    reden = "verwijderd via het Dashboard door " & CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    If Not RegelsWeghalen(wv, rijen, reden) Then Exit Sub
    KeuzelijstenVerversen True
    KaartWissen
    ControleLijstVernieuwen
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
        ControleLijstVernieuwen
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
    Dim rv As Variant, grpBedrag As Object, g As Variant, vAD As Variant, rij0 As Long, bew() As Variant, nb As Long
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
                    nb = nb + 1
                    bew(nb, 1) = nrs(i): bew(nb, 2) = lj - y: bew(nb, 3) = Round(bU(i, y) + bI(i, y) + bO(i, y), 2)
                End If
            End If
        Next y
    Next i
    VorigeUploadSchrijven bew, nb, Dir(pad)

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
Public Sub Meld(ByVal prio As Long, ByVal soort As String, ByVal jaar As Long, ByVal nr As String, ByVal oms As String, _
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
' KredietoverzichtJaarrekening  : stand per 31-12 vorig jaar uit het blad 'Verzamellijst 31-12-<jaar>' (ontstaat bij de jaarovergang):
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
    rol = Trim$(CStr(wd.Range("C" & NM_ROL).Value))
    naam = Trim$(CStr(wd.Range("C" & NM_NAAM).Value))
    Select Case LCase$(rol)
        Case "budgetbeheerder": kol = 15
        Case "projectleider": kol = 16
        Case "tijdschrijver": kol = 17
        Case Else
            MsgBox "Kies op het Dashboard in cel C" & NM_ROL & " de rol: Budgetbeheerder, Projectleider of Tijdschrijver.", vbExclamation
            Exit Sub
    End Select
    If Len(naam) = 0 Then
        MsgBox "Kies op het Dashboard in cel C" & NM_NAAM & " een naam.", vbExclamation
        Exit Sub
    End If
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
Private Function IsHoofddeel(ByRef v As Variant, ByVal r As Long) As Boolean
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

' Blok 3 op het Dashboard: keuze rol en naam en de knoppen voor de overzichten.
Public Sub UitdraaiBlokPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, naam As Variant, i As Long, titels As Variant, acties As Variant, namen As Variant, rijOff As Variant, kolOff As Variant
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("A" & NM_KOP & ":H" & (FA_KOP - 1)).Clear
    ' blok 2: financieel adviseur (overzicht op naam en alle knoppen van de financieel adviseur)
    ws.Range("B" & NM_KOP).Value = "2. Financieel adviseur"
    ws.Range("B" & NM_KOP).Font.Bold = True
    ws.Range("B" & NM_KOP).Font.Size = 12
    ws.Range("B" & NM_ROL).Value = "Overzicht op naam: rol"
    ws.Range("B" & NM_NAAM).Value = "Overzicht op naam: naam"
    With ws.Range("C" & NM_ROL)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Budgetbeheerder,Projectleider,Tijdschrijver"
        If Leeg(.Value) Then .Value = "Budgetbeheerder"
        .Interior.Color = RGB(255, 255, 204)
    End With
    With ws.Range("C" & NM_NAAM)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(""Lst_""&$C$" & NM_ROL & ")"
        .Interior.Color = RGB(255, 255, 204)
    End With
    On Error Resume Next
    For Each naam In Array("btnOverzichtRaad", "btnOverzichtNaam", "btnOverzichtJaarrekening", "btnAlleBudgethouders", "btnTijdschrijven", _
                           "btnUitvraag", "btnActualisatie", "btnJaarovergang", "btnUitgaven", "btnAfgekeurd", "btnHercontrole", _
                           "btnBhVervangen", "btnPlVervangen", "btnTsVervangen", "btnUvStand", "btnUvControle")
        ws.Shapes(CStr(naam)).Delete
        ws.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    On Error GoTo 0
    For i = NM_NAAM + 1 To ADV_KNOP + 3
        ws.Rows(i).RowHeight = 26
    Next i
    ' rij onder de naam: overzicht op naam en uitvraag verwerken; daaronder vier rijen knoppen
    namen = Array("btnOverzichtNaam", "btnActualisatie", "btnUvStand", "btnUvControle", _
                  "btnUitgaven", "btnOverzichtRaad", "btnOverzichtJaarrekening", "btnUitvraag", _
                  "btnAlleBudgethouders", "btnTijdschrijven", "btnJaarovergang", "btnAfgekeurd", _
                  "btnHercontrole", _
                  "btnBhVervangen", "btnPlVervangen", "btnTsVervangen")
    titels = Array("Kredietoverzicht op naam", "Uitvraag verwerken", "Stand uitvraag", "Begroting vastgesteld", _
                   "Verwerken werkelijk", "Kredietoverzicht raad", "Kredietoverzicht jaarrekening", "Uitvraag kadernota/begroting", _
                   "Overzichten alle budgethouders", "Overzicht tijdschrijven", "Jaarovergang uitvoeren", "Afgekeurd krediet herstellen", _
                   "Aanbieden voor hercontrole", _
                   "Budgethouder vervangen", "Projectleider vervangen", "Tijdschrijver vervangen")
    acties = Array("KredietoverzichtOpNaam", "UitvraagInlezen", "UitvraagStand", "UitvraagNaarControle", _
                   "UitgavenBijwerken", "KredietoverzichtMaken", "KredietoverzichtJaarrekening", "UitvraagBegroting", _
                   "KredietoverzichtAlleBudgethouders", "OverzichtTijdschrijven", "JaarovergangUitvoeren", "AfgekeurdHerstellen", _
                   "HercontroleAanbieden", _
                   "BudgethouderVervangen", "ProjectleiderVervangen", "TijdschrijverVervangen")
    rijOff = Array(NM_NAAM + 1, NM_NAAM + 1, NM_NAAM + 1, NM_NAAM + 1, ADV_KNOP, ADV_KNOP, ADV_KNOP, ADV_KNOP, ADV_KNOP + 1, ADV_KNOP + 1, ADV_KNOP + 1, ADV_KNOP + 1, _
                   ADV_KNOP + 2, ADV_KNOP + 3, ADV_KNOP + 3, ADV_KNOP + 3)
    kolOff = Array(0, 1, 2, 3, 0, 1, 2, 3, 0, 1, 2, 3, 0, 0, 1, 2)
    For i = 0 To UBound(namen)
        Set b = ws.Buttons.Add(ws.Range("B1").Left + kolOff(i) * 205, ws.Range("B" & rijOff(i)).Top + 2, 195, 22)
        b.Name = namen(i)
        b.Caption = titels(i)
        b.OnAction = acties(i)
    Next i
End Sub

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
    Dim wb As Workbook, t As Variant, fout As String
    If Len(map) = 0 Then Exit Function           ' map kon niet worden aangemaakt (melding is al gegeven)
    For Each t In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
        bestand = Replace(bestand, CStr(t), "-")
    Next t
    On Error GoTo Mislukt
    ws.Copy
    Set wb = ActiveWorkbook
    ' knoppen (Terug naar dashboard) horen niet in het losse bestand
    On Error Resume Next
    wb.Worksheets(1).Unprotect WACHTWOORD
    wb.Worksheets(1).Buttons.Delete
    ' kolom Melding is alleen voor financien: niet in het opgeslagen of verzonden bestand
    If KolomMelding(wb.Worksheets(1)) > 0 Then wb.Worksheets(1).Columns(KolomMelding(wb.Worksheets(1))).Delete
    If ws.ProtectContents Then wb.Worksheets(1).Protect Password:=WACHTWOORD, AllowFiltering:=True, AllowFormattingRows:=True
    On Error GoTo Mislukt
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=map & bestand & ".xlsx", FileFormat:=51
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    BladOpslaan = True
    Exit Function

Mislukt:
    fout = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
    If Not mStil Then MsgBox "Opslaan is niet gelukt:" & vbCrLf & map & bestand & ".xlsx" & vbCrLf & vbCrLf & fout & vbCrLf & vbCrLf & _
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

Private Function RolNaam(ByVal kol As Long) As String
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
    Application.EnableEvents = True
    On Error Resume Next
    For Each naam In Array(BL_VZ, "Logboek", "Vorige upload", "Stand tijdschrijven", BL_MUTREG, BL_UITVRAAG, "Parameters", "Dashboard", "Controle", _
                           BL_MELD, BL_OVERZICHT, BL_OVERZICHT_NAAM, BL_KENG, BL_UVSTAND)
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
            ws.Range("C" & NM_ROL & ":C" & NM_NAAM).Locked = False
            ws.Range("G2").Locked = False                         ' tabbladkiezer
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
Public Sub BladOpmaken(ByVal ws As Worksheet, ByVal kopRij As Long, ByVal van As Long, ByVal tot As Long)
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

