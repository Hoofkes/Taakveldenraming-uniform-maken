Attribute VB_Name = "Kredieten_8_DashboardAdm"
Option Explicit
Public Const KNOP_B As Long = 158               ' breedte van een knop op de dashboards
Public Const KNOP_STAP As Long = 165            ' afstand tussen knoppen naast elkaar (4 knoppen passen in kolom B tot en met E)
Private Const BL_AFB As String = "Afbeeldingen"      ' onzichtbaar blad met het logo (zie AppLogoInlezen)

' Twee dashboards: 'Dashboard FinAdv' (financieel adviseurs: invoerkaart, blok 2 en controles) en 'Dashboard Adm'
' (financiele administratie: controle van kredieten, kredietoverzichten en dezelfde controles).
' Iedere groep ziet alleen zijn eigen dashboard; de applicatiebeheerder ziet ze allebei.
' Blok 4 (controles) staat op Dashboard FinAdv; Dashboard Adm toont dezelfde cellen via een formule (op dezelfde rijen).
' Overzichten die de financiele administratie maakt, gaan naar eigen mappen: 'FinAdm Kredietoverzicht op naam' en
' 'FinAdm Kredietoverzicht jaarrekening'.

' Het oude blad 'Dashboard' krijgt de nieuwe naam (eenmalig, bij het openen). Formules en namen gaan vanzelf mee.
Public Sub DashboardNaamZorgen()
    On Error Resume Next
    If BladBestaatStil("Dashboard") And Not BladBestaatStil(BL_DASH) Then
        ThisWorkbook.Worksheets("Dashboard").Unprotect WACHTWOORD
        ThisWorkbook.Worksheets("Dashboard").Name = BL_DASH
        LogRegel "dashboard", , , "Bladnaam", "Dashboard", BL_DASH, "nieuwe naam (twee dashboards)"
    End If
End Sub

' Het dashboard van de huidige gebruiker: financiele administratie (zonder recht Toevoegen) = Dashboard Adm, anders FinAdv.
Public Function EigenDashboard() As Worksheet
    On Error Resume Next
    If BladBestaatStil(BL_ADM) And HeeftRecht("Bevestigen") And Not HeeftRecht("Toevoegen") Then
        Set EigenDashboard = ThisWorkbook.Worksheets(BL_ADM)
    Else
        Set EigenDashboard = ThisWorkbook.Worksheets(BL_DASH)
    End If
End Function

Public Sub NaarEigenDashboard()
    On Error Resume Next
    EigenDashboard.Visible = xlSheetVisible
    EigenDashboard.Activate
    EigenDashboard.Range("A1").Select
End Sub

' Zichtbaarheid van de twee dashboards. Zonder rechten: alleen Dashboard FinAdv (alleen lezen).
Public Sub DashboardsZichtbaar()
    Dim adv As Boolean, fa As Boolean, beheer As Boolean, toonAdv As Boolean, toonAdm As Boolean
    On Error Resume Next
    If Not BladBestaatStil(BL_ADM) Or Not BladBestaatStil(BL_DASH) Then Exit Sub
    adv = HeeftRecht("Toevoegen")
    fa = HeeftRecht("Bevestigen")
    beheer = HeeftRecht("Beheer")
    toonAdm = fa Or beheer
    toonAdv = adv Or beheer Or Not toonAdm
    ' eerst tonen wat zichtbaar moet zijn (er moet altijd een blad zichtbaar blijven), dan pas verbergen
    If toonAdv Then ThisWorkbook.Worksheets(BL_DASH).Visible = xlSheetVisible
    If toonAdm Then ThisWorkbook.Worksheets(BL_ADM).Visible = xlSheetVisible
    If ActiveSheet.Name = BL_DASH And Not toonAdv Then ThisWorkbook.Worksheets(BL_ADM).Activate
    If ActiveSheet.Name = BL_ADM And Not toonAdm Then ThisWorkbook.Worksheets(BL_DASH).Activate
    If Not toonAdv Then ThisWorkbook.Worksheets(BL_DASH).Visible = xlSheetHidden
    If Not toonAdm Then ThisWorkbook.Worksheets(BL_ADM).Visible = xlSheetHidden
End Sub

' Bouwt Dashboard Adm (opnieuw) op. Wordt aangeroepen vanuit DashboardKnoppenPlaatsen; kan ook los worden gedraaid.
Public Sub DashboardAdmOpbouwen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, wa As Worksheet, b As Object, r As Long, c As Long, naam As Variant, kopieer As Boolean, actief As Object
    Dim stap As String
    Dim rr As Variant
    On Error GoTo Fout
    stap = "begin"
    Set actief = ActiveSheet
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    Application.ScreenUpdating = False
    stap = "blad Dashboard Adm leegmaken of maken"
    If BladBestaatStil(BL_ADM) Then
        Set wa = ThisWorkbook.Worksheets(BL_ADM)
        wa.Visible = xlSheetVisible
        wa.Unprotect WACHTWOORD
        Do While wa.Shapes.Count > 0
            wa.Shapes(1).Delete
        Loop
        wa.Cells.Clear
        wa.Rows.Hidden = False
    Else
        Set wa = ThisWorkbook.Worksheets.Add(After:=wd)
        wa.Name = BL_ADM
    End If
    wa.Tab.Color = RGB(112, 173, 71)
    wa.Activate
    ActiveWindow.DisplayGridlines = False
    wa.Cells.Font.Name = wd.Cells(1, 1).Font.Name
    wa.Cells.Font.Size = wd.Cells(1, 1).Font.Size
    For c = 1 To 11
        wa.Columns(c).ColumnWidth = wd.Columns(c).ColumnWidth
    Next c

    stap = "kop en controles overnemen"
    ' kop (rij 1 tot en met 4) en blok 4 (controles) overnemen van Dashboard FinAdv, zonder knoppen
    kopieer = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = False
    wd.Range("A1:K4").Copy Destination:=wa.Range("A1")
    wd.Range("A" & CB_KOP & ":K" & (CB_LAATSTE + 2)).Copy Destination:=wa.Range("A" & CB_KOP)
    Application.CopyObjectsWithCells = kopieer
    For r = 1 To 4
        wa.Rows(r).RowHeight = wd.Rows(r).RowHeight
    Next r
    wa.Range("B1").Value = "Dashboard financi" & ChrW(235) & "le administratie"
    On Error Resume Next
    wa.Range("E3:F3").ClearContents                              ' keuzeopties ververst op: alleen voor de financieel adviseur
    wa.Range("E2:H2").ClearContents
    wa.Range("E2:H2").Validation.Delete
    wa.Range("E2:H2").Interior.Pattern = xlNone
    On Error GoTo Fout
    TitelOpmaak wa
    stap = "controles koppelen met formules"
    ' controles: dezelfde cellen als op Dashboard FinAdv (formule), zodat ze altijd gelijk zijn
    For r = CB_KOP To CB_LAATSTE + 2
        wa.Rows(r).RowHeight = wd.Rows(r).RowHeight
        For c = 1 To 11
            If Len(CStr(wd.Cells(r, c).Formula)) > 0 Then
                wa.Cells(r, c).Formula = "=IF('" & BL_DASH & "'!" & wd.Cells(r, c).Address(False, False) & "="""",""""," & _
                                         "'" & BL_DASH & "'!" & wd.Cells(r, c).Address(False, False) & ")"
            End If
        Next c
    Next r
    KopOpmaak wa, CB_KOP, "4. Controles"

    stap = "blok 1 Controle kredieten"
    ' 1. Controle door de financiele administratie (de knoppen die eerst in blok 3 van het Dashboard stonden)
    FinAdmBlokPlaatsen
    stap = "oud blok 3 op Dashboard FinAdv leegmaken"
    ' oud blok 3 op Dashboard FinAdv leegmaken (de rijen worden verborgen door DashboardNaarRol)
    wd.Unprotect WACHTWOORD
    wd.Range("A" & FA_KOP & ":H" & (CB_KOP - 1)).Clear
    On Error Resume Next
    For Each naam In Array("btnFaVernieuwen", "btnFaOpenen", "btnFaBevestigen", "btnFaAfkeuren", "btnFaAanpassen", "btnFaVolgende")
        wd.Shapes(CStr(naam)).Delete
        wd.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    On Error GoTo Fout

    stap = "blok 2 Kredietoverzichten"
    ' 2. Kredietoverzichten
    KopOpmaak wa, ADM_OV_KOP, "2. Kredietoverzichten"
    wa.Range("B" & (ADM_OV_KOP + 1)).Value = "Bestanden komen in de eigen FinAdm-mappen."
    wa.Range("B" & (ADM_OV_KOP + 1)).Font.Italic = True
    wa.Range("B" & (ADM_OV_KOP + 1)).Font.Color = RGB(89, 89, 89)
    wa.Rows(ADM_OV_KOP + 2).RowHeight = 26
    Set b = wa.Buttons.Add(wa.Range("B" & (ADM_OV_KOP + 2)).Left, wa.Range("B" & (ADM_OV_KOP + 2)).Top + 2, KNOP_B, 22)
    b.Name = "btnAdmOverzichtNaam"
    b.Caption = "Kredietoverzicht naar rol"
    b.OnAction = "AdmKredietoverzichtOpNaam"
    Set b = wa.Buttons.Add(wa.Range("B" & (ADM_OV_KOP + 2)).Left + KNOP_STAP, wa.Range("B" & (ADM_OV_KOP + 2)).Top + 2, KNOP_B, 22)
    b.Name = "btnAdmJaarrekening"
    b.Caption = "Kredietoverzicht bij jaarafsluiting"
    b.OnAction = "AdmKredietoverzichtJaarrekening"
    KnopUitleg wa, "btnAdmOverzichtNaam", "Kredietoverzicht naar rol: je kiest eerst 1 = budgetbeheerder, 2 = projectleider of 3 = tijdschrijver " & _
               "en daarna de naam. Het overzicht komt als blad 'Laatste overzicht op naam' en wordt ook als los bestand met " & _
               "datum bewaard in de map 'FinAdm Kredietoverzicht op naam', in de submap van die persoon."
    KnopUitleg wa, "btnAdmJaarrekening", "Kredietoverzicht bij jaarafsluiting: maakt het overzicht met de stand per 31 december, uit de back-up die " & _
               "bij de jaarovergang is gemaakt. Het wordt als los bestand bewaard in de map 'FinAdm Kredietoverzicht jaarrekening'. " & _
               "Bestaat er nog geen back-up, dan kun je een proef zien op basis van de huidige lijst (wordt niet bewaard)."

    stap = "blok 3 Eerstvolgende vrije kredietnummers"
    KredietnummersBlokPlaatsen
    Set wa = ThisWorkbook.Worksheets(BL_ADM)
    wa.Unprotect WACHTWOORD

    stap = "knop Controles vernieuwen"
    ' knop Controles vernieuwen bij blok 3
    Set b = wa.Buttons.Add(wa.Range("E" & (CB_KOP + 1)).Left + 2, wa.Range("E" & (CB_KOP + 1)).Top + 3, KNOP_B, 18)
    b.Name = "btnAdmControles"
    b.Caption = "Controles vernieuwen"
    b.OnAction = "ControlesVernieuwen"
    KnopUitleg wa, "btnAdmControles", "Controles vernieuwen: rekent de controles opnieuw uit (en werkt de vrije kredietnummers bij). Het zijn dezelfde controles als op het " & _
               "Dashboard van de financieel adviseurs. Er wordt niets gewijzigd in de kredieten."

    stap = "rijen verbergen"
    ' lege rijen tussen blok 2 en de controles verbergen
    ' lege rijen weg: tussen de blokken een smalle rij, na de nummertabel alles verbergen tot de controles
    For Each rr In Array(5, ADM_OV_KOP - 1, ADM_NR_KOP - 1, CB_KOP - 1)
        wa.Rows(CLng(rr)).RowHeight = 8
    Next rr
    r = ADM_NR_KOP + 3
    Do While Len(CStr(wa.Cells(r, 4).Value)) > 0
        r = r + 1
    Loop
    wa.Rows((r + 1) & ":" & (CB_KOP - 2)).Hidden = True        ' rij r: knop Kredietnummer aanpassen
    stap = "tabbladkiezer en Krediet raadplegen"
    ' knoppen bovenaan: tabbladkiezer en krediet raadplegen
    TabbladkiezerWeg wa
    RaadplegenKnopPlaatsen BL_ADM
    AppLogoPlaatsen wa, wa.Range("E1").Left + wa.Range("E1").Width, 44, 3, True   ' logo rechtsboven, binnen kolom E
    wa.Unprotect WACHTWOORD
    Set wa = ThisWorkbook.Worksheets(BL_ADM)
    wa.Unprotect WACHTWOORD
    wa.Range("A1").Select
    BladBeveiligen wa
    If Not actief Is Nothing Then actief.Activate
    Application.ScreenUpdating = True
    Exit Sub
Fout:
    Application.CopyObjectsWithCells = True
    Application.ScreenUpdating = True
    MsgBox "Dashboard Adm kon niet worden opgebouwd." & vbCrLf & vbCrLf & "Stap: " & stap & vbCrLf & _
           "Fout " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Stuur deze melding door.", vbExclamation, BL_ADM
End Sub

' Knoppen op Dashboard Adm: zelfde overzichten als de financieel adviseur, maar in eigen mappen (FinAdm ...).
Public Sub AdmKredietoverzichtOpNaam()
    mFinAdm = True
    On Error GoTo Klaar
    KredietoverzichtOpNaam
Klaar:
    mFinAdm = False
End Sub

Public Sub AdmKredietoverzichtJaarrekening()
    mFinAdm = True
    On Error GoTo Klaar
    KredietoverzichtJaarrekening
Klaar:
    mFinAdm = False
End Sub

' Map voor een overzicht: bij de financiele administratie (knoppen op Dashboard Adm) een eigen map.
Public Function OverzichtMap(ByVal standaard As String, ByVal finAdm As String) As String
    OverzichtMap = MapStambestand(IIf(mFinAdm, finAdm, standaard))
End Function

' ======================================================================
' Indeling Dashboard FinAdv (november 2026)
' ======================================================================
'   1. Applicatiebeheer (rij 6-10, alleen applicatiebeheerder)   2. Kredieten verwerken en raadplegen (rij 11-17)
'   invoerkaart (rij 18-48, alleen zichtbaar na 'Krediet invoeren' of 'Krediet wijzigen')
'   3. Kredietoverzichten   4. Uitvraag kadernota/begroting   5. Controles
' Het oude Dashboard had de invoerkaart direct onder de kop (rij 7). Eenmalig worden boven de kaart 12 rijen
' ingevoegd (hele rijen: formules, namen en opmaak gaan mee); daarna worden alle knoppen opnieuw geplaatst.
Public Sub DashboardRijenInvoegen()
    Dim wd As Worksheet
    On Error GoTo Fout
    If Not BladBestaatStil(BL_DASH) Then Exit Sub
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    If Not (CStr(wd.Range("B7").Value) Like "*Invoerkaart*") Then Exit Sub          ' al omgezet
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    wd.Unprotect WACHTWOORD
    Application.EnableEvents = False
    Do While wd.Shapes.Count > 0                 ' alle knoppen weg: ze worden op de nieuwe plek opnieuw gemaakt
        wd.Shapes(1).Delete
    Loop
    wd.Rows("6:" & (5 + DASH_EXTRA)).Insert Shift:=xlDown
    wd.Rows("6:" & (5 + DASH_EXTRA)).Hidden = False
    wd.Range("B" & (eerste - 3)).Value = "Invoerkaart kredieten"
    Application.EnableEvents = True
    LogRegel "dashboard", , , "Indeling", "invoerkaart bovenaan", "applicatiebeheer, kredieten, overzichten, uitvraag, controles"
    Exit Sub
Fout:
    Application.EnableEvents = True
    MsgBox "De nieuwe indeling van Dashboard FinAdv kon niet worden gemaakt: " & Err.Description, vbExclamation, BL_DASH
End Sub

' Staat de invoerkaart open (na Krediet invoeren of Krediet wijzigen)?
Public Function KaartOpen() As Boolean
    On Error Resume Next
    KaartOpen = (CStr(Evaluate(ThisWorkbook.Names("Par_KaartOpen").RefersTo)) = "1")
End Function

' Invoerkaart tonen of verbergen (de rijen gaan open of dicht via DashboardNaarRol).
Public Sub KaartTonen(ByVal tonen As Boolean)
    ThisWorkbook.Names.Add Name:="Par_KaartOpen", RefersTo:="=" & IIf(tonen, "1", "0"), Visible:=False
    DashboardNaarRol
    On Error Resume Next
    If ActiveSheet.Name = BL_DASH Then Application.GoTo ActiveSheet.Range("A1"), True   ' bovenaan beginnen
End Sub

' Knop Krediet invoeren: de invoerkaart komt tevoorschijn.
Public Sub KredietInvoeren()
    Dim wd As Worksheet
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs voeren kredieten in.", vbExclamation
        Exit Sub
    End If
    KaartTonen True
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    wd.Activate
    Application.GoTo wd.Range("A1"), True
    wd.Cells(eerste, 3).Select
End Sub

' Knop Terug naar het dashboard onder de invoerkaart (naast Scherm leegmaken).
Public Sub KaartSluitenKnopPlaatsen()
    Dim ws As Worksheet, b As Object, verborgen As Boolean, beveiligd As Boolean
    Set ws = ThisWorkbook.Worksheets(BL_DASH)
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    On Error Resume Next
    ws.Shapes("btnKaartSluiten").Delete
    On Error GoTo 0
    verborgen = ws.Rows(KAART_KNOP).Hidden
    ws.Rows(KAART_KNOP).Hidden = False                       ' op een zichtbare rij plaatsen, anders komt de knop verkeerd
    Set b = ws.Buttons.Add(ws.Range("B" & KAART_KNOP).Left + 430, ws.Range("B" & KAART_KNOP).Top + 8, 180, 22)
    b.Name = "btnKaartSluiten"
    b.Caption = "Terug naar het dashboard"
    b.OnAction = "KaartSluiten"
    b.Placement = xlMove
    b.PrintObject = False
    ws.Rows(KAART_KNOP).Hidden = verborgen
    b.Visible = Not verborgen
    If beveiligd Then BladBeveiligen ws
End Sub

' Knop Terug naar het dashboard (onder de invoerkaart): kaart dicht, alle onderdelen weer zichtbaar.
Public Sub KaartSluiten()
    Dim wd As Worksheet
    KaartTonen False
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    wd.Activate
    Application.GoTo wd.Range("A1"), True
End Sub

' Met de hand te starten (Alt+F8) als de nieuwe indeling na het openen niet vanzelf is gemaakt.
Public Sub DashboardNieuweIndeling()
    Dim wd As Worksheet
    DashboardNaamZorgen
    DashboardRijenInvoegen
    ThisWorkbook.Names.Add Name:="Par_KaartOpen", RefersTo:="=0", Visible:=False
    KaartIndelingProject
    If BladBestaatStil(BL_MELD) Then                 ' oud blad Meldingen uitgaven weghalen (controles staan op het dashboard)
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(BL_MELD).Unprotect WACHTWOORD
        ThisWorkbook.Worksheets(BL_MELD).Delete
        Application.DisplayAlerts = True
    End If
    DashboardKnoppenPlaatsen
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    MsgBox "Klaar." & vbCrLf & "Rij " & AB_KOP & ": " & CStr(wd.Range("B" & AB_KOP).Value) & vbCrLf & _
           "Rij " & KR_KOP & ": " & CStr(wd.Range("B" & KR_KOP).Value) & vbCrLf & _
           "Rij " & NM_KOP & ": " & CStr(wd.Range("B" & NM_KOP).Value) & vbCrLf & _
           "Rij " & (eerste - 3) & ": " & CStr(wd.Range("B" & (eerste - 3)).Value), vbInformation, BL_DASH
End Sub

' ======================================================================
' Logo HoofDidThis (doorzichtige achtergrond)
' ======================================================================
' Het logo staat eenmalig in de werkmap, op het onzichtbare blad 'Afbeeldingen' (vorm 'AppLogo'), zodat het geen
' bestand in een map nodig heeft. Inlezen: macro AppLogoInlezen (Alt+F8) en het bestand HoofDidThis_logo.png kiezen.
' Daarna komt het rechtsboven op Dashboard FinAdv, Dashboard Adm en het blad Krediet raadplegen.

Public Sub AppLogoInlezen()
    Dim f As Variant, wa As Worksheet, shp As Object
    f = Application.GetOpenFilename("Afbeelding (*.png),*.png", , "Kies het logo (HoofDidThis_logo.png)")
    If VarType(f) = vbBoolean Then Exit Sub
    On Error GoTo Fout
    If Not BladBestaatStil(BL_AFB) Then
        Set wa = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wa.Name = BL_AFB
    Else
        Set wa = ThisWorkbook.Worksheets(BL_AFB)
        wa.Visible = xlSheetVisible
        On Error Resume Next
        wa.Shapes("AppLogo").Delete
        On Error GoTo Fout
    End If
    Set shp = wa.Shapes.AddPicture(CStr(f), msoFalse, msoTrue, 10, 10, -1, -1)   ' ingesloten, niet gekoppeld
    shp.Name = "AppLogo"
    wa.Visible = xlSheetVeryHidden
    AppLogoOveral
    LogRegel "logo", , , "AppLogo", "", CStr(f), "logo ingelezen"
    MsgBox "Het logo staat nu in de werkmap en op de dashboards.", vbInformation, "Logo"
    Exit Sub
Fout:
    MsgBox "Het logo kon niet worden ingelezen: " & Err.Description, vbExclamation, "Logo"
End Sub

' Logo op beide dashboards zetten (na het opnieuw plaatsen van de knoppen).
Public Sub AppLogoOveral()
    Dim w As Worksheet, shp As Object, beveiligd As Boolean
    On Error Resume Next
    For Each w In ThisWorkbook.Worksheets                ' eerst losse kopieen op andere bladen opruimen
        If w.Name <> BL_AFB And w.Name <> BL_DASH And w.Name <> BL_ADM And w.Name <> BL_RAADPLEGEN Then
            Set shp = Nothing
            Set shp = w.Shapes("AppLogo")
            If Not shp Is Nothing Then
                beveiligd = w.ProtectContents
                If beveiligd Then w.Unprotect WACHTWOORD
                shp.Delete
                If beveiligd Then BladBeveiligen w
            End If
        End If
    Next w
    On Error GoTo 0
    Dim blad As Variant
    For Each blad In Array(BL_DASH, BL_ADM)                 ' rechtsboven, rechts uitgelijnd op kolom E, klein (44 punten)
        If BladBestaatStil(CStr(blad)) Then
            With ThisWorkbook.Worksheets(CStr(blad))
                AppLogoPlaatsen ThisWorkbook.Worksheets(CStr(blad)), .Range("E1").Left + .Range("E1").Width, 44, 3, True
            End With
        End If
    Next blad
End Sub

' Logo direct onder de knop Ga naar tabblad (zelfde linkerkant). Geen knop: rechts van de kop.
Public Sub LogoOnderGaNaar(ByVal ws As Worksheet, ByVal knop As String)
    Dim b As Object
    On Error Resume Next
    Set b = ws.Shapes(knop)
    On Error GoTo 0
    If b Is Nothing Then
        AppLogoPlaatsen ws, RechtsVanKop(ws), 60
    Else
        AppLogoPlaatsen ws, b.Left, 60, b.Top + b.Height + 4
    End If
End Sub

' Plek voor het logo: rechts van alles wat in rij 1 tot en met 4 staat (knoppen en tekst), met wat ruimte ertussen.
Public Function RechtsVanKop(ByVal ws As Worksheet) As Double
    Dim shp As Object, c As Long, r As Long, x As Double
    x = ws.Range("H1").Left + ws.Range("H1").Width
    For Each shp In ws.Shapes
        If shp.Name <> "AppLogo" And shp.Top < ws.Rows(5).Top Then
            If shp.Left + shp.Width > x Then x = shp.Left + shp.Width
        End If
    Next shp
    For r = 1 To 4
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(r, c + 1).Left > x Then x = ws.Cells(r, c + 1).Left
    Next r
    RechtsVanKop = x + 20
End Function

' Zet een kopie van het logo op een blad (links = positie in punten, hoogte in punten). Geen logo ingelezen: niets.
Public Sub AppLogoPlaatsen(ByVal ws As Worksheet, ByVal links As Double, ByVal hoogte As Double, Optional ByVal boven As Double = 3, _
                           Optional ByVal rechtsUitlijnen As Boolean = False)
    Dim src As Object, shp As Object, actief As Object, scherm As Boolean, beveiligd As Boolean, zicht As Long
    On Error GoTo klaar
    If Not BladBestaatStil(BL_AFB) Then Exit Sub
    Set src = ThisWorkbook.Worksheets(BL_AFB).Shapes("AppLogo")
    Set actief = ActiveSheet
    scherm = Application.ScreenUpdating
    Application.ScreenUpdating = False
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    On Error Resume Next
    ws.Shapes("AppLogo").Delete
    On Error GoTo klaar
    zicht = ws.Visible
    ws.Visible = xlSheetVisible
    src.Copy
    ws.Activate
    ws.Range("A1").Select
    ws.Paste
    Set shp = ws.Shapes(ws.Shapes.Count)
    If shp.Type <> msoPicture Then GoTo klaar          ' plakken mislukt: niets doen
    shp.Name = "AppLogo"
    shp.ScaleHeight 1, msoTrue                          ' eerst terug naar de echte maat, dan pas verkleinen
    shp.ScaleWidth 1, msoTrue
    shp.LockAspectRatio = msoTrue
    shp.Height = hoogte
    shp.Top = boven
    shp.Left = IIf(rechtsUitlijnen, links - shp.Width, links)   ' rechts uitlijnen: links = rechterkant
    shp.Placement = xlFreeFloating               ' blijft bovenaan staan (rij 1 tot en met 5 gaan nooit dicht)
    shp.OnAction = ""
    Application.CutCopyMode = False
    If Not actief Is Nothing Then actief.Activate
    ws.Visible = zicht
    If beveiligd Then BladBeveiligen ws
klaar:
    Application.CutCopyMode = False
    Application.ScreenUpdating = scherm
End Sub

' Eenheidsopmaak van een kop op de dashboards (blok 1, 2, 3 ...): donkerblauwe balk van B tot en met G, witte vette
' letters, maat 12. Zelfde lettertype als de rest van het blad.
Public Sub KopOpmaak(ByVal ws As Worksheet, ByVal rij As Long, ByVal tekst As String)
    ws.Range("B" & rij).Value = tekst
    With ws.Range("B" & rij & ":G" & rij)
        .Interior.Color = RGB(0, 73, 118)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Italic = False
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
    End With
    ws.Rows(rij).RowHeight = 22
End Sub

' Titel in B1 van een dashboard.
Public Sub TitelOpmaak(ws As Worksheet)
    With ws.Range("B1").Font
        .Bold = True
        .Size = 16
        .Color = RGB(0, 73, 118)
    End With
    ws.Rows(1).RowHeight = 28
End Sub

' Smaller dashboard (afspraak november 2026): geen tabbladkiezer meer bovenaan (Naar tabblad, keuzelijst en knop),
' kolom F weer smal en kolom G als smalle rand. Alles past dan binnen kolom A tot en met F.
Public Sub TabbladkiezerWeg(ByVal ws As Worksheet)
    Dim naam As Variant, beveiligd As Boolean
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    On Error Resume Next
    For Each naam In Array("btnNaarTabblad", "btnAdmNaarTabblad")
        ws.Shapes(CStr(naam)).Delete
        ws.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    ws.Range("E2:H2").Validation.Delete
    On Error GoTo 0
    ws.Range("E2:H2").ClearContents
    ws.Range("E2:H2").Interior.Pattern = xlNone
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 3
    If beveiligd Then BladBeveiligen ws
End Sub
