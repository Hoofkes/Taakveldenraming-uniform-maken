Attribute VB_Name = "Kredieten_8_DashboardAdm"
Option Explicit

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
    On Error GoTo Fout
    Set actief = ActiveSheet
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    Application.ScreenUpdating = False
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
    wa.Range("E2:F2,E3:F4").ClearContents                         ' alleen voor de financieel adviseur
    wa.Range("F2").Value = "Naar tabblad:"
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
    wa.Range("B" & CB_KOP).Value = "3. Controles"

    ' 1. Controle door de financiele administratie (de knoppen die eerst in blok 3 van het Dashboard stonden)
    FinAdmBlokPlaatsen
    ' oud blok 3 op Dashboard FinAdv leegmaken (de rijen worden verborgen door DashboardNaarRol)
    wd.Unprotect WACHTWOORD
    wd.Range("A" & FA_KOP & ":H" & (CB_KOP - 1)).Clear
    On Error Resume Next
    For Each naam In Array("btnFaVernieuwen", "btnFaOpenen", "btnFaBevestigen", "btnFaAfkeuren", "btnFaAanpassen", "btnFaVolgende")
        wd.Shapes(CStr(naam)).Delete
        wd.Shapes("uitleg_" & CStr(naam)).Delete
    Next naam
    On Error GoTo Fout

    ' 2. Kredietoverzichten
    wa.Range("B" & ADM_OV_KOP).Value = "2. Kredietoverzichten"
    wa.Range("B" & ADM_OV_KOP).Font.Bold = True
    wa.Range("B" & ADM_OV_KOP).Font.Size = 12
    wa.Range("B" & (ADM_OV_KOP + 1)).Value = "Bestanden gaan naar de mappen 'FinAdm Kredietoverzicht op naam' en 'FinAdm Kredietoverzicht jaarrekening'."
    wa.Range("B" & (ADM_OV_KOP + 1)).Font.Italic = True
    wa.Range("B" & (ADM_OV_KOP + 1)).Font.Color = RGB(89, 89, 89)
    wa.Rows(ADM_OV_KOP + 2).RowHeight = 26
    Set b = wa.Buttons.Add(wa.Range("B" & (ADM_OV_KOP + 2)).Left, wa.Range("B" & (ADM_OV_KOP + 2)).Top + 2, 210, 22)
    b.Name = "btnAdmOverzichtNaam"
    b.Caption = "Kredietoverzicht op naam"
    b.OnAction = "AdmKredietoverzichtOpNaam"
    Set b = wa.Buttons.Add(wa.Range("B" & (ADM_OV_KOP + 2)).Left + 220, wa.Range("B" & (ADM_OV_KOP + 2)).Top + 2, 210, 22)
    b.Name = "btnAdmJaarrekening"
    b.Caption = "Kredietoverzicht jaarrekening"
    b.OnAction = "AdmKredietoverzichtJaarrekening"
    KnopUitleg wa, "btnAdmOverzichtNaam", "Kredietoverzicht op naam: je kiest eerst 1 = budgetbeheerder, 2 = projectleider of 3 = tijdschrijver " & _
               "en daarna de naam. Het overzicht komt als blad 'Laatste overzicht op naam' en wordt ook als los bestand met " & _
               "datum bewaard in de map 'FinAdm Kredietoverzicht op naam', in de submap van die persoon."
    KnopUitleg wa, "btnAdmJaarrekening", "Kredietoverzicht jaarrekening: maakt het overzicht met de stand per 31 december, uit de back-up die " & _
               "bij de jaarovergang is gemaakt. Het wordt als los bestand bewaard in de map 'FinAdm Kredietoverzicht jaarrekening'. " & _
               "Bestaat er nog geen back-up, dan kun je een proef zien op basis van de huidige lijst (wordt niet bewaard)."

    ' knop Controles vernieuwen bij blok 3
    Set b = wa.Buttons.Add(wa.Range("E" & (CB_KOP + 1)).Left + 2, wa.Range("E" & (CB_KOP + 1)).Top + 3, 170, 18)
    b.Name = "btnAdmControles"
    b.Caption = "Controles vernieuwen"
    b.OnAction = "ControlesVernieuwen"
    KnopUitleg wa, "btnAdmControles", "Controles vernieuwen: rekent de controles opnieuw uit. Het zijn dezelfde controles als op het " & _
               "Dashboard van de financieel adviseurs. Er wordt niets gewijzigd in de kredieten."

    ' lege rijen tussen blok 2 en de controles verbergen
    wa.Rows((ADM_OV_KOP + 4) & ":" & (CB_KOP - 2)).Hidden = True
    ' knoppen bovenaan: tabbladkiezer en krediet raadplegen
    TabbladkiezerPlaatsen BL_ADM
    RaadplegenKnopPlaatsen BL_ADM
    BeheerKnoppenPlaatsen                        ' knoppen applicatiebeheerder (alleen zichtbaar voor de applicatiebeheerder)
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
    MsgBox "Dashboard Adm kon niet worden opgebouwd: " & Err.Description, vbExclamation, BL_ADM
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
