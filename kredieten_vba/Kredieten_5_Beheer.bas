Attribute VB_Name = "Kredieten_5_Beheer"
Option Explicit
Private Const WACHTWOORD_JAAROVERGANG As String = "Best"   ' wachtwoord knop Jaarovergang uitvoeren (hoofdletter B)
Private Const MAX_WAAR As Long = 5            ' blok 4: zoveel vindplaatsen tonen, daarna 'en .. meer'
Private Const MAX_NUMMERS As Long = 25         ' blok 4, te controleren: zoveel kredietnummers tonen

' Alle tekst uit de toelichtingskolom van het Logboek (om dubbel verwerken te herkennen).
Public Function LogboekTekst() As String
    Dim lg As Worksheet, v As Variant, i As Long, t As String
    On Error Resume Next
    Set lg = ThisWorkbook.Worksheets("Logboek")
    v = lg.Range("K1:K" & lg.Cells(lg.Rows.Count, 1).End(xlUp).Row).Value
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            t = t & CStr(v(i, 1)) & vbLf
        Next i
    End If
    LogboekTekst = t
End Function

' Legt elke verwerkte wijziging per jaarschijf vast op het blad 'Mutaties kadernota-begroting'.
' Met taakveld, balanscode en afschrijftermijn is dit de basis voor de berekening van de lastenaanpassingen
' (kapitaallasten) in de kadernota/begroting. Het register wordt aangevuld, nooit leeggemaakt.
Public Sub MutatieRegistreren(ByVal naam As String, ByVal volgnr As Variant, ByVal nr As String, ByVal oms As String, _
                               ByVal taakveld As Variant, ByVal balans As Variant, ByVal afschr As Variant, ByVal jaar As Long, _
                               ByVal oud As Double, ByVal nieuw As Double, ByVal act As String, ByVal toel As String, ByVal bestand As String)
    Dim ws As Worksheet, rij As Long
    If Not BladBestaatStil(BL_MUTREG) Then
        Set ws = NieuwBlad()
        ws.Name = BL_MUTREG
        ws.Cells.Font.Name = "Arial"
        ws.Cells.Font.Size = 10
        ws.Range("A1").Value = "Mutaties kadernota / begroting (basis voor de lastenaanpassingen)"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Color = RGB(0, 73, 118)
        ws.Range("A4:O4").Value = Array("Verwerkt op", "Verwerkt door", "Budgetbeheerder", "Volgnr", "Kredietnummer", "Omschrijving", _
                                        "Taakveld", "Balanscode", "Afschr", "Jaar", "Oud budget", "Nieuw budget", "Mutatie", _
                                        "Actualisatie", "Toelichting")
        With ws.Range("A4:O4")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 73, 118)
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        ws.Rows(4).RowHeight = 30
        ws.Columns("A").ColumnWidth = 16: ws.Columns("B:C").ColumnWidth = 18: ws.Columns("D").ColumnWidth = 8
        ws.Columns("E").ColumnWidth = 14: ws.Columns("F").ColumnWidth = 40: ws.Columns("G:H").ColumnWidth = 11
        ws.Columns("I:J").ColumnWidth = 7: ws.Columns("K:M").ColumnWidth = 15: ws.Columns("N").ColumnWidth = 26
        ws.Columns("O").ColumnWidth = 50
        ws.Columns("E").NumberFormat = "@"
        ws.Range("A4:O4").AutoFilter
        KnopTerug ws, ws.Range("A2")
    Else
        Set ws = ThisWorkbook.Worksheets(BL_MUTREG)
        ws.Unprotect WACHTWOORD
    End If
    rij = Application.Max(5, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1)
    ws.Cells(rij, 1).Value = Now
    ws.Cells(rij, 1).NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Cells(rij, 2).Value = ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value
    ws.Cells(rij, 3).Value = naam
    ws.Cells(rij, 4).Value = volgnr
    ws.Cells(rij, 5).Value = nr
    ws.Cells(rij, 6).Value = oms
    ws.Cells(rij, 7).NumberFormat = "@"
    ws.Cells(rij, 7).Value = CStr(taakveld)
    ws.Cells(rij, 8).Value = balans
    ws.Cells(rij, 9).Value = afschr
    ws.Cells(rij, 10).Value = jaar
    ws.Cells(rij, 11).Value = oud
    ws.Cells(rij, 12).Value = nieuw
    ws.Cells(rij, 13).Value = nieuw - oud
    ws.Cells(rij, 14).Value = act
    ws.Cells(rij, 15).Value = toel
    ws.Range(ws.Cells(rij, 11), ws.Cells(rij, 13)).NumberFormat = "#,##0;-#,##0;-"
    ws.Range(ws.Cells(rij, 1), ws.Cells(rij, 15)).VerticalAlignment = xlTop
    ws.Cells(rij, 6).WrapText = True
    ws.Cells(rij, 15).WrapText = True
End Sub

Public Function IsOpbrengstActiva(ByVal toel As Variant) As Boolean
    IsOpbrengstActiva = (InStr(1, CStr(toel), OPBRENGST, vbTextCompare) > 0)
End Function

' Eenmalig: markeert de bekende opbrengsten activa (verkoop vastgoed Sportlaan 10 en voertuigen V06 en V19) door
' "Opbrengst activa." voor de toelichting te zetten op de regels met een negatief budget. Oude tekst gaat naar het Logboek.
Public Sub OpbrengstenActivaMarkeren()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, n As Long, nr As Variant, c As Long, negatief As Boolean, toel As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        For Each nr In Array("7203107", "7206072", "7205045", "7205041")
            If NrTekst(wv.Cells(r, 3).Value) = nr Then
                negatief = (Nul(wv.Cells(r, 26).Value) < 0)
                For c = 31 To 35
                    If Nul(wv.Cells(r, c).Value) < 0 Then negatief = True
                Next c
                toel = Trim$(CStr(wv.Cells(r, 36).Value))
                If negatief And Not IsOpbrengstActiva(toel) Then
                    LogRegel "opbrengst activa markeren", wv.Cells(r, 1).Value, nr, "Toelichting", toel, OPBRENGST & ". " & toel, "rij " & r
                    wv.Cells(r, 36).Value = OPBRENGST & IIf(Len(toel) > 0, ". " & toel, "")
                    n = n + 1
                End If
            End If
        Next nr
    Next r
    MsgBox n & " regel(s) gemarkeerd als opbrengst activa.", vbInformation
End Sub

' Nieuw tabblad voor Voertuigen en Controle (die blijven achteraan). Zo hoeven Voertuigen en Controle niet steeds
' verplaatst te worden; verplaatsen maakt een blad actief, waardoor het Controle-blad even in beeld kwam.
Public Function NieuwBlad() As Worksheet
    If BladBestaatStil("Voertuigen") Then
        Set NieuwBlad = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets("Voertuigen"))
    ElseIf BladBestaatStil("Controle") Then
        Set NieuwBlad = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets("Controle"))
    Else
        Set NieuwBlad = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
End Function

' ======================================================================
' Controle en kengetallen (stap 12)
' ======================================================================
' Stap 12 - Blok 5 op het Dashboard (vanaf rij 69) en de statusregel in C4 (functioneel ontwerp par. 8).
' ControleblokPlaatsen : zet de controles en kengetallen op het Dashboard. Waar het kan als formule (altijd actueel).
' ControlesVernieuwen  : rekent de controles uit die niet met een formule kunnen (spiegelbladen, verdeelsleutel,
'                        dubbele nummers, handmatige formules, SUBTOTAL-regels). Draait bij openen, na een upload en via de knop.
' Statusregel C4: "OK - alles sluit aan" als alle controles 0 zijn, anders "LET OP".


' Bereik van een kolom van de verzamellijst over de kredietregels (rij 2 t/m Par_LaatsteRij), als formuletekst.
Private Function Rg(ByVal kol As String) As String
    Rg = VL & "$" & kol & "$2:INDEX(" & VL & "$" & kol & ":$" & kol & ",Par_LaatsteRij)"
End Function

' Tabblad Kengetallen (achteraan): de kengetallen als formules op de Verzamellijst (altijd actueel).
Private Sub KengetallenBladMaken()
    Dim ws As Worksheet, r As Long, k As Long
    If BladBestaatStil(BL_KENG) Then
        Set ws = ThisWorkbook.Worksheets(BL_KENG)
        ws.Unprotect WACHTWOORD
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = BL_KENG
    End If
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 18
    ws.Range("B1").Value = "Kengetallen investeringskredieten"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B2").Value = "Formules op de Verzamellijst: altijd actueel."
    ws.Range("B2").Font.Italic = True
    r = 4
    ' lopend: alleen t/m het lopend jaar (geen kredieten die pas in een toekomstig jaar starten)
    Kengetal ws, r, "=""Aantal lopende kredieten (L t/m ""&Par_LopendJaar&"")""", "=SUMPRODUCT((LEFT(" & Rg("H") & ",1)=""L"")*(IFERROR(--MID(" & _
             Rg("H") & ",2,4),9999)<=Par_LopendJaar))", "0"
    ' afgesloten per jaar: lopend jaar, de twee jaren ervoor, en eerder
    For k = 0 To 2
        Kengetal ws, r, "=""Aantal afgesloten in ""&(Par_LopendJaar-" & k & ")", "=COUNTIF(" & Rg("H") & ",""A""&(Par_LopendJaar-" & k & "))", "0"
    Next k
    Kengetal ws, r, "=""Aantal afgesloten voor ""&(Par_LopendJaar-2)", "=SUMPRODUCT((LEFT(" & Rg("H") & ",1)=""A"")*(IFERROR(--MID(" & _
             Rg("H") & ",2,4),9999)<Par_LopendJaar-2))", "0"
    r = r + 1
    Kengetal ws, r, "=""Budget t/m ""&Par_LopendJaar", "=SUM(" & Rg("Z") & ")", ""
    Kengetal ws, r, "=""Werkelijke uitgaven t/m ""&Par_LopendJaar", "=SUM(" & Rg("AA") & ")+SUM(" & Rg("AB") & ")+SUM(" & Rg("AC") & ")", ""
    Kengetal ws, r, "Restantkrediet", "=SUM(" & Rg("AD") & ")", ""
    For k = 1 To 5
        Kengetal ws, r, "=""Budget ""&(Par_LopendJaar+" & k & ")", "=SUM(" & Rg(Choose(k, "AE", "AF", "AG", "AH", "AI")) & ")", ""
    Next k
    r = r + 1
    Kengetal ws, r, "Te controleren door de administratie", "=COUNTIF(" & Rg("AP") & ",""" & CtlTeControleren() & """)", "0"
    Kengetal ws, r, "Afgekeurd", "=COUNTIF(" & Rg("AP") & ",""Afgekeurd"")", "0"
    ActiveWindow.DisplayGridlines = False
    KnopTerug ws, ws.Range("D1")
    BladBeveiligen ws
End Sub

' Kolom 'Beheer' in de rollentabel: ja = applicatiebeheerder (ziet alles en beheert de applicatiebeheerders).
' Ontbreekt de kolom, dan wordt hij toegevoegd; Michiel wordt dan applicatiebeheerder, de rest niet.
' Het recht Beheer kan alleen via de knoppen Applicatiebeheerder toevoegen / verwijderen worden gewijzigd (de kolom staat op slot).
Public Function BeheerKolomZorgen() As Boolean
    Dim wp As Worksheet, lo As ListObject, lc As ListColumn, r As ListRow, nm As String, wasBeveiligd As Boolean
    On Error GoTo klaar
    Set wp = ThisWorkbook.Worksheets("Parameters")
    Set lo = wp.ListObjects("Tbl_Rollen")
    On Error Resume Next
    Set lc = lo.ListColumns("Beheer")
    On Error GoTo klaar
    If Not lc Is Nothing Then BeheerKolomZorgen = True: Exit Function
    If CStr(lo.HeaderRowRange.Cells(1, 1).Value) Like "Kolom#*" Then Exit Function   ' tabel door elkaar: eerst RollentabelHerstellen
    wasBeveiligd = wp.ProtectContents
    wp.Unprotect WACHTWOORD
    Set lc = KolomBeheerToevoegen(lo)
    For Each r In lo.ListRows
        nm = LCase$(Trim$(CStr(r.Range.Cells(1, 1).Value)))
        r.Range.Cells(1, lc.Index).Value = IIf(nm Like "michiel*", "ja", "nee")
    Next r
    lc.DataBodyRange.Interior.Color = RGB(217, 217, 217)      ' grijs: niet zelf aanpassen, alleen via de knoppen
    If wasBeveiligd Then BladBeveiligen wp
    LogRegel "rollentabel", , , "Beheer", "", "kolom toegevoegd", "applicatiebeheerder: Michiel"
    BeheerKolomZorgen = True
    Exit Function
klaar:
    MsgBox "De kolom Beheer kon niet aan de rollentabel worden toegevoegd: " & Err.Description & vbCrLf & vbCrLf & _
           "Is de rollentabel door elkaar geraakt (koppen Kolom1, Kolom2 ...)? Start dan via Alt+F8 de macro RollentabelHerstellen.", vbExclamation
End Function

' Voegt kolom Beheer rechts aan de rollentabel toe zonder cellen te verschuiven (ListColumns.Add schuift cellen,
' dat mag niet als er rechts of onder nog een tabel staat): de tabel wordt een kolom breder gemaakt.
' Voorwaarde: de kolom rechts naast de tabel is leeg op de rijen van de tabel.
Private Function KolomBeheerToevoegen(ByVal lo As ListObject) As ListColumn
    Dim naast As Range
    Set naast = lo.Range.Offset(0, lo.Range.Columns.Count).Resize(, 1)
    If Application.WorksheetFunction.CountA(naast) > 0 Then
        Err.Raise vbObjectError + 1, , "de cellen " & naast.Address(False, False) & " rechts naast de rollentabel zijn niet leeg. " & _
                                       "Maak ze leeg (of verplaats wat daar staat) en probeer het opnieuw."
    End If
    Dim nKol As Long
    nKol = lo.Range.Columns.Count
    naast.Cells(1, 1).Value = "Beheer"
    If lo.Range.Columns.Count = nKol Then lo.Resize lo.Range.Resize(, nKol + 1)   ' Excel kan de tabel al zelf hebben uitgebreid
    lo.HeaderRowRange.Font.Color = RGB(0, 0, 0)
    lo.HeaderRowRange.Font.Bold = True
    Set KolomBeheerToevoegen = lo.ListColumns("Beheer")
End Function

' Herstelt de rollentabel als die door elkaar is geraakt (koppen 'Kolom1, Kolom2 ...', de kopregel tussen de namen,
' de lijst Lst_Soort uit kolom I in de tabel). Bewaart alle personen met hun rechten, zet de tabel terug op
' A14:G (Naam, Windows-naam, Rol, Toevoegen, Wijzigen, Bevestigen, Beheer) en ververst daarna de keuzelijsten.
' Eenmalig met de hand te starten (Alt+F8).
Public Sub RollentabelHerstellen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wp As Worksheet, lo As ListObject, kop As Variant, eerste As Long, laatste As Long, r As Long, c As Long, n As Long
    Dim d() As Variant, nm As String, nJa As Long, lijst As String, stap As String
    On Error GoTo Fout
    stap = "rollentabel zoeken"
    Set wp = ThisWorkbook.Worksheets("Parameters")
    Set lo = wp.ListObjects("Tbl_Rollen")
    wp.Unprotect WACHTWOORD
    kop = Array("Naam", "Windows-naam", "Rol", "Toevoegen", "Wijzigen", "Bevestigen", "Beheer")
    eerste = lo.Range.Row
    laatste = eerste + lo.Range.Rows.Count - 1
    If lo.Range.Column <> 1 Then Err.Raise vbObjectError + 2, , "de rollentabel begint niet in kolom A"

    stap = "personen verzamelen"
    ReDim d(1 To laatste - eerste + 1, 1 To 7)
    For r = eerste To laatste
        nm = Trim$(CStr(wp.Cells(r, 1).Value))
        If Len(nm) > 0 And LCase$(nm) <> "naam" And Not (nm Like "Kolom#*") Then
            n = n + 1
            For c = 1 To 7
                d(n, c) = wp.Cells(r, c).Value
            Next c
            If LCase$(Trim$(CStr(d(n, 7)))) <> "ja" Then d(n, 7) = "nee" Else nJa = nJa + 1
            lijst = lijst & "- " & nm & " (" & d(n, 3) & IIf(d(n, 7) = "ja", ", applicatiebeheerder", "") & ")" & vbCrLf
        End If
    Next r
    If n = 0 Then Err.Raise vbObjectError + 3, , "er zijn geen personen in de rollentabel gevonden"
    If MsgBox("Deze " & n & " personen worden in de rollentabel teruggezet:" & vbCrLf & vbCrLf & lijst & vbCrLf & _
              "Doorgaan?", vbYesNo + vbQuestion, "Rollentabel herstellen") <> vbYes Then Exit Sub

    stap = "tabel terugzetten op A" & eerste & ":G" & (eerste + n)
    lo.Resize wp.Range(wp.Cells(eerste, 1), wp.Cells(eerste + n, 7))
    wp.Range(wp.Cells(eerste + 1, 1), wp.Cells(Application.Max(laatste, eerste + n), 7)).ClearContents
    wp.Range(wp.Cells(eerste + n + 1, 1), wp.Cells(Application.Max(laatste, eerste + n) + 1, 7)).Interior.Pattern = xlNone
    wp.Range(wp.Cells(eerste, 8), wp.Cells(laatste, 9)).ClearContents          ' kolom H (buffer) en I (Lst_Soort, wordt opnieuw gevuld)
    stap = "koppen en personen schrijven"
    For c = 1 To 7
        wp.Cells(eerste, c).Value = kop(c - 1)
    Next c
    For r = 1 To n
        For c = 1 To 7
            wp.Cells(eerste + r, c).Value = d(r, c)
        Next c
    Next r
    If nJa = 0 Then
        For r = 1 To n
            If LCase$(Trim$(CStr(d(r, 1)))) Like "michiel*" Then wp.Cells(eerste + r, 7).Value = "ja"
        Next r
    End If
    With lo.HeaderRowRange
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    lo.ListColumns("Beheer").DataBodyRange.Interior.Color = RGB(217, 217, 217)
    LogRegel "rollentabel hersteld", , , "Tbl_Rollen", "", lo.Range.Address(False, False), n & " personen teruggezet"

    stap = "keuzelijsten verversen"
    KeuzelijstenVerversen True
    stap = "dashboards bijwerken"
    GebruikerRegelZetten
    DashboardNaarRol
    BladBeveiligen wp
    MsgBox "De rollentabel is hersteld (" & lo.Range.Address(False, False) & ", " & n & " personen)." & vbCrLf & _
           "Rechten van de huidige gebruiker: toevoegen " & IIf(HeeftRecht("Toevoegen"), "ja", "nee") & _
           ", wijzigen " & IIf(HeeftRecht("Wijzigen"), "ja", "nee") & ", applicatiebeheer " & IIf(HeeftRecht("Beheer"), "ja", "nee") & ".", _
           vbInformation, "Rollentabel herstellen"
    Exit Sub
Fout:
    MsgBox "Niet gelukt bij de stap: " & stap & vbCrLf & vbCrLf & "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Rollentabel herstellen"
End Sub

' Hulp bij problemen met rechten (Alt+F8): toont wat de macro's in de rollentabel lezen voor de huidige gebruiker.
Public Sub RechtenTonen()
    Dim lo As ListObject, r As ListRow, naam As String, t As String, c As Long, gevonden As Boolean, recht As Variant
    On Error GoTo Fout
    naam = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    t = "Huidige gebruiker: [" & naam & "]" & vbCrLf & "Windows-naam: " & Environ("USERNAME") & vbCrLf & vbCrLf
    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects("Tbl_Rollen")
    t = t & "Rollentabel: " & lo.Range.Address(False, False) & ", kolommen: "
    For c = 1 To lo.ListColumns.Count
        t = t & "[" & lo.ListColumns(c).Name & "] "
    Next c
    t = t & vbCrLf & vbCrLf
    For Each r In lo.ListRows
        If StrComp(Trim$(CStr(r.Range.Cells(1, 1).Text)), Trim$(naam), vbTextCompare) = 0 Then
            gevonden = True
            t = t & "Regel van " & naam & ":" & vbCrLf
            For c = 1 To lo.ListColumns.Count
                t = t & "  " & lo.ListColumns(c).Name & " = [" & CStr(r.Range.Cells(1, c).Text) & "]" & vbCrLf
            Next c
        End If
    Next r
    If Not gevonden Then t = t & "Deze naam staat NIET in de rollentabel." & vbCrLf
    t = t & vbCrLf & "Uitkomst rechten:" & vbCrLf
    For Each recht In Array("Toevoegen", "Wijzigen", "Bevestigen", "Beheer")
        t = t & "  " & recht & ": " & IIf(HeeftRecht(CStr(recht)), "ja", "nee") & vbCrLf
    Next recht
    MsgBox t, vbInformation, "Rechten"
    Exit Sub
Fout:
    MsgBox t & vbCrLf & "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Rechten"
End Sub

' Knop Overzicht gebruikers (blok 1 Applicatiebeheer): alle gebruikers uit de rollentabel, per rol.
Public Sub GebruikersOverzicht()
    Dim lo As ListObject, r As ListRow, rollen As Object, rol As Variant, t As String, regel As String, cB As Long, n As Long
    On Error GoTo Fout
    Set lo = ThisWorkbook.Worksheets("Parameters").ListObjects("Tbl_Rollen")
    Set rollen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    cB = lo.ListColumns("Beheer").Index
    On Error GoTo Fout
    For Each r In lo.ListRows
        If Len(Trim$(CStr(r.Range.Cells(1, 1).Value))) > 0 Then
            rol = Trim$(CStr(r.Range.Cells(1, 3).Value))
            If Len(rol) = 0 Then rol = "(geen rol)"
            regel = "   - " & Trim$(CStr(r.Range.Cells(1, 1).Value)) & _
                    IIf(Len(Trim$(CStr(r.Range.Cells(1, 2).Value))) > 0, " (" & Trim$(CStr(r.Range.Cells(1, 2).Value)) & ")", " (geen Windows-naam)")
            If cB > 0 Then
                If LCase$(Trim$(CStr(r.Range.Cells(1, cB).Value))) = "ja" Then regel = regel & " - applicatiebeheerder"
            End If
            rollen(rol) = rollen(rol) & regel & vbCrLf
            n = n + 1
        End If
    Next r
    For Each rol In rollen.Keys
        t = t & UCase$(Left$(rol, 1)) & Mid$(rol, 2) & ":" & vbCrLf & rollen(rol) & vbCrLf
    Next rol
    MsgBox n & " gebruikers (naam en Windows-naam):" & vbCrLf & vbCrLf & t, vbInformation, "Overzicht gebruikers"
    Exit Sub
Fout:
    MsgBox "Het overzicht kon niet worden gemaakt: " & Err.Description, vbExclamation
End Sub

' Knop Gebruiker toevoegen (blok 1 Applicatiebeheer): een nieuwe medewerker in de rollentabel zetten.
Public Sub GebruikerToevoegen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wp As Worksheet, lo As ListObject, lr As ListRow, r As ListRow, naam As String, windows As String, k As String, rol As String
    On Error GoTo Fout
    If Not HeeftRecht("Beheer") Then
        MsgBox "Dit mag alleen een applicatiebeheerder.", vbExclamation
        Exit Sub
    End If
    Set wp = ThisWorkbook.Worksheets("Parameters")
    Set lo = wp.ListObjects("Tbl_Rollen")
    naam = Trim$(InputBox("Voornaam van de nieuwe gebruiker (zoals collega's hem of haar kennen):", "Gebruiker toevoegen"))
    If Len(naam) = 0 Then Exit Sub
    For Each r In lo.ListRows
        If StrComp(Trim$(CStr(r.Range.Cells(1, 1).Value)), naam, vbTextCompare) = 0 Then
            MsgBox naam & " staat al in de rollentabel.", vbExclamation
            Exit Sub
        End If
    Next r
    windows = Trim$(InputBox("Windows-naam van " & naam & " (mag leeg blijven; wordt dan bij de eerste keer openen gevraagd):", "Gebruiker toevoegen"))
    k = Trim$(InputBox("Rol van " & naam & ":" & vbCrLf & "1 = financieel adviseur" & vbCrLf & "2 = financi" & ChrW(235) & "le administratie", "Gebruiker toevoegen", "1"))
    If k <> "1" And k <> "2" Then Exit Sub
    rol = IIf(k = "1", "financieel adviseur", "financi" & ChrW(235) & "le administratie")
    wp.Unprotect WACHTWOORD
    Set lr = lo.ListRows.Add
    lr.Range.Cells(1, lo.ListColumns("Naam").Index).Value = naam
    lr.Range.Cells(1, lo.ListColumns("Windows-naam").Index).Value = windows
    lr.Range.Cells(1, lo.ListColumns("Rol").Index).Value = rol
    lr.Range.Cells(1, lo.ListColumns("Toevoegen").Index).Value = IIf(k = "1", "ja", "nee")
    lr.Range.Cells(1, lo.ListColumns("Wijzigen").Index).Value = "ja"
    lr.Range.Cells(1, lo.ListColumns("Bevestigen").Index).Value = IIf(k = "1", "nee", "ja")
    lr.Range.Cells(1, lo.ListColumns("Beheer").Index).Value = "nee"
    BladBeveiligen wp
    LogRegel "gebruiker toegevoegd", , , "Rol", , rol, naam & " toegevoegd door de applicatiebeheerder"
    MsgBox naam & " is toegevoegd als " & rol & ".", vbInformation
    Exit Sub
Fout:
    MsgBox "Gebruiker toevoegen is niet gelukt: " & Err.Description, vbExclamation
End Sub

' Eenmalig met de hand te starten (Alt+F8): zet de kolom Beheer in de rollentabel en maakt Michiel applicatiebeheerder
' als er nog geen applicatiebeheerder is. Meldt precies wat er gebeurt (ook een foutmelding).
Public Sub ApplicatiebeheerInstellen()
    Dim wp As Worksheet, lo As ListObject, lc As ListColumn, r As ListRow, stap As String, n As Long, gevonden As Boolean
    On Error GoTo Fout
    stap = "blad Parameters en rollentabel zoeken"
    Set wp = ThisWorkbook.Worksheets("Parameters")
    Set lo = wp.ListObjects("Tbl_Rollen")
    stap = "beveiliging van Parameters uitzetten"
    wp.Unprotect WACHTWOORD
    stap = "kolom Beheer zoeken"
    On Error Resume Next
    Set lc = lo.ListColumns("Beheer")
    On Error GoTo Fout
    If lc Is Nothing Then
        stap = "kolom Beheer toevoegen aan de rollentabel"
        Set lc = KolomBeheerToevoegen(lo)
        For Each r In lo.ListRows
            r.Range.Cells(1, lc.Index).Value = "nee"
        Next r
        lc.DataBodyRange.Interior.Color = RGB(217, 217, 217)
        LogRegel "rollentabel", , , "Beheer", "", "kolom toegevoegd", "via ApplicatiebeheerInstellen"
    End If
    stap = "applicatiebeheerders tellen"
    For Each r In lo.ListRows
        If LCase$(Trim$(CStr(r.Range.Cells(1, lc.Index).Value))) = "ja" Then n = n + 1
    Next r
    If n = 0 Then
        stap = "Michiel applicatiebeheerder maken"
        For Each r In lo.ListRows
            If LCase$(Trim$(CStr(r.Range.Cells(1, 1).Value))) Like "michiel*" Then
                r.Range.Cells(1, lc.Index).Value = "ja"
                gevonden = True
                LogRegel "applicatiebeheerder toegevoegd", , , "Beheer", "nee", "ja", CStr(r.Range.Cells(1, 1).Value) & " (eerste applicatiebeheerder)"
            End If
        Next r
    End If
    stap = "kop van de rollentabel leesbaar maken"
    With lo.HeaderRowRange
        .Font.Color = RGB(0, 0, 0)                   ' zwarte letters (witte tekst op geel was onleesbaar)
        .Font.Bold = True
    End With
    stap = "Parameters weer beveiligen"
    BladBeveiligen wp
    stap = "knoppen applicatiebeheerder plaatsen"
    BeheerKnoppenPlaatsen
    stap = "dashboard bijwerken"
    DashboardNaarRol
    MsgBox "Klaar. Kolom Beheer staat in de rollentabel (kolom " & Split(lc.Range.Cells(1, 1).Address(True, False), "$")(0) & ")." & vbCrLf & _
           IIf(n > 0, "Er waren al " & n & " applicatiebeheerder(s).", IIf(gevonden, "Michiel is applicatiebeheerder.", _
           "Let op: Michiel staat niet in de rollentabel; niemand is applicatiebeheerder.")) & vbCrLf & _
           "Huidige gebruiker: " & CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value) & _
           IIf(HeeftRecht("Beheer"), " (applicatiebeheerder).", " (geen applicatiebeheerder)."), vbInformation, "Applicatiebeheer"
    Exit Sub
Fout:
    MsgBox "Niet gelukt bij de stap: " & stap & vbCrLf & vbCrLf & "Fout " & Err.Number & ": " & Err.Description, vbExclamation, "Applicatiebeheer"
End Sub

' Knoppen Applicatiebeheerder toevoegen en verwijderen bovenaan het Dashboard (rij 3, vanaf kolom G; niet in rij 4,
' daar staat de controlestatus), alleen zichtbaar voor een applicatiebeheerder (dat regelt DashboardNaarRol).
Public Sub BeheerKnoppenPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, b As Object, naam As Variant, blad As Variant, pre As String
    For Each blad In Array(BL_DASH)                         ' alleen op Dashboard FinAdv (niet op Dashboard Adm)
        If BladBestaatStil(CStr(blad)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(blad))
            pre = IIf(CStr(blad) = BL_ADM, "btnAdm", "btn")
            ws.Unprotect WACHTWOORD
            On Error Resume Next
            For Each naam In Array(pre & "BeheerToevoegen", pre & "BeheerVerwijderen")
                ws.Shapes("uitleg_" & CStr(naam)).Delete
                ws.Shapes(CStr(naam)).Delete
            Next naam
            On Error GoTo 0
            Set b = ws.Buttons.Add(ws.Range("B1").Left + 410, ws.Range("B" & (AB_KOP + 1)).Top + 2, 195, 22)   ' blok 1 Applicatiebeheer
            b.Name = pre & "BeheerToevoegen"
            b.Caption = "Applicatiebeheerder toevoegen"
            b.OnAction = "BeheerderToevoegen"
            b.Placement = xlMove
            b.PrintObject = False
            Set b = ws.Buttons.Add(ws.Range("B1").Left + 615, ws.Range("B" & (AB_KOP + 1)).Top + 2, 195, 22)
            b.Name = pre & "BeheerVerwijderen"
            b.Caption = "Applicatiebeheerder verwijderen"
            b.OnAction = "BeheerderVerwijderen"
            b.Placement = xlMove
            b.PrintObject = False
            KnopUitleg ws, pre & "BeheerToevoegen", "Applicatiebeheerder toevoegen (alleen applicatiebeheerder): toont een genummerde lijst met de " & _
                       "gebruikers uit de rollentabel die nog geen applicatiebeheerder zijn. Typ het nummer. Die persoon krijgt " & _
                       "dezelfde bevoegdheden als jij: alle blokken zien en applicatiebeheerders toevoegen of verwijderen. " & _
                       "Zijn eigen rol (financieel adviseur of financi" & ChrW(235) & "le administratie) blijft. Komt in het Logboek."
            KnopUitleg ws, pre & "BeheerVerwijderen", "Applicatiebeheerder verwijderen (alleen applicatiebeheerder): toont de applicatiebeheerders. " & _
                       "Typ het nummer. Die persoon blijft gebruiker met zijn eigen rol, maar is geen applicatiebeheerder meer. " & _
                       "Er moet altijd minstens " & ChrW(233) & ChrW(233) & "n applicatiebeheerder overblijven. Komt in het Logboek."
            BeheerKnoppenZichtbaar ws, pre
            BladBeveiligen ws
        End If
    Next blad
    GebruikerRegelZetten
End Sub

' Knoppen applicatiebeheerder alleen tonen aan een applicatiebeheerder.
Public Sub BeheerKnoppenZichtbaar(ByVal ws As Worksheet, ByVal pre As String)
    Dim nmK As Variant, zien As Boolean
    On Error Resume Next
    zien = HeeftRecht("Beheer")
    For Each nmK In Array(pre & "BeheerToevoegen", "uitleg_" & pre & "BeheerToevoegen", pre & "BeheerVerwijderen", "uitleg_" & pre & "BeheerVerwijderen")
        ws.Shapes(CStr(nmK)).Visible = zien
    Next nmK
End Sub

' C2 op de dashboards: naam (rol, applicatiebeheerder).
Public Sub GebruikerRegelZetten()
    Dim blad As Variant, ws As Worksheet, beveiligd As Boolean
    On Error Resume Next
    For Each blad In Array(BL_DASH, BL_ADM)
        If BladBestaatStil(CStr(blad)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(blad))
            beveiligd = ws.ProtectContents
            If beveiligd Then ws.Unprotect WACHTWOORD
            ws.Range("C2").Formula = "=Par_HuidigeGebruiker&IF(Par_HuidigeRol="""",""""," & _
                                     """ (""&Par_HuidigeRol&IF(COUNTIFS(Tbl_Rollen[Naam],Par_HuidigeGebruiker,Tbl_Rollen[Beheer],""ja"")>0," & _
                                     """, applicatiebeheerder"","""")&"")"")"
            If beveiligd Then BladBeveiligen ws
        End If
    Next blad
End Sub

Public Sub BeheerderToevoegen()
    BeheerderWijzigen True
End Sub

Public Sub BeheerderVerwijderen()
    BeheerderWijzigen False
End Sub

' toevoegen = True: iemand applicatiebeheerder maken; False: het recht weghalen (minstens een beheerder blijft over).
Private Sub BeheerderWijzigen(ByVal toevoegen As Boolean)
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wp As Worksheet, lo As ListObject, cB As Long, i As Long, n As Long, k As Long, nBeheer As Long
    Dim idx() As Long, lijst As String, antw As String, naam As String, titel As String, ik As String, isBeheer As Boolean
    titel = IIf(toevoegen, "Applicatiebeheerder toevoegen", "Applicatiebeheerder verwijderen")
    On Error GoTo Fout
    If Not BeheerKolomZorgen() Then Exit Sub
    If Not HeeftRecht("Beheer") Then
        MsgBox "Dit mag alleen een applicatiebeheerder.", vbExclamation, titel
        Exit Sub
    End If
    Set wp = ThisWorkbook.Worksheets("Parameters")
    Set lo = wp.ListObjects("Tbl_Rollen")
    cB = lo.ListColumns("Beheer").Index
    ik = CStr(ThisWorkbook.Names("Par_HuidigeGebruiker").RefersToRange.Value)
    For i = 1 To lo.ListRows.Count
        naam = Trim$(CStr(lo.ListRows(i).Range.Cells(1, 1).Value))
        isBeheer = (LCase$(Trim$(CStr(lo.ListRows(i).Range.Cells(1, cB).Value))) = "ja")
        If isBeheer Then nBeheer = nBeheer + 1
        If Len(naam) > 0 And isBeheer <> toevoegen Then
            n = n + 1
            ReDim Preserve idx(1 To n)
            idx(n) = i
            lijst = lijst & n & " = " & naam & " (" & CStr(lo.ListRows(i).Range.Cells(1, 3).Value) & ")" & vbCrLf
        End If
    Next i
    If n = 0 Then
        MsgBox IIf(toevoegen, "Alle gebruikers zijn al applicatiebeheerder.", "Er zijn geen applicatiebeheerders."), vbInformation, titel
        Exit Sub
    End If
    If Not toevoegen And nBeheer <= 1 Then
        MsgBox "Er is maar " & ChrW(233) & ChrW(233) & "n applicatiebeheerder. Die kan niet worden verwijderd." & vbCrLf & _
               "Voeg eerst een andere applicatiebeheerder toe.", vbExclamation, titel
        Exit Sub
    End If
    antw = Trim$(InputBox(IIf(toevoegen, "Wie wordt applicatiebeheerder? Typ het nummer.", "Wie is geen applicatiebeheerder meer? Typ het nummer.") & _
                          vbCrLf & vbCrLf & lijst, titel))
    If Len(antw) = 0 Then Exit Sub
    If Not IsNumeric(antw) Then
        MsgBox "Dat is geen nummer. Er is niets gewijzigd.", vbExclamation, titel
        Exit Sub
    End If
    k = CLng(antw)
    If k < 1 Or k > n Then
        MsgBox "Dat nummer staat niet in de lijst. Er is niets gewijzigd.", vbExclamation, titel
        Exit Sub
    End If
    naam = Trim$(CStr(lo.ListRows(idx(k)).Range.Cells(1, 1).Value))
    If Not toevoegen And StrComp(naam, ik, vbTextCompare) = 0 Then
        If MsgBox("Je haalt je eigen recht als applicatiebeheerder weg. Daarna kun je dit niet zelf terugzetten." & vbCrLf & _
                  "Doorgaan?", vbYesNo + vbExclamation + vbDefaultButton2, titel) <> vbYes Then Exit Sub
    ElseIf MsgBox(naam & IIf(toevoegen, " applicatiebeheerder maken?", " is dan geen applicatiebeheerder meer. Doorgaan?"), _
                  vbYesNo + vbQuestion + vbDefaultButton2, titel) <> vbYes Then
        Exit Sub
    End If
    wp.Unprotect WACHTWOORD
    lo.ListRows(idx(k)).Range.Cells(1, cB).Value = IIf(toevoegen, "ja", "nee")
    BladBeveiligen wp
    LogRegel IIf(toevoegen, "applicatiebeheerder toegevoegd", "applicatiebeheerder verwijderd"), , , "Beheer", _
             IIf(toevoegen, "nee", "ja"), IIf(toevoegen, "ja", "nee"), naam
    DashboardNaarRol
    MsgBox IIf(toevoegen, naam & " is nu applicatiebeheerder.", naam & " is geen applicatiebeheerder meer."), vbInformation, titel
    Exit Sub
Fout:
    MsgBox "Fout: " & Err.Description, vbCritical, titel
End Sub

' Dashboard per rol: financieel adviseur ziet blok 1 (invoerkaart) en blok 3, niet blok 4; financiele administratie
' ziet blok 4, niet blok 1 en 3 (zij controleert en past aan via de Verzamellijst). Zonder rechten: alleen blok 2 en 5.
' Applicatiebeheerder (kolom Beheer = ja in de rollentabel, eerste keer: Michiel; daarna via de knoppen): ziet alles.
' Draait bij openen (na het kiezen van de gebruiker), bij Wie werkt er en bij het opbouwen van het Dashboard.
Public Sub DashboardNaarRol()
    Dim ws As Worksheet, adv As Boolean, fa As Boolean, beheer As Boolean, shp As Object, beveiligd As Boolean, kaart As Boolean
    On Error Resume Next
    If Not BladBestaatStil(BL_DASH) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(BL_DASH)
    If CStr(ws.Range("B" & NM_KOP).Value) <> "3. Kredietoverzichten" Then Exit Sub   ' oude indeling: eerst DashboardKnoppenPlaatsen
    BeheerKolomZorgen                            ' kolom Beheer in de rollentabel (applicatiebeheerder)
    Set shp = Nothing
    Set shp = ws.Shapes("btnGebruikerVerwijderen")
    If shp Is Nothing Then GebruikerKnopPlaatsen: KnoppenUitlegPlaatsen       ' knop ontbreekt (nog): alsnog plaatsen
    Set shp = Nothing
    Set shp = ws.Shapes("btnBeheerToevoegen")
    If shp Is Nothing Then BeheerKnoppenPlaatsen
    Set shp = Nothing
    Set shp = ws.Shapes("btnKaartSluiten")
    If shp Is Nothing Then KaartSluitenKnopPlaatsen                           ' knop onder de invoerkaart ontbreekt: alsnog plaatsen
    adv = HeeftRecht("Toevoegen")
    fa = HeeftRecht("Bevestigen")
    beheer = HeeftRecht("Beheer")
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    kaart = adv And KaartOpen()                                               ' invoerkaart open: alleen de kaart tonen
    ws.Rows(AB_KOP & ":" & (KR_KOP - 1)).Hidden = Not beheer Or kaart        ' 1. Applicatiebeheer
    ws.Rows(KR_KOP & ":" & (eerste - 5)).Hidden = Not adv Or kaart           ' 2. Kredieten verwerken en raadplegen
    ws.Rows((eerste - 4) & ":" & (KAART_KNOP + 1)).Hidden = Not kaart        ' invoerkaart: pas na Krediet invoeren/wijzigen
    ws.Rows(eerste + 3).Hidden = True                                         ' korte omschrijving blijft altijd verborgen
    ws.Rows(NM_KOP & ":" & (FA_KOP - 1)).Hidden = Not adv Or kaart           ' 3. Kredietoverzichten en 4. Uitvraag
    ws.Rows(FA_KOP & ":" & (CB_KOP - 2)).Hidden = True                       ' oud blok financiele administratie (staat op Dashboard Adm)
    ws.Rows((KR_KOP + 5) & ":" & (KR_KOP + 6)).Hidden = True                 ' lege rijen onder blok 2
    ws.Rows((NM_KOP + 5) & ":" & (FA_KOP - 1)).Hidden = True                 ' lege rijen onder blok 4
    ws.Rows(CB_KOP & ":" & (CB_LAATSTE + 10)).Hidden = kaart                 ' 5. Controles
    For Each shp In ws.Shapes                                                 ' knoppen op verborgen rijen ook verbergen
        shp.Visible = Not ws.Rows(shp.TopLeftCell.Row).Hidden
    Next shp
    If beveiligd Then BladBeveiligen ws
    DashboardsZichtbaar                                                       ' alleen het eigen dashboard tonen
End Sub

' Knop Naam vervangen (bovenaan het Dashboard): vervangt in een keer een budgetbeheerder, projectleider of
' tijdschrijver door een opvolger, in alle regels van de Verzamellijst waar die naam staat. Bij tijdschrijvers
' (meerdere namen in een cel) wordt alleen die ene naam vervangen. Elke wijziging gaat naar het Logboek.
Public Sub BudgethouderVervangen()
    NaamVervangenSoort "1"
End Sub

Public Sub ProjectleiderVervangen()
    NaamVervangenSoort "2"
End Sub

Public Sub TijdschrijverVervangen()
    NaamVervangenSoort "3"
End Sub

Public Sub NaamVervangen()
    NaamVervangenSoort ""
End Sub

Private Sub NaamVervangenSoort(ByVal keuze As String)
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, kols As Variant, oud As String, nieuw As String, alleenL As VbMsgBoxResult, rol As String
    Dim r As Long, laatste As Long, c As Variant, n As Long, Waarde As String, nw As String, deel As Variant, gevonden As Boolean
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "U mag geen kredieten wijzigen.", vbExclamation
        Exit Sub
    End If
    If Len(keuze) = 0 Then
        keuze = Trim$(InputBox("Welke naam vervangen?" & vbCrLf & "1 = budgetbeheerder (kolom O)" & vbCrLf & "2 = projectleider (kolom P)" & vbCrLf & _
                               "3 = tijdschrijver (kolom Q)" & vbCrLf & "4 = in alle drie", "Naam vervangen", "4"))
    End If
    Select Case keuze
        Case "1": kols = Array(15): rol = "budgethouder"
        Case "2": kols = Array(16): rol = "projectleider"
        Case "3": kols = Array(17): rol = "tijdschrijver"
        Case "4": kols = Array(15, 16, 17): rol = "naam"
        Case "": Exit Sub
        Case Else
            MsgBox "Typ 1, 2, 3 of 4.", vbExclamation
            Exit Sub
    End Select
    oud = Trim$(InputBox("Welke " & rol & " moet worden vervangen? (naam precies zoals in de lijst)", "Vervangen: " & rol))
    If Len(oud) = 0 Then Exit Sub
    nieuw = Trim$(InputBox("Door wie wordt " & oud & " opgevolgd? (naam van de nieuwe " & rol & ")", "Vervangen: " & rol))
    If Len(nieuw) = 0 Then Exit Sub
    alleenL = MsgBox("Alleen bij lopende kredieten (L) vervangen?" & vbCrLf & vbCrLf & "Ja = alleen lopende kredieten (aanbevolen; " & _
                     "bij afgesloten kredieten blijft de oude naam staan)" & vbCrLf & "Nee = ook bij afgesloten kredieten (A)", _
                     vbYesNoCancel + vbQuestion, "Naam vervangen")
    If alleenL = vbCancel Then Exit Sub
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    ' eerst tellen
    For r = 2 To laatste
        If Not IsEmpty(wv.Cells(r, 1).Value) And (alleenL = vbNo Or UCase$(Left$(Trim$(CStr(wv.Cells(r, 8).Value)), 1)) = "L") Then
            For Each c In kols
                If NaamKomtVoor(CStr(wv.Cells(r, CLng(c)).Value), oud) Then n = n + 1
            Next c
        End If
    Next r
    If n = 0 Then
        MsgBox "De naam '" & oud & "' komt niet voor in de gekozen kolom(men).", vbInformation
        Exit Sub
    End If
    If MsgBox("'" & oud & "' vervangen door '" & nieuw & "' in " & n & " cel(len)?", vbYesNo + vbQuestion, "Naam vervangen") <> vbYes Then Exit Sub
    n = 0
    For r = 2 To laatste
        If Not IsEmpty(wv.Cells(r, 1).Value) And (alleenL = vbNo Or UCase$(Left$(Trim$(CStr(wv.Cells(r, 8).Value)), 1)) = "L") Then
            For Each c In kols
                Waarde = CStr(wv.Cells(r, CLng(c)).Value)
                If NaamKomtVoor(Waarde, oud) Then
                    nw = "": gevonden = False
                    For Each deel In Split(Replace(Waarde, ";", ","), ",")
                        If StrComp(Trim$(CStr(deel)), oud, vbTextCompare) = 0 Then deel = nieuw
                        If InStr(1, ", " & nw & ",", ", " & Trim$(CStr(deel)) & ",", vbTextCompare) = 0 Then
                            nw = nw & IIf(Len(nw) > 0, ", ", "") & Trim$(CStr(deel))
                        End If
                    Next deel
                    wv.Cells(r, CLng(c)).Value = nw
                    LogRegel "naam vervangen", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, CStr(wv.Cells(1, CLng(c)).Value), Waarde, nw, "rij " & r
                    n = n + 1
                End If
            Next c
        End If
    Next r
    KeuzelijstenVerversen True
    MsgBox "'" & oud & "' is vervangen door '" & nieuw & "' in " & n & " cel(len). Alles staat in het Logboek.", vbInformation
End Sub

Public Sub ControleblokPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet, r As Long, k As Long, fc As Object, b As Object, deel As Variant, aantalAf As String
    Set ws = ThisWorkbook.Worksheets(BL_DASH)
    ws.Range("A" & CB_KOP & ":H" & (CB_LAATSTE + 40)).Clear
    KopOpmaak ws, CB_KOP, "5. Controles"
    ws.Range("B" & (CB_KOP + 1)).Value = "Controles laatst vernieuwd"
    ws.Range("C" & (CB_KOP + 1)).NumberFormat = "dd-mm-yyyy hh:mm"
    ws.Range("C" & (CB_KOP + 1)).HorizontalAlignment = xlLeft
    On Error Resume Next
    ws.Buttons("btnControles").Delete
    ws.Shapes("uitleg_btnControles").Delete
    On Error GoTo 0
    ws.Rows(CB_KOP + 1).RowHeight = 24               ' knop past binnen de rij en loopt niet door de kopregel eronder
    Set b = ws.Buttons.Add(ws.Range("E" & (CB_KOP + 1)).Left + 2, ws.Range("E" & (CB_KOP + 1)).Top + 3, 170, 18)
    b.Name = "btnControles"
    b.Caption = "Controles vernieuwen"
    b.OnAction = "ControlesVernieuwen"
    ws.Range("B" & (CB_KOP + 2) & ":C" & (CB_KOP + 2)).Value = Array("Controle (rood = nakijken)", "Aantal")
    ws.Range("E" & (CB_KOP + 2)).Value = "Waar (rij in de Verzamellijst) / wat te doen"
    ws.Range("B" & (CB_KOP + 2) & ":F" & (CB_KOP + 2)).Font.Bold = True
    ws.Range("B" & (CB_KOP + 2) & ":F" & (CB_KOP + 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous

    ' de aantallen zijn formules (altijd actueel); de vindplaatsen in kolom D vult ControlesVernieuwen
    r = CB_EERSTE
    Regel ws, r, "Omschrijving langer dan toegestaan (kolom E en G)", "=COUNTIF(" & Rg("E") & ",""fout"")+COUNTIF(" & Rg("G") & ",""fout"")"
    Regel ws, r, "Verdeelsleutel samengesteld krediet niet 100%", ""
    Regel ws, r, "Afgesloten (A) met restantkrediet", "=COUNTIFS(" & Rg("H") & ",""A*""," & Rg("AD") & ",""<>0"")"
    Regel ws, r, "Zonder budgetbeheerder (O)", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("O") & ","""")"
    Regel ws, r, "Zonder projectleider (P)", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("P") & ","""")"
    Regel ws, r, "Zonder activanummer (X) - tijdelijk toegestaan", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("X") & ","""")"
    Regel ws, r, "Zonder afschrijftermijn (Y)", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("Y") & ","""")"
    Regel ws, r, "Dubbel: kredietnummer + activanummer", ""
    Regel ws, r, "Te controleren of afgekeurd", "=COUNTIF(" & Rg("AP") & ",""" & CtlTeControleren() & """)+COUNTIF(" & Rg("AP") & ",""Afgekeurd"")"
    Regel ws, r, "Hoofdletters of spaties in kolom I of J", "=SUMPRODUCT((" & Rg("A") & "<>"""")*(" & Rg("I") & "<>"""")*(1-EXACT(" & Rg("I") & _
          "&"""",LOWER(TRIM(" & Rg("I") & "&"""")))))+SUMPRODUCT((" & Rg("A") & "<>"""")*(" & Rg("J") & "<>"""")*(1-EXACT(" & Rg("J") & _
          "&"""",LOWER(TRIM(" & Rg("J") & "&"""")))))"
    Regel ws, r, "Werkbereik bijna vol (binnen 100 rijen)", "=IF(Par_LaatsteRij>Par_MaxRij-100,1,0)"
    Regel ws, r, "Status (H) niet L of A met jaartal", "=SUMPRODUCT((" & Rg("A") & "<>"""")*(1-(LEN(" & Rg("H") & ")=5)*((LEFT(" & Rg("H") & _
          ",1)=""L"")+(LEFT(" & Rg("H") & ",1)=""A""))*ISNUMBER(--MID(" & Rg("H") & ",2,4))))"
    Regel ws, r, "Zonder kredietnummer (C) - tijdelijk toegestaan", "=COUNTIFS(" & Rg("A") & ",""<>""," & Rg("C") & ","""")"
    Regel ws, r, "7999998 in het begrotingsjaar of eerder", "=SUMPRODUCT((" & Rg("C") & "&""""=""7999998"")*(LEFT(" & Rg("H") & ",1)=""L"")*(IFERROR(--MID(" & _
          Rg("H") & ",2,4),9999)<=Par_LopendJaar+1))"
    Regel ws, r, "Handmatige formules in bedragkolommen", ""
    Regel ws, r, "SUBTOTAL-regels onder de kredieten", ""

    KengetallenBladMaken

    ' rood bij een aantal > 0; zonder activanummer en zonder kredietnummer geel (tijdelijk toegestaan, telt niet mee)
    For Each deel In Array(CB_EERSTE + 5, CB_EERSTE + 12)
        Set fc = VoegCF(ws.Range(ws.Cells(deel, 3), ws.Cells(deel, 3)), "=AND(ISNUMBER(C" & deel & "),C" & deel & ">0)")
        If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 235, 156)
    Next deel
    For Each deel In Array(Array(CB_EERSTE, CB_EERSTE + 4), Array(CB_EERSTE + 6, CB_EERSTE + 11), Array(CB_EERSTE + 13, CB_LAATSTE))
        Set fc = VoegCF(ws.Range(ws.Cells(deel(0), 3), ws.Cells(deel(1), 3)), "=AND(ISNUMBER(C" & deel(0) & "),C" & deel(0) & ">0)")
        If Not fc Is Nothing Then fc.Interior.Color = RGB(255, 150, 150)
    Next deel
    ws.Range(ws.Cells(CB_EERSTE, 3), ws.Cells(CB_LAATSTE, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(CB_EERSTE, 5), ws.Cells(CB_LAATSTE, 5)).WrapText = False      ' tekst loopt door naar rechts
    ws.Range(ws.Cells(CB_EERSTE, 2), ws.Cells(CB_LAATSTE, 5)).VerticalAlignment = xlTop
    ' omschrijving rood als de controle afwijkt (zodat direct te zien is wat nagekeken moet worden)
    For Each deel In Array(Array(CB_EERSTE, CB_EERSTE + 4), Array(CB_EERSTE + 6, CB_EERSTE + 11), Array(CB_EERSTE + 13, CB_LAATSTE))
        Set fc = VoegCF(ws.Range(ws.Cells(deel(0), 2), ws.Cells(deel(1), 2)), "=AND(ISNUMBER($C" & deel(0) & "),$C" & deel(0) & ">0)")
        If Not fc Is Nothing Then fc.Font.Color = RGB(192, 0, 0): fc.Font.Bold = True
    Next deel

    ' statusregel
    aantalAf = "(COUNTIF(C" & CB_EERSTE & ":C" & CB_LAATSTE & ","">0"")-N(C" & (CB_EERSTE + 5) & ">0)-N(C" & (CB_EERSTE + 12) & ">0))"
    ws.Range("C4").Formula = "=IF(" & aantalAf & "=0,""OK - alles sluit aan"",""Te controleren: ""&" & aantalAf & ")"
    ws.Range("C4").Font.Bold = True
    ' klik op 'Controlestatus' (B4) = naar blok 5
    ws.Range("B4").Hyperlinks.Delete
    ws.Hyperlinks.Add Anchor:=ws.Range("B4"), Address:="", SubAddress:="'Dashboard'!B" & CB_KOP, TextToDisplay:="Controlestatus", _
                      ScreenTip:="Naar blok 4: controles"
    ' kleur loopt door tot en met kolom F, zodat de hele melding in het groene of rode blok staat
    ws.Range("C4:F4").FormatConditions.Delete
    Set fc = ws.Range("C4:F4").FormatConditions.Add(Type:=xlExpression, Formula1:="=$C$4=""OK - alles sluit aan""")
    fc.Interior.Color = RGB(198, 239, 206)
    Set fc = ws.Range("C4:F4").FormatConditions.Add(Type:=xlExpression, Formula1:="=$C$4<>""OK - alles sluit aan""")
    fc.Interior.Color = RGB(255, 150, 150)
    ControlesVernieuwen
End Sub

Public Sub Regel(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal formule As String)
    ws.Cells(r, 2).Value = label
    If Len(formule) > 0 Then ws.Cells(r, 3).Formula2 = formule Else ws.Cells(r, 3).ClearContents   ' Formula2: rekent met hele bereiken
    ws.Cells(r, 3).NumberFormat = "0"
    ws.Cells(r, 4).Font.Color = RGB(89, 89, 89)
    r = r + 1
End Sub

Private Sub Kengetal(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal formule As String, ByVal notatie As String)
    If Left$(label, 1) = "=" Then ws.Cells(r, 2).Formula = label Else ws.Cells(r, 2).Value = label
    ws.Cells(r, 3).Formula2 = formule
    ws.Cells(r, 3).NumberFormat = IIf(Len(notatie) > 0, notatie, "#,##0;-#,##0;-")
    ws.Cells(r, 3).HorizontalAlignment = xlRight
    r = r + 1
End Sub

' Rekent de controles uit die niet met een formule kunnen en zet bij elke controle in kolom D waar het zit
' (rij en kredietnummer in de Verzamellijst) en wat te doen.
' Kredieten met 'Te controleren' waarbij sinds de laatste goedkeuring alleen velden zijn gewijzigd die geen controle nodig
' hebben (zie ZonderControle: omschrijving, namen, afschrijftermijn, budgetten, uitgaven, toelichting) krijgen hun status
' 'Gecontroleerd' terug. Bron is het Logboek. Nieuwe kredieten, bijdrageregels en afgekeurde kredieten blijven staan.
Public Sub StatusZonderControleHerstellen(ByVal wv As Worksheet, ByVal laatste As Long)
    Dim lg As Worksheet, lv As Variant, nL As Long, i As Long, r As Long, c As Long, vn As String, actie As String
    Dim vrij As Object, open1 As Object, telt As Object, nHersteld As Long
    On Error GoTo Klaar
    If Not BladBestaat("Logboek") Then Exit Sub
    Set lg = ThisWorkbook.Worksheets("Logboek")
    nL = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row
    If nL < 2 Then Exit Sub
    lv = lg.Range("A1:K" & nL).Value
    ' kopteksten van de kolommen zonder controle (zo staan ze als veld in het Logboek)
    Set vrij = CreateObject("Scripting.Dictionary")
    vrij.CompareMode = vbTextCompare
    For c = 1 To 49
        If ZonderControle(c) Then vrij(Trim$(CStr(wv.Cells(1, c).Value))) = True
    Next c
    vrij("Afschr") = True: vrij("Afschrijving") = True        ' kop Y heette eerder 'Afschr'
    ' per volgnummer: kan de status terug (True) of niet (False), gerekend vanaf de laatste goedkeuring
    Set open1 = CreateObject("Scripting.Dictionary")
    Set telt = CreateObject("Scripting.Dictionary")
    For i = 2 To nL
        vn = Trim$(CStr(lv(i, 6)))
        If Len(vn) > 0 Then
            actie = LCase$(Trim$(CStr(lv(i, 5))))
            Select Case actie
                Case "goedkeuren"
                    open1(vn) = True: telt(vn) = False                    ' opnieuw beginnen na een goedkeuring
                Case "wijzigen"
                    If Not open1.Exists(vn) Then open1(vn) = True
                    If Not vrij.Exists(Trim$(CStr(lv(i, 8)))) Then open1(vn) = False
                    telt(vn) = True
                Case "ter controle na uitvraag", "uitvraag verwerkt"
                    If Not open1.Exists(vn) Then open1(vn) = True
                    telt(vn) = True
                Case "toevoegen", "toevoegen bijdrage", "afkeuren", "aangeboden voor hercontrole"
                    open1(vn) = False: telt(vn) = True
            End Select
        End If
    Next i
    For r = 2 To laatste
        vn = Trim$(CStr(wv.Cells(r, 1).Value))
        If Len(vn) > 0 And CStr(wv.Cells(r, 42).Value) = CtlTeControleren() Then
            If open1.Exists(vn) And telt.Exists(vn) Then
                If open1(vn) And telt(vn) Then
                    wv.Cells(r, 42).Value = CtlGecontroleerd()
                    wv.Cells(r, 43).Value = "geen controle nodig"
                    wv.Cells(r, 44).NumberFormat = "dd-mm-yyyy"
                    wv.Cells(r, 44).Value = Date
                    wv.Cells(r, 45).Value = "alleen wijziging zonder controle (status hersteld)"
                    LogRegel "status hersteld", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Controlestatus", CtlTeControleren(), CtlGecontroleerd(), "rij " & r
                    nHersteld = nHersteld + 1
                End If
            End If
        End If
    Next r
    If nHersteld > 0 Then MsgBox nHersteld & " krediet(en) hoeven niet meer gecontroleerd te worden (alleen namen, budgetten, uitgaven, " & _
                                 "afschrijftermijn, omschrijving of toelichting gewijzigd). De status is weer 'Gecontroleerd'.", vbInformation
Klaar:
End Sub

Public Sub ControlesVernieuwen()
    KredietnummersBijwerken                      ' eerstvolgende vrije kredietnummers (Dashboard Adm, blok 3)
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wd As Worksheet, wv As Worksheet, laatste As Long, maxRij As Long, lj As Long, v As Variant, r As Long, i As Long
    Dim waar(0 To 15) As String, aant(0 To 15) As Long, hint As Variant, c As Variant, t As String, st As String, jr As Long
    Dim grp As Object, grpRij As Object, k As Variant, dubbel As Object, sleutel As String, cel As Range, eerste As String
    On Error GoTo Fout
    Set wd = ThisWorkbook.Worksheets(BL_DASH)
    If CStr(wd.Cells(CB_KOP, 2).Value) <> "5. Controles" Then Exit Sub   ' blok 4 nog niet op deze plek (eerst DashboardAanmaken)
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    StatusZonderControleHerstellen wv, laatste
    v = wv.Range("A1:AW" & laatste).Value
    JGetalNormaliseren v
    hint = Array("macro OmschrijvingenInkorten (max 36 tekens)", "kolom K van de groep (moet samen 100% zijn)", _
                 "restant moet 0 zijn; macro RestantFormuleBijwerken", "invullen via Krediet wijzigen", "invullen via Krediet wijzigen", _
                 "invullen zodra bekend", "invullen via Krediet wijzigen", "kredietnummer en activanummer controleren", _
                 "zoeken met 'Bestaand krediet aanpassen' of 'Volgend krediet controleren' (blok 3)", "macro LijstwaardenNormaliseren", "macro WerkbereikVerruimen", "status moet L of A met jaartal zijn", _
                 "invullen zodra bekend", "echt kredietnummer invullen", "macro FormulesNaarWaarden", "macro SubtotaalRegelsVerwijderen")

    Set grp = CreateObject("Scripting.Dictionary")
    Set grpRij = CreateObject("Scripting.Dictionary")
    Set dubbel = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not Leeg(v(r, 1)) Then
            t = "rij " & r & " (" & IIf(Leeg(v(r, 3)), "volgnr " & v(r, 1), NrTekst(v(r, 3))) & ")"
            st = UCase$(Trim$(Tkst(v(r, 8))))
            If Tkst(v(r, 5)) = "fout" Or Tkst(v(r, 7)) = "fout" Then Voeg waar(0), aant(0), t
            If IsGetal(v(r, 10)) And IsGetal(v(r, 11)) And Not HeeftNegatiefBudget(v, r) Then   ' bijdrageregels tellen niet mee
                grp(Tkst(v(r, 10))) = grp(Tkst(v(r, 10))) + CDbl(v(r, 11))
                grpRij(Tkst(v(r, 10))) = grpRij(Tkst(v(r, 10))) & IIf(Len(grpRij(Tkst(v(r, 10)))) > 0, ", ", "") & r
            End If
            If Left$(st, 1) = "A" And Nul(v(r, 30)) <> 0 Then Voeg waar(2), aant(2), t
            If Leeg(v(r, 15)) Then Voeg waar(3), aant(3), t
            If Leeg(v(r, 16)) Then Voeg waar(4), aant(4), t
            If Leeg(v(r, 24)) Then Voeg waar(5), aant(5), t
            If Leeg(v(r, 25)) Then Voeg waar(6), aant(6), t
            If Not Leeg(v(r, 3)) And IsGetal(v(r, 24)) And NrTekst(v(r, 3)) <> "7999998" And Not HeeftNegatiefBudget(v, r) Then
                sleutel = NrTekst(v(r, 3)) & "|" & Tkst(v(r, 24))
                If dubbel.Exists(sleutel) Then Voeg waar(7), aant(7), t & " = rij " & dubbel(sleutel) Else dubbel.Add sleutel, r
            End If
            If Tkst(v(r, 42)) = CtlTeControleren() Or Tkst(v(r, 42)) = "Afgekeurd" Then
                ' alleen het kredietnummer (zonder kredietnummer: het volgnummer), elk nummer een keer: daarop zoekt de administratie
                eerste = IIf(Len(NrTekst(v(r, 3))) > 0, NrTekst(v(r, 3)), "volgnr " & Tkst(v(r, 1)))
                If InStr(1, "; " & waar(8) & ";", "; " & eerste & ";") = 0 Then Voeg waar(8), aant(8), eerste, MAX_NUMMERS
            End If
            For Each c In Array(9, 10)
                If Not Leeg(v(r, c)) And Not IsGetal(v(r, c)) Then
                    If StrComp(Tkst(v(r, c)), LCase$(Trim$(Tkst(v(r, c)))), vbBinaryCompare) <> 0 Then Voeg waar(9), aant(9), t: Exit For
                End If
            Next c
            If Not (Len(st) = 5 And (Left$(st, 1) = "L" Or Left$(st, 1) = "A") And IsAlleenCijfers(Mid$(st, 2))) Then Voeg waar(11), aant(11), t
            If Leeg(v(r, 3)) Then Voeg waar(12), aant(12), "rij " & r & " (volgnr " & v(r, 1) & ")"
            If NrTekst(v(r, 3)) = "7999998" And Left$(st, 1) = "L" Then
                jr = Val(Mid$(st, 2))
                If jr > 0 And jr <= lj + 1 Then Voeg waar(13), aant(13), "rij " & r & " (volgnr " & v(r, 1) & ", " & st & ")"
            End If
            For Each c In Array(26, 27, 28, 29, 31, 32, 33, 34, 35)
                If wv.Cells(r, c).HasFormula Then Voeg waar(14), aant(14), wv.Cells(r, c).Address(False, False)
            Next c
        End If
    Next r
    For Each k In grp.Keys
        If Abs(grp(k) - 1) > 0.0005 Then Voeg waar(1), aant(1), "groep " & k & " = " & Format$(grp(k) * 100, "0.0") & "% (rijen " & grpRij(k) & ")"
    Next k
    If laatste > maxRij - 100 Then Voeg waar(10), aant(10), "laatste rij " & laatste & ", werkbereik t/m " & maxRij
    Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).Find(What:="SUBTOTAL", LookIn:=xlFormulas, LookAt:=xlPart)
    If Not cel Is Nothing Then
        eerste = cel.Address
        Do
            Voeg waar(15), aant(15), cel.Address(False, False)
            Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).FindNext(cel)
        Loop While Not cel Is Nothing And cel.Address <> eerste
    End If

    For i = 0 To 15
        If aant(i) > 0 Then
            wd.Cells(CB_EERSTE + i, 5).Value = waar(i) & IIf(aant(i) > MaxWaar(i), " en " & (aant(i) - MaxWaar(i)) & " meer", "") & "  >  " & hint(i)
        Else
            wd.Cells(CB_EERSTE + i, 5).Value = ""
        End If
    Next i
    ' aantallen van de controles zonder formule
    wd.Cells(CB_EERSTE + 1, 3).Value = aant(1)
    wd.Cells(CB_EERSTE + 7, 3).Value = aant(7)
    wd.Cells(CB_EERSTE + 14, 3).Value = aant(14)
    wd.Cells(CB_EERSTE + 15, 3).Value = aant(15)
    wd.Range(wd.Cells(CB_EERSTE, 2), wd.Cells(CB_LAATSTE, 2)).Rows.AutoFit
    wd.Range("C" & (CB_KOP + 1)).Value = Now
    If ActiveSheet.Name = "Controle" Then wd.Activate
    Exit Sub
Fout:
    wd.Range("C" & (CB_KOP + 1)).Value = "fout bij vernieuwen: " & Err.Description & " (rij " & r & ")"
End Sub

' Voegt een vindplaats toe aan de lijst (maximaal MAX_WAAR tonen, anders wordt de regel onleesbaar lang) en telt mee.
Private Sub Voeg(ByRef lijst As String, ByRef n As Long, ByVal t As String, Optional ByVal maxN As Long = MAX_WAAR)
    n = n + 1
    If n <= maxN Then lijst = lijst & IIf(Len(lijst) > 0, "; ", "") & t
End Sub

' Aantal vindplaatsen dat blok 4 toont: bij 'te controleren of afgekeurd' (8) meer, want dat zijn korte kredietnummers.
Private Function MaxWaar(ByVal i As Long) As Long
    MaxWaar = IIf(i = 8, MAX_NUMMERS, MAX_WAAR)
End Function

' Regel in de verzamellijst met een negatief budget (bijdrage van derden).
Public Function NegatiefBudgetRij(ByVal wv As Worksheet, ByVal r As Long) As Boolean
    Dim c As Variant
    For Each c In Array(26, 31, 32, 33, 34, 35)
        If Nul(wv.Cells(r, CLng(c)).Value) < 0 Then NegatiefBudgetRij = True: Exit Function
    Next c
End Function

Private Function HeeftNegatiefBudget(ByVal v As Variant, ByVal r As Long) As Boolean
    Dim c As Variant
    For Each c In Array(26, 31, 32, 33, 34, 35)
        If Nul(v(r, c)) < 0 Then HeeftNegatiefBudget = True: Exit Function
    Next c
End Function

' Zet handmatige formules in de bedragkolommen (Z, AA t/m AC, AE t/m AI) om naar hun uitkomst.
' De uitkomst verandert niet; de oude formule gaat naar het Logboek. Let op: een formule die naar een andere cel
' verwijst (bijv. =-ROUND(Z109*20%;0)) beweegt daarna niet meer mee met die cel.
Public Sub FormulesNaarWaarden()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, c As Variant, n As Long, Waarde As Variant, voor As Variant, na As Variant, k As Long
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        For Each c In Array(26, 27, 28, 29, 31, 32, 33, 34, 35)
            If wv.Cells(r, c).HasFormula Then n = n + 1
        Next c
    Next r
    If n = 0 Then
        MsgBox "Er staan geen handmatige formules in de bedragkolommen.", vbInformation
        Exit Sub
    End If
    If MsgBox(n & " handmatige formules omzetten naar hun uitkomst? De bedragen veranderen niet; de formules gaan naar het Logboek.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Formules naar waarden") <> vbYes Then Exit Sub
    Application.Calculate
    voor = MeetStap8(wv, laatste)
    Application.ScreenUpdating = False
    n = 0
    For r = 2 To laatste
        For Each c In Array(26, 27, 28, 29, 31, 32, 33, 34, 35)
            If wv.Cells(r, c).HasFormula Then
                Waarde = wv.Cells(r, c).Value
                LogRegel "formule naar waarde", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, wv.Cells(r, c).Address(False, False), _
                         "'" & wv.Cells(r, c).Formula, Waarde, ""      ' apostrof: in het Logboek als tekst, niet als formule
                wv.Cells(r, c).Value = Waarde
                n = n + 1
            End If
        Next c
    Next r
    Application.Calculate
    na = MeetStap8(wv, laatste)
    Application.ScreenUpdating = True
    For k = 0 To 9
        If Round(na(k) - voor(k), 2) <> 0 Then
            MsgBox "Let op: de totalen wijken af na het omzetten (meting " & k & "). Controleer het Logboek.", vbCritical
            Exit For
        End If
    Next k
    ControlesVernieuwen
    MsgBox n & " formules omgezet naar hun uitkomst. Totalen ongewijzigd.", vbInformation
End Sub

' Herstelt omschrijvingen (kolom D) die langer zijn dan 36 tekens. Per regel een voorstel (straat -> str, daarna afkappen)
' dat u kunt aanpassen. Oude tekst gaat naar het Logboek.
Public Sub OmschrijvingenInkorten()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, r As Long, laatste As Long, oud As String, nieuw As String, n As Long
    If Not HeeftRecht("Wijzigen") Then
        MsgBox "U heeft geen recht om dit te doen.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    For r = 2 To laatste
        oud = CStr(wv.Cells(r, 4).Value)
        If Not Leeg(wv.Cells(r, 1).Value) And Len(oud) > 36 Then
            nieuw = Replace(Replace(oud, "straat", "str"), "  ", " ")
            If Len(nieuw) > 36 Then nieuw = Left$(nieuw, 36)
            Do
                nieuw = InputBox("Volgnr " & wv.Cells(r, 1).Value & " (" & wv.Cells(r, 3).Value & ")" & vbCrLf & _
                                 "Huidig (" & Len(oud) & " tekens): " & oud & vbCrLf & vbCrLf & _
                                 "Nieuwe omschrijving, maximaal 36 tekens:", "Omschrijving te lang", nieuw)
                If Len(nieuw) = 0 Then Exit Do               ' annuleren = deze regel overslaan
            Loop While Len(nieuw) > 36
            If Len(nieuw) > 0 And nieuw <> oud Then
                LogRegel "omschrijving ingekort", wv.Cells(r, 1).Value, wv.Cells(r, 3).Value, "Omschrijving", oud, nieuw, "rij " & r
                wv.Cells(r, 4).Value = nieuw
                n = n + 1
            End If
        End If
    Next r
    Application.Calculate
    MsgBox n & " omschrijving(en) ingekort.", vbInformation
End Sub

' Haalt de oude SUBTOTAL-regels onder de kredieten weg (bijv. rij 1474). De kengetallen staan nu in blok 5 op het Dashboard.
Public Sub SubtotaalRegelsVerwijderen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, laatste As Long, maxRij As Long, cel As Range, lijst As Collection, c As Variant, n As Long, eerste As String
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Alleen financieel adviseurs mogen dit.", vbExclamation
        Exit Sub
    End If
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)
    maxRij = CLng(ThisWorkbook.Names("Par_MaxRij").RefersToRange.Value)
    Set lijst = New Collection
    Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).Find(What:="SUBTOTAL", LookIn:=xlFormulas, LookAt:=xlPart)
    If Not cel Is Nothing Then
        eerste = cel.Address
        Do
            lijst.Add cel.Address(False, False)
            Set cel = wv.Range("A" & (laatste + 1) & ":AW" & (maxRij + 100)).FindNext(cel)
        Loop While Not cel Is Nothing And cel.Address <> eerste
    End If
    If lijst.Count = 0 Then
        MsgBox "Er staan geen SUBTOTAL-regels onder de kredieten.", vbInformation
        Exit Sub
    End If
    If MsgBox(lijst.Count & " SUBTOTAL-cellen onder de kredieten leegmaken?", vbYesNo + vbQuestion, "SUBTOTAL-regels") <> vbYes Then Exit Sub
    For Each c In lijst
        LogRegel "SUBTOTAL-regel verwijderd", , , CStr(c), "'" & wv.Range(c).Formula, "", ""
        wv.Range(c).ClearContents
        n = n + 1
    Next c
    ControlesVernieuwen
    MsgBox n & " SUBTOTAL-cellen leeggemaakt.", vbInformation
End Sub

' ======================================================================
' Jaarovergang (stap 14)
' ======================================================================
' Stap 14 - Jaarovergang (functioneel ontwerp par. 2.7, aangevuld met latere afspraken)
' JaarovergangUitvoeren : knop in blok 4 (alleen financieel adviseur). Bij openen in een nieuw kalenderjaar wordt erom gevraagd.
' 1. Back-up 'Verzamellijst 31-12-<oud jaar>' als harde waarden (= stand per 31-12; bron voor het jaarrekeningoverzicht).
' 2. Per regel: Z = Z + AE; AA = AA + AB; AB = AC; AC leeg; AE..AH = AF..AI; AI leeg. AD blijft formule.
' 3. Status L<oud jaar of eerder> wordt L<nieuw jaar>. Staat er voor de jaarovergang al een bedrag in Budget nieuw jaar (AE), dan blijven
'    Startjaar (M) en Eindjaar (N) ongewijzigd; anders wordt een waarde gelijk aan het oude jaar in M of N het nieuwe jaar (andere jaartallen blijven).
'    Status A en toekomstige L blijven ongewijzigd.
' 4. Par_LopendJaar + 1 (koppen en Dashboard schuiven vanzelf mee).
' 5. Controle op de totalen; bij een verschil alles terugzetten.

Public Sub JaarovergangUitvoeren()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, wb As Worksheet, oud As Long, nieuw As Long, laatste As Long, r As Long, k As Long, nm As String
    Dim z As Variant, aa As Variant, ab As Variant, ac As Variant, ae As Variant, h As Variant, m As Variant, n As Variant
    Dim oZ As Variant, oAA As Variant, oAB As Variant, oAC As Variant, oAE As Variant, oH As Variant, oStart As Variant, oEind As Variant
    Dim sZ As Double, sAE1 As Double, sW As Double, sAEtot As Double, sAFtot As Double, st As String, jr As Long, nRol As Long
    Dim cZ As Double, cW As Double, cAE As Double, afwijking As Boolean
    Dim geschreven As Boolean, klaar As Boolean, keuzeFout As String

    If Not HeeftRecht("Toevoegen") Then
        MsgBox "De jaarovergang wordt uitgevoerd door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    ' extra slot: wachtwoord voor de jaarovergang
    Dim jw As String
    jw = InputBox("Wachtwoord voor de jaarovergang:", "Jaarovergang uitvoeren")
    If StrPtr(jw) = 0 Or Len(jw) = 0 Then Exit Sub
    If StrComp(jw, WACHTWOORD_JAAROVERGANG, vbBinaryCompare) <> 0 Then
        MsgBox "Wachtwoord onjuist. De jaarovergang is niet uitgevoerd.", vbExclamation, "Jaarovergang uitvoeren"
        LogRegel "jaarovergang geweigerd", , , "", , , "wachtwoord onjuist"
        Exit Sub
    End If
    oud = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    nieuw = oud + 1
    nm = "Verzamellijst 31-12-" & oud              ' een / mag niet in een tabbladnaam
    If Year(Date) <= oud Then
        If MsgBox("LET OP: het jaar " & oud & " is nog niet verstreken." & vbCrLf & _
                  "De jaarovergang mag pas vanaf 1 januari " & nieuw & " worden uitgevoerd." & vbCrLf & vbCrLf & _
                  "Alleen doorgaan als u dit op een KOPIE van het bestand test. Doorgaan?", _
                  vbYesNo + vbCritical + vbDefaultButton2, "Jaarovergang te vroeg") <> vbYes Then Exit Sub
    End If
    If UvOnverwerkt() > 0 Then
        MsgBox "De jaarovergang kan nog niet worden uitgevoerd: op het blad '" & BL_UITVRAAG & "' staan nog " & UvOnverwerkt() & " regel(s) die niet zijn verwerkt." & vbCrLf & _
               "Verwerk ze met '3. Uitvraag verwerken', of kies 'afwijzen'.", vbExclamation, "Jaarovergang"
        Exit Sub
    End If
    If BladBestaatStil(nm) Then
        MsgBox "Het blad '" & nm & "' bestaat al: de jaarovergang van " & oud & " naar " & nieuw & " is al uitgevoerd?" & vbCrLf & _
               "Er is niets gewijzigd.", vbExclamation
        Exit Sub
    End If
    Dim totBud As Variant, tekstBud As String
    totBud = MeetStap8(Vz(), CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value))
    tekstBud = "- budget t/m " & oud & ": " & Format$(totBud(1), "#,##0;-#,##0;0") & " wordt budget t/m " & nieuw & ": " & _
               Format$(totBud(1) + totBud(5), "#,##0;-#,##0;0") & " (budget " & nieuw & " komt erbij);" & vbCrLf & _
               "- budget " & (oud + 2) & " t/m " & (oud + 5) & " blijft hetzelfde, het staat daarna alleen in andere kolommen; budget " & (oud + 6) & " begint leeg;" & vbCrLf
    If MsgBox("Jaarovergang van " & oud & " naar " & nieuw & " uitvoeren?" & vbCrLf & vbCrLf & _
              "- er komt eerst een back-up '" & nm & "' (stand per 31-12-" & oud & ");" & vbCrLf & _
              "- alle bedragkolommen schuiven een jaar op;" & vbCrLf & tekstBud & _
              "- lopende kredieten L" & oud & " en eerder worden L" & nieuw & "; startjaar en eindjaar " & oud & " worden " & nieuw & " (tenzij er al budget in " & nieuw & " staat)." & vbCrLf & vbCrLf & _
              "Doe dit eerst op een kopie van het bestand.", vbYesNo + vbExclamation + vbDefaultButton2, "Jaarovergang") <> vbYes Then Exit Sub

    On Error GoTo Fout
    Application.ScreenUpdating = False
    Application.Calculate
    Set wv = Vz()
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)

    ' 1. back-up als harde waarden
    wv.Copy Before:=ThisWorkbook.Worksheets(BL_DASH)
    Set wb = ActiveSheet
    wb.Unprotect WACHTWOORD
    wb.Name = nm
    On Error Resume Next
    wb.Buttons.Delete
    If wb.FilterMode Then wb.ShowAllData
    On Error GoTo Fout
    wb.UsedRange.Value = wb.UsedRange.Value
    wb.Range("A1").AddComment "Stand per 31-12-" & oud & ", gemaakt bij de jaarovergang op " & Format$(Now, "dd-mm-yyyy hh:mm") & ". Alleen lezen."
    wb.Tab.Color = RGB(128, 128, 128)
    wb.Protect Password:=WACHTWOORD, AllowFiltering:=True
    wb.Move Before:=ThisWorkbook.Worksheets("Voertuigen")

    ' 2. inlezen (oude waarden bewaren om terug te kunnen zetten)
    oZ = wv.Range("Z2:Z" & laatste).Value: oAA = wv.Range("AA2:AA" & laatste).Value
    oAB = wv.Range("AB2:AB" & laatste).Value: oAC = wv.Range("AC2:AC" & laatste).Value
    oAE = wv.Range("AE2:AI" & laatste).Value: oH = wv.Range("H2:H" & laatste).Value
    oStart = wv.Range("M2:M" & laatste).Value: oEind = wv.Range("N2:N" & laatste).Value
    z = oZ: aa = oAA: ab = oAB: ac = oAC: ae = oAE: h = oH: m = oStart: n = oEind

    For r = 1 To laatste - 1
        sZ = sZ + Nul(oZ(r, 1)): sAE1 = sAE1 + Nul(oAE(r, 1))
        sW = sW + Nul(oAA(r, 1)) + Nul(oAB(r, 1)) + Nul(oAC(r, 1))
        For k = 2 To 5
            sAFtot = sAFtot + Nul(oAE(r, k))
        Next k
        If Not Leeg(wv.Cells(r + 1, 1).Value) Then
            z(r, 1) = Waarde(Nul(oZ(r, 1)) + Nul(oAE(r, 1)), oZ(r, 1), oAE(r, 1))
            aa(r, 1) = Waarde(Nul(oAA(r, 1)) + Nul(oAB(r, 1)), oAA(r, 1), oAB(r, 1))
            ab(r, 1) = oAC(r, 1)
            ac(r, 1) = Empty
            For k = 1 To 4
                ae(r, k) = oAE(r, k + 1)
            Next k
            ae(r, 5) = Empty
            st = UCase$(Trim$(Tkst(oH(r, 1))))
            jr = Val(Mid$(st, 2))
            If Left$(st, 1) = "L" And jr > 0 And jr <= oud Then
                h(r, 1) = "L" & nieuw
                nRol = nRol + 1
                If Nul(oAE(r, 1)) = 0 Then       ' geen budget in het nieuwe jaar: de uitgaven starten pas in het nieuwe jaar
                    If IsGetal(oStart(r, 1)) Then
                        If CLng(oStart(r, 1)) = oud Then m(r, 1) = nieuw
                    End If
                    If IsGetal(oEind(r, 1)) Then
                        If CLng(oEind(r, 1)) = oud Then n(r, 1) = nieuw
                    End If
                End If
            End If
        End If
    Next r

    ' 3. wegschrijven en lopend jaar ophogen
    geschreven = True                            ' vanaf hier wordt bij een fout alles teruggezet
    wv.Range("Z2:Z" & laatste).Value = z: wv.Range("AA2:AA" & laatste).Value = aa
    wv.Range("AB2:AB" & laatste).Value = ab: wv.Range("AC2:AC" & laatste).Value = ac
    wv.Range("AE2:AI" & laatste).Value = ae
    wv.Range("H2:H" & laatste).Value = h
    wv.Range("M2:M" & laatste).Value = m: wv.Range("N2:N" & laatste).Value = n
    ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value = nieuw
    Application.Calculate

    ' 4. controle op de totalen
    For r = 2 To laatste
        cZ = cZ + Nul(wv.Cells(r, 26).Value)
        cW = cW + Nul(wv.Cells(r, 27).Value) + Nul(wv.Cells(r, 28).Value) + Nul(wv.Cells(r, 29).Value)
        For k = 31 To 35
            cAE = cAE + Nul(wv.Cells(r, k).Value)
        Next k
    Next r
    afwijking = (Round(cZ - (sZ + sAE1), 2) <> 0) Or (Round(cW - sW, 2) <> 0) Or (Round(cAE - sAFtot, 2) <> 0)
    If afwijking Then
        wv.Range("Z2:Z" & laatste).Value = oZ: wv.Range("AA2:AA" & laatste).Value = oAA
        wv.Range("AB2:AB" & laatste).Value = oAB: wv.Range("AC2:AC" & laatste).Value = oAC
        wv.Range("AE2:AI" & laatste).Value = oAE: wv.Range("H2:H" & laatste).Value = oH
        wv.Range("M2:M" & laatste).Value = oStart: wv.Range("N2:N" & laatste).Value = oEind
        ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value = oud
        Application.DisplayAlerts = False
        wb.Unprotect WACHTWOORD
        wb.Delete
        Application.DisplayAlerts = True
        Application.Calculate
        Application.ScreenUpdating = True
        MsgBox "De controle op de totalen klopt niet. Alles is teruggezet; er is niets gewijzigd.", vbCritical
        Exit Sub
    End If

    LogRegel "jaarovergang", , , "Par_LopendJaar", oud, nieuw, nRol & " lopende kredieten naar L" & nieuw & "; back-up '" & nm & "'"
    klaar = True                                 ' de jaarovergang is gelukt; fouten hierna draaien die niet meer terug
    On Error Resume Next
    KeuzelijstenVerversen True
    If Err.Number <> 0 Then keuzeFout = Err.Description: Err.Clear
    On Error GoTo Fout
    ControlesVernieuwen
    ' overzichten van voor de jaarovergang tonen nog het oude jaar ('start na <oud jaar>', budgetkolommen): weghalen.
    ' Een nieuw overzicht rekent altijd met het nieuwe lopende jaar. Opgeslagen bestanden in de mappen blijven bewaard.
    Dim oudBlad As Variant
    Application.DisplayAlerts = False
    For Each oudBlad In Array(BL_OVERZICHT, BL_OVERZICHT_NAAM, BL_UITVRAAG)
        If BladBestaatStil(CStr(oudBlad)) Then
            ThisWorkbook.Worksheets(CStr(oudBlad)).Unprotect WACHTWOORD
            ThisWorkbook.Worksheets(CStr(oudBlad)).Delete
        End If
    Next oudBlad
    Application.DisplayAlerts = True
    On Error Resume Next
    ThisWorkbook.Names("Uitvraag_Jaar").Delete                                ' de uitvraag van het oude jaar is afgerond
    On Error GoTo Fout
    Application.ScreenUpdating = True
    NaarEigenDashboard
    MsgBox "Jaarovergang naar " & nieuw & " uitgevoerd." & vbCrLf & vbCrLf & _
           "- de oude bladen 'Laatste kredietoverzicht', 'Laatste overzicht op naam' en 'Uitvraag' zijn weggehaald; " & _
           "maak ze opnieuw (dan met " & nieuw & ")" & vbCrLf & _
           "- back-up: blad '" & nm & "' (bron voor het kredietoverzicht jaarrekening)" & vbCrLf & _
           "- " & nRol & " lopende kredieten naar L" & nieuw & vbCrLf & _
           "- totalen gecontroleerd: budget t/m " & nieuw & " = oud budget + budget " & nieuw & "; werkelijke uitgaven gelijk" & _
           IIf(Len(keuzeFout) > 0, vbCrLf & vbCrLf & "Let op: de keuzelijsten konden niet worden ververst (" & keuzeFout & "). Gebruik de knop Keuzelijsten verversen.", ""), vbInformation
    Exit Sub
Fout:
    Dim foutTekst As String
    foutTekst = Err.Description
    On Error Resume Next
    If klaar Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox "De jaarovergang is gelukt, maar daarna ging iets mis: " & foutTekst & vbCrLf & _
               "Ververs de keuzelijsten en de controles met de knoppen op het Dashboard.", vbExclamation
        Exit Sub
    End If
    If geschreven Then
        ' alles terugzetten: bedragen, status, startjaar, eindjaar en het lopend jaar
        wv.Range("Z2:Z" & laatste).Value = oZ: wv.Range("AA2:AA" & laatste).Value = oAA
        wv.Range("AB2:AB" & laatste).Value = oAB: wv.Range("AC2:AC" & laatste).Value = oAC
        wv.Range("AE2:AI" & laatste).Value = oAE: wv.Range("H2:H" & laatste).Value = oH
        wv.Range("M2:M" & laatste).Value = oStart: wv.Range("N2:N" & laatste).Value = oEind
        ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value = oud
        Application.Calculate
    End If
    ' de (half gemaakte) back-up weer weghalen
    If Not wb Is Nothing Then
        If wb.Name <> wv.Name Then
            Application.DisplayAlerts = False
            wb.Unprotect WACHTWOORD
            wb.Delete
        End If
    End If
    If geschreven Then LogRegel "jaarovergang teruggedraaid", , , "Par_LopendJaar", , , "fout: " & foutTekst
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Fout bij de jaarovergang: " & foutTekst & vbCrLf & "Alles is teruggezet en de back-up is weggehaald. Er is niets gewijzigd.", vbCritical
End Sub

' Leeg laten als beide bronnen leeg waren, anders het bedrag.
Private Function Waarde(ByVal bedrag As Double, ByVal a As Variant, ByVal b As Variant) As Variant
    If Leeg(a) And Leeg(b) Then Waarde = Empty Else Waarde = bedrag
End Function

' Bij openen: is het kalenderjaar verder dan het lopend jaar, dan vragen om de jaarovergang.
Public Sub JaarovergangSignaleren()
    Dim lj As Long
    On Error Resume Next
    lj = CLng(ThisWorkbook.Names("Par_LopendJaar").RefersToRange.Value)
    If lj = 0 Or Year(Date) <= lj Then Exit Sub
    If HeeftRecht("Toevoegen") Then
        If MsgBox("Het is " & Year(Date) & ", maar het lopend jaar in dit bestand is nog " & lj & "." & vbCrLf & _
                  "Jaarovergang nu uitvoeren?", vbYesNo + vbQuestion, "Jaarovergang") = vbYes Then JaarovergangUitvoeren
    Else
        MsgBox "Het lopend jaar in dit bestand is nog " & lj & ". De financieel adviseur moet de jaarovergang nog uitvoeren.", vbInformation
    End If
End Sub

' ======================================================================
' Uitleg bij de knoppen
' ======================================================================
' Naast elke knop op het Dashboard een knopje "?"; klikken toont de uitleg van die knop.
' Is de uitleg onder de vraagtekens van de huidige versie? (anders zet Auto_Open hem opnieuw neer)
Public Function UitlegVersieKlopt() As Boolean
    On Error Resume Next
    UitlegVersieKlopt = (ThisWorkbook.Names("Uitleg_Versie").RefersTo = "=""" & UITLEG_VERSIE & """")
End Function

Public Sub KnoppenUitlegPlaatsen()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(BL_DASH)
    ws.Range("J3:K40").Clear                       ' eerder blok 'Uitleg knoppen' opruimen
    KnopUitleg ws, "btnVerwerken", "Ingevuld? Verwerken!: slaat op wat op de invoerkaart staat. Staat er geen krediet geladen, dan wordt " & _
               "een nieuw krediet onderaan de Verzamellijst toegevoegd met de status 'Te controleren door financi" & ChrW(235) & "le " & _
               "administratie'. Heb je eerst een krediet geladen met 'Krediet wijzigen', dan worden alleen de velden " & _
               "opgeslagen die je hebt aangepast; ook dat krediet moet daarna opnieuw gecontroleerd worden. Eerst " & _
               "controleert de macro of alle verplichte (gele) velden zijn ingevuld. Zo niet, dan krijg je een lijst " & _
               "met wat ontbreekt en wordt er niets opgeslagen. Is het kredietnummer leeg, dan krijgt een nieuw krediet " & _
               "het eerstvolgende vrije nummer van de balanscode. Alles komt in het Logboek. Gaat er tijdens het " & _
               "opslaan iets mis, dan wordt alles teruggezet. Na het opslaan gaat de kaart dicht en sta je weer op het dashboard."
    KnopUitleg ws, "btnKaartWissen", "Scherm leegmaken: maakt de hele invoerkaart leeg, ook de bijdragen van derden, en vergeet een " & _
               "krediet dat je had geladen. Er wordt niets opgeslagen en er verandert niets in de Verzamellijst. Er " & _
               "komt geen vraag vooraf."
    KnopUitleg ws, "btnGebruikersOverzicht", "Overzicht gebruikers (alleen applicatiebeheerder): toont alle gebruikers uit de rollentabel, " & _
               "per rol, met hun Windows-naam en of ze applicatiebeheerder zijn. Er wordt niets gewijzigd."
    KnopUitleg ws, "btnKaartSluiten", "Terug naar het dashboard: de invoerkaart gaat dicht en je ziet weer alle onderdelen. Wat op de kaart staat, " & _
               "blijft bewaard tot je op Scherm leegmaken klikt of een ander krediet opent."
    KnopUitleg ws, "btnKaartOpenen", "Krediet invoeren: de invoerkaart komt tevoorschijn. Vul de velden in en klik onder de kaart op " & _
               "'Ingevuld? Verwerken!'. Na het opslaan gaat de kaart weer dicht. Krediet wijzigen opent de kaart ook."
    KnopUitleg ws, "btnGebruikerToevoegen", "Gebruiker toevoegen (alleen applicatiebeheerder): typ de voornaam, eventueel de Windows-naam, " & _
               "en kies de rol (financieel adviseur of financi" & ChrW(235) & "le administratie). De rechten horen bij de rol. Komt in het Logboek."
    KnopUitleg ws, "btnWijzigen", "Krediet wijzigen: je kiest eerst waarop je zoekt (1 = kredietnummer, 2 = omschrijving, 3 = " & _
               "activanummer, 4 = volgnummer) en typt dan wat je zoekt. Zijn er meer treffers, dan kies je er een " & _
               "uit de lijst. De invoerkaart gaat open met het krediet erop. Er is dan nog niets gewijzigd: pas de velden aan " & _
               "en klik op 'Ingevuld? Verwerken!' om op te slaan. Staat er al een krediet op de kaart, dan vraagt de " & _
               "macro eerst of je dat wilt vervangen."
    KnopUitleg ws, "btnVerwijderen", "Krediet verwijderen (alleen financieel adviseur): staat er een krediet op de invoerkaart, dan wordt " & _
               "dat krediet gebruikt. Anders kies je eerst waarop je zoekt (kredietnummer, omschrijving, " & _
               "activanummer of volgnummer). De macro toont het volgnummer, de omschrijving en het aantal regels en " & _
               "vraagt of je het helemaal wilt verwijderen (standaard Nee). Een krediet met werkelijke uitgaven of " & _
               "inkomsten wordt nooit verwijderd: sluit het dan af met status A. Na Ja verdwijnt de regel uit de " & _
               "Verzamellijst en schuiven de regels eronder een plek op. Bijdrageregels van dat krediet gaan mee als " & _
               "er geen andere regel met dat kredietnummer overblijft. De gegevens blijven bewaard in het Logboek. " & _
               "Gaat er iets mis, dan wordt alles teruggezet."
    KnopUitleg ws, "btnTijdschrijvers", "Tijdschrijvers kiezen: toont de namen van de tijdschrijvers, 20 per pagina. Typ de nummers van de " & _
               "namen die je wilt, gescheiden door komma's. Namen die al in het veld staan zijn voorgevuld en namen " & _
               "die niet in de lijst staan blijven behouden. Daarna kun je nog een naam typen die niet in de lijst " & _
               "staat. Het resultaat komt in het veld Tijdschrijvers op de invoerkaart. Er wordt pas iets opgeslagen " & _
               "als je op 'Ingevuld? Verwerken!' klikt."
    KnopUitleg ws, "btnExport", "Export naar financieel pakket (alleen applicatiebeheerder): maakt een bestand waarmee de kredieten en budgetten in het " & _
               "boekhoudpakket (Coda, later Centric) worden geladen. Wordt nog gebouwd (stap 10)."
    KnopUitleg ws, "btnUitgaven", "Import werkelijke uitgaven (alleen applicatiebeheerder): werkt de werkelijke uitgaven en inkomsten van het lopende jaar bij (tot en met " & _
               "31 maart ook die van het vorige jaar). De macro zoekt in de map 'export coda' het nieuwste " & _
               "Excelbestand en vraagt of je dat wilt uploaden (Ja of Nee). Daarna worden de bedragen per " & _
               "kredietnummer opgeteld en gecontroleerd. Sluit een jaar niet aan, dan wordt er niets bijgewerkt. " & _
               "Anders komen de bedragen in de Verzamellijst. Wat opvalt, zie je in het venster na afloop; blijvende " & _
               "fouten staan bij de controles op het dashboard. Ten slotte wordt het kredietoverzicht raad opnieuw " & _
               "gemaakt en bewaard in map 03."
    KnopUitleg ws, "btnOverzichtNaam", "Kredietoverzicht naar rol: je kiest eerst 1 = budgetbeheerder, 2 = projectleider of 3 = tijdschrijver " & _
               "en daarna de naam (typ een nummer of een deel van de naam). Het overzicht toont de lopende kredieten " & _
               "van die persoon, de kredieten die dit jaar zijn afgesloten en de kredieten die nog niet zijn gestart " & _
               "(in een eigen sectie). Het komt als blad 'Laatste overzicht op naam' (een vorig overzicht wordt " & _
               "vervangen) en wordt ook als los bestand met datum bewaard in map 02, in de submap van die persoon. " & _
               "De budgethouder kan daarin de projectleider (kolom C) aanpassen en onder elk krediet een actuele " & _
               "aanvulling op de toelichting geven (groene cel)."
    KnopUitleg ws, "btnOverzichtRaad", "Kredietoverzicht raad: maakt het overzicht voor de raad met de stand van vandaag. Het toont de " & _
               "lopende kredieten tot en met dit jaar en de kredieten die dit jaar zijn afgesloten. Subsidies en " & _
               "bijdragen (negatief budget) en kredieten zonder budget en zonder uitgaven laat het weg; opbrengsten " & _
               "van activa blijven staan. Het komt als blad 'Laatste kredietoverzicht' (een vorig overzicht wordt " & _
               "vervangen) en wordt ook als los bestand met datum bewaard in map 03."
    KnopUitleg ws, "btnOverzichtJaarrekening", "Kredietoverzicht bij jaarafsluiting: maakt het overzicht met de stand per 31 december, uit de back-up die " & _
               "bij de jaarovergang is gemaakt (blad 'Verzamellijst 31-12-<jaar>', het nieuwste jaar). Het gebruikt " & _
               "dezelfde regels als het overzicht voor de raad. Het wordt als los bestand bewaard in map 03. Bestaat " & _
               "er nog geen back-up, dan vraagt de macro of je een proef wilt zien op basis van de huidige lijst. " & _
               "Die proef wordt niet bewaard."
    KnopUitleg ws, "btnAlleBudgethouders", "Overzichten alle budgethouders: maakt voor elke budgetbeheerder die in de Verzamellijst staat een " & _
               "kredietoverzicht naar rol (zelfde inhoud als bij 'Kredietoverzicht naar rol') en bewaart elk overzicht " & _
               "als los bestand met datum in map 02, in de submap van die persoon. Eerst vraagt de macro of het mag, " & _
               "met het aantal budgetbeheerders. Er worden geen e-mails verstuurd. Staat een bestand open of is een " & _
               "map niet bereikbaar, dan meldt de macro aan het eind welke niet zijn gelukt."
    KnopUitleg ws, "btnTijdschrijven", "Overzicht tijdschrijven (alleen applicatiebeheerder): voor de beheerder tijdschrijven. Het toont alleen de kredieten waar iets " & _
               "mee moet gebeuren: nieuwe, gewijzigde en afgesloten kredieten, met de kolommen Kredietnummer, " & _
               "Omschrijving, Status, Tijdschrijvers, Actie en Toegevoegd op. Het vergelijkt met de stand van de " & _
               "vorige keer; de eerste keer is alles nieuw. Bij een project met deelkredieten staat alleen het hoofddeel " & _
               "in de lijst. Het komt als blad 'Laatste overzicht tijdschrijven' (vervangt het vorige) en als los " & _
               "bestand in map 01. De nieuwe stand wordt onthouden voor de volgende keer."
    KnopUitleg ws, "btnUitvraag", "1. Uitvraag versturen: eerst kies je voor wie: 0 = alle budgetbeheerders, of het nummer van een " & _
               "budgetbeheerder (of een deel van de naam). Daarna vraagt de macro om bevestiging. Per gekozen " & _
               "budgetbeheerder met lopende kredieten komt een bestand in map 02, in de submap van die persoon; " & _
               "daarin kan de teamleider de budgetten van de jaren na dit jaar aanpassen en toelichten. In deze " & _
               "werkmap komt het blauwe blad 'Uitvraag' met de budgetbeheerder(s) en de datum van versturen. Kies je " & _
               "een enkele budgetbeheerder, dan worden alleen diens regels vervangen en blijven de anderen staan; de " & _
               "macro waarschuwt eerst als er al iets van die persoon staat. Er wordt niets gemaild: je geeft het " & _
               "bestand zelf door. Het lopende jaar wordt vastgelegd."
    KnopUitleg ws, "btnUvOphalen", "2. Aanleveringen ophalen: de macro kijkt in map 02 welke budgetbeheerders het bestand na het " & _
               "versturen hebben opgeslagen en sinds het vorige ophalen iets nieuws hebben aangeleverd. Alleen die " & _
               "staan in de lijst; je kiest er een (typ het nummer). Alleen de regels van die budgetbeheerder komen " & _
               "op het blad 'Uitvraag'; regels van anderen blijven ongemoeid. Regels die jij al hebt aangepast of " & _
               "waarover je al hebt besloten worden niet overschreven. Wijkt de nieuwe aanlevering dan af, dan " & _
               "krijgt de regel het signaal 'aanlevering gewijzigd na beoordeling'. De statusregel van elke " & _
               "budgetbeheerder (rood = nog niet ingeleverd) wordt hierbij bijgewerkt. Er wordt niets in de " & _
               "Verzamellijst gewijzigd. Ophalen kan alleen in het jaar waarin de uitvraag is gemaakt."
    KnopUitleg ws, "btnUvVerwerkAlle", "3. Aanleveringen verwerken: verwerkt alle regels op het blad 'Uitvraag' (van alle budgetbeheerders) die " & _
               "nog geen status hebben en waarbij het besluit goedkeuren of afwijzen is. De macro toont eerst de " & _
               "aantallen en vraagt om bevestiging. Bij goedkeuren gaan de budgetten per jaar, de omschrijving en de " & _
               "toelichting naar de Verzamellijst. Is er iets gewijzigd, dan krijgt dat krediet direct de status 'Te " & _
               "controleren door financi" & ChrW(235) & "le administratie'. Nieuwe kredieten gaan naar het blad 'Voorstellen nieuwe " & _
               "kredieten'; die voer je zelf in via de invoerkaart. Alles komt in het register 'Mutaties " & _
               "kadernota-begroting' en het Logboek. Klopt de nameting niet of gaat er iets mis, dan wordt alles " & _
               "teruggezet. Regels met 'nog niet' blijven open."
    KnopUitleg ws, "btnJaarovergang", "Jaarovergang (alleen applicatiebeheerder, met wachtwoord): pas uitvoeren vanaf 1 januari en test het eerst op een " & _
               "kopie. De macro vraagt eerst om bevestiging. Dan: 1) er komt een back-up als blad 'Verzamellijst " & _
               "31-12-<jaar>' met de stand per 31 december; 2) alle bedragen schuiven een jaar op; 3) lopende " & _
               "kredieten van het oude jaar krijgen status L met het nieuwe jaar; startjaar en eindjaar met het oude " & _
               "jaar worden het nieuwe jaar, tenzij er al budget in het nieuwe jaar stond; 4) het lopende jaar gaat " & _
               "met 1 omhoog; 5) de totalen worden gecontroleerd. Klopt iets niet, dan wordt alles teruggezet. De " & _
               "bladen 'Laatste ...' met overzichten worden weggehaald; maak ze daarna opnieuw. De jaarovergang " & _
               "wordt geblokkeerd zolang er regels op het blad 'Uitvraag' staan die nog niet zijn verwerkt."
    KnopUitleg ws, "btnControles", "Controles vernieuwen: rekent de controles in blok 5 opnieuw uit (en werkt de vrije kredietnummers bij). Per controle zie je het aantal en " & _
               "waar in de Verzamellijst het zit (rood = nakijken). Er wordt niets gewijzigd."
    KnopUitleg ws, "btnNaarTabblad", "Ga naar tabblad: opent het tabblad dat je links ervan (cel F2) hebt gekozen. Er wordt niets gewijzigd."
    KnopUitleg ws, "btnAfgekeurd", "Afgekeurd krediet herstellen (financieel adviseur): de financi" & ChrW(235) & "le administratie kan een krediet " & _
               "afkeuren. Met deze knop kies je zo'n krediet (bij meer dan een uit een lijst). Het komt op de " & _
               "invoerkaart en een melding toont wie het afkeurde en waarom. Pas het aan en klik op 'Ingevuld? " & _
               "Verwerken!'. Daarna vraagt de macro of je het opnieuw aanbiedt ter controle. Zijn er geen afgekeurde " & _
               "kredieten, dan meldt de macro dat."
    KnopUitleg ws, "btnHercontrole", "Aanbieden voor hercontrole (financieel adviseur): biedt een hersteld, afgekeurd krediet opnieuw aan. " & _
               "Staat er een afgekeurd krediet op de invoerkaart, dan wordt dat gebruikt; anders kies je uit een " & _
               "lijst. De macro vraagt om bevestiging. Na Ja krijgt het krediet weer de status 'Te controleren door " & _
               "financi" & ChrW(235) & "le administratie'. De reden van de afkeuring blijft bij het krediet staan met 'hersteld na " & _
               "afkeuring'. De financi" & ChrW(235) & "le administratie heeft altijd het laatste woord. Het komt in het Logboek."
    KnopUitleg ws, "btnBhVervangen", "Budgethouder vervangen: je typt de naam die vervangen moet worden (precies zoals in de lijst) en de " & _
               "naam van de opvolger. Daarna kies je of het alleen voor lopende kredieten (L) geldt, of ook voor " & _
               "afgesloten kredieten (A). De macro toont in hoeveel cellen de naam voorkomt en vraagt om " & _
               "bevestiging. Na Ja wordt de naam in de kolom Budgetbeheerder van de Verzamellijst vervangen. Elke " & _
               "wijziging komt in het Logboek."
    KnopUitleg ws, "btnPlVervangen", "Projectleider vervangen: je typt de naam die vervangen moet worden (precies zoals in de lijst) en de " & _
               "naam van de opvolger. Daarna kies je of het alleen voor lopende kredieten (L) geldt, of ook voor " & _
               "afgesloten kredieten (A). De macro toont in hoeveel cellen de naam voorkomt en vraagt om " & _
               "bevestiging. Na Ja wordt de naam in de kolom Projectleider van de Verzamellijst vervangen. Elke " & _
               "wijziging komt in het Logboek."
    KnopUitleg ws, "btnTsVervangen", "Tijdschrijver vervangen (alleen applicatiebeheerder): je typt de naam die vervangen moet worden (precies zoals in de lijst) en de " & _
               "naam van de opvolger. Daarna kies je of het alleen voor lopende kredieten (L) geldt, of ook voor " & _
               "afgesloten kredieten (A). De macro toont in hoeveel cellen de naam voorkomt en vraagt om " & _
               "bevestiging. Na Ja wordt de naam vervangen in de kolom Tijdschrijvers van de Verzamellijst; andere " & _
               "tijdschrijvers in dezelfde cel blijven staan. Elke wijziging komt in het Logboek."
    KnopUitleg ws, "btnGebruikerVerwijderen", "Gebruiker verwijderen (alleen applicatiebeheerder): toont een genummerde lijst met de gebruikers uit " & _
               "de rollentabel (blad Parameters) met hun functie. Je typt het nummer van de persoon die je wilt " & _
               "verwijderen. Jezelf verwijderen kan niet. Daarna vraagt de macro om bevestiging (standaard Nee). Na " & _
               "Ja verdwijnt die persoon uit de rollentabel. Er blijft altijd minstens " & ChrW(233) & ChrW(233) & "n applicatiebeheerder over. " & _
               "Wie niet in de tabel staat, kan zichzelf bij het openen " & _
               "van het bestand weer toevoegen. Het wordt vastgelegd in het Logboek."
    ThisWorkbook.Names.Add Name:="Uitleg_Versie", RefersTo:="=""" & UITLEG_VERSIE & """", Visible:=False
End Sub

' Klein knopje "?" naast de knop. Klikken toont de uitleg in een venster. Dit is een gewone formulierknop, net als de
' andere knoppen, zodat het ook op het beveiligde Dashboard werkt (scherminfo bij zweven is in Excel niet betrouwbaar).
Public Sub KnopUitleg(ByVal ws As Worksheet, ByVal knop As String, ByVal tekst As String)
    Dim b As Object, q As Object, had As Boolean
    On Error Resume Next
    Set b = ws.Buttons(knop)
    If b Is Nothing Then Exit Sub
    Set q = ws.Shapes("uitleg_" & knop)
    had = Not q Is Nothing
    ' knop smaller maken voor het vraagteken, tenzij dat al gebeurd is (vraagteken staat dan al rechts naast de knop)
    If had Then
        If q.Left < b.Left + b.Width Then b.Width = b.Width - 20
        q.Delete
    Else
        b.Width = b.Width - 20
    End If
    Set q = ws.Buttons.Add(b.Left + b.Width + 2, b.Top, 18, b.Height)
    q.Name = "uitleg_" & knop
    q.Caption = "?"
    q.OnAction = "KnopUitlegTonen"
    q.Placement = xlMove                         ' beweegt mee met de knop als rijen open of dicht gaan
    UitlegBewaren ws, "uitleg_" & knop, tekst
End Sub

' De uitleg wordt bewaard in een verborgen naam per knop, zodat hij ook na sluiten en openen beschikbaar is.
' Een tekst in een formule mag maximaal 255 tekens zijn: langere uitleg wordt in stukken opgeslagen ("..."&"...").
Private Sub UitlegBewaren(ByVal ws As Worksheet, ByVal naam As String, ByVal tekst As String)
    Dim f As String
    tekst = Replace(tekst, """", "'")
    Do While Len(tekst) > 250
        f = f & """" & Left$(tekst, 250) & """&"
        tekst = Mid$(tekst, 251)
    Loop
    f = "=" & f & """" & tekst & """"
    ThisWorkbook.Names.Add Name:="Uitleg_" & Mid$(naam, Len("uitleg_") + 1), RefersTo:=f, Visible:=False
End Sub

' Klik op een vraagteken: toont de uitleg van de knop ernaast.
Public Sub KnopUitlegTonen()
    Dim naam As String, knop As String, tekst As String, titel As String
    On Error Resume Next
    naam = CStr(Application.Caller)
    knop = Mid$(naam, Len("uitleg_") + 1)
    tekst = ThisWorkbook.Names("Uitleg_" & knop).RefersTo          ' ="tekst" of ="deel 1"&"deel 2"
    tekst = Replace(Mid$(tekst, 3, Len(tekst) - 3), """&""", "")
    titel = ActiveSheet.Buttons(knop).Caption
    On Error GoTo 0
    If Len(tekst) = 0 Then Exit Sub
    MsgBox tekst, vbInformation, "Uitleg" & IIf(Len(titel) > 0, ": " & titel, "")
End Sub

' ======================================================================
' Blauwe overzichtbladen: pdf maken en mailen
' ======================================================================
' Blauwe tabbladen: Laatste kredietoverzicht, Laatste overzicht op naam, Uitvraag en
' Laatste overzicht tijdschrijven. Op elk blad staan de knoppen "Opslaan als pdf" en "Mailen als pdf".
' De kolom Melding (alleen voor financien) staat nooit in de pdf.

Public Function IsBlauwBlad(ByVal naam As String) As Boolean
    IsBlauwBlad = (naam = BL_OVERZICHT Or naam = BL_OVERZICHT_NAAM Or naam = BL_UITVRAAG Or naam = BL_TIJD)
End Function

' Kolomnummer van de kolom 'Melding' (kopregel 4 of 5); 0 als die er niet is.
Public Function KolomMelding(ByVal ws As Worksheet) As Long
    Dim c As Long, r As Long
    For r = 4 To 5
        For c = 1 To ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            If StrComp(Trim$(CStr(ws.Cells(r, c).Value)), "Melding", vbTextCompare) = 0 Then KolomMelding = c: Exit Function
        Next c
    Next r
End Function

' Kopregel van het blad (de regel met de kolomnamen).
Private Function KopRijBlad(ByVal ws As Worksheet) As Long
    KopRijBlad = IIf(ws.Name = BL_TIJD, 4, 5)
End Function

' Voorgestelde bestandsnaam (zonder .pdf) en submap van 'Stambestand kredieten'.
Private Sub PdfNaam(ByVal ws As Worksheet, ByRef naam As String, ByRef mapNaam As String)
    Dim t As Variant
    Select Case ws.Name
        Case BL_OVERZICHT
            naam = "Kredietoverzicht " & CStr(ws.Range("A2").Value)
            mapNaam = "03 kredietoverzichten"
        Case BL_OVERZICHT_NAAM
            naam = "Kredietoverzicht " & Replace(CStr(ws.Range("A2").Value), ": ", " ")
            mapNaam = "02 Budgetbeheerders en projectleiders"
        Case BL_UITVRAAG
            naam = "Uitvraag kadernota-begroting " & Format$(Date, "dd-mm-yyyy")
            mapNaam = "02 Budgetbeheerders en projectleiders"
        Case Else
            naam = "Tijdschrijven kredieten " & Format$(Date, "dd-mm-yyyy")
            mapNaam = "01 Applicatiebeheer Tijdschrijven"
    End Select
    ' financiele administratie (werkt vanuit Dashboard Adm): pdf van een overzicht naar de eigen FinAdm-map
    If EigenDashboard.Name = BL_ADM Then
        If ws.Name = BL_OVERZICHT_NAAM Then mapNaam = "FinAdm Kredietoverzicht op naam"
        If ws.Name = BL_OVERZICHT Then mapNaam = "FinAdm Kredietoverzicht jaarrekening"
    End If
    For Each t In Array("\", "/", "?", "*", "[", "]", ":", """", "<", ">", "|")
        naam = Replace(naam, CStr(t), "-")
    Next t
    naam = Left$(Trim$(naam), 120)
End Sub

' Maakt de pdf van het blad: afdrukbereik, kopregel op elke pagina, liggend, passend op de breedte, kolom Melding verborgen.
Private Function PdfMaken(ByVal ws As Worksheet, ByVal pad As String) As Boolean
    Dim beveiligd As Boolean, cM As Long, mVerborgen As Boolean, kop As Long, lc As Long, lr As Long, c As Long, b As Object
    On Error GoTo Fout
    Application.ScreenUpdating = False
    beveiligd = ws.ProtectContents
    If beveiligd Then ws.Unprotect WACHTWOORD
    kop = KopRijBlad(ws)
    cM = KolomMelding(ws)
    If cM > 0 Then
        mVerborgen = ws.Columns(cM).Hidden
        ws.Columns(cM).Hidden = True
    End If
    ' laatste kolom en regel van het overzicht
    lc = ws.Cells(kop, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lc
        If ws.Cells(ws.Rows.Count, c).End(xlUp).Row > lr Then lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    Next c
    ' vervolgregels van een vorige pdf weghalen
    VervolgRegelsWeg ws, kop, lr
    ' rijhoogte opnieuw bepalen (een toelichting kan met de hand zijn verbeterd)
    BladOpmaken ws, kop, kop + 1, lr
    For Each b In ws.Buttons
        b.PrintObject = False
    Next b
    On Error Resume Next                         ' paginainstellingen kunnen per printer verschillen
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Address
        .PrintTitleRows = IIf(ws.Name = BL_UITVRAAG, "$4:$5", "$" & kop & ":$" & kop)
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .HeaderMargin = Application.CentimetersToPoints(0.6)
        .FooterMargin = Application.CentimetersToPoints(0.6)
        .CenterHorizontally = True
        .PrintGridlines = False
        .CenterFooter = "Pagina &P van &N"
    End With
    Application.PrintCommunication = True
    On Error GoTo Fout
    ' pagina's netjes afbreken: elke pagina begint met de donkerblauwe kopregel en een lichtblauwe regel
    PaginasIndelen ws, kop, lr, lc
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pad, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    PdfMaken = True
Einde:
    On Error Resume Next
    If cM > 0 Then ws.Columns(cM).Hidden = mVerborgen
    If beveiligd Then BladBeveiligen ws
    Application.ScreenUpdating = True
    Exit Function
Fout:
    Application.PrintCommunication = True
    MsgBox "De pdf kon niet worden gemaakt: " & Err.Description, vbCritical
    Resume Einde
End Function

' Lichtblauwe regel (programma, budgetbeheerder, 'Nog niet gestart', blok samengestelde kredieten).
Private Function IsLichtblauw(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    IsLichtblauw = (ws.Cells(r, 1).Interior.Color = RGB(214, 228, 240)) And Len(CStr(ws.Cells(r, 1).Value)) > 0
End Function

' Kopje van een investeringsgroep of project (kolom A leeg, kolom B vet en schuin).
Private Function IsGroepKopje(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    IsGroepKopje = Leeg(ws.Cells(r, 1).Value) And Len(CStr(ws.Cells(r, 2).Value)) > 0 And _
                   ws.Cells(r, 2).Font.Bold = True And ws.Cells(r, 2).Font.Italic = True
End Function

' Haalt de regels '... (vervolg)' van een vorige pdf weg (alleen lichtblauwe regels met die tekst).
Private Sub VervolgRegelsWeg(ByVal ws As Worksheet, ByVal kop As Long, ByRef lr As Long)
    Dim r As Long
    ws.ResetAllPageBreaks
    For r = lr To kop + 1 Step -1
        If IsLichtblauw(ws, r) And CStr(ws.Cells(r, 1).Value) Like "* (vervolg)" Then
            ws.Rows(r).Delete
            lr = lr - 1
        End If
    Next r
End Sub

' Zet de pagina-einden zelf (A3 liggend, passend op de breedte). Elke pagina begint met de kopregel (herhaald) en
' daaronder een lichtblauwe regel. Afbreken gebeurt bij voorkeur voor een lichtblauwe regel. Past een programma niet
' op een pagina, dan wordt afgebroken voor een groepskopje (of anders een gewone regel) en komt bovenaan de nieuwe
' pagina een lichtblauwe regel '<programma> (vervolg)'.
Private Sub PaginasIndelen(ByVal ws As Worksheet, ByVal kop As Long, ByRef lr As Long, ByVal lc As Long)
    Dim c As Long, r As Long, breed As Double, schaal As Double, titelH As Double, ruimte As Double, gebruikt As Double
    Dim h As Double, start As Long, lb As Long, gk As Long, brk As Long, bron As Long, eersteTitel As Long, tekst As String
    Dim nPag As Long
    On Error GoTo klaar
    ws.ResetAllPageBreaks
    For c = 1 To lc
        If Not ws.Columns(c).Hidden Then breed = breed + ws.Columns(c).Width
    Next c
    If breed <= 0 Then Exit Sub
    ' A3 liggend 1190,6 x 841,9 punten; marges links/rechts 1 cm, boven/onder 1,5 cm
    schaal = (1190.6 - 2 * 28.35) / breed
    If schaal > 1 Then schaal = 1
    schaal = Int(schaal * 100) / 100
    eersteTitel = IIf(ws.Name = BL_UITVRAAG, 4, kop)
    For r = eersteTitel To kop
        titelH = titelH + ws.Rows(r).Height
    Next r
    ruimte = (841.9 - 2 * 42.5) / schaal * 0.96 - titelH      ' 4% veiligheidsmarge
    If ruimte < 100 Then Exit Sub
    For r = 1 To eersteTitel - 1                              ' titel bovenaan de eerste pagina
        gebruikt = gebruikt + ws.Rows(r).Height
    Next r
    start = kop + 1: r = kop + 1
    Do While r <= lr
        h = ws.Rows(r).Height
        If r > start Then
            If IsLichtblauw(ws, r) Then lb = r
            If IsGroepKopje(ws, r) Then gk = r
        End If
        If gebruikt + h > ruimte And r > start Then
            If lb > start Then
                brk = lb
            Else
                If gk > start Then brk = gk Else brk = r
                ' lichtblauwe vervolgregel met het programma (of de budgetbeheerder) erboven
                bron = 0
                For c = brk - 1 To kop + 1 Step -1
                    If IsLichtblauw(ws, c) Then bron = c: Exit For
                Next c
                If bron > 0 Then
                    tekst = CStr(ws.Cells(bron, 1).Value)
                    If Not tekst Like "* (vervolg)" Then tekst = tekst & " (vervolg)"
                    ws.Rows(brk).Insert
                    ws.Rows(bron).Copy Destination:=ws.Rows(brk)
                    ws.Range(ws.Cells(brk, 2), ws.Cells(brk, lc)).ClearContents
                    ws.Cells(brk, 1).Value = tekst
                    ws.Rows(brk).RowHeight = ws.Rows(bron).Height
                    lr = lr + 1
                End If
            End If
            ws.HPageBreaks.Add Before:=ws.Rows(brk)
            nPag = nPag + 1
            If nPag > 500 Then Exit Do                           ' noodrem
            start = brk: r = brk: gebruikt = 0: lb = 0: gk = 0
        Else
            gebruikt = gebruikt + h
            r = r + 1
        End If
    Loop
klaar:
End Sub

' Knop "Opslaan als pdf" op een blauw blad.
Public Sub BladNaarPdf()
    Dim ws As Worksheet, naam As String, sm As String, map As String, pad As Variant
    Set ws = ActiveSheet
    If Not IsBlauwBlad(ws.Name) Then Exit Sub
    PdfNaam ws, naam, sm
    map = MapStambestand(sm)
    pad = Application.GetSaveAsFilename(InitialFileName:=map & naam & ".pdf", _
                                        FileFilter:="PDF-bestand (*.pdf), *.pdf", Title:="Pdf opslaan")
    If VarType(pad) = vbBoolean Then Exit Sub    ' geannuleerd
    If LCase$(Right$(CStr(pad), 4)) <> ".pdf" Then pad = pad & ".pdf"
    If PdfMaken(ws, CStr(pad)) Then
        LogRegel "pdf gemaakt", , , "", , , ws.Name & " | " & pad
        MsgBox "Pdf opgeslagen:" & vbCrLf & pad, vbInformation
    End If
End Sub

' Knop "Mailen als pdf" op een blauw blad: maakt de pdf en opent een nieuw Outlook-bericht met de pdf als bijlage.
' Het bericht wordt niet vanzelf verstuurd: vul de ontvanger in en klik zelf op Verzenden.
Public Sub BladMailen()
    Dim ws As Worksheet, naam As String, sm As String, pad As String, ol As Object, m As Object
    Set ws = ActiveSheet
    If Not IsBlauwBlad(ws.Name) Then Exit Sub
    PdfNaam ws, naam, sm
    pad = Environ$("TEMP") & "\" & naam & ".pdf"
    If Not PdfMaken(ws, pad) Then Exit Sub
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        MsgBox "Outlook kon niet worden gestart (de nieuwe Outlook ondersteunt dit niet)." & vbCrLf & _
               "De pdf staat klaar in:" & vbCrLf & pad & vbCrLf & "Voeg die zelf toe als bijlage.", vbExclamation
        Exit Sub
    End If
    Set m = ol.CreateItem(0)
    m.Subject = naam
    m.Attachments.Add pad
    m.Display                                    ' eerst tonen, dan staat de handtekening er al in
    m.HTMLBody = "<p style='font-family:Arial;font-size:10pt'>Beste collega,<br><br>Bijgaand het " & _
                 LCase$(Left$(CStr(ws.Range("A1").Value), 1)) & Mid$(CStr(ws.Range("A1").Value), 2) & _
                 IIf(ws.Name = BL_UITVRAAG, "", " (" & CStr(ws.Range("A2").Value) & ")") & ".</p>" & m.HTMLBody
    LogRegel "pdf gemaild", , , "", , , ws.Name & " | " & naam
End Sub

' ======================================================================
' Stap 8: dynamische bereiken robuust maken (eenmalig, via Ontwikkelaars > Macro's)
' ======================================================================
' 1. Krediet_Aantal telt niet meer de gevulde cellen in kolom A, maar gaat uit van de laatste gevulde rij (Par_LaatsteRij).
'    Een lege rij tussen de kredieten maakt de bereiken Krediet_J, Krediet_AL, Krediet_AM dan niet meer te kort.
' 2. Harde getallen in kolom K worden weer de formule, maar alleen als de uitkomst precies gelijk blijft.
' 3. Koppen U, X en Y krijgen een duidelijke naam; kolom E en G worden verborgen.
' Nulmeting vooraf en nameting achteraf; bij elk verschil gaat alles terug.
Public Sub Stap8DynamischeBereiken()
    WerkStart                                    ' beveiliging even uit, daarna vanzelf weer aan
    Dim wv As Worksheet, laatste As Long, r As Long, i As Long, rf As Long
    Dim voor As Variant, na As Variant, kVoor As Variant, amVoor As Variant, kNa As Variant, amNa As Variant
    Dim oudNaam As String, formR1C1 As String, hard As Object, hersteld As String, nOm As Long, nHard As Long, verschil As String
    Dim kopOud(1 To 3) As Variant, kopKol As Variant

    On Error GoTo Fout
    If Not HeeftRecht("Toevoegen") Then
        MsgBox "Dit wordt gedaan door de financieel adviseur.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Stap 8: dynamische bereiken robuust maken, harde getallen in kolom K weer formule (alleen als de uitkomst gelijk blijft), " & _
              "koppen U, X en Y hernoemen en kolom E en G verbergen." & vbCrLf & vbCrLf & _
              "Vooraf wordt alles gemeten; bij elk verschil gaat alles terug. Doorgaan?", vbYesNo + vbQuestion, "Stap 8") <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    Set wv = Vz()
    wv.Unprotect WACHTWOORD
    Application.Calculate
    laatste = CLng(ThisWorkbook.Names("Par_LaatsteRij").RefersToRange.Value)

    ' 1. nulmeting
    voor = MeetStap8(wv, laatste)
    kVoor = wv.Range("K2:K" & laatste).Value
    amVoor = wv.Range("AM2:AM" & laatste).Value
    oudNaam = ThisWorkbook.Names("Krediet_Aantal").RefersTo
    kopKol = Array(21, 24, 25)
    For i = 1 To 3
        kopOud(i) = wv.Cells(1, kopKol(i - 1)).Formula
    Next i

    ' 2. Krediet_Aantal vanaf de laatste gevulde rij
    ThisWorkbook.Names("Krediet_Aantal").RefersTo = "=Par_LaatsteRij-1"
    Application.Calculate
    kNa = wv.Range("K2:K" & laatste).Value
    amNa = wv.Range("AM2:AM" & laatste).Value
    For r = 1 To UBound(kVoor, 1)
        If Not GelijkStap8(kVoor(r, 1), kNa(r, 1)) Or Not GelijkStap8(amVoor(r, 1), amNa(r, 1)) Then
            verschil = "rij " & (r + 1) & " kolom K of AM": Exit For
        End If
    Next r
    If Len(verschil) > 0 Then GoTo Terug

    ' 3. harde getallen in K: formule terug als de uitkomst gelijk blijft
    For r = 2 To laatste
        If wv.Cells(r, 11).HasFormula Then rf = r: Exit For
    Next r
    If rf = 0 Then verschil = "geen formule in kolom K gevonden om te kopieren": GoTo Terug
    formR1C1 = wv.Cells(rf, 11).FormulaR1C1
    Set hard = CreateObject("Scripting.Dictionary")
    For r = 2 To laatste
        If Not wv.Cells(r, 11).HasFormula And Not IsEmpty(wv.Cells(r, 1).Value) Then
            hard(r) = wv.Cells(r, 11).Value
            nHard = nHard + 1
            wv.Cells(r, 11).FormulaR1C1 = formR1C1
        End If
    Next r
    Application.Calculate
    For Each kopKol In hard.Keys
        r = CLng(kopKol)
        If GelijkStap8(hard(r), wv.Cells(r, 11).Value) Then
            nOm = nOm + 1
        Else
            wv.Cells(r, 11).Value = hard(r)                 ' uitkomst wijkt af: hard getal blijft staan
            hersteld = hersteld & IIf(Len(hersteld) > 0, ", ", "") & "K" & r
        End If
    Next kopKol
    Application.Calculate

    ' 4. koppen en verborgen kolommen
    wv.Cells(1, 21).Value = "Code organisatieonderdeel"
    wv.Cells(1, 24).Value = "Activanr"
    wv.Cells(1, 25).Value = "Afschrijving"
    wv.Columns("E").Hidden = True
    wv.Columns("G").Hidden = True
    mStil = True
    KoppenVerzamellijstPassend                   ' kopregel past na het hernoemen
    mStil = False
    wv.Unprotect WACHTWOORD

    ' 5. nameting: alles moet gelijk zijn aan de nulmeting
    Application.Calculate
    na = MeetStap8(wv, laatste)
    For i = 0 To 9
        If Round(na(i) - voor(i), 2) <> 0 Then verschil = "som of aantal (meting " & i & ")": Exit For
    Next i
    If Len(verschil) = 0 Then
        kNa = wv.Range("K2:K" & laatste).Value
        amNa = wv.Range("AM2:AM" & laatste).Value
        For r = 1 To UBound(kVoor, 1)
            If Not GelijkStap8(kVoor(r, 1), kNa(r, 1)) Or Not GelijkStap8(amVoor(r, 1), amNa(r, 1)) Then
                verschil = "rij " & (r + 1) & " kolom K of AM": Exit For
            End If
        Next r
    End If
    If Len(verschil) > 0 Then GoTo Terug

    LogRegel "stap 8 dynamische bereiken", , , "Krediet_Aantal", oudNaam, "=Par_LaatsteRij-1", _
             nOm & " van " & nHard & " harde getallen in K weer formule" & IIf(Len(hersteld) > 0, "; hard gebleven: " & hersteld, "")
    BladBeveiligen wv
    Application.ScreenUpdating = True
    MsgBox "Stap 8 is klaar. De nameting is gelijk aan de nulmeting." & vbCrLf & vbCrLf & _
           "Kolom K: " & nOm & " van de " & nHard & " harde getallen zijn weer een formule." & _
           IIf(Len(hersteld) > 0, vbCrLf & "Hard gebleven (de formule geeft een andere uitkomst): " & hersteld, "") & vbCrLf & _
           "Koppen U, X en Y hernoemd; kolom E en G verborgen.", vbInformation, "Stap 8"
    Exit Sub

Terug:
    On Error Resume Next
    mStil = False
    ThisWorkbook.Names("Krediet_Aantal").RefersTo = oudNaam
    If Not hard Is Nothing Then
        For Each kopKol In hard.Keys
            wv.Cells(CLng(kopKol), 11).Value = hard(kopKol)
        Next kopKol
    End If
    kopKol = Array(21, 24, 25)
    For i = 1 To 3
        If Not IsEmpty(kopOud(i)) Then wv.Cells(1, kopKol(i - 1)).Formula = kopOud(i)
    Next i
    Application.Calculate
    BladBeveiligen wv
    Application.ScreenUpdating = True
    LogRegel "stap 8 teruggedraaid", , , "", , , verschil
    MsgBox "Stap 8 is teruggedraaid: er was een verschil (" & verschil & "). Er is niets veranderd.", vbCritical, "Stap 8"
    Exit Sub
Fout:
    verschil = "fout: " & Err.Description
    Resume Terug
End Sub

' Twee celwaarden gelijk (getallen tot 6 decimalen, anders als tekst).
Private Function GelijkStap8(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then GelijkStap8 = (IsError(a) And IsError(b)): Exit Function
    If IsNumeric(a) And IsNumeric(b) And Not IsEmpty(a) And Not IsEmpty(b) And CStr(a) <> "" And CStr(b) <> "" Then
        GelijkStap8 = (Round(CDbl(a) - CDbl(b), 6) = 0)
    Else
        GelijkStap8 = (CStr(a) = CStr(b))
    End If
End Function

' Knop Export naar financieel pakket (blok 2b). Wordt gebouwd in stap 10 (functioneel ontwerp paragraaf 14).
Public Sub ExportNaarFinancieelPakket()
    MsgBox "De export naar het financieel pakket wordt nog gebouwd (stap 10).", vbInformation, "Export naar financieel pakket"
End Sub
