Attribute VB_Name = "mSupport"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   #####   #   *
' * #     #     #   # #   # #   # #   #                         #       #  ##   *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # #   *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #       #   *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        mSupport
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Support functies
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.21         2019-03-02              Conform nieuw protocol
' 02.47         2019-05-17              ~ Sub KleurBlokken()
'                                       ~ Sub DatumTijdKolomSplitser()
' 02.48         2019-05-18              ~ Function KiesMap()
'                                       ~ Function FontIsInstalled()
'                                       ~ Function WorksheetExists()
' 02.62         2019-06-13              + Private Sub OpmaakTitels()
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Sub KleurBlokken(Optional nKolom As Variant, Optional nCombi As Variant)
' * Sub DatumTijdKolomSplitser(Optional nKol As Variant)
' * Private Sub OpmaakTitels()

' Functions:
' ~~~~~~~~~~
' * Function KiesMap(sMap As String, sTitel As String) As String
' * Function FontIsInstalled(sFont) As Boolean
' * Function WorksheetExists(sSheet As String) As Boolean
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Private Module

' **********************************************************************************************
' * Procedure:      Sub KleurBlokken(Optional nKolom As Variant, Optional nCombi As Variant)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Kleur cellen per blok met gelijke inhoud
' * Aanroep van:    multi
' * Argumenten:     nKolom              kolom om te kleuren
' *                 nCombi              kleurcombinatie
' * Gebruikt:       LaatsteRij
' * Scope:          Option Private Module
' * Aanroep:        KleurBlokken(5, 2)
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-03-02      Eerste Release
' *
Sub KleurBlokken(Optional nKolom As Variant, Optional nCombi As Variant)
    Dim aKleuren() As Variant           ' kleurarray
    Dim nRijen As Double                ' aantal rijen
    Dim lEerste As Boolean              ' eerste kleur of andere
    Dim lSkipLeeg As Boolean            ' behoud kleur bij lege cel
    Dim sVorige As String               ' inhoud ter vergelijking
    Dim sActueel As String              ' inhoud van deze cel
    Dim nX As Double                    ' lusteller
    Dim nT As Long                      ' Timer
    
    nT = GetTickCount
    If IsMissing(nKolom) Then
        nKolom = ActiveCell.Column
    End If
    If IsMissing(nCombi) Then
        nCombi = 1
    End If
    
    aKleuren = IIf(nCombi = 1, Array(gnGROEN, gnZALMROZE), Array(gnLICHTBLAUW, gnGOUD))
    nRijen = LaatsteRij()
    lEerste = False
    lSkipLeeg = True
    sVorige = ""
    
    For nX = 2 To nRijen
        sActueel = Cells(nX, nKolom)
        If Len(Trim(sActueel)) = 0 And lSkipLeeg Then
            ' status quo
        Else
            If sActueel <> sVorige Then
                sVorige = sActueel
                lEerste = Not lEerste
            End If
        End If
        Cells(nX, nKolom).Interior.Color = aKleuren(IIf(lEerste, 0, 1))
    Next nX
    DC_Journal "Banden kleuren in kolom " & nKolom & " in " & DC_Verloop(nT)
End Sub

' hulpfunctie voor Tandem
' 08/06/2019
Sub TintShade()
    Dim nRijen As Double                ' totaal aantal rijen
    Dim nRij As Double                  ' rij waar gewerkt wordt
    Dim nBlok As Integer                ' rijen per blok
    Dim lKleur As Boolean               ' blok kleuren?
    Dim nT As Long                      ' GetTickCount
    
    nT = GetTickCount
    nRijen = LaatsteRij()
    nRij = 2
    lKleur = True
    While nRij < nRijen
        nBlok = Cells(nRij, 15)
        If lKleur Then
            Cells(nRij, 1).Resize(nBlok, 4).Interior.TintAndShade = -0.1
            Cells(nRij, 5).Resize(nBlok, 4).Interior.TintAndShade = -0.1
            Cells(nRij, 9).Resize(nBlok, 2).Interior.TintAndShade = -0.1
            Cells(nRij, 11).Resize(nBlok, 3).Interior.Color = gnLICHTGRIJS
        Else
            Cells(nRij, 1).Resize(nBlok, 4).Interior.TintAndShade = 0.3
            Cells(nRij, 5).Resize(nBlok, 4).Interior.TintAndShade = 0.3
            Cells(nRij, 9).Resize(nBlok, 2).Interior.TintAndShade = 0.3
            Cells(nRij, 11).Resize(nBlok, 3).Interior.Color = vbWhite
        End If
        lKleur = Not lKleur
        nRij = nRij + nBlok
    Wend
    DC_Journal "Tandem - Tint and shade in " & DC_Verloop(nT)
End Sub

Sub TandemBalkjes(nKolom As Integer)
    Dim nT As Long                      ' GetTickCount
    Dim nRijen As Double                ' aantal rijen
    
    nT = GetTickCount
    nRijen = LaatsteRij()
    
    Cells(2, nKolom).Resize(nRijen - 1).Select
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    Selection.FormatConditions(1).BarColor.Color = gnBALKBLAUW
    Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
    Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Color.Color = gnBALKBLAUW
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    Selection.FormatConditions(1).AxisColor.Color = 0
    Selection.FormatConditions(1).NegativeBarFormat.Color.Color = 255
    Selection.FormatConditions(1).NegativeBarFormat.BorderColor.Color = 255
    Cells(2, nKolom).Select
    DC_Journal "Tandem - balkjes in kolom " & nKolom & " in " & DC_Verloop(nT)
End Sub

' **********************************************************************************************
' * Procedure:      Sub DatumTijdKolomSplitser(Optional nKol As Variant)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Splits datum en tijd naar aparte kolommen
' *                 ===> 2019-05-17T21:12:03+0100
' *                 ===> 17/05/2019  21:12:03
' * Aanroep van:    multi
' * Argumenten:     nKolom              kolom waar de datum-tijd staat
' *                 -                   -
' * Gebruikt:       LaatsteRij
' *                 cfgPlusToepassen
' * Scope:          Option Private Module
' * Aanroep:        DatumTijdKolomSplitser(3)
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-03-02      Eerste Release
' * 2019-05-17      Refactoring, stroomlijnen, testen
' * 2019-06-08      NumberFormat moet altijd toegepast worden, was bij refactoring veranderd

Sub DatumTijdKolomSplitser(Optional nKol As Variant)
    Dim nKolom As Integer               ' kolom met datum en tijd
    Dim nRijen As Double                ' aantal rijen
    Dim nExtraTijd As Integer           ' extra bij te tellen tijd
    Dim nTopRij As Integer              ' rij met eerste gegevens
    
    If IsMissing(nKol) Then
        nKolom = ActiveCell.Column
        If MsgBox("Datum splitsen?" & vbCrLf & "Staat de cursor in de juiste kolom", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    Else
        nKolom = nKol
    End If
    nRijen = LaatsteRij()
    
    ' test of bovenste rij titel bevat (val(string) geeft 0)
    nTopRij = IIf(Val(Cells(1, nKolom)) > 0, 1, 2)
    Cells(nTopRij, nKolom).Select
    If Mid(ActiveCell, 11, 1) = "T" Then
        ' [2018-11-15T16:30:00+0100]
        Columns(nKolom + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns(nKolom + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(nTopRij, nKolom + 1).FormulaR1C1 = "=DATE(LEFT(RC[-1],4),MID(RC[-1],6,2),MID(RC[-1],9,2))"
        Cells(nTopRij, nKolom + 1).AutoFill Destination:=Cells(2, nKolom + 1).Resize(nRijen - 1, 1)
        Cells(nTopRij, nKolom + 2).FormulaR1C1 = "=MID(RC[-2],12,8)"
        Cells(nTopRij, nKolom + 2).AutoFill Destination:=Cells(nTopRij, nKolom + 2).Resize(nRijen - 1, 1)
        Cells(nTopRij, nKolom + 1).Resize(nRijen - 1, 2).Copy
        Cells(nTopRij, nKolom + 1).Resize(nRijen - 1, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells(1, nKolom + 1).Resize(1, 2).EntireColumn.AutoFit
        If Len(Cells(nTopRij, nKolom)) = 24 Then
            ' UTC-aanduiding nog toepassen?
            If [cfgPlusToepassen] Then
                nExtraTijd = Val(Mid(Cells(2, nKolom), 22, 1))
            End If
            ' voorzie drie extra kolommen
            Columns(nKolom + 3).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Columns(nKolom + 3).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Columns(nKolom + 3).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(nTopRij, nKolom + 3).FormulaR1C1 = "=RC[-2]+RC[-1]+" & nExtraTijd & "/24"
            Cells(nTopRij, nKolom + 3).AutoFill Destination:=Cells(2, nKolom + 3).Resize(nRijen - 1, 1)
            Cells(nTopRij, nKolom + 4).FormulaR1C1 = "=TRUNC(RC[-1])"
            Cells(nTopRij, nKolom + 4).AutoFill Destination:=Cells(2, nKolom + 4).Resize(nRijen - 1, 1)
            Cells(nTopRij, nKolom + 5).FormulaR1C1 = "=RC[-2]-RC[-1]"
            Cells(nTopRij, nKolom + 5).AutoFill Destination:=Cells(2, nKolom + 5).Resize(nRijen - 1, 1)
            Cells(nTopRij, nKolom + 4).Resize(nRijen - 1, 2).Copy
            Cells(nTopRij, nKolom + 4).Resize(nRijen - 1, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Cells(1, nKolom + 4).Resize(1, 2).EntireColumn.AutoFit
            Cells(nTopRij, nKolom + 1).Resize(1, 3).EntireColumn.Delete shift:=xlToLeft
        End If
        'Cells(nTopRij, nKolom + 1).Resize(nRijen - 1, 2).Select
        
        Application.CutCopyMode = False
        Cells(nTopRij, nKolom + 1).Resize(nRijen - 1, 1).NumberFormat = "dd/mm/yyyy;@"
        Cells(nTopRij, nKolom + 2).Resize(nRijen - 1, 1).NumberFormat = "h:mm:ss"
        Cells(1, nKolom + 1) = "Datum"
        Cells(1, nKolom + 2) = "Tijd"
    Else
        ' [17/05/2019  21:12:03]
        ' probeer DatumTijd in twee te delen: geheel en rest
        Columns(nKolom + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns(nKolom + 1).Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(nTopRij, nKolom + 1).FormulaR1C1 = "=TRUNC(RC[-1])"
        Cells(nTopRij, nKolom + 2).FormulaR1C1 = "=RC[-2]-TRUNC(RC[-2])"
        Cells(nTopRij, nKolom + 1).NumberFormat = "dd/mm/yyyy;@"
        Cells(nTopRij, nKolom + 2).NumberFormat = "hh:mm:ss;@"
        Cells(nTopRij, nKolom + 1).AutoFill Destination:=Cells(nTopRij, nKolom + 1).Resize(nRijen - 1, 1)
        Cells(nTopRij, nKolom + 2).AutoFill Destination:=Cells(nTopRij, nKolom + 2).Resize(nRijen - 1, 1)
        Cells(nTopRij, nKolom + 1).Resize(nRijen - 1, 2).Copy
        Cells(nTopRij, nKolom + 1).Resize(nRijen - 1, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells(1, nKolom + 1).Resize(1, 2).EntireColumn.AutoFit
        Cells(1, nKolom + 1) = "Datum"
        Cells(1, nKolom + 2) = "Tijd"
    End If
End Sub

' **********************************************************************************************
' * Procedure:      Function KiesMap(sMap As String, sTitel As String) As String
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Een map kiezen, vertrekkend van bepaalde map
' * Aanroep van:    multi
' * Argumenten:     sMap                map waar het zoeken begint
' *                 sTitel              titel voor het dialoogvenster
' * Gebruikt:       -
' *                 -
' * Resultaat:      Gekozen map of vbNullString bij annulatie
' * Scope:          Option Private Module
' * Aanroep:        KiesMap("C:\DATA","Kies een map")
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-03-02      Eerste Release
' * 2019-05-17      Refactoring, stroomlijnen, testen
' *
Function KiesMap(sMap As String, sTitel As String) As String
    ' voeg \ toe indien nodig
    sMap = Replace(sMap & "\", "\\", "\")
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = sTitel
        .ButtonName = "Kies Map"
        .InitialFileName = sMap
        
        ' -1 = bestand gekozen
        If .Show = -1 And .SelectedItems.count = 1 Then
            KiesMap = .SelectedItems(1)
        Else
            KiesMap = ""
        End If
    End With
End Function

' **********************************************************************************************
' * Procedure:      Function FontIsInstalled(sFont) As Boolean
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Testen of een bepaald lettertype beschikbaar is
' * Aanroep van:    multi
' * Argumenten:     sFont               naam van het lettertype
' *                 -                   -
' * Gebruikt:       -
' *                 -
' * Resultaat:      Waar of Onwaar
' * Scope:          Option Private Module
' * Aanroep:        FontIsInstalled("Segoe UI")
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-03-02      Eerste Release
' * 2019-05-18      Refactoring
' *
Function FontIsInstalled(sFont) As Boolean
    Dim FontList As CommandBarControl   ' Control met opmaak
    Dim TempBar As CommandBar           ' tijdelijke werkbalk voor noodgevallen
    Dim nX As Integer                   ' lusteller

    FontIsInstalled = False
    Set FontList = Application.CommandBars("Formatting").FindControl(ID:=1728)
    
    ' Als Font Control niet bestaat: temp CommandBar maken
    If FontList Is Nothing Then
        Set TempBar = Application.CommandBars.Add
        Set FontList = TempBar.Controls.Add(ID:=1728)
    End If
    
    For nX = 0 To FontList.ListCount - 1
        If FontList.List(nX + 1) = sFont Then
            FontIsInstalled = True
            On Error Resume Next
            TempBar.Delete
            Exit Function
        End If
    Next nX

    On Error Resume Next
    TempBar.Delete
    Set FontList = Nothing
    Set TempBar = Nothing
End Function

' **********************************************************************************************
' * Procedure:      Function WorksheetExists(sSheet As String) As Boolean
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Testen of een bepaald werkblad bestaat
' * Aanroep van:    multi
' * Argumenten:     sSheet              naam van het werkblad
' *                 -                   -
' * Gebruikt:       -
' *                 -
' * Resultaat:      Waar of Onwaar
' * Scope:          Option Private Module
' * Aanroep:        WorksheetExists("Users")
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-03-02      Eerste Release
' * 2019-05-18      Refactoring
' *
Function WorksheetExists(sSheet As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sSheet & "'!A1)")
End Function

' 12/06/2019
Function IsDataSheet(ws As Worksheet) As Boolean
    Dim sCN As String                   ' CodeName
    Dim s As String                     ' naam van werkblad
    sCN = ws.CodeName
    s = UCase(ws.Name)
    If Left(sCN, 2) = "G_" Or Left(s, 6) = "TANDEM" Or Left(s, 6) = "INVENT" Or Left(s, 5) = "THEMA" Or Left(s, 6) = "PUZZEL" Or Left(s, 1) = "_" Or ws.[K1] <> "Combi" Then
        IsDataSheet = False
    Else
        IsDataSheet = True
    End If
End Function

' 13/06/2019
Function IsTandemSheet(ws As Worksheet) As Boolean
    Dim s As String                     ' naam van werkblad
    s = UCase(ws.Name)
    If Left(s, 6) = "TANDEM" Or Left(s, 6) = "PUZZEL" Then
        IsTandemSheet = True
    Else
        IsTandemSheet = False
    End If
End Function

' **********************************************************************************************
' * Procedure:      Sub OpmaakTitels()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Generieke functie voor opmaak van titels op werkblad
' * Aanroep van:    divers
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Public
' * Aanroep:        OpmaakTitels()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-06-13      Eerste Release
' *
Sub OpmaakTitels()
    Dim nCellen As Integer              ' aantal cellen in titelrij
    Dim sAdres As String                ' cel bij begin
    nCellen = Application.WorksheetFunction.CountA(Rows(1))
    sAdres = ActiveCell.Address
    If IsDataSheet(ActiveSheet) Then
        [cfgOpmaakTitelData].Copy
    Else
        [cfgOpmaakTitelOverig].Copy
    End If
    [A1].Resize(1, nCellen).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(sAdres).Select
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - mSupport                                                                      '
'-----------------------------------------------------------------------------------------------
