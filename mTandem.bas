Attribute VB_Name = "mTandem"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   #####   #   *
' * #     #     #   # #   # #   # #   #                         #       #  ##   *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # #   *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #       #   *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        mTandem
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Opbouwen van Tandems
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.30         2019-03-05              eerste release
' 02.54         2019-06-06              Dictionary opbouwen op basis van zichtbare cellen
'                                       # 34187 >=4 hits: van 8,7" => 0,5" (# 69)
'                                       # 34187 >=2 hits: van 12,8" => 4,8" (# 10584)
' 02.56         2019-09-10              wissen van te klein aantal Tandems aangepast
'                                       van wissen van onderaf naar wissen met AutoFilter
'                                       nieuwe versie: scharnierpunt zoeken en dan EntireRow.Delete
'                                       #890 => #130 van 11,35" naar 0,05"
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Sub Tandem()

' Functions:
' ~~~~~~~~~~

'-----------------------------------------------------------------------------------------------

Option Explicit
Option Private Module
Option Base 1

' TANDEM
' maak een Dictionary met alle Combi's die meermaals voorkomen, zie Hits in Schema
' enkel Combi's die in die Dictionary zitten zijn interessant voor Tandems
' doorloop alle werkbladen (behalve G_*), doorloop combi's die in de Dictionary voorkomen
' koppel die Combi's aan Combi's die ook in de Dictionary voorkomen en zet die Tandem in G_Tandem

Sub ZoekTandem()
    Dim nTTT As Double                  ' aantal acties bij vergelijken
    Dim dictT As scripting.Dictionary   ' Dictionary voor Tandems
    Dim nX As Double                    ' lusteller
    Dim nRijen As Double                ' aantal rijen in Schema
    Dim nHits As Double                 ' aantal kolommmen in Schema (= kolom met Hits)
    Dim nMaxHits As Integer             ' maximum aantal Hits in Schema
    '
    Dim ws As Worksheet                 ' lus
    Dim nY As Integer                   ' lusteller
    Dim nVerschil As Double             ' tijdsverschil
    Dim nTandemRij As Double            ' rij in Tandem werkblad
    Dim nShift1 As Integer              ' kolomshift voor kenteken 1
    Dim nShift2 As Integer              ' kolomshift voor kenteken 2
    Dim lNxVoorop As Boolean            ' rijdt eerste voertuig voorop?
    Dim nLaatsteRij As Double           ' laatste rij in Tandem
    Dim rVisible As Range               ' zichtbare cellen in AutoFilter
    Dim cVisible As Range               ' zichtbare cellen in AutoFilter
    
    ' gegevens van frmTandem, variabelen sneller in lussen
    Dim nMinimumHits As Integer         '
    Dim nMaximumAfstand As Integer      '
    Dim lVerwijderTeLang As Boolean     '
    Dim nMaxTijd As Integer             '
    Dim lVerwijderTeWeinig As Boolean   '
    Dim nTandemKeten As Integer         '
    Dim lMarkeer As Boolean             ' brongegevens markeren (groen)?
    Dim nDrempel As Double              ' drempel voor wissen van te beperkte Tandems
    
    Dim nT As Long                      ' start van clock
    Dim rSheets As Range                ' bereik met namen van werkbladen
    Dim nSheets As Integer              ' aantal werkbladen om te doorlopen
    Dim nSheet As Integer               ' lusteller
    Dim sSh As String                   ' naam van werkblad
    
    nT = GetTickCount
    nMinimumHits = frmTandem.txtHits
    nMaximumAfstand = frmTandem.txtAfstand + 1
    lVerwijderTeLang = frmTandem.chkVerwijderTeLang
    nMaxTijd = frmTandem.txtInterval
    lVerwijderTeWeinig = frmTandem.chkVerwijderBeperkt
    nTandemKeten = frmTandem.txtMinimum
    lMarkeer = frmTandem.chkMarkeren
    
    DC_Journal gsSTERRETJES
    DC_Journal "| Tandem start... " & Format(Now, "hh:mm:ss")
    DC_Journal "| " & nMinimumHits & " hits nodig | " & IIf(lMarkeer, "Markeren", "Niet markeren") & " in Brongegevens"
    DC_Journal "| tot " & nMaximumAfstand - 1 & " voertuigen en max " & nMaxTijd & " minuten interval"
    DC_Journal gsSTERRETJES
    Sheets("Tandem").UsedRange.Offset(1, 0).Delete shift:=xlUp
    
    VerzamelSets ("DOSSIERTANDEM")
    If frmTandem.optVolledig Then
        ' gebruik alle Sets
        [DossierTandem].Offset(1, 1).Resize([DossierTandem].End(xlDown).Row - 2).Value = 1
    Else
        ' gebruik de Sets die gebruikt werden voor het Schema
        SetsInSchema ("DOSSIERTANDEM")
    End If
    
    FiltersUit
    
    Sheets("Schema").Activate
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    nRijen = LaatsteRij()
    nHits = frmTandem.txtHitsKolom
    
    nMaxHits = frmTandem.txtMaxHits
    If nMinimumHits > nMaxHits Then
        nMinimumHits = nMaxHits
    End If
    DC_Journal "Tandem - verzamel kentekens - "
    frmTandem.lblInfo = "verzamel kentekens..."
    DoEvents
    
    nT = GetTickCount
    Set dictT = New scripting.Dictionary
    ActiveSheet.UsedRange.AutoFilter Field:=nHits, Criteria1:=">=" & frmTandem.txtHits, Operator:=xlAnd
    Set rVisible = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible)
    For Each cVisible In rVisible
        dictT.Add Key:=cVisible.Text, Item:=cVisible.Offset(0, nHits - 1).Value
    Next cVisible
        
    DC_Journal "Tandem - Dictionary samengesteld in " & DC_Verloop(nT, True) & dictT.count & " items"
    nTandemRij = 2
    'frmTandem.lblInfo = "doorloop registraties..."
    Set rSheets = Range("DOSSIERTANDEM")
    nSheets = Sheets("Dossier").Cells(Rows.count, rSheets.Column).End(xlUp).Row
    For nSheet = 1 To nSheets - 2
        If rSheets.Offset(nSheet, 1) = 1 Then
            sSh = rSheets.Offset(nSheet, 0)
            Set ws = Sheets(sSh)
            'Set ws = Sheets("01.06 LAM 19.40 -20.05")
          '''For Each ws In ActiveWorkbook.Sheets
            nT = GetTickCount
            ws.Activate
            frmTandem.lblInfo = "doorloop " & ws.Name
            Application.ScreenUpdating = False
            ' doorloop alle registraties
            nRijen = LaatsteRij()
            For nX = 2 To nRijen - 1
                Application.StatusBar = ws.Name & " - " & nX & " / " & nRijen
                '''''''DoEvents
                ' zit Combi in Dictionary?
                nTTT = nTTT + 1
                If dictT(Cells(nX, 11).Text) <> 0 Then
                    ' noteer eventuele tandems met Combi's binnen bepaalde afstand
                    For nY = 1 To nMaximumAfstand
                        If nX + nY <= nRijen Then
                            nTTT = nTTT + 1
                            If dictT(Cells(nX + nY, 11).Text) <> vbNullString Then
                                nVerschil = (Cells(nX, 3) + Cells(nX, 4)) - (Cells(nX + nY, 3) + Cells(nX + nY, 4))
                                If Not lVerwijderTeLang Or (nVerschil <= nMaxTijd / 24 / 60) Then
                                    ' alfabetisch ordenen
                                    If Cells(nX, 11) < Cells(nX + nY, 11) Then
                                        nShift1 = 0
                                        nShift2 = 4
                                        lNxVoorop = (Cells(nX, 3) + Cells(nX, 4)) < (Cells(nX + nY, 3) + Cells(nX + nY, 4))
                                    Else
                                        nShift1 = 4
                                        nShift2 = 0
                                        lNxVoorop = (Cells(nX, 3) + Cells(nX, 4)) > (Cells(nX + nY, 3) + Cells(nX + nY, 4))
                                    End If
                                    With Sheets("Tandem")
                                        .Cells(nTandemRij, 1 + nShift1) = Cells(nX, 1)
                                        .Cells(nTandemRij, 2 + nShift1).NumberFormat = "dd/mm/yyyy;@"
                                        .Cells(nTandemRij, 2 + nShift1) = Cells(nX, 3).Value
                                        .Cells(nTandemRij, 3 + nShift1) = Format(Cells(nX, 4), "hh:mm:ss")
                                        .Cells(nTandemRij, 4 + nShift1) = Cells(nX, 11)
                                        .Cells(nTandemRij, 1 + nShift2) = Cells(nX + nY, 1)
                                        .Cells(nTandemRij, 2 + nShift2).NumberFormat = "dd/mm/yyyy;@"
                                        .Cells(nTandemRij, 2 + nShift2) = Cells(nX + nY, 3).Value
                                        .Cells(nTandemRij, 3 + nShift2) = Format(Cells(nX + nY, 4), "hh:mm:ss")
                                        .Cells(nTandemRij, 4 + nShift2) = Cells(nX + nY, 11)
                                        .Cells(nTandemRij, 9) = Format(nVerschil, "hh:mm:ss")
                                        .Cells(nTandemRij, 10) = nY - 1
                                        .Cells(nTandemRij, 11) = Cells(nX, 9)
                                        .Cells(nTandemRij, 12) = Cells(nX, 10)
                                        .Cells(nTandemRij, 13) = ws.Name
                                        .Cells(nTandemRij, IIf(lNxVoorop, 3, 7)).Font.Bold = True
                                    End With
                                    nTandemRij = nTandemRij + 1
                                    If lMarkeer Then
                                        Cells(nX, 3).Interior.Color = vbGreen
                                        Cells(nX, 4).Interior.Color = vbGreen
                                        Cells(nX, 6).Interior.Color = vbGreen
                                        Cells(nX + nY, 3).Interior.Color = vbGreen
                                        Cells(nX + nY, 4).Interior.Color = vbGreen
                                        Cells(nX + nY, 6).Interior.Color = vbGreen
                                    End If
                                End If
                            End If
                        End If
                    Next nY
                End If
            Next nX
            DC_Journal "Tandem - " & ws.Name & " ( " & DC_Verloop(nT, True) & nRijen & " rijen)"
            Application.ScreenUpdating = True
            frmTandem.Repaint
        End If
    Next nSheet
    Debug.Print "vergelijkingen: " & nTTT
    
    Sheets("Tandem").Activate
    nLaatsteRij = LaatsteRij()
    If nLaatsteRij > 1 Then
        nT = GetTickCount
        Application.ScreenUpdating = False
        frmTandem.lblInfo = "Groepeer tandems..."
        DoEvents
        ' Tandem-voertuigen groeperen
        [N2].FormulaR1C1 = "=RC[-10]&""-""&RC[-6]"
        [N2].AutoFill Destination:=Range("N2:N" & nLaatsteRij)
        Range("N2:N" & nLaatsteRij).Copy
        [N2].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        DC_Journal "Tandem - Tandems gegroepeerd - " & nLaatsteRij & " rijen - " & DC_Verloop(nT)

        ' Tandems tellen
        nT = GetTickCount
        frmTandem.lblInfo = "Tandems tellen"
        DoEvents
        [O2].FormulaR1C1 = "=COUNTIF(R2C[-1]:R" & nLaatsteRij & "C[-1],RC[-1])"
        Range("O2").AutoFill Destination:=Range("O2:O" & nLaatsteRij)
        Range("O2:O" & nLaatsteRij).Copy
        [P2].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Columns("O:O").Delete shift:=xlToLeft
        Columns("O:O").EntireColumn.AutoFit
        [O1] = "Aantal"
        [N1].Copy
        [O1].PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        DC_Journal "Tandem - Tandems geteld - " & DC_Verloop(nT)
        
        nT = GetTickCount
        frmTandem.lblInfo = "Verwijder beperkte Tandems..."
        DoEvents

        nLaatsteRij = LaatsteRij()
        With Sheets("Tandem").Sort
            With .SortFields
                .Clear
                .Add Key:=Range("O2:O" & nLaatsteRij), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .Add Key:=Range("N2:N" & nLaatsteRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .Add Key:=Range("B2:B" & nLaatsteRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
                .Add Key:=Range("C2:C" & nLaatsteRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            End With
            .SetRange Range("A1:O" & nLaatsteRij)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With


        If ActiveSheet.AutoFilterMode Then
            [A1:O1].AutoFilter
        End If
        [A1:O1].AutoFilter

        ' zoek met AutoFilter naar aantal dat lager ligt dan drempel
        If lVerwijderTeWeinig Then
            ActiveSheet.UsedRange.AutoFilter Field:=15, Criteria1:="<" & nTandemKeten, Operator:=xlAnd
            If Range("O1:O" & nLaatsteRij).SpecialCells(xlCellTypeVisible).Cells.count > 1 Then
                ' meer dan alleen titelrij gevonden, dus wissen
                On Error Resume Next
                ' zoek eerste rij met te laag aantal Tandems
                nDrempel = Range("O2:O" & nLaatsteRij).SpecialCells(xlCellTypeVisible).Cells(1).Row
                If Err.Number > 0 Then
                    Err.Clear
                    nDrempel = 0
                End If
                On Error GoTo 0
                If nDrempel > 0 Then
                    ' kleintjes wissen, eerst AutoFilter weer afzetten voor kolom 15
                    ActiveSheet.Range("A1:O" & nLaatsteRij).AutoFilter Field:=15
                    Cells(nDrempel, 1).Resize(nLaatsteRij - nDrempel + 1, 1).EntireRow.Delete shift:=xlUp
                End If
            End If
        End If
        [A1].Select
        
        DC_Journal "Tandem - beperkte Tandems verwijderd - " & DC_Verloop(nT)
        WisOverbodigeCellen
        nLaatsteRij = DC_LaatsteRij()
        [A1].Select
        Range("A2:D" & nLaatsteRij).Interior.Color = gnGROEN
        Range("E2:H" & nLaatsteRij).Interior.Color = gnLICHTBLAUW
        Range("I2:J" & nLaatsteRij).Interior.Color = gnZALMROZE
        Application.ScreenUpdating = False
        
        nT = GetTickCount
        frmTandem.lblInfo = "Sorteer Tandems..."
        DoEvents
        
        KleurBlokken 15
        KleurBlokken 14, nCombi:=2
        If nLaatsteRij < 5000 Then
            TintShade
            If [cfgTandemBalken] Then
                TandemBalkjes (9)
                TandemBalkjes (10)
            End If
        End If
        
        DC_Journal "Tandem - Tandems gesorteerd - " & DC_Verloop(nT)
    End If
    
    WisOverbodigeCellen
    Columns("A:O").EntireColumn.AutoFit
    DC_Journal "Tandem verzamelde " & LaatsteRij() - 1 & " rijen gegevens..."
    DC_Journal "Tandem einde"
    frmTandem.lblInfo = LaatsteRij() - 1 & " rijen gegevens verzameld..."
    DoEvents
    frmTandem.Repaint
    Application.ScreenUpdating = True
    'Debug.Print "einde: " & GetTickCount - nT
End Sub

Sub jfkd()
    WisOverbodigeCellen
End Sub

Sub jfdskfjkdsfsdmk()
    On Error Resume Next
    Debug.Print Range("02:120").SpecialCells(xlCellTypeVisible).Cells(1).Address
    Debug.Print Err.Number
    Debug.Print Err.Description
    On Error GoTo 0
End Sub
