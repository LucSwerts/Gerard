VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventaris 
   Caption         =   "Inventaris"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   OleObjectBlob   =   "frmInventaris.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInventaris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   #####   #   *
' * #     #     #   # #   # #   # #   #                         #       #  ##   *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # #   *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #       #   *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        frmInventaris
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Inventaris opbouwen van 1 kenteken of deel van kenteken of kenteken met jokers
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.22         2019-03-03              Nieuw
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Private Sub UserForm_Initialize()
' * Private Sub cmdExit_Click()
' * Private Sub cmdPlakThema_Click()
' * Private Sub ThemasVerzamelen(sKenteken As String, lExact As Boolean, wsThema As Worksheet)

' Functions:
' ~~~~~~~~~~

'-----------------------------------------------------------------------------------------------

Option Explicit

Private Sub cmdPlakInventaris_Click()

End Sub

' **********************************************************************************************
' * Procedure:      Private Sub UserForm_Activate()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   UserForm initialiseren
' * Aanroep van:    Ribbon
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event driven
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-03-03      Eerste Release      creatie, v 2.22
' *
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim sInvent As String
    Me.Caption = gsAPP & " - Inventaris"
    Me.lstInventaris.AddItem "[nieuwe inventaris]"
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name Like "Invent*" Then
            Me.lstInventaris.AddItem ws.Name
        End If
    Next
    'If Me.lstInventaris.ListCount = 1 Then
    '    Me.lstInventaris.ListIndex = 0
    'End If
    If ActiveSheet.Name = "Schema" Then
        Me.txtKenteken = Cells(ActiveCell.Row, 1).Text
    ElseIf ActiveSheet.Name = "Tandem" Then
        If ActiveCell.Column = 4 Or ActiveCell.Column = 8 Then
            Me.txtKenteken = ActiveCell.Text
        Else
            Me.txtKenteken = Cells(ActiveCell.Row, 4)
        End If
    Else
        Me.txtKenteken = Cells(ActiveCell.Row, 11).Text
    End If
    sInvent = "Invent-" & Me.txtKenteken
    If WorksheetExists(sInvent) Then
        ' ???
    Else
        Me.lstInventaris.AddItem "Invent-" & Me.txtKenteken
    End If
    Me.lstInventaris.ListIndex = Me.lstInventaris.ListCount - 1
    Me.optExact = True
    Me.optAlleSets = True
End Sub

' **********************************************************************************************
' * Procedure:      Private Sub cmdExit_Click()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   UserForm verlaten
' * Aanroep van:    cmdExit
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event driven
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-03-03      Eerste Release      creatie, v 2.22
' *
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdInventaris_Click()
    Dim sNaam As String
    Dim ws As Worksheet
    Dim sInvent As String
    
    If Me.lstInventaris.Text = "[nieuwe inventaris]" Then
        sNaam = InputBox("Naam voor de nieuwe inventaris?")
        DC_Journal "Nieuwe inventaris: " & "[" & sNaam & "] sleutel: [" & Me.txtKenteken & "] " & IIf(Me.optExact, "zoek exact", "zoek op deel")
        Set ws = Worksheets.Add
        ws.Name = "Invent_" & sNaam
        ws.Tab.Color = gnLICHTBLAUW
    Else
        sInvent = Me.lstInventaris.Text
        If WorksheetExists(sInvent) Then
            Worksheets(sInvent).Activate
            Set ws = ActiveSheet
            DC_Journal "Inventaris aangevuld: " & "[" & Me.lstInventaris.Text & "] sleutel: [" & Me.txtKenteken & "] " & IIf(Me.optExact, "zoek exact", "zoek op deel")
        Else
            Set ws = Worksheets.Add
            ws.Name = sInvent
            ws.Tab.Color = gnLICHTBLAUW
            DC_Journal "Nieuwe inventaris: " & "[" & sNaam & "] sleutel: [" & Me.txtKenteken & "] " & IIf(Me.optExact, "zoek exact", "zoek op deel")
        End If
    End If
    InventarisVerzamelen Me.txtKenteken, Me.optExact, ws
End Sub

Private Sub InventarisVerzamelen(sKenteken As String, lExact As Boolean, wsInventaris As Worksheet)
    Dim ws As Worksheet
    Dim nLaatsteRij As Double
    Dim rCel As Range
    Dim rRange As Range
    Dim nRijen As Double
    Dim lTitels As Boolean
    Dim nTotaal As Double
    Dim nWerkbladen As Integer
    
    Dim rSheets As Range                ' bereik met namen van werkbladen
    Dim nSheets As Integer              ' aantal werkbladen om te doorlopen
    Dim nSheet As Integer               ' lusteller
    Dim sSh As String                   ' naam van werkblad
    
    lTitels = True
    ' welke Sets zijn er allemaal beschikbaar
    VerzamelSets ("INVENTARIS")
    ' alle Sets of alleen die in Schema
    If Me.optAlleSets Then
        [INVENTARIS].Offset(1, 1).Resize([INVENTARIS].End(xlDown).Row - 2).Value = 1
    Else
        SetsInSchema ("INVENTARIS")
    End If
    Set rSheets = Range("INVENTARIS")
    nSheets = Sheets("Dossier").Cells(Rows.count, rSheets.Column).End(xlUp).Row
    For nSheet = 1 To nSheets - 2
        If rSheets.Offset(nSheet, 1) = 1 Then
            sSh = rSheets.Offset(nSheet, 0)
            Set ws = Sheets(sSh)
    
            nWerkbladen = nWerkbladen + 1
            ws.Activate
            If lTitels Then
                Range("A1:L1").Copy Destination:=wsInventaris.Range("A1:L1")
                wsInventaris.Range("M1") = "Bron"
                OpmaakTitels
                lTitels = False
            End If
            nLaatsteRij = LaatsteRij()
            nRijen = 0
            ' select A1 omdat anders AutoFilter actief kan worden op andere beperkte selectie
            [A1].Select
            Selection.AutoFilter
            If lExact Then
                ActiveSheet.Range("$A$1:$L$" & nLaatsteRij).AutoFilter Field:=11, Criteria1:=sKenteken
            Else
                ActiveSheet.Range("$A$1:$L$" & nLaatsteRij).AutoFilter Field:=11, Criteria1:="=*" & sKenteken & "*"
            End If
            
            Set rRange = ActiveSheet.UsedRange.Offset(1, 0).Resize(ActiveSheet.UsedRange.Rows.count - 1, 1)
            On Error Resume Next
            If rRange.SpecialCells(xlCellTypeVisible).Cells.count > 0 Then
                If Err.Number = 0 Then
                    For Each rCel In rRange.Resize(, 1).SpecialCells(xlCellTypeVisible)
                        nRijen = nRijen + 1
                    Next rCel
                    nTotaal = nTotaal + (nRijen)
                End If
            End If
            On Error GoTo 0
            Me.lblInfo = ws.Name
            Me.lblAantal = nTotaal
            DoEvents
            ActiveSheet.UsedRange.Offset(1, 0).Copy
            wsInventaris.Activate
            
            If nRijen > 0 Then
                nLaatsteRij = LaatsteRij()
                Cells(nLaatsteRij + 1, 1).Select
                ActiveSheet.Paste
                Cells(nLaatsteRij + 1, 13).Resize(nRijen, 1) = ws.Name
            End If
        End If
    Next nSheet
    Me.lblInfo = nWerkbladen & " werkbladen verwerkt..."
    nLaatsteRij = LaatsteRij()
    Cells(2, 12).FormulaR1C1 = "=COUNTIF(R2C[-1]:R" & nLaatsteRij & "C[-1],RC[-1])"
    If nLaatsteRij > 2 Then
        Range("L2").AutoFill Destination:=Range("L2:L" & nLaatsteRij)
        ' TODO values
    End If

    Columns("A:M").EntireColumn.AutoFit
    Application.CutCopyMode = False
    wsInventaris.Sort.SortFields.Clear
    wsInventaris.Sort.SortFields.Add Key:=Range("B1:B" & nLaatsteRij), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsInventaris.Sort
        .SetRange ActiveSheet.Range("A1:M" & nLaatsteRij)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    [A1].Select
    DC_Journal "Inventaris verwerkt: " & nWerkbladen & " werkbladen => " & nTotaal & " rijen"
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmThema                                                                      '
'-----------------------------------------------------------------------------------------------

