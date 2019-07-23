VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmThema 
   Caption         =   "Thema"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   OleObjectBlob   =   "frmThema.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmThema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #  ## *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # # # *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #   # ##  # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        frmThemas
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Thema's opbouwen
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
    Me.lstThemas.AddItem "[nieuw thema]"
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name Like "Thema*" Then
            Me.lstThemas.AddItem ws.Name
        End If
    Next
    If Me.lstThemas.ListCount = 1 Then
        Me.lstThemas.ListIndex = 0
    End If
    If IsDataSheet(ActiveSheet) Then
        Me.txtKenteken = Cells(ActiveCell.Row, 11)
    ElseIf ActiveSheet.Name = "Tandem" Then
        If ActiveCell.Column = 4 Or ActiveCell.Column = 8 Then
            Me.txtKenteken = ActiveCell.Text
        End If
    End If
    Me.optExact = True
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

Private Sub cmdPlakThema_Click()
    Dim sNaam As String
    Dim ws As Worksheet
    
    If Me.lstThemas.Text = "[nieuw thema]" Then
        sNaam = InputBox("Naam voor het nieuwe thema?")
        DC_Journal "Nieuw thema: " & "[" & sNaam & "] sleutel: [" & Me.txtKenteken & "] " & IIf(Me.optExact, "zoek exact", "zoek op deel")
        Set ws = Worksheets.Add
        ws.Name = "Thema_" & sNaam
        With ws.Tab
           .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.4
        End With
    Else
        Worksheets(Me.lstThemas.Text).Activate
        Set ws = ActiveSheet
        DC_Journal "Thema aangevuld: " & "[" & Me.lstThemas.Text & "] sleutel: [" & Me.txtKenteken & "] " & IIf(Me.optExact, "zoek exact", "zoek op deel")
    End If
    ThemasVerzamelen Me.txtKenteken, Me.optExact, ws
    [A2].Select
End Sub

Private Sub ThemasVerzamelen(sKenteken As String, lExact As Boolean, wsThema As Worksheet)
    Dim ws As Worksheet
    Dim nLaatsteRij As Double
    Dim rCel As Range
    Dim rRange As Range
    Dim nRijen As Double
    Dim lTitels As Boolean
    Dim nTotaal As Double
    Dim nWerkbladen As Integer
    
    lTitels = True
    For Each ws In ActiveWorkbook.Worksheets
        If IsDataSheet(ws) Then
            nWerkbladen = nWerkbladen + 1
            ws.Activate
            If lTitels Then
                Range("A1:L1").Copy Destination:=wsThema.Range("A1:L1")
                wsThema.Range("M1") = "Bron"
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
            
            Set rRange = ActiveSheet.UsedRange.Offset(1, 0).Resize(ActiveSheet.UsedRange.Rows.count - 1, 12)
            On Error Resume Next
            If rRange.SpecialCells(xlCellTypeVisible).Cells.count > 0 Then
                If Err.Number = 0 Then
                    For Each rCel In rRange.Resize(, 1).SpecialCells(xlCellTypeVisible)
                        nRijen = nRijen + 1
                    Next rCel
                    nTotaal = nTotaal + nRijen
                End If
            End If
            On Error GoTo 0
            Me.lblInfo = ws.Name
            Me.lblAantal = nTotaal
            DoEvents
            ActiveSheet.UsedRange.Offset(1, 0).Copy
            wsThema.Activate
            
            If nRijen > 0 Then
                nLaatsteRij = LaatsteRij()
                Cells(nLaatsteRij + 1, 1).Select
                ActiveSheet.Paste
                Cells(nLaatsteRij + 1, 13).Resize(nRijen, 1) = ws.Name
                'Me.lblAantal = DC_LaatsteRij() - 1
            End If
        End If
    Next ws
    Me.lblInfo = nWerkbladen & " werkbladen verwerkt..."
    nLaatsteRij = LaatsteRij()
    Me.lblAantal = DC_LaatsteRij() - 1
    Cells(2, 12).FormulaR1C1 = "=COUNTIF(R2C[-1]:R" & nLaatsteRij & "C[-1],RC[-1])"
    If nLaatsteRij > 2 Then
        Range("L2").AutoFill Destination:=Range("L2:L" & nLaatsteRij)
    End If

    Columns("A:M").EntireColumn.AutoFit
    Application.CutCopyMode = False
    wsThema.Sort.SortFields.Clear
    wsThema.Sort.SortFields.Add Key:=Range("B2:B" & nLaatsteRij), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsThema.Sort
        .SetRange ActiveSheet.Range("A2:M" & nLaatsteRij)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    DC_Journal "Thema verwerkt: " & nWerkbladen & " werkbladen => " & nTotaal & " rijen"
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmThema                                                                      '
'-----------------------------------------------------------------------------------------------

