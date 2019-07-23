VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSorteren 
   Caption         =   "Sorteren"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   OleObjectBlob   =   "frmSorteren.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSorteren"
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
' Project       ATLAS - Allerhande Tools ter Lichtere Arbeid van de Speurder
' Module        ftmSorteren
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Diverse sorteerroutines, in functie van omgeving
' References    None

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.20         2019-02-28              Eerste Release
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' -

' Functions:
' ~~~~~~~~~~
' -
'-----------------------------------------------------------------------------------------------

Option Base 1
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = gsAPP & " - Sorteren"
    Me.optDatumTijd = True
    Me.opt1 = True
    Me.chkKleuren = False
    If IsTandemSheet(ActiveSheet) Then
        Me.optCombiDatumTijd.Caption = "Tandem, datum en tijd"
        Me.optAantalCombiDatumTijd.Caption = "Aantal, Tandem, Datum en Tijd"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub optAantalCombiDatumTijd_Click()
    OptiesUpdate
End Sub

Private Sub optCombiDatumTijd_Click()
    OptiesUpdate
End Sub

Private Sub optDatumTijd_Click()
    OptiesUpdate
End Sub

Sub OptiesUpdate()
    Dim lTandem As Boolean
        lTandem = ActiveSheet.CodeName = "G_Tandem"
    If Me.optDatumTijd Then
        Me.opt1.Caption = "Oud naar nieuw"
        Me.opt2.Caption = "Nieuw naar oud"
    ElseIf Me.optCombiDatumTijd Then
        If lTandem Then
            Me.opt1.Caption = "Tandem A-Z + chronologisch"
            Me.opt2.Caption = "Tandem Z-A + chronologisch"
        Else
            Me.opt1.Caption = "Combi A-Z + chronologisch"
            Me.opt2.Caption = "Combi Z-A + chronologisch"
        End If
    ElseIf Me.optAantalCombiDatumTijd Then
        If lTandem Then
            Me.opt1.Caption = "Aantal 9=>1 + Tandem A-Z + chronologisch"
            Me.opt2.Caption = "Aantal 1=>9 + Tandem A-Z + chronologisch"
        Else
                Me.opt1.Caption = "Aantal 9=>1 + Combi A-Z + chronologisch"
            Me.opt2.Caption = "Aantal 1=>9 + Combi A-Z + chronologisch"
        End If
    End If
End Sub

Private Sub cmdGo_Click()
    Dim nRij As Double
    Dim sRange As String
    Dim sLogSleutel As String
    Dim sLogRichting As String
    Dim lTandem As Boolean
    Dim nRichting As Integer
    Dim nRichtingAantal As Integer
    Dim ws As Worksheet
    Set ws = ActiveSheet
    lTandem = ws.CodeName = "G_Tandem" Or Left(UCase(ws.Name), 6) = "PUZZEL"

    nRij = LaatsteRij()
    If lTandem Then
        sRange = "A1:O" & nRij
        ActiveSheet.Sort.SortFields.Clear
        nRichting = IIf(Me.opt1, xlAscending, xlDescending)
        nRichtingAantal = IIf(Me.opt1, xlDescending, xlAscending)
        If Me.optDatumTijd Then
            sLogSleutel = Me.optDatumTijd.Caption
            ws.Sort.SortFields.Add Key:=Range("B2:B" & nRij), SortOn:=xlSortOnValues, Order:=nRichting, DataOption:=xlSortTextAsNumbers
            ws.Sort.SortFields.Add Key:=Range("C2:C" & nRij), SortOn:=xlSortOnValues, Order:=nRichting, DataOption:=xlSortNormal
        ElseIf Me.optCombiDatumTijd Then
            sLogSleutel = Me.optCombiDatumTijd.Caption
            ws.Sort.SortFields.Add Key:=Range("N1:N" & nRij), SortOn:=xlSortOnValues, Order:=nRichting, DataOption:=xlSortNormal
            ws.Sort.SortFields.Add Key:=Range("B2:B" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            ws.Sort.SortFields.Add Key:=Range("C2:C" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ElseIf Me.optAantalCombiDatumTijd Then
            sLogSleutel = Me.optAantalCombiDatumTijd.Caption
            ws.Sort.SortFields.Add Key:=Range("O2:O" & nRij), SortOn:=xlSortOnValues, Order:=nRichtingAantal, DataOption:=xlSortNormal
            ws.Sort.SortFields.Add Key:=Range("N1:N" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ws.Sort.SortFields.Add Key:=Range("B2:B" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            ws.Sort.SortFields.Add Key:=Range("C2:C" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End If
    Else
        nRichting = IIf(Me.opt1, xlAscending, xlDescending)
        nRichtingAantal = IIf(Me.opt1, xlDescending, xlAscending)
        sRange = "A1:L" & nRij
        ActiveSheet.Sort.SortFields.Clear
        If Me.optDatumTijd Then
            sLogSleutel = Me.optDatumTijd.Caption
            ActiveSheet.Sort.SortFields.Add Key:=Range("B1:B" & nRij), SortOn:=xlSortOnValues, Order:=nRichting, DataOption:=xlSortNormal
        ElseIf Me.optCombiDatumTijd Then
            sLogSleutel = Me.optCombiDatumTijd.Caption
            ActiveSheet.Sort.SortFields.Add Key:=Range("K1:K" & nRij), SortOn:=xlSortOnValues, Order:=nRichting, DataOption:=xlSortNormal
            ActiveSheet.Sort.SortFields.Add Key:=Range("B1:B" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ElseIf Me.optAantalCombiDatumTijd Then
            sLogSleutel = Me.optAantalCombiDatumTijd.Caption
            ActiveSheet.Sort.SortFields.Add Key:=Range("L2:L" & nRij), SortOn:=xlSortOnValues, Order:=nRichtingAantal, DataOption:=xlSortNormal
            ActiveSheet.Sort.SortFields.Add Key:=Range("K2:K" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ActiveSheet.Sort.SortFields.Add Key:=Range("B2:B" & nRij), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End If
    End If
    sLogRichting = IIf(Me.opt1, Me.opt1.Caption, Me.opt2.Caption)
    With ActiveSheet.Sort
        .SetRange Range(sRange)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    DC_Journal "Sorteren: " & sRange & " - " & sLogSleutel & " - " & sLogRichting & " - "
    
    If Me.chkKleuren Then
        Range(sRange).Offset(1, 0).Interior.Color = vbWhite
        If lTandem Then
            If Me.optDatumTijd Then
                '
            ElseIf Me.optCombiDatumTijd Then
                KleurBlokken 14
            ElseIf Me.optAantalCombiDatumTijd Then
                KleurBlokken 15
                KleurBlokken 14, 2
                Range("A2:D" & nRij).Interior.Color = gnGROEN
                Range("E2:H" & nRij).Interior.Color = gnLICHTBLAUW
                Range("I2:J" & nRij).Interior.Color = gnZALMROZE
                TintShade
            End If
        Else
            If Me.optDatumTijd Then
                KleurBlokken 3
            ElseIf Me.optCombiDatumTijd Then
                KleurBlokken 11
            ElseIf Me.optAantalCombiDatumTijd Then
                KleurBlokken 12
                KleurBlokken 11, 2
            End If
        End If
    End If
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmSorteren                                                                   '
'-----------------------------------------------------------------------------------------------
