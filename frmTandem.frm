VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTandem 
   Caption         =   "Tandem"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   OleObjectBlob   =   "frmTandem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTandem"
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
' Module        frmTandem
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Tandems verzamelen
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.30         2019-03-05              Nieuw
' 02.54         2019-06-06              Refactoring ifv snelheid
' 02.71         2019-07-09              Keuze tussen alle DataSets en enkel die in Schema
'                                       Layout aangepast, optioneel meer info aan rechterzijde
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Private Sub UserForm_Initialize()
' * Private Sub cmdExit_Click()

' Functions:
' ~~~~~~~~~~

'-----------------------------------------------------------------------------------------------

Option Explicit

Dim lEnableEvents As Boolean

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
' * 2019-06-06      spinbutton voor Me.txtHits
' *
Private Sub UserForm_Initialize()
    Me.Caption = gsAPP & " - Tandem"
    G_Schema.Activate
    
    Me.txtInterval = 5
    Me.txtMinimum = 2
    Me.txtInterval.Enabled = False
    Me.txtMinimum.Enabled = False
    
    Me.txtHitsKolom = LaatsteKolom()
    Me.txtMaxHits = Application.WorksheetFunction.Max(Columns(Me.txtHitsKolom))
    Me.txtTotaalRecs = Range("DOSSIERTANDEM").Offset(-1, 2)
    
    Me.spnHits.Max = Me.txtMaxHits
    Me.spnHits.Min = 2
    Me.spnHits.Value = 2
    ' bepaal txtAfstand, pas daarna txtHits en dus event voor update testmogelijkheden
    lEnableEvents = False
    Me.txtAfstand = 50
    lEnableEvents = True
    Me.txtHits = 2

    Me.chkVerwijderTeLang = False
    Me.chkVerwijderBeperkt = False
    Me.chkMarkeren = True
    lEnableEvents = True
    
    Me.optVolledig = True
    Me.cmdExtra.Caption = "Meer Info =>"
    Me.Width = 351
    ' eerst éénmaal DataSheets verzamelen
    VerzamelSets ("DOSSIERTANDEM")
    SetsInSchema ("DOSSIERTANDEM")
    Me.lblSetsAlles = Application.WorksheetFunction.count(Range("dossiertandem").Offset(1, 1).Resize(100, 1))
    Me.lblSetsSchema = Application.WorksheetFunction.Sum(Range("dossiertandem").Offset(1, 1).Resize(100, 1))
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

Private Sub chkVerwijderBeperkt_Change()
    Me.txtMinimum.Enabled = Me.chkVerwijderBeperkt
End Sub

Private Sub chkVerwijderTeLang_Change()
    Me.txtInterval.Enabled = Me.chkVerwijderTeLang
End Sub

Private Sub cmdGo_Click()
    ZoekTandem
End Sub

Private Sub spnHits_Change()
    Me.txtHits = spnHits.Value
End Sub

Private Sub txtHits_Change()
    Dim nOK As Double                  ' aantal kentekens dat voldoet aan gevraagde aantal hits
    If lEnableEvents = False Then
        Exit Sub
    End If
    lEnableEvents = False
    
    If Me.txtHits > Me.txtMaxHits Then
        Me.txtHits = Me.txtMaxHits
    End If
    If Me.txtHits < 1 Then
        Me.txtHits = 1
    End If
    ActiveSheet.AutoFilter.Range.AutoFilter Field:=Me.txtHitsKolom, Criteria1:=">=" & Me.txtHits, Operator:=xlAnd
    nOK = ActiveSheet.UsedRange.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count - 1
    Me.lblRecords = "==> " & nOK & " rec."
    'On Error Resume Next
    Me.lblCombinaties = Format(nOK * Val(Me.txtTotaalRecs) * Val(Me.txtAfstand), "#,##0") & " test"
    'On Error GoTo 0
    lEnableEvents = True
End Sub

Private Sub txtAfstand_Change()
    txtHits_Change
End Sub

Private Sub cmdExtra_Click()
    If Me.cmdExtra.Caption = "Meer Info =>" Then
        Me.cmdExtra.Caption = "<= Minder Info"
        Me.Width = 453
    Else
        Me.cmdExtra.Caption = "Meer Info =>"
        Me.Width = 351
    End If
End Sub

Private Sub optSelectie_Change()
    BerekenAantallen
End Sub

Private Sub optVolledig_Change()
    'BerekenAantallen
End Sub


Sub BerekenAantallen()
    Dim nOK As Double
    Me.lblInfo = "Even geduld..."
    DoEvents
    If frmTandem.optVolledig Then
        ' gebruik alle Sets
        [DossierTandem].Offset(1, 1).Resize([DossierTandem].End(xlDown).Row - 2).Value = 1
    Else
        ' gebruik de Sets die gebruikt werden voor het Schema
        SetsInSchema ("DOSSIERTANDEM")
    End If
    Me.txtTotaalRecs = Range("DOSSIERTANDEM").Offset(-1, 2)
    
    nOK = ActiveSheet.UsedRange.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count - 1
    Me.lblRecords = "==> " & nOK & " rec."
    'Me.lblSetsSchema = Application.WorksheetFunction.Sum(Range("dossiertandem").Offset(1, 1).Resize(100, 1))
    'On Error Resume Next
    Me.lblCombinaties = Format(nOK * Val(Me.txtTotaalRecs) * Val(Me.txtAfstand), "#,##0") & " test"
    Me.lblInfo = vbNullString
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmTandem                                                                     '
'-----------------------------------------------------------------------------------------------

