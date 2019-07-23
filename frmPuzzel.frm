VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPuzzel 
   Caption         =   "Puzzel"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   OleObjectBlob   =   "frmPuzzel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPuzzel"
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
' Module        frmPuzzel
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Puzzel samenstellen met onderdelen uit Tandem of andere Puzzel(s)
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.59         2019-06-11              Nieuw
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
' * 2019-06-11      Eerste Release      creatie, v 2.59
' *
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Me.Caption = gsAPP & " - Puzzel"
    Me.lstPuzzel.AddItem "[nieuwe Puzzel]"
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name Like "Puzzel*" Then
            Me.lstPuzzel.AddItem ws.Name
        End If
    Next
    If Me.lstPuzzel.ListCount = 1 Then
        Me.lstPuzzel.ListIndex = 0
    End If
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
' * 2019-06-11      Eerste Release      creatie, v 2.59
' *
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPlakPuzzel_Click()
    Dim sNaam As String
    Dim wsVan As Worksheet
    Dim wsNaar As Worksheet
    
    Set wsVan = ActiveSheet
    If Me.lstPuzzel.Text = "[nieuwe Puzzel]" Then
        sNaam = InputBox("Naam voor de nieuwe Puzzel?")
        DC_Journal "Nieuwe Puzzel: " & "[" & sNaam & "] "
        Worksheets.Add after:=ActiveSheet
        Set wsNaar = ActiveSheet
        wsNaar.Name = "Puzzel_" & sNaam
        With wsNaar.Tab
           .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.4
        End With
    Else
        Worksheets(Me.lstPuzzel.Text).Activate
        Set wsNaar = ActiveSheet
        DC_Journal "Puzzel aangevuld: "
    End If
    If wsVan.Name = wsNaar.Name Then
        MsgBox "Bron en Doel zijn hetzelfde werkblad...", vbInformation
    Else
        PuzzelVerzamelen wsVan, wsNaar
    End If
End Sub

Private Sub PuzzelVerzamelen(wsVan As Worksheet, wsPuzzel As Worksheet)
    Dim wsActive As Worksheet
    Dim nLaatsteRij As Double
    Dim rCel As Range
    Dim rRange As Range
    Dim nRijen As Double
    Dim nWerkbladen As Integer
    
    ' kopieer titels als Puzzel nog geen titels heeft (A1 = leeg)
    If wsPuzzel.Range("A1") = vbNullString Then
        wsVan.Range("A1:O1").Copy Destination:=wsPuzzel.Range("A1:O1")
    End If
    wsVan.Activate
    Set rRange = Selection.EntireRow
    nRijen = Selection.Rows.count
    rRange.Copy
    
    wsPuzzel.Activate
    
    nLaatsteRij = DC_LaatsteRij()
    Cells(nLaatsteRij + 1, 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    Columns("A:O").EntireColumn.AutoFit
    Columns("I:J").ColumnWidth = 10
    DC_Journal "Puzzel verwerkt: " & nRijen & " rijen"
    [A2].Select
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmPuzzel                                                                     '
'-----------------------------------------------------------------------------------------------
