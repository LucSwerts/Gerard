VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLog 
   Caption         =   "Log"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9210
   OleObjectBlob   =   "frmLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLog"
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
' Doel          ---
' Module        frmLog
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Doel van deze module
' References    None

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 01.00         2019-04-26              Eerste Release
' 01.00 b001    2019-04-26              update...
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' Private Sub UserForm_Activate()
' Private Sub cmdExit_Click()
' Private Sub cmdKleiner_Click()
' Private Sub cmdGroter_Click()
' Private Sub cmdStandaard_Click()
'-----------------------------------------------------------------------------------------------

Option Explicit

' *****************************************************************************
' * procedure:  Private Sub UserForm_Activate()                               *
' * ---------------------------------------------------------------------------
' * doel:       UserForm initialiseren                                        *
' * gebruikt:   Sub DrillDown()                                               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
'*
Private Sub UserForm_Activate()
    Me.Caption = gsAPP & " - Log"
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdExit_Click()                                   *
' * ---------------------------------------------------------------------------
' * doel:       UserForm verlaten                                             *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdExit_Click()
    Unload Me
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdKleiner_Click()                                *
' *             Private Sub cmdGroter_Click()                                 *
' *             Private Sub cmdStandaard_Click()                              *
' * ---------------------------------------------------------------------------
' * doel:       tools                                                         *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdKleiner_Click()
    txtLog.Font.Size = Application.WorksheetFunction.Max(txtLog.Font.Size - 2, 8)
    cmdStandaard.Caption = txtLog.Font.Size
End Sub

Private Sub cmdGroter_Click()
    txtLog.Font.Size = Application.WorksheetFunction.Min(txtLog.Font.Size + 2, 24)
    cmdStandaard.Caption = txtLog.Font.Size
End Sub

Private Sub cmdStandaard_Click()
    Me.txtLog.Font.Size = Range("cfgZoomPuntGrootte")
    cmdStandaard.Caption = txtLog.Font.Size
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmLog                                                                       '
'-----------------------------------------------------------------------------------------------
