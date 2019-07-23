VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #   # *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### ##### *
' * #   # #     #  #  #   # #  #  #   #            # #      #           # #   # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        clsUpdate
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        class procedure voor realtime update van Forms
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.58         2019-06-11              nieuw
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Private Sub GERARD_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
' * Private Sub Updater()

'-----------------------------------------------------------------------------------------------

Option Explicit
Public WithEvents GERARD As Application
Attribute GERARD.VB_VarHelpID = -1

Private Sub GERARD_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
    Call Updater
End Sub

Private Sub Updater()
    On Error Resume Next
    If frmAccentueren.Visible = True Then
        If lUpdatingAccent Then Exit Sub    'to prevent cascading events
        Call UpdateAccent
        On Error GoTo 0
    End If
    On Error GoTo 0
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - clsUpdate                                                                     '
'-----------------------------------------------------------------------------------------------

