Attribute VB_Name = "mAccent"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #   # *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### ##### *
' * #   # #     #  #  #   # #  #  #   #            # #      #           # #   # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        mAccent
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        procedures voor frmAccentueren
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.58         2019-06-11              nieuw
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Sub ShowFrmAccentueren()
' * Sub UpdateAccent()

'-----------------------------------------------------------------------------------------------

Option Explicit

Public myGERARD As New clsUpdate
Public CheckingAccent As Boolean
Public lUpdatingAccent As Boolean

Sub ShowFrmAccentueren()
    If ActiveWorkbook Is Nothing Then Exit Sub
    If frmAccentueren.Visible Then
        Unload frmAccentueren
        Exit Sub
    End If
    Set myGERARD.GERARD = Application
    frmAccentueren.Show 0
End Sub

Sub UpdateAccent()
    If ActiveCell.Column = 4 Or ActiveCell.Column = 8 Then
        If ActiveCell <> vbNullString Then
            If ActiveCell.Row > 1 Then
                frmAccentueren.txtZoekterm = ActiveCell.Text
            End If
        End If
    End If
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - mAccent                                                                       '
'-----------------------------------------------------------------------------------------------

