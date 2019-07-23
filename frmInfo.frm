VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfo 
   Caption         =   "Info"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   OleObjectBlob   =   "frmInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        frmInfo
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Info overGERARD
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.21         2019-03-02              Conform nieuw protocol
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' Private Sub UserForm_Activate()
' Private Sub cmdExit_Click()
' Private Sub imgAtlas_Click()

' Functions:
' ~~~~~~~~~~
' -
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Base 1

' **********************************************************************************************
' * Procedure:      Private Sub UserForm_Activate()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   UserForm initialiseren, instellingen ophalen
' * Aanroep van:    Ribbon
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event driven
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release      creatie, v 1.00
' *
Private Sub UserForm_Activate()
    DC_Journal "Info " & gsAPP & gsVER
    Me.Caption = gsAPP & " - Info"
    Me.lblApp = gsAPP & gsVER
    Me.lblCopyRight = gsCOPYRIGHT
    Me.lblInfo = gsINFO
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Top = Application.Height / 2 - Me.Height / 2
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
' * 2017-04-11      Eerste Release      creatie, v 1.00
' *
Private Sub cmdExit_Click()
    Unload Me
End Sub

' **********************************************************************************************
' * Procedure:      Private Sub imgAtlas_Click()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Info over Atlas-project
' * Aanroep van:    imgAtlas
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event driven
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release      creatie, v 1.00
' *
Private Sub imgAtlas_Click()
    MsgBox gsATLAS, vbInformation, "ATLAS"
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmInfo                                                                       '
'-----------------------------------------------------------------------------------------------
