VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOnderhoud 
   Caption         =   "Onderhoud"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   OleObjectBlob   =   "frmOnderhoud.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOnderhoud"
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
' Module        frmOnderhoud
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Onderhoudsmodule, wissen en opkuisen van allerlei werkbladen en bereiken
' References    None

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.61         2019-06-12              Eerste Release
' 02.62         2019-06-13              revisie
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' Private Sub cmdExit_Click()
' Private Sub cmdOnderhoud_Click()
'-----------------------------------------------------------------------------------------------

Option Base 1
Option Explicit

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
' * 2019-06-12      Eerste Release      creatie, v 2.61
' *
Private Sub cmdExit_Click()
    Unload Me
End Sub

' **********************************************************************************************
' * Procedure:      Private Sub cmdOnderhoud_Click()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Startpunt
' * Aanroep van:    cmdOnderhoud
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event driven
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-06-12      Eerste Release      creatie, v 2.61
' *
Private Sub cmdOnderhoud_Click()
    Dim ws As Worksheet
    Dim sWS As String                   ' naam van werkblad
    
    Application.DisplayAlerts = False
    
    ' TODO: Autofilter uitzetten bij Schema en Tandem
    
    If Me.chkWisDatasets Then
        For Each ws In ActiveWorkbook.Worksheets
            If IsDataSheet(ws) Then
                ws.Delete
            End If
        Next
        G_Dossier.UsedRange.Offset(2, 0).Delete shift:=xlUp
        Me.chkWisDatasets = False
    End If
    
    If Me.chkWisOutput Then
        For Each ws In ActiveWorkbook.Worksheets
            sWS = UCase(ws.Name)
            If Left(sWS, 6) = "INVENT" Or Left(sWS, 6) = "PUZZEL" Then
                ws.Delete
            End If
        Next
        Me.chkWisOutput = False
    End If
    
    If Me.chkWisSchema Then
        Worksheets("Schema").UsedRange.Offset(1, 0).Delete shift:=xlUp
        Me.chkWisSchema = False
    End If
    
    If Me.chkWisTandem Then
        Worksheets("Tandem").UsedRange.Offset(1, 0).Delete shift:=xlUp
        Me.chkWisTandem = False
    End If
    
    InhoudsTafel
    Application.DisplayAlerts = False
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmOnderhoud                                                                  '
'-----------------------------------------------------------------------------------------------
