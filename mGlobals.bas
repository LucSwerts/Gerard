Attribute VB_Name = "mGlobals"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #  ## *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # # # *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #   # ##  # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        mGlobals
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Algemene definities
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.21         2019-03-02              Conform nieuw protocol
' 02.68         2019-07-05              TabKleuren
'-----------------------------------------------------------------------------------------------

' API-calls:
' ~~~~~~~~~~
' GetTickCount Lib KERNEL32
' GetKeyState Lib USER32
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Private Module

' ----------------------------------------------------------------------------------------------
' Windows API calls
' ----------------------------------------------------------------------------------------------
#If Win64 Then
    Public Declare PtrSafe Function GetTickCount Lib "KERNEL32" () As Long
#Else
    Public Declare Function GetTickCount Lib "KERNEL32" () As Long
#End If

#If VBA7 Then
    Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#End If

' ----------------------------------------------------------------------------------------------
' Publics                                                                                      '
' ----------------------------------------------------------------------------------------------
Public gTcStart As Long                 ' start voor timer
Public gTcEinde As Long                 ' einde voor timer

' ------------------------------------------------------------
' Algemeen APP
' ------------------------------------------------------------
Public Const gsATLAS As String = "Allerhande Tools ter Lichtere Arbeid van de Speurder..."
Public Const gsAPP As String = "GERARD"
Public Const gsVER As String = " v2.73"
Public Const gsBUILD As String = "001"
Public Const gsCOPYRIGHT As String = "(c) Luc S. apr 2017 - jul 2019"
Public Const gsINFO As String = "Gericht Efficiënt Rechercheren" & vbCrLf & _
            "met ANPR tegen Rondtrekkende Daders"

' ------------------------------------------------------------
' schermteksten
' ------------------------------------------------------------
Public Const gsVRAAGDOSSIER As String = "Geef de naam van het dossier: "
Public Const gsKIESANPRTITEL As String = "Kies de map met de ANPR-gegevens"
Public Const gsGEGEVENSLADEN As String = "Gegevens worden geladen..."

' ------------------------------------------------------------
' kleuren
' ------------------------------------------------------------
Public Const gnGROEN As Double = 11854022
Public Const gnBLAUW As Double = 9851952
Public Const gnLICHTBLAUW As Double = 15123099
Public Const gnZALMROZE As Double = 8696052
Public Const gnGOUD As Double = 6740479
Public Const gnLICHTGRIJS As Double = 15132390
Public Const gnMUISGRIJS As Double = 12566463
Public Const gnMOKKA As Double = 1137094
Public Const gnGEBROKENWIT As Double = 13431551
Public Const gnSCHAALGROEN As Double = 8109667
Public Const gnSCHAALROOD As Double = 7039480
Public Const gnBALKBLAUW As Double = 13012579
' TabKleuren
'''Public Const gnTABINHOUD As Double = 14277081
'''Public Const gnTABGERARD As Double = 1137094
'''Public Const gnTABTANDEM As Double = 8696052
'''Public Const gnTABDATASET As Double = 11854022
'''Public Const gnTABINVENTARIS As Double = 16247773
'''Public Const gnTABPUZZEL As Double = 15652797
'''Public Const gnTABTHEMA As Double = 15123099
'''Public Const gnTABOVERIG As Double = 10092543

Public Const gnTABINHOUD As Double = 6908265
Public Const gnTABGERARD As Double = 36095
Public Const gnTABSCHEMA As Double = 128
Public Const gnTABTANDEM As Double = 3937500
Public Const gnTABDATASET As Double = 27464
Public Const gnTABINVENTARIS As Double = 10388072
Public Const gnTABTHEMA As Double = 8874803
Public Const gnTABPUZZEL As Double = 6439469
Public Const gnTABUNDERSCORE As Double = 7371150
Public Const gnTABOVERIG As Double = 10047907
' ------------------------------------------------------------
' overige
' ------------------------------------------------------------
Public Const gsSTERRETJES = "************************************************************"
'
'-----------------------------------------------------------------------------------------------
' Einde module - mGlobals                                                                      '
'-----------------------------------------------------------------------------------------------
