VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfig 
   Caption         =   "Configuratie"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10080
   OleObjectBlob   =   "frmConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfig"
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
' Module        frmConfig
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Configuratie van GERARD
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
' Private Sub cmdBewaar_Click()
'-----------------------------------------------------------------------------------------------

Option Explicit

' *****************************************************************************
' * procedure:  Private Sub UserForm_Initialize()                             *
' * ---------------------------------------------------------------------------
' * doel:       UserForm initialiseren                                        *
' *             instellingen ophalen en tonen                                 *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
'*
Private Sub UserForm_Activate()
    Me.Caption = gsAPP & " - Configuratie"
        
    ' importeren
    Me.chkVorigeMap = [cfgvorigemap]
    Me.txtNaamVorigeMap = [cfgnaamvorigemap]
    Me.txtFormatDatum = [cfgFormatDatum]
    Me.txtFormatTijd = [cfgFormatTijd]
    
    ' algemeen
    Me.chkLog = [cfgLog]
    Me.txtLogboek = [cfgLogBestand]
    Me.txtLogAlternatief = [cfgLogAlternatief]
    Me.chkDevModus = [cfgDevModus]
    
    ' zoom
    Me.txtZoomLettertype = [cfgZoomLetterType]
    Me.txtZoomPuntgrootte = [cfgZoomPuntgrootte]
    Me.cfgZoomUitlijnen = [cfgZoomUitlijnen]
    Me.chkSelectie = [cfgZoomSelectie]
    Me.chkOokGelijken = [cfgZoomNoteerAlles]
    Me.chkZoomDumpClip = [cfgZoomDumpNaarKlembord]
    Me.txtZoomDumpNaam = [cfgZoomDumpNaam]
    
    ' schema
    Me.chkBelgen = [cfgSchemaBelgen]
    Me.chkKleur = [cfgSchemaKleur]
    
    ' rest
    Me.chkWisSchemaFilter = [cfgWisSchemaFilter]
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
' * procedure:  Private Sub cmdBewaar_Click()                                 *
' * ---------------------------------------------------------------------------
' * doel:       instellingen opslaan                                          *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdBewaar_Click()
    ' importeren
    [cfgvorigemap] = Me.chkVorigeMap
    [cfgnaamvorigemap] = Me.txtNaamVorigeMap
    [cfgFormatDatum] = Me.txtFormatDatum
    [cfgFormatTijd] = Me.txtFormatTijd
    
    ' algemeen
    [cfgLog] = Me.chkLog
    [cfgLogBestand] = Me.txtLogboek
    [cfgLogAlternatief] = Me.txtLogAlternatief
    [cfgDevModus] = Me.chkDevModus
    
    ' zoom
    If FontIsInstalled(Me.txtZoomLettertype) Then
        [cfgZoomLetterType] = Me.txtZoomLettertype
    Else
        MsgBox Me.txtZoomLettertype & " bestaat niet"
    End If
    [cfgZoomPuntgrootte] = Me.txtZoomPuntgrootte
    [cfgZoomUitlijnen] = Me.cfgZoomUitlijnen
    [cfgZoomSelectie] = Me.chkSelectie
    [cfgZoomNoteerAlles] = Me.chkOokGelijken
    [cfgZoomDumpNaarKlembord] = Me.chkZoomDumpClip
    [cfgZoomDumpNaam] = Me.txtZoomDumpNaam
    
    ' schema
    [cfgSchemaBelgen] = Me.chkBelgen
    [cfgSchemaKleur] = Me.chkKleur
    
    ' rest
    [cfgWisSchemaFilter] = Me.chkWisSchemaFilter
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmConfig                                                                     '
'-----------------------------------------------------------------------------------------------
