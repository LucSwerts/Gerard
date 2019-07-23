Attribute VB_Name = "mDC_Log"
' *******************************************************************************************
' * ####  ##### ##### ##### ##### ##### ##### #####                       #     ##### ##### *
' * #   # #     #     #   # #     #   # #   # #                          ##     #  ## #  ## *
' * #   # ####  ####  ##### #     #   # ##### ####    #####   #   #     # #     # # # # # # *
' * #   # #     #     #     #     #   # #  #  #                # #        #     ##  # ##  # *
' * ####  ##### ##### #     ##### ##### #   # #####             #   #   ##### # ##### ##### *
' *******************************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       DeepCore - Logging
' Module        mDC_Log
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Procedures voor logging
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 01.00 b001    2019-04-13              Eerste Release
'                                       + Sub DC_Journal()
'                                       + Sub DC_AtlasLog()
'                                       + Function DC_WhoDunIt()
'                                       + Function DC_LogKlok()
' 01.00 b002    2019-06-18              + Function DC_Dossier()
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' DC_Journal(sLog As String)
' DC_AtlasLog()
'

' Functions:
' ~~~~~~~~~~
' DC_WhoDunIt() As String
' DC_LogKlok() As String
' DC_Dossier() As String
'
'-----------------------------------------------------------------------------------------------
Option Explicit

' **********************************************************************************************
' * Procedure:      DC_Journal(sLog As String)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Logging, lokaal en eventueel op netwerk
' * Aanroep van:    multi
' * Argumenten:     sLog                Tekst voor logging
' * Gebruikt:       DC_Klok
' *                 DC_WhoDunIt
' *                 Config: cfgLogActief cfgLogMap cfgLogalternatief cfgLogboek cfgLogNetwerk
' * Resultaat:      -
' * Scope:          Public
' * Aanroep:        DC_Journal("Gegevens zijn verwerkt")
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' *
Sub DC_Journal(sLog As String)
    Dim sBewaar As String               ' tijdstip van bewaren van document
    Dim sBestand As String              ' naam van bestand dat open was
    Dim sPad As String                  ' pad voor logbestand
    Dim sLogBoek As String              ' naam van het logbestand
    Dim sNetwerkLog As String           ' volledige naam voor netwerklog
    
    On Error Resume Next
    ' is dit bestand ooit bewaard? Anders geen pad beschikbaar...
    sBewaar = ThisWorkbook.BuiltinDocumentProperties("last save time")
    On Error GoTo 0
    
    ' alleen loggen indien geen AddIn en indien ooit bewaard en dus pad beschikbaar
    If Not ThisWorkbook.IsAddin And sBewaar <> vbNullString Then
        ' logging terwijl ander bestand open is?
        If ThisWorkbook.Name <> ActiveWorkbook.Name Then
            sBestand = ActiveWorkbook.Name
            ThisWorkbook.Activate
        End If
            
        If [cfgLogActief] Then
            sPad = IIf([cfgLogMap] <> vbNullString, [cfgLogMap] & "\", ThisWorkbook.path & "\")
            sLogBoek = sPad & IIf([cfgLogAlternatief] <> vbNullString, [cfgLogAlternatief], [cfgLogboek])
            Open sLogBoek For Append Shared As #1
            Width #1, 150
            Print #1, DC_LogKlok & DC_WhoDunIt; Tab(44); sLog
            Close #1
        End If
        If [cfgLogNetwerk] <> vbNullString Then
            If Dir([cfgLogNetwerk], vbDirectory) <> vbNullString Then
                sNetwerkLog = [cfgLogNetwerk] & "\" & [cfgLogBestand]
                Open sNetwerkLog For Append Shared As #1
                Width #1, 150
                Print #1, DC_LogKlok() & DC_WhoDunIt; Tab(44); sLog
                Close #1
            End If
        End If
        
        If sBestand <> vbNullString Then
            Workbooks(sBestand).Activate
        End If
    End If
End Sub

' **********************************************************************************************
' * Procedure:      Function DC_WhoDunIt() As String
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Geeft de naam van de gebruiker met juiste opmaak voor logging
' * Aanroep van:    DC_Journal
' * Argumenten:     -                   -
' * Gebruikt:       -
' * Resultaat:      Naam van de gebruiker
' * Scope:          Public
' * Aanroep:        DC_WhoDunIt()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' *
Function DC_WhoDunIt() As String
    DC_WhoDunIt = "[" & Left(Application.UserName & Application.Rept(" ", 20), 20) & "]"
End Function

' **********************************************************************************************
' * Procedure:      Function DC_LogKlok() As String
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Geeft datum en tijd met DevModus-toestand met juiste opmaak voor logging
' * Aanroep van:    multi
' * Argumenten:     -                   -
' * Gebruikt:       -
' * Resultaat:      Datum en tijd met DevModus-toestand
' * Scope:          Public
' * Aanroep:        DC_LogKlok()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' *
Function DC_LogKlok() As String
    Dim sDev As String
    sDev = IIf([cfgDevModus], "x", "-")
    DC_LogKlok = Format(Date, "yyyymmdd ") & sDev & Format(Now, " hh:mm:ss ")
End Function

' **********************************************************************************************
' * Procedure:      Sub DC_AtlasLog()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   log in ATLASLOG.TXT
' * Aanroep van:    OOXML
' * Argumenten:     -                   -
' * Gebruikt:       -
' * Resultaat:      Naam van de gebruiker
' * Scope:          Public
' * Aanroep:        OOMXL
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' *
Sub DC_AtlasLog()
    Dim sFile As String                 ' Atlas-log
    Dim sLog As String                  ' tekst voor logging
    Dim nLengte As Integer              ' lengte van build-nummer
    Dim sBuild As String                ' build-tekst
    
    sFile = Left(ActiveWorkbook.path, 2) & "\###Atlas\Atlas\ATLASLOG.TXT"
    sLog = vbNullString
    If Dir(sFile) <> vbNullString Then
        sLog = InputBox("Log?", "Log")
    End If
    If sLog <> vbNullString Then
        sLog = " - LOG [" & sLog & "]"
        On Error Resume Next
        nLengte = Len(ThisWorkbook.names("cfgBuild"))
        If Err.Number <> 0 Then
            On Error GoTo 0
            sBuild = " 000"
        Else
            sBuild = Format([cfgBuild], "000")
        End If
        Open sFile For Append Shared As #1
        Width #1, 150
        Print #1, DC_LogKlok() & gsAPP & " -" & gsVER & " - build " & sBuild & sLog
        Close #1
    End If
End Sub

' **********************************************************************************************
' * Procedure:      Function DC_Dossier() As String
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Geeft de naam van het opgegeven dossier
' * Aanroep van:    DC_Journal
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Resultaat:      Naam van het dossier
' * Scope:          Option Private Module
' * Aanroep:        DC_Dossier()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-26      Eerste Release
' *
Function DC_Dossier() As String
    DC_Dossier = "[" & Left([cfgDossier] & Application.Rept(" ", 16), 16) & "]"
End Function
'
'-----------------------------------------------------------------------------------------------
' Einde module - mDC_Log                                                                       '
'-----------------------------------------------------------------------------------------------
