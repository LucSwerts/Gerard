Attribute VB_Name = "m_start"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #  ## *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # # # *
' * #   # #     #  #  #   # #  #  #   #            # #      #           # ##  # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        m_Start
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Ingang van GERARD
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.21         2019-03-02              Conform nieuw protocol
' 02.52         2019-05-24              + Sub Immo
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Sub Start()
' * Sub Journal(sLog As String)
' * Sub Immo()


' Functions:
' ~~~~~~~~~~
' * Function Dossier() As String
' * Function Nu() As String
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Private Module
Option Base 1

' **********************************************************************************************
' * Procedure:      Sub Start()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Entrypoint voor GERARD
' * Aanroep van:    ThisWorkbook - Workbook_Open()
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       Immo
' * Scope:          Option Private Module
' * Aanroep:        Start
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release      creatie, v 1.00
' *
Sub Start()
    [cfgDossier] = InputBox(gsVRAAGDOSSIER, gsAPP)
    ActiveWindow.Caption = gsAPP & " - build " & gsBUILD
    DC_Journal gsSTERRETJES
    DC_Journal "| " & gsAPP & gsVER & " b-" & gsBUILD & " start..."
    DC_Journal "| " & gsCOPYRIGHT & IIf([cfgDevModus], " - DevModus", "")
    DC_Journal "| " & "Dossier: " & Dossier() & " - " & "User: " & DC_WhoDunIt
    DC_Journal gsSTERRETJES
    Immo
End Sub

' **********************************************************************************************
' * Procedure:      Sub Journal()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Logboek bijhouden, eventueel met alternatieve naam
' * Aanroep van:    Diverse
' * Argumenten:     sLog                tekst voor logboek
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Journal("Tekst voor logboek")
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release      creatie, v 1.00
' * 2019-03-02      Volledige revisie
' * 2019-03-11      Log centraal op netwerk
'
Sub Journal(sLog As String)
    Dim sPad As String                  ' pad van GERARD
    Dim sLogBoek As String              ' naam van log
    Dim sLink As String                 ' link tussen logschakels
    Dim sDatum As String                ' datum na format
    Dim sTijd As String                 ' tijd na format
    Dim sNetwerkLog As String           ' volledige naam voor netwerklog

    ' Niet loggen als AddIn
    If ThisWorkbook.IsAddin Then
        Exit Sub
    End If
        
    sPad = ThisWorkbook.path & "\"
    If [cfgLogActief] Then
        If [cfgLogAlternatief] = vbNullString Then
            sLogBoek = sPad & [cfgLogBestand]
        Else
            sLogBoek = sPad & [cfgLogAlternatief]
        End If
        sLink = IIf([cfgDevModus], " x ", " - ")
        sDatum = Format(Date, "yyyymmdd")
        sTijd = Format(Now, "hh:mm:ss")
        
        Open sLogBoek For Append Shared As #1
        Width #1, 150
        Print #1, sDatum & sLink & sTijd & " " & DC_WhoDunIt; Tab(39); sLog
        Close #1
    End If
    
    If [cfgLogNetwerk] <> vbNullString Then
        If Dir([cfgLogNetwerk], vbDirectory) <> vbNullString Then
            sNetwerkLog = [cfgLogNetwerk] & "\" & [cfgLogBestand]
            Open sNetwerkLog For Append Shared As #1
            Width #1, 150
            Print #1, sDatum & sLink & sTijd & " " & DC_WhoDunIt; Tab(39); sLog
            Close #1
        End If
    End If
End Sub

' ****************************************************************************
' * Groep:          Hulpfuncties voor Journal
' * --------------------------------------------------------------------------
' * Procedures:     Function WhoDunIt()
' *                 Function Dossier()
' *                 Function Nu()
' * --------------------------------------------------------------------------
' * 2017-04-11      Eerste Release      creatie, v 1.00
'
Function WhoDunIt() As String
    WhoDunIt = "[" & Left(Application.UserName & Application.Rept(" ", 15), 15) & "]"
End Function

Function Dossier() As String
    Dossier = "[" & Left([cfgDossier] & Application.Rept(" ", 12), 12) & "]"
End Function

Function Nu() As String
    Nu = Format(Now, "hh:mm:ss") & " "
End Function

' **********************************************************************************************
' * Procedure:      Sub Immo()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Namen definiëren
' * Aanroep van:    Start
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Immo()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-05-24      Eerste Release      creatie, v 2.52
'
Sub Immo()
    On Error Resume Next
    ActiveWorkbook.names.Add Name:="CHECK", RefersToR1C1:="=Dossier!R2C13"
    ActiveWorkbook.names.Add Name:="IMPORT", RefersToR1C1:="=Dossier!R2C1"
    ActiveWorkbook.names.Add Name:="SCHEMA", RefersToR1C1:="=Dossier!R2C5"
    ActiveWorkbook.names.Add Name:="ZOOMIN", RefersToR1C1:="=Dossier!R2C9"
    On Error GoTo 0
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - m_Start                                                                       '
'-----------------------------------------------------------------------------------------------
