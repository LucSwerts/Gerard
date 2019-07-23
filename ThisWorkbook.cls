VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #  ## *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # # # *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #   # ##  # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       ATLAS - Allerhande Tools ter Lichtere Arbeid van de Speurder
' Module        ThisWorkbook
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Event-afhandeling van ThisWorkbook
' References    None

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 01.00         2017-04-11              Eerste Release
' 02.62         2019-06-13              Refactoring
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' Private Sub Workbook_Open()
' Private Sub Workbook_BeforeClose(Cancel As Boolean)
' Private Sub Workbook_WindowActivate(ByVal Wn As Window)
' Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
' Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'-----------------------------------------------------------------------------------------------

Option Explicit

' ----------------------------------------------------------------------------------------------
' Publics en Constanten                                                                        '
' ----------------------------------------------------------------------------------------------
Const KeyCtrl = &H11                    ' Control-toets

' **********************************************************************************************
' * Procedure:      Private Sub Workbook_Open()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   openen werkboek en voorbereidingen
' * Aanroep van:    Event driven
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release
' *
Private Sub Workbook_Open()
    Start
End Sub

' **********************************************************************************************
' * Procedure:      Private Sub Workbook_BeforeClose(Cancel As Boolean)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   hoofding aanpassen
' * Aanroep van:    Event driven
' * Argumenten:     Cancel              -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release
' *
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DC_Journal Application.Rept("*", 60)
    DC_Journal Nu & gsAPP & " einde..."
    ActiveWindow.Caption = ""
End Sub

' **********************************************************************************************
' * Procedure:      Private Sub Workbook_WindowActivate(ByVal Wn As Window)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   hoofding aanpassen
' * Aanroep van:    Event driven
' * Argumenten:     Wn                  -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release
' *
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    Application.Caption = " "
    ActiveWindow.Caption = ActiveWorkbook.Name & " - " & gsAPP & " " & gsVER & " - build " & gsBUILD
End Sub

' **********************************************************************************************
' * Procedure:      Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   hoofding aanpassen
' * Aanroep van:    Event driven
' * Argumenten:     Wn                  -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release
' *
Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    Application.StatusBar = False
    Application.Caption = ""
    ActiveWindow.Caption = ""
End Sub

' **********************************************************************************************
' * Procedure:      Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Eventueel loggen van info als linkse Alt-toets ingedrukt is
' * Aanroep van:    Event driven
' * Argumenten:     SaveAsUi            -
' *                 Cancel              -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Event
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release
' *
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim sTekst As String
    DC_Journal gsSTERRETJES
    DC_Journal gsAPP & gsVER & " " & gsBUILD & " bewaren. [" & Format(FileLen(ThisWorkbook.FullName), "#,##0") & "]"
    If GetKeyState(KeyCtrl) < 0 Then
        sTekst = InputBox("Log?", "Log")
        DC_Journal "LOG [" & sTekst & "]"
    End If
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - ThisWorkbook                                                                  '
'-----------------------------------------------------------------------------------------------
