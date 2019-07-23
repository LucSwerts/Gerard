Attribute VB_Name = "mGit"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #  ## *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # # # *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #   # ##  # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        mGit
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        voor GitHub
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.73         2019-07-23              + Sub ExportCode
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' ExportCode()

' Functions:
' ~~~~~~~~~~
' -
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Private Module
Option Base 1

' ----------------------------------------------------------------------------------------------
' Publics en Constanten                                                                        '
' ----------------------------------------------------------------------------------------------
Const Module = 1
Const ClassModule = 2
Const Form = 3
Const Document = 100

' **********************************************************************************************
' * Procedure:      Sub ExportCode()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Exporteert alle code naar bepaalde map
' * Aanroep van:    -
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        ExportCode()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-07-23      Eerste Release
' *
Public Sub ExportCode()
    Dim sMap As String
    Dim nAantal As Integer
    Dim objComponent As Object
    Dim sExtensie As String
    Dim fso As New FileSystemObject
    Dim sPad As String
    
    sMap = ActiveWorkbook.path & "\VBA"
    nAantal = 0
    
    If Not fso.FolderExists(sMap) Then
        Call fso.CreateFolder(sMap)
    End If
    Set fso = Nothing
    
    For Each objComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case objComponent.Type
            Case ClassModule, Document
                sExtensie = ".cls"
            Case Form
                sExtensie = ".frm"
            Case Module
                sExtensie = ".bas"
            Case Else
                sExtensie = ".txt"
        End Select
                
        On Error Resume Next
        Err.Clear
        
        sPad = sMap & "\" & objComponent.Name & sExtensie
        Call objComponent.Export(sPad)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & objComponent.Name & " to " & sPad, vbCritical)
        Else
            nAantal = nAantal + 1
            Debug.Print "Exported " & Left$(objComponent.Name & ":" & Space(24), 24) & sPad
        End If
        On Error GoTo 0
    Next
    Debug.Print "Successfully exported " & CStr(nAantal) & " VBA files to " & sMap
End Sub
