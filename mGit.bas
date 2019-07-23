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
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    directory = ActiveWorkbook.path & "\VBA"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub
