Attribute VB_Name = "mTools"
' ******************************************************************
' ##### ##### #####  ###  ##### ####             #     ##### ##### *
' #     #     #   # #   # #   # #   #           ##         # #   # *
' # ### ###   ##### ##### ##### #   #  #   #   # #     ##### #   # *
' #   # #     #  #  #   # #  #  #   #   # #      #     #     #   # *
' ##### ##### #   # #   # #   # ####     #   # ##### # ##### ##### *
' ******************************************************************
' Gerard v. 1.20 - ANPR vs. Rondtrekkende Daders            mTools *
' ******************************************************************
Option Explicit

Function TelCodeNamen()
    Dim nAantal As Integer
    Dim nX As Integer
    For nX = 1 To Sheets.count
        If UCase(Left(Sheets(nX).CodeName, 2)) = "G_" Then
            nAantal = nAantal + 1
        End If
    Next nX
    TelCodeNamen = nAantal
End Function

' *****************************************************************************
' * procedure:  Sub VerzamelSets(sRange As String)                            *
' * ---------------------------------------------------------------------------
' * doel:       Verzamel de namen van de beschikbare Sets                     *
' * gebruiker:  Private Sub UserForm_Initialize                               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 24/05/19                    selectie verfijnd

Sub fdkskf()
    VerzamelSets ("SETS")
End Sub

Sub VerzamelSets(sRange As String)
    Dim rRange As Range                 ' werkplaats
    Dim nOffset As Integer              ' offset in lijst
    Dim sh As Worksheet
    
    Set rRange = Range(sRange)
    nOffset = 1
    rRange.Offset(1, 0).Resize(1000, 3).ClearContents
    For Each sh In ActiveWorkbook.Sheets
        If IsDataSheet(sh) And sh.Visible Then
        'If Left(sh.CodeName, 2) <> "G_" And UCase(sh.Name) <> "INHOUD" And UCase(Left(sh.Name, 6)) <> "INVENT" And Left(sh.Name, 1) <> "_" Then
            rRange.Offset(nOffset, 0) = sh.Name
            rRange.Offset(nOffset, 1) = 0
            ' gebruik UsedRange.Rows.Count ipv LaatsteRij omwille van mogelijke AutoFilter
            rRange.Offset(nOffset, 2) = sh.UsedRange.Rows.count
            nOffset = nOffset + 1
        End If
    Next sh
End Sub

' *****************************************************************************
' * procedure:  Sub VerzamelSets(sRange As String)                            *
' * ---------------------------------------------------------------------------
' * doel:       Verzamel de namen van de beschikbare Sets                     *
' * gebruiker:  Private Sub UserForm_Initialize                               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 24/05/19                    selectie verfijnd

Sub fjdskfjks()
    SetsInSchema ("SETS")
End Sub

Sub SetsInSchema(sRange As String)
    Dim nSets As Integer                ' aantal sets in Schema
    Dim dictSets As Object              ' array met namen van sets in Schema
    Dim nX As Integer                   ' lusteller
    Dim rRange As Range                 ' werkplaats
    Dim nSet As Integer                 ' aantal sets in range
    Dim nOffset As Integer              ' offset in lijst
    Dim sh As Worksheet
    
    If Len(Trim(G_Schema.Range("B1"))) = 0 Then
        MsgBox "geen Sets opgenomen in Schema..."
        Exit Sub
    End If
    ' plaats alle sets in Schema in een Dictionary
    Set rRange = Range(sRange)
    nSets = Sheets("Schema").Range("B1").End(xlToRight).Column - 2
    Set dictSets = CreateObject("Scripting.Dictionary")
    For nX = 1 To nSets
        dictSets.Add G_Schema.Range("A1").Offset(0, nX).Value, 1
    Next nX
        
    Set rRange = Range(sRange)
    nSets = Worksheets(rRange.Parent.Name).Cells(Rows.count, rRange.Column).End(xlUp).Row
    For nSet = 1 To nSets - 2
        If dictSets.Exists(rRange.Offset(nSet, 0).Text) = True Then
            rRange.Offset(nSet, 1) = 1
        Else
            rRange.Offset(nSet, 1) = 0
        End If
    Next nSet
End Sub

Sub fjksdsd()
    Dim dictSets As Object              ' array met namen van sets in Schema
    Set dictSets = CreateObject("Scripting.Dictionary")
    dictSets.Add "aaa", 1
    dictSets.Add "bbb", 1
    dictSets.Add "ccc", 1
    
    Debug.Print "bbb" & " - "; dictSets.Exists("bbb")

End Sub


' *****************************************************************************
' * procedure:  Sub FiltersUit()                                              *
' * ---------------------------------------------------------------------------
' * doel:       Schakel AutoFilters uit (AutoFilter zelf niet verwijderen     *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *

Sub FiltersUit()
    Dim nSheets As Integer              ' aantal werkbladen in Gerard
    Dim nX As Integer                   ' lusteller

    nSheets = ThisWorkbook.Sheets.count
    For nX = 1 To nSheets
        ' niet voor de core-sheets die met "G_" beginnen
        If Left(Worksheets(nX).CodeName, 2) <> "G_" Then
            If Worksheets(nX).AutoFilterMode Then
                If Worksheets(nX).FilterMode Then
                    Worksheets(nX).ShowAllData
                End If
            End If
        End If
    Next nX
    If Range("cfgWisSchemaFilter") Then
        If Worksheets("SCHEMA").AutoFilterMode Then
            If Worksheets("SCHEMA").FilterMode Then
                Worksheets("SCHEMA").ShowAllData
            End If
        End If
    End If
End Sub

' 02/01/2019
Function LaatsteRij() As Double
    LaatsteRij = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row
End Function

Function LaatsteKolom() As Double
    LaatsteKolom = ActiveSheet.Cells.Find("*", SearchOrder:=xlByColumns, LookIn:=xlValues, SearchDirection:=xlPrevious).Column
End Function

Sub WisOverbodigeCellen(Optional sSheet As Variant)
    Dim nRij As Double
    Dim nKolom As Double
    Dim sOldSheet As String
    If IsMissing(sSheet) Then
        sSheet = ActiveSheet.Name
    End If
    sOldSheet = ActiveSheet.Name
    Sheets(sSheet).Activate
    nRij = DC_LaatsteRij()
    nKolom = DC_LaatsteKolom()
    Range(Cells(nRij + 1, 1), Cells(Rows.count, 1)).EntireRow.Delete
    Range(Cells(1, nKolom + 1), Cells(1, Columns.count)).EntireColumn.Delete
    Sheets(sOldSheet).Activate
End Sub

Sub DevTijd()
    Debug.Print GetTickCount
End Sub

' 07/06/2019
Function DC_Verloop(nT As Long, Optional lKoppel As Variant) As String
    If IsMissing(lKoppel) Then
        lKoppel = False
    End If
    DC_Verloop = Format((GetTickCount - nT) / 1000, "0.000 sec ") & IIf(lKoppel, "- ", "")
End Function
'
' EINDE mTools ****************************************************************
