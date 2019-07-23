Attribute VB_Name = "mDC"
' *******************************************************************************************
' * ####  ##### ##### ##### ##### ##### ##### #####                       #     ##### ##### *
' * #   # #     #     #   # #     #   # #   # #                          ##     #  ## #  ## *
' * #   # ####  ####  ##### #     #   # ##### ####    #####   #   #     # #     # # # # # # *
' * #   # #     #     #     #     #   # #  #  #                # #        #     ##  # ##  # *
' * ####  ##### ##### #     ##### ##### #   # #####             #   #   ##### # ##### ##### *
' *******************************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       DeepCore - Basisprocedures voor modulair gebruik
' Module        mDC
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Basisprocedures
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 01.00 b001    2019-04-13              Eerste Release
'                                       + Function DC_LaatsteRij()
'                                       + Function DC_LaatsteKolom()
'                                       + Sub DC_ResetUsedRange()
'                                       + Sub DC_Kaders()
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' Sub DC_ResetUsedRange()
' Sub DC_Kaders(r As Range, Optional nType As Integer = 1)
'

' Functions:
' ~~~~~~~~~~
' Function DC_LaatsteRij(Optional ws As Worksheet) As Long
' Function DC_LaatsteKolom(Optional ws As Worksheet) As Long
'-----------------------------------------------------------------------------------------------
Option Explicit

' **********************************************************************************************
' * Procedure:      Function DC_LaatsteRij(Optional ws As Worksheet) As Long
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Geeft nummer van laatste rij met inhoud
' * Aanroep van:    multi
' * Argumenten:     Optional ws         Worksheet om te testen
' * Gebruikt:       -
' * Resultaat:      Nummer van laatste rij met inhoud, 0 indien leeg werkblad
' * Scope:          Public
' * Aanroep:        DC_LaatsteRij(ActiveSheet)
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' *
Function DC_LaatsteRij(Optional ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    On Error Resume Next
    DC_LaatsteRij = ws.Cells.Find(What:="*", after:=ws.Range("A1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    On Error GoTo 0
End Function

' **********************************************************************************************
' * Procedure:      Function DC_LaatsteKolom(Optional ws As Worksheet) As Long
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Geeft nummer van laatste kolom met inhoud
' * Aanroep van:    multi
' * Argumenten:     Optional ws         Worksheet om te testen
' * Gebruikt:       -
' * Resultaat:      Nummer van laatste kolom met inhoud, 0 indien leeg werkblad
' * Scope:          Public
' * Aanroep:        DC_LaatsteKolom(ActiveSheet)
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' *
Function DC_LaatsteKolom(Optional ws As Worksheet) As Long
    If ws Is Nothing Then Set ws = ActiveSheet
    On Error Resume Next
    DC_LaatsteKolom = ws.Cells.Find(What:="*", after:=ws.Range("A1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    On Error GoTo 0
End Function

' **********************************************************************************************
' * Procedure:      Sub DC_ResetUsedRange()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Reset van UsedRange op elk werkblad
' * Aanroep van:    multi
' * Argumenten:     -                   -
' * Gebruikt:       DC_LaatsteRij
' *                 DC_LaatsteKolom
' * Resultaat:      -
' * Scope:          Public
' * Aanroep:        DC_ResetUsedRange()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' *
Sub DC_ResetUsedRange()
    Dim nR As Long                      ' Laatste rij met gegevens
    Dim nK As Long                      ' Laatste kolom met gegevens
    Dim ws As Worksheet                 ' Worksheeobject voor lus
    Dim rDummy As Range                 ' Dummy Range voor reset
        
    For Each ws In ActiveWorkbook.Sheets
        With ws
            On Error Resume Next
            nR = DC_LaatsteRij(ws)
            nK = DC_LaatsteKolom(ws)
            On Error GoTo 0
        
            If nR * nK = 0 Then
                .Columns.Delete
            Else
                .Range(.Cells(nR + 1, 1), .Cells(.Rows.count, 1)).EntireRow.Delete
                .Range(.Cells(1, nK + 1), .Cells(1, .Columns.count)).EntireColumn.Delete
            End If
            Set rDummy = ws.UsedRange
        End With
    Next ws
End Sub

' **********************************************************************************************
' * Procedure:      Sub DC_Kaders(r As Range, Optional nType As Integer = 1)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Kaders tekenen of verwijderen
' *                 nType 0 is geen kader
' *                 nType 1 is buitenkader
' *                 nType 2 is met binnenkaders
' * Aanroep van:    multi
' * Argumenten:     r                   Range waar kaders moeten getekend / verwijderd worden
' *                 nType               Type kaders (0 = geen, 1 = buiten, 2 = binnen + buiten)
' * Gebruikt:       -
' * Resultaat:      -
' * Scope:          Public
' * Aanroep:        DC_Kaders(rRange, 2)
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht
' * 2019-04-13      Eerste Release
' * 2019-04-27      Optional nType as Integer = 1
' *
Sub DC_Kaders(r As Range, Optional nType As Integer = 2)
    With r
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        Select Case nType
            Case 0
                .BorderAround xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            Case 1
                .BorderAround xlContinuous
            Case 2
                .BorderAround xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End Select
    End With
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - mDC                                                                           '
'-----------------------------------------------------------------------------------------------
