Attribute VB_Name = "mInhoud"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   #####   #   *
' * #     #     #   # #   # #   # #   #                         #       #  ##   *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # #   *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #       #   *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Doel          maak een Inhoudstafel van alle werkbladen in de actieve werkmap
'               plaats vooraan een werkblad Inhoud, wis eventueel bestaande Inhoudstafel
'               plaats de naam van elk werkblad met een hyperlink in de Inhoudstafel
'               zet daar het aantal rijen en kolommen naast
'               kleur de achtergrond naargelang het type werkblad (Basis, Tandem, Inventaris, ...)
' Module        mInhoud
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Inhoudstafel opbouwen van alle werkbladen
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.49         2019-05-21              eerste release
' 02.51         2019-05-22              Refactoring
' 02.53         2019-06-05              diverse type werkbladen
' 02.68         2019-07-05              + WerkbladenOrdenen
' 02.69         2019-07-06              nieuwe kleuren, apart voor Schema en Tandem
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Sub InhoudsTafel()
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Private Module
Option Base 1

' ----------------------------------------------------------------------------------------------
' Publics en Constanten                                                                        '
' ----------------------------------------------------------------------------------------------
Public Const INHOUD As String = "___INHOUDSTAFEL___"

' **********************************************************************************************
' * Procedure:      Sub InhoudsTafel()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Inhoudstabel maken van alle werkbladen, op een werkblad [Inhoud]
' * Aanroep van:    OOXML
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        Inhoudstafel()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-05-21      Eerste Release
' * 2019-07-06      Nieuwe kleuren, apart voor Schema en Tandem
'
Sub InhoudsTafel()
    Dim nRij As Long
    Dim shtInhoud As Worksheet
    Dim sht As Object
    FiltersUit
    WerkbladenOrdenen
    On Error Resume Next
    
    Application.DisplayAlerts = False
    ' wis werkblad waar INHOUD staat, onafhankelijk van Sheet.Name
    Sheets(Range(INHOUD).Parent.Name).Delete
    Set shtInhoud = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Sheets(1))
    shtInhoud.Name = "Inhoud"
    shtInhoud.Tab.Color = gnTABINHOUD
    Application.DisplayAlerts = True
        
    On Error Resume Next
    ActiveWorkbook.Styles("InhoudHyperlink").Delete
    On Error GoTo 0
    ActiveWorkbook.Styles.Add Name:="InhoudHyperlink"
    With ActiveWorkbook.Styles("InhoudHyperlink")
        .IncludeNumber = True
        .IncludeFont = True
        .IncludeAlignment = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
        .IndentLevel = 1
        With .Font
            .Name = "Arial"
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = xlUnderlineStyleNone
            .Strikethrough = False
            .ColorIndex = 5
        End With
    End With
    On Error GoTo 0

    [A1:C1].Merge
    [A1] = "Inhoudstafel"
    With [A1]
        .HorizontalAlignment = xlCenter
        .Interior.Color = gnMOKKA
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = False
        .Font.ColorIndex = 6
        .Name = INHOUD
    End With
    [A2:C2] = Array("Werkblad", "R", "K")
    With [A2:C2]
        .HorizontalAlignment = xlCenter
        .Interior.Color = gnGOUD
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Color = vbBlue
    End With
    
    nRij = 3
    For Each sht In ActiveWorkbook.Sheets
        If sht.Visible = True Then
            If sht.Name <> shtInhoud.Name Then
                If TypeName(sht) = "Worksheet" Then
                    shtInhoud.Hyperlinks.Add _
                        Anchor:=Cells(nRij, 1), _
                        Address:="", _
                        SubAddress:="'" & Replace(sht.Name, "'", "''") & "'!A1", _
                        ScreenTip:="Ga naar " & sht.Name, _
                        TextToDisplay:=Chr(149) & " " & sht.Name
                Else
                    Cells(nRij, 1) = Chr(149) & " " & sht.Name
                End If
                Cells(nRij, 1).Style = "InhoudHyperlink"

                If UCase(sht.Name) = "SCHEMA" Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABSCHEMA
                    sht.Tab.Color = gnTABSCHEMA
                ElseIf UCase(Left(sht.CodeName, 2)) = "G_" Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABGERARD
                    sht.Tab.Color = gnTABGERARD
                ElseIf UCase(Left(sht.Name, 6)) = "TANDEM" Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABTANDEM
                    sht.Tab.Color = gnTABTANDEM
                ElseIf UCase(Left(sht.Name, 6)) = "INVENT" Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABINVENTARIS
                    sht.Tab.Color = gnTABINVENTARIS
                ElseIf UCase(Left(sht.Name, 5)) = "THEMA" Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABTHEMA
                    sht.Tab.Color = gnTABTHEMA
                ElseIf UCase(Left(sht.Name, 6)) = "PUZZEL" Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABPUZZEL
                    sht.Tab.Color = gnTABPUZZEL
                ElseIf UCase(Left(sht.Name, 1)) = "_" Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABUNDERSCORE
                    sht.Tab.Color = gnTABUNDERSCORE
                ElseIf IsDataSheet(sht) Then
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABDATASET
                    sht.Tab.Color = gnTABDATASET
                Else
                    Cells(nRij, 1).Resize(1, 3).Interior.Color = gnTABOVERIG
                    sht.Tab.Color = gnTABOVERIG
                End If
                Cells(nRij, 1).Resize(1, 3).Font.Color = vbWhite
                Cells(nRij, 2) = DC_LaatsteRij(sht)
                Cells(nRij, 3) = DC_LaatsteKolom(sht)
                nRij = nRij + 1
            End If
        End If
        DoEvents
    Next sht
    DC_Kaders ActiveSheet.UsedRange, 2
    shtInhoud.Columns("A").AutoFit
    shtInhoud.Columns("B:C").ColumnWidth = 8
    shtInhoud.Activate
    [D1].Select
    On Error GoTo 0
End Sub

' **********************************************************************************************
' * Procedure:      Sub WerkbladenOrdenen()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Ordent alle werkbladen volgens type
' * Aanroep van:    InhoudsTafel()
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        WerkbladenOrdenen()
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2019-07-05      Eerste Release
'
Sub WerkbladenOrdenen()
    Dim nInvoeg As Integer              ' invoegplek
    Dim nStart As Integer               ' eerste werkblad
    Dim nEinde As Integer               ' laatste werkblad
    Dim nX As Integer                   ' lusteller
    
    ' vermijd problemen met niet-bestaande werkbladen
    On Error Resume Next
    G_Dossier.Move Before:=Sheets(1)
    G_Config.Move after:=G_Dossier
    G_Schema.Move after:=G_Config
    Sheets("Tandem").Move after:=Sheets("Schema")
    Sheets("Inhoud").Move Before:=G_Dossier
    
    nInvoeg = Sheets("Tandem").Index
    nStart = nInvoeg + 1
    nEinde = Sheets.count
    For nX = nStart To nEinde
        If UCase(Left(Sheets(nX).Name, 6)) = "INVENT" Then
            Sheets(nX).Move after:=Sheets(nInvoeg)
            nInvoeg = nInvoeg + 1
        End If
    Next nX
    nStart = nInvoeg + 1
    For nX = nStart To nEinde
        If UCase(Left(Sheets(nX).Name, 5)) = "THEMA" Then
            Sheets(nX).Move after:=Sheets(nInvoeg)
            nInvoeg = nInvoeg + 1
        End If
    Next nX
    nStart = nInvoeg + 1
    For nX = nStart To nEinde
        If UCase(Left(Sheets(nX).Name, 6)) = "PUZZEL" Then
            Sheets(nX).Move after:=Sheets(nInvoeg)
            nInvoeg = nInvoeg + 1
        End If
    Next nX
    nStart = nInvoeg + 1
    For nX = nStart To nEinde
        If IsDataSheet(Sheets(nX)) Then
            Sheets(nX).Move after:=Sheets(nInvoeg)
            nInvoeg = nInvoeg + 1
        End If
    Next nX
    
    nStart = nInvoeg + 1
    For nX = nStart To nEinde
        If IsDataSheet(Sheets(nX)) Then
            Sheets(nX).Move after:=Sheets(nInvoeg)
            nInvoeg = nInvoeg + 1
        End If
    Next nX
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - mInhoud                                                                       '
'-----------------------------------------------------------------------------------------------
