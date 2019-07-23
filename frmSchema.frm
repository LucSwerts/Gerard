VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSchema 
   Caption         =   "Schema"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10590
   OleObjectBlob   =   "frmSchema.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******************************************************************
' ##### ##### #####  ###  ##### ####             #     ##### ##### *
' #     #     #   # #   # #   # #   #           ##         # #   # *
' # ### ###   ##### ##### ##### #   #  #   #   # #     ##### #   # *
' #   # #     #  #  #   # #  #  #   #   # #      #     #     #   # *
' ##### ##### #   # #   # #   # ####     #   # ##### # ##### ##### *
' ******************************************************************
' Gerard v. 1.20 - ANPR vs. Rondtrekkende Daders         frmSchema *
' ******************************************************************

' ************************
' * overzicht procedures *
' ************************
' * Private Sub UserForm_Initialize()
' * Private Sub cmdExit_Click()
' * Sub VulListBox()
' * Private Sub cmdKeuzeOmkeren_Click()
' * Private Sub cmdKiesAlle_Click()
' * Private Sub lstNamen_Change()
' * Function ListCounter(lst As MSForms.ListBox) As Integer
' * Sub Status(sTekst As String)
' * Sub StatusCSV()
' * Private Sub cmdSchema_Click()
' * Sub SchemaNaarUniek()
' * Sub KleurAsterisk()

Option Explicit

' *****************************************************************************
' * procedure:  Private Sub UserForm_Initialize()                             *
' * ---------------------------------------------------------------------------
' * doel:       UserForm initialiseren                                        *
' *             gewenste instellingen activeren                               *
' * gebruikt:   Sub VerzamelSets(sRange As String)                            *
' *             Sub VulListBox()                                              *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub UserForm_Initialize()
    Me.Caption = gsAPP & " - Schema"
    Me.chkBelgen = True
    Me.chkKleurInSchema = True
    ' verzamel beschikbare Sets voor Schema
    VerzamelSets ("SCHEMA")
    VulListBox
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
' * procedure:  Sub VulListBox()                                              *
' * ---------------------------------------------------------------------------
' * doel:       laadt de namen van de CSV-bestanden in de ListBox             *
' * gebruiker:  Private Sub UserForm_Initialize                               *
' * gebruikt:   Sub StatusCSV()                                               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Sub VulListBox()
    Dim nOffset As Integer              ' offset in lijst
    Dim sSheet As String                ' naam van werkblad
    nOffset = 1
    Dim n As Integer
    n = 0
    With lstNamen
        .Clear
        While Len(Trim([Schema].Offset(nOffset, 0))) > 0
            sSheet = [Schema].Offset(nOffset, 0)
            If Sheets(sSheet).Visible Then
                .AddItem
                .List(n, 0) = [Schema].Offset(nOffset, 0)
                .List(n, 1) = [Schema].Offset(nOffset, 2)
                n = n + 1
            End If
            nOffset = nOffset + 1
        Wend
    End With
    StatusCSV
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdKeuzeOmkeren_Click()                           *
' *             Private Sub cmdKiesAlle_Click()                               *
' *             Private Sub lstNamen_Change()                                 *
' * ---------------------------------------------------------------------------
' * doel:       Eventafhandeling                                              *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdKeuzeOmkeren_Click()
    Dim nX As Integer                   ' lusteller
    For nX = 1 To lstNamen.ListCount
        lstNamen.Selected(nX - 1) = Not lstNamen.Selected(nX - 1)
    Next
End Sub

Private Sub cmdKiesAlle_Click()
    Dim nX As Integer                   ' lusteller
    For nX = 1 To lstNamen.ListCount
        lstNamen.Selected(nX - 1) = True
    Next
End Sub

Private Sub lstNamen_Change()
    StatusCSV
End Sub

' *****************************************************************************
' * procedure:  Function ListCounter(lst As MSForms.ListBox) As Integer       *
' *             Sub Status(sTekst As String)                                  *
' *             Sub StatusCSV()                                               *
' * argumenten: lst                     naam van de listbox                   *
' * ---------------------------------------------------------------------------
' * doel:       tools                                                         *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Function ListCounter(lst As MSForms.ListBox) As Integer
    Dim nX As Integer                   ' lusteller
    Dim nN As Integer                   ' aantal geselecteerde items
    For nX = 0 To lst.ListCount - 1
        nN = nN + IIf(lst.Selected(nX), 1, 0)
    Next nX
    ListCounter = nN
End Function

Function NrplCounter(lst As MSForms.ListBox) As Long
    Dim nX As Integer                   ' lusteller
    Dim nNrpl As Long                   ' totaal aantal kentekens in gekozen sets
    nNrpl = 0
    For nX = 0 To lst.ListCount - 1
        If lst.Selected(nX) Then
            nNrpl = nNrpl + lst.List(nX, 1)
        End If
    Next nX
    NrplCounter = nNrpl
End Function


Sub Status(sTekst As String)
    Me.txtStatus = "==> " & sTekst
    DoEvents
    Application.EnableEvents = True
End Sub

Sub StatusCSV()
    Me.txtCSVStatus = ListCounter(frmSchema.lstNamen) & " / " & frmSchema.lstNamen.ListCount & " - " & NrplCounter(frmSchema.lstNamen)
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdSchema_Click()                                 *
' * ---------------------------------------------------------------------------
' * doel:       werkt een overzichtsschema uit met filtermogelijkheden        *
' * gebruikt:   Sub SchemaNaarUniek()                                         *
' *             Sub KleurAsterisk()                                           *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 09/05/19                    nieuwe versie met Collection, Array, ...
' *
Private Sub cmdSchema_Click()
    Dim nT As Long                      ' TickCount
    Dim nX As Double                    ' lusteller
    Dim nSets As Integer                ' gekozen sets voor Schema
    Dim nNaarRechts As Integer          ' kolom voor plaatsen van titels
    Dim sSheet As String                ' naam van worksheet in verwerking
    Dim nRecsInSet As Double            ' aantal records in worksheet in verwerking
    Dim aSet As Variant                 ' array voor sterretjes
    Dim CollAlles As Object             ' Collection met alle kentekens

    Dim sRangeAdres As String           ' bereik met alle gegevens in worksheet
    Dim nInvoegRij As Double            ' aantal records in schema tot nu toe
    Dim nTestRecs As Double             ' aantal te vergelijken records
    Dim rVergelijk As Range             ' range per set met combis land-nummerplaat
    Dim nY As Double                    ' lusteller
    Dim collSet As Object               ' Collection met gegevens van één set
    Dim z As Double

    Dim rVisible As Range
    Dim cVisible As Range
    Dim oColItem As Variant
    
    Me.cmdSchema.Caption = "Bezig..."
    Status ("gegevens verzamelen...")
    DC_Journal ("Schema - gegevens verzamelen")
    ' controleer of AutoFilter actief is, schakel uit indien nodig
    G_Schema.Activate
    If ActiveSheet.AutoFilterMode Then
        Selection.AutoFilter
    End If
    
    ' noteer de vinkjes uit ListBox en tel het aantal gekozen sets
    For nX = 1 To lstNamen.ListCount
        [Schema].Offset(nX, 1) = IIf(Me.lstNamen.Selected(nX - 1), 1, 0)
    Next nX
    nSets = Application.Sum([Schema].Offset(1, 1).Resize(1000, 1))
    If nSets = 0 Then
        MsgBox "Geen sets gekozen..."
        Exit Sub
    End If
    
    ' maak plaats voor nieuw schema, leegmaken inclusief voorwaardelijke opmaak
    G_Schema.Cells.Clear

    ' hoofdingen
    [A1] = "nrpl"
    For nX = 1 To lstNamen.ListCount
        If [Schema].Offset(nX, 1) = 1 Then
            nNaarRechts = nNaarRechts + 1
            Sheets("Schema").Cells(1, nNaarRechts + 1) = [Schema].Offset(nX, 0)
        End If
    Next nX
    ' alle kolommen behalve eerste
    With Columns(2).Resize(, nNaarRechts + 1 + 1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ColumnWidth = 3
    End With
    ' titelcellen behalve A1
    With Cells(1, 2).Resize(1, nNaarRechts + 1)
        .Orientation = 45
        .Borders.LineStyle = xlContinuous
        With .Interior
            .ThemeColor = 10
            .TintAndShade = 0.2
        End With
    End With
    With Cells(1, nNaarRechts + 2)
        With .Interior
            .ThemeColor = 8
            .TintAndShade = 0.2
        End With
    End With
    
    DoEvents
    ' verzamel eerst ALLE nummerplaten van de gekozen sets
    Application.ScreenUpdating = False
    nInvoegRij = 2
    
    For nX = 1 To nSets
        sSheet = G_Schema.Cells(1, nX + 1)
        Sheets(sSheet).Activate
        If ActiveSheet.AutoFilterMode = True Then
            Selection.AutoFilter
        End If
        nRecsInSet = LaatsteRij
        sRangeAdres = "A1:K" & nRecsInSet
        
        ' ook Belgische nummerplaten opnemen?
        If frmSchema.chkBelgen Then
            If ActiveSheet.AutoFilterMode Then
                Selection.AutoFilter
            End If
        Else
            If ActiveSheet.AutoFilterMode = False Then
                Selection.AutoFilter
            End If
            ActiveSheet.Range(sRangeAdres).AutoFilter Field:=7, Criteria1:="<>BE"
        End If
        
        ' kopieer per set de range met combinatie Landcode-Nummerplaat
        ActiveSheet.UsedRange.Range("K2:K" & nRecsInSet).SpecialCells(xlCellTypeVisible).Copy _
            Destination:=Sheets("Schema").Range("A" & nInvoegRij)
        nInvoegRij = nInvoegRij + ActiveSheet.UsedRange.Columns(11).SpecialCells(xlCellTypeVisible).count - 1
    Next nX
    Status ("naar unieke gegevens en sorteren...")
    DC_Journal ("naar unieke gegevens en sorteren...")
    ' *********************
    ' werk verder in Schema
    G_Schema.Activate
    SchemaNaarUniek
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = False
    DC_Journal ("unieke gegevens klaar...")
    ' vergelijk de hele reeks met elke set
    Status ("vergelijken met sets...")
    DC_Journal ("vergelijken met sets...")
    
    nTestRecs = G_Schema.Cells(Rows.count, 1).End(xlUp).Row
    ' stockeer alle kentekens uit Schema in een Collection, met hun rijnummer
    Set CollAlles = New Collection
    For nX = 2 To nTestRecs
        CollAlles.Add CStr(nX), Cells(nX, 1)
    Next nX
        
    ' stockeer elke set in zijn eigen Collection (unieke items), vergelijk die met Schema-Collection
    On Error Resume Next
    For nX = 1 To nSets
        gTcStart = GetTickCount
        sSheet = G_Schema.Cells(1, nX + 1)
        Sheets(sSheet).Activate
        nRecsInSet = LaatsteRij
        Set rVergelijk = Range("K2:K" & nRecsInSet)
        Status ("[" & nX & "] vergelijken met " & sSheet)
        G_Schema.Activate
        nT = GetTickCount
        If [cfgDictionary] Then
            On Error Resume Next
            Set collSet = New Collection
            ReDim aSet(nTestRecs)
            Sheets(sSheet).Activate
            
            ' selecteer de zichtbare cellen in de Set (houdt rekening met bv AutoFilter
            Set rVisible = ActiveSheet.UsedRange.Columns(11).SpecialCells(xlCellTypeVisible).Cells
            nRecsInSet = rVisible.Cells.count
            For Each cVisible In rVisible
                collSet.Add cVisible.Text, cVisible.Text
            Next cVisible
            gTcEinde = GetTickCount
            G_Schema.Activate
            On Error Resume Next
            
            ' zoek elk item uit set in Schema en plaats sterretjes
            For Each oColItem In collSet
                z = CollAlles(oColItem)
                If Err.Number = 0 Then
                    aSet(z) = "*"
                    Cells(z, nX + 1) = "*"
                Else
                    Err.Clear
                End If
            Next oColItem
        Else
            ' te controleren gegevens niet in array plaatsen
            ' array is te beperkt voor grote hoeveelheden...
            For nY = 2 To nTestRecs
                If Application.CountIf(rVergelijk, Cells(nY, 1)) > 0 Then
                    Cells(nY, nX + 1) = "*"
                End If
            Next nY
        
        End If
        gTcEinde = GetTickCount
        DC_Journal ("Set " & nX & " vergeleken in " & (gTcEinde - gTcStart) / 1000 & " sec")
        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False
    Next nX
    
    ' schema inkleuren?
    If Me.chkKleurInSchema Then
        Status ("kleur aanbrengen")
        DC_Journal ("kleur aanbrengen")
        [B2].Resize(nTestRecs, nSets).Select
        KleurAsterisk
    End If
    [B2].Resize(nTestRecs, nSets + 1).HorizontalAlignment = xlCenter
    
    ' hits tellen
    Cells(1, nNaarRechts + 2) = "Hits"
    Cells(2, nNaarRechts + 2).FormulaR1C1 = "=COUNTIF(RC[-" & nNaarRechts & "]:RC[-1],""*"")"
    Cells(2, nNaarRechts + 2).Select
    Selection.AutoFill Destination:=Cells(2, nNaarRechts + 2).Resize(nTestRecs - 1, 1)
        
    Cells(2, nNaarRechts + 2).Resize(nTestRecs, 1).Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = gnSCHAALGROEN
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = vbWhite
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = gnSCHAALROOD
    
    ' autofilter toevoegen, breedte eerste kolom en titels blokkeren
    With [A1].Interior
        .ThemeColor = 9
        .TintAndShade = 0.2
    End With
    Columns(1).AutoFit
    With Cells(1, 1).Resize(nTestRecs + 1, nNaarRechts + 2)
        .AutoFilter
        .Font = "Calibri"
        .Font.Size = 11
    End With
    'Cells(1, nNaarRechts + 3).Resize(1, 100000).Delete shift:=xlToLeft
    [A2].Select
    
    ActiveWindow.FreezePanes = True
    WisOverbodigeCellen
    DC_Kaders (ActiveSheet.UsedRange)
    Status "Klaar... " & Sheets("Schema").Cells(Rows.count, 1).End(xlUp).Row & " rijen"
    Me.cmdSchema.Caption = "Schema"
    DC_Journal ("Schema einde" & " - " & LaatsteRij & " rijen inclusief hoofding")
    Application.ScreenUpdating = True
    [A2].Activate
End Sub

' *****************************************************************************
' * procedure:  Sub SchemaNaarUniek()                                         *
' * ---------------------------------------------------------------------------
' * doel:       herleid de nummerplaten in het Schema tot unieke waarden      *
' *             sorteer de nummerplaten op landcode & nummerplaat
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Sub SchemaNaarUniek()
    Dim nRijen As Double                ' aantal te verwerken rijen
    Dim rR As Range                     ' range voor sterretjes
    Dim nT As Double
    
    nT = Timer
    nRijen = LaatsteRij
    ActiveSheet.Range("$A$1:$BZ$" & nRijen).RemoveDuplicates Columns:=1, Header:=xlYes
    
    nRijen = Cells(Rows.count, 1).End(xlUp).Row
    Set rR = Sheets("Schema").Range("A2:A" & nRijen)
    ' sorteer de nummerplaten op landcode + nummerplaat
    G_Schema.Sort.SortFields.Clear
    G_Schema.Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With G_Schema.Sort
        .SetRange rR
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' *****************************************************************************
' * procedure:  Sub KleurAsterisk()                                           *
' * ---------------------------------------------------------------------------
' * doel:       Schema inkleuren                                              *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Sub KleurAsterisk()
    With Selection
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlTextString, String:="*", TextOperator:=xlContains
        .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.399945066682943
        End With
        .FormatConditions(1).StopIfTrue = False
    End With
End Sub
'
' EINDE frmSchema *************************************************************

