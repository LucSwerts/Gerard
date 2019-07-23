VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "Importeren van ANPR-gegevens"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10590
   OleObjectBlob   =   "frmImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport"
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
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        frmImport
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Importeren van Excel of CSV-bestanden in GERARD
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.21         2019-03-02              Conform nieuw protocol
' 02.44         2019-05-13              Refactoring
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' Private Sub UserForm_Initialize()
' Private Sub cmdExit_Click()
' Private Sub cmdMap_Click()
' Private Sub txtMapNaam_Change()
' Private Sub VulListBox()
' Private Sub cmdKeuzeOmkeren_Click()
' Private Sub cmdKiesAlle_Click()
' Private Sub lstNamen_Change()
' Sub Status(sTekst As String)
' Sub StatusImport()
' Private Sub cmdImporteren_Click()
' Private Sub CSV_Import(sMap As String, sBestand As String)
' Private Sub VerzamelSetInfo()

' Functions:
' ~~~~~~~~~~
' Function ListCounter(lst As MSForms.ListBox) As Integer
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Base 1


Dim sLog As String

' *****************************************************************************
' * procedure:  Private Sub UserForm_Initialize()                             *
' * ---------------------------------------------------------------------------
' * doel:       UserForm initialiseren                                        *
' *             gewenste instellingen activeren                               *
' * gebruikt:   Sub VulListBox()                                              *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 13/05/19    Luc S           herstel vorig type
' *
Private Sub UserForm_Initialize()
    Me.Caption = gsAPP & " - Importeren"
    If [cfgvorigemap] Then
        If Len(Trim([cfgnaamvorigemap])) > 0 Then
            If Dir([cfgnaamvorigemap], vbDirectory) <> vbNullString Then
                Me.txtMapNaam = Trim([cfgnaamvorigemap])
            Else
                Me.txtMapNaam = ThisWorkbook.path
            End If
        End If
    Else
        Me.txtMapNaam = ThisWorkbook.path
    End If
    If [cfgVorigTypeOnthouden] Then
        Select Case [cfgVorigTypeType]
            Case 1
                Me.optExcel = True
            Case 2
                Me.optComma = True
            Case 3
                Me.optSemiColon = True
        End Select
    Else
        Me.optExcel = True
    End If
    
    lstNamen.Enabled = True
    VulListBox
    sLog = ""
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdExit_Click()                                   *
' * ---------------------------------------------------------------------------
' * doel:       UserForm verlaten                                             *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 13/05/19    Luc S           onthoud type
' *
Private Sub cmdExit_Click()
    If [cfgVorigTypeOnthouden] Then
        If Me.optExcel Then
            [cfgVorigTypeType] = 1
        Else
            If Me.optComma Then
                [cfgVorigTypeType] = 2
            Else
                [cfgVorigTypeType] = 3
            End If
        End If
    End If
    Unload Me
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdMap_Click()                                    *
' * ---------------------------------------------------------------------------
' * doel:       UserForm verlaten                                             *
' * gebruikt:   Function KiesMap()                                            *
' * gebruikt:   Sub VulListBox()                                              *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdMap_Click()
    Dim sMap As String
    sMap = KiesMap(Me.txtMapNaam, gsKIESANPRTITEL)
    If sMap <> Me.txtMapNaam And Len(Trim(sMap)) > 0 Then
        Me.txtMapNaam = sMap
        VulListBox
        Status ("map [ " & sMap & "] gekozen...")
    End If
End Sub

Private Sub cmdDezeMap_Click()
    Me.txtMapNaam = ThisWorkbook.path
    Status ("map [ " & Me.txtMapNaam & "] gekozen...")
End Sub

' *****************************************************************************
' * procedure:  Private Sub txtMapNaam_Change()                               *
' * ---------------------------------------------------------------------------
' * doel:       namen van Excel / CSV-bestanden laden na keuze van map        *
' * gebruikt:   Sub VulListBox()                                              *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 13/12/18    Luc S           LaadGegevens als aparte procedure
' *                             ook aanroepbaar van CommandButton
' *
Private Sub txtMapNaam_Change()
    LaadGegevens
End Sub
        
Private Sub cmdLaadAlles_Click()
    LaadGegevens
End Sub
        
Private Sub optExcel_Change()
    LaadGegevens
End Sub
        
Private Sub optComma_Change()
    LaadGegevens
End Sub

Private Sub optSemiColon_Change()
    LaadGegevens
End Sub
        
Private Sub LaadGegevens()
    Dim sZoek As String                 ' te zoeken bestanden
    Dim sBestand As String              ' gevonden bestand
    Dim nAantal As Integer              ' aantal gevonden bestanden
    
    ThisWorkbook.Sheets("Dossier").Range("IMPORT").Offset(1, 0).Resize(500, 3).ClearContents
    If Me.optExcel Then
        sZoek = Me.txtMapNaam & "\*.XLS*"
    Else
        sZoek = Me.txtMapNaam & "\*.CSV"
    End If
    sBestand = Dir(sZoek)
    Do While sBestand <> "" And sBestand <> ActiveWorkbook.Name And UCase(Left(sBestand, 6)) <> "GERARD"
        nAantal = nAantal + 1
        [Import].Offset(nAantal, 0) = Left(sBestand, (InStrRev(sBestand, ".", -1, vbTextCompare) - 1))
        [Import].Offset(nAantal, 1) = 0
        sBestand = Dir()
    Loop
    VulListBox
End Sub

' *****************************************************************************
' * procedure:  Sub VulListBox()                                              *
' * ---------------------------------------------------------------------------
' * doel:       laadt de namen van de Excel / CSV-bestanden in de ListBox     *
' * gebruiker:  Private Sub UserForm_Initialize                               *
' *             Private Sub cmdMap_Click                                      *
' *             Private Sub txtMapNaam_Change                                 *
' * gebruikt:   Sub StatusImport()                                            *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Sub VulListBox()
    Dim nOffset As Integer              ' offset in lijst
    Dim nTeLang As Integer              ' aantal bestand met te lange naam
    nOffset = 1
    lstNamen.Clear
    Me.txtStatus = vbNullString
    While Len(Trim([Import].Offset(nOffset, 0))) > 0
        lstNamen.AddItem [Import].Offset(nOffset, 0)
        If Len([Import].Offset(nOffset, 0)) > 30 Then
            nTeLang = nTeLang + 1
        End If
        nOffset = nOffset + 1
    Wend
    StatusImport
    If nTeLang > 0 Then
        Me.txtStatus = nTeLang & " bestanden met te lange naam! Wijzig de naam tot 30 karakters of minder..."
    End If
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
    StatusImport
End Sub

' *****************************************************************************
' * procedure:  Function ListCounter(lst As MSForms.ListBox) As Integer       *
' *             Sub Status(sTekst As String)                                  *
' *             Sub StatusImport()                                            *
' * argumenten: lst                     naam van de listbox                   *
' * ---------------------------------------------------------------------------
' * doel:       telt geselecteerde items in listbox                           *
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

Sub Status(sTekst As String)
    Me.txtStatus = "==> " & sTekst
    Application.EnableEvents = True
End Sub

Sub StatusImport()
    Me.txtImportStatus = ListCounter(frmImport.lstNamen) & " / " & frmImport.lstNamen.ListCount
End Sub


' *****************************************************************************
' * procedure:  Private Sub cmdImporteren_Click()                                  *
' * ---------------------------------------------------------------------------
' * doel:       laadt de CSV-bestanden die gekozen werden in frmDataKeuze     *
' * gebruikt:   Sub Status()                                                  *
' *             Private Sub CSV_Import()                                      *
' *             Private Sub VerzamelSetInfo()                                 *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 13/12/18    Luc S           aanvulling voor laden van xls-bestanden
Private Sub cmdImporteren_Click()
    Dim nX As Integer                   ' lusteller
    Dim sBestand As String              ' naam van het te laden bestand
    Dim nGeladen As Integer             ' aantal geladen bestanden
    Dim nRecs As Double                 ' aantal records in bestand
    Dim nTotaalRecs As Double           ' totaal aantal geladen records
    
    Status (gsGEGEVENSLADEN)
    Application.ScreenUpdating = False
    
    ' onthoud map voor volgende keer
    [cfgnaamvorigemap] = frmImport.txtMapNaam
    
    ' noteer de vinkjes uit ListBox
    For nX = 1 To lstNamen.ListCount
        [Import].Offset(nX, 1) = IIf(Me.lstNamen.Selected(nX - 1), 1, 0)
    Next nX
    
    For nX = 1 To lstNamen.ListCount
        ' laad de aangevinkte bestanden
        If Me.lstNamen.Selected(nX - 1) Then
            nGeladen = nGeladen + 1
            sBestand = Range("IMPORT").Offset(nX, 0)
            Status ("[" & nGeladen & "] " & sBestand & " wordt geladen...")
            DoEvents
            If Me.optExcel Then
                Excel_Import frmImport.txtMapNaam, sBestand
            Else
                CSV_Import frmImport.txtMapNaam, sBestand
            End If
            
            nRecs = Sheets(sBestand).Cells(Rows.count, 1).End(xlUp).Row
            nRecs = DC_LaatsteRij(Sheets(sBestand))
            nTotaalRecs = nTotaalRecs + nRecs
        End If
    Next nX
    [A2].Select
    ActiveWindow.FreezePanes = True
    Status (nGeladen & " bestanden geladen met " & Format(nTotaalRecs, "#,##0") & " gegevens...")
    ' noteer gegevens van geladen sets
    VerzamelSetInfo
    Application.ScreenUpdating = True
    If Len(Trim(sLog)) > 0 Then
        frmLog.txtLog = sLog
        frmLog.Show
    End If
End Sub

' *****************************************************************************
' * procedure:  CSV_Import                                                    *
' * ---------------------------------------------------------------------------
' * doel:       laadt een CSV-bestand in een bepaalde map                     *
' * argumenten: sMap: map waar het bestand staat                              *
' *             sNaam: te laden bestand                                       *
' * gebruikt:   Sub DatumTijdKolomSplitser()                                  *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub CSV_Import(sMap As String, sBestand As String)
    Dim ws As Worksheet                 ' object nieuw worksheet
    Dim sBestandExt As String           ' bestandsnaam met extensie .csv
    
    ' controleer eerst of het werkblad nog niet bestaat...
    If WorksheetExists(sBestand) Then
        DC_Journal "import => " & frmImport.txtMapNaam & " | " & sBestand & " - bestaat al - niet geladen"
        sLog = sLog & sBestand & " bestaat al... - niet geladen!" & vbCrLf
        Exit Sub
    End If
    DC_Journal "import => " & frmImport.txtMapNaam & " | " & sBestand
    Set ws = Sheets.Add(after:=Sheets(Sheets.count))

    ' naam toekennen en tabkleur groen
    ws.Name = sBestand
    ws.Tab.ThemeColor = xlThemeColorAccent6
    ws.Tab.TintAndShade = 0.4
    
    ' voeg eventueel extensie toe
    sBestandExt = sBestand
    If UCase(Right(sBestandExt, 4)) <> ".CSV" Then
        sBestandExt = sBestandExt & ".CSV"
    End If

    ' laad CSV-gegevens met juiste delimiter (, of ;)
    With ws.QueryTables.Add(Connection:="TEXT;" & sMap & "\" & sBestandExt, Destination:=ws.Range("A1"))
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = frmImport.optSemiColon
        .TextFileCommaDelimiter = frmImport.optComma
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ' datum tijd splitsen
    DatumTijdKolomSplitser (2)
    Combineer
    BerekenAantallen
    OpmaakTitels
End Sub

' *****************************************************************************
' * procedure:  Excel_Import                                                  *
' * ---------------------------------------------------------------------------
' * doel:       laadt een Excel-bestand in een bepaalde map                   *
' * argumenten: sMap: map waar het bestand staat                              *
' *             sNaam: te laden bestand                                       *
' * gebruikt:   Sub DatumTijdKolomSplitser()                                  *
' * ---------------------------------------------------------------------------
' * 13/12/18    Luc S           creatie, v 1.00                               *
' *
Private Sub Excel_Import(sMap As String, sBestand As String)
    Dim ws As Worksheet                 ' object nieuw worksheet
    Dim wb As Workbook                  ' Gerard
    
    Set wb = ThisWorkbook
    ' controleer eerst of het werkblad nog niet bestaat...
    If WorksheetExists(sBestand) Then
        DC_Journal "import => " & frmImport.txtMapNaam & " | " & sBestand & " - bestaat al - niet geladen"
        sLog = sLog & sBestand & " bestaat al... - niet geladen!" & vbCrLf
        Exit Sub
    End If
    DC_Journal "import => " & frmImport.txtMapNaam & " | " & sBestand
    Set ws = Sheets.Add(after:=Sheets(Sheets.count))

    ' naam toekennen en tabkleur groen
    ws.Name = sBestand
    ws.Tab.ThemeColor = xlThemeColorAccent6
    ws.Tab.TintAndShade = 0.4
    
    ' laad Excel-gegevens
    Workbooks.Open Filename:=sMap & "\" & sBestand
    Sheets(1).UsedRange.Copy Destination:=ws.Range("A1")
    ActiveWorkbook.Close SaveChanges:=False
    ' terug naar GERARD - datum tijd splitsen
    DatumTijdKolomSplitser (2)
    Combineer
    BerekenAantallen
    OpmaakTitels
End Sub

Sub BerekenAantallen()
    Dim nAantal As Double
    nAantal = LaatsteRij()
    Cells(2, 12).FormulaR1C1 = "=COUNTIF(R2C[-1]:R" & nAantal & "C[-1],RC[-1])"
    Range("L2").AutoFill Destination:=Range("L2:L" & nAantal)
    Range("L2:L" & nAantal).Copy
    [L2].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Columns("L:L").EntireColumn.AutoFit
    [L1] = "Aantal"
End Sub

Sub Combineer()
    Dim sAdres As String
    
    ' combineer landcode en nummerplaat
    [K2].FormulaR1C1 = "=RC[-4]&""-""&RC[-5]"
    sAdres = "K2:K" & Cells(Rows.count, 1).End(xlUp).Row
    [K2].AutoFill Destination:=Range(sAdres)
    Range(sAdres).Copy
    [L2].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("K:K").Delete shift:=xlToLeft
    [K1] = "Combi"
    [L1] = "Aantal"
    Columns("M:Z").Delete
End Sub

' *****************************************************************************
' * procedure:  VerzamelSetInfo                                               *
' * ---------------------------------------------------------------------------
' * doel:       plaatst info over geladen bestanden in werkblad Dossier       *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub VerzamelSetInfo()
    Dim nX As Integer                   ' lusteller
    Dim nRijen As Integer               ' aantal dossiers
    Dim nRecs As Double                 ' aantal records
    Dim sSheet As String                ' naam van worksheet
    Dim nLijst As Integer               ' bestanden in lijst
    
    ' verwijder oude informatie over de beschikbare bestanden
    [Schema].Offset(1, 0).Resize(1000, 4).ClearContents
    ' tel dossiers na Import
    nRijen = Sheets("Dossier").Cells(Rows.count, 1).End(xlUp).Row - 2
    For nX = 1 To nRijen
         If Range("IMPORT").Offset(nX, 1) = 1 Then
            nLijst = nLijst + 1
            sSheet = Range("IMPORT").Offset(nX, 0)
            [Schema].Offset(nLijst, 0) = sSheet
            nRecs = Sheets(sSheet).Cells(Rows.count, 1).End(xlUp).Row
            [Schema].Offset(nLijst, 2) = nRecs
            ' onthoud alle ImportSets ook voor eventuele DrillDown
            [ZoomIn].Offset(nLijst, 0) = sSheet
            [ZoomIn].Offset(nLijst, 1) = 1
        End If
    Next nX
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmImport                                                                     '
'-----------------------------------------------------------------------------------------------
