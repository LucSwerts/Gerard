VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZoom 
   Caption         =   "Zoom In"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20190
   OleObjectBlob   =   "frmZoom.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #  ## *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # # # *
' * #   # #     #  #  #   # #  #  #   #            # #      #           # ##  # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       ATLAS - Allerhande Tools ter Lichtere Arbeid van de Speurder
' Doel          Zoom in op een kenteken en zoek alle passages
'               zoek met jokers of in CombiModus
' Module        frmZoom
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Overzicht van passages van een kenteken in meer / alle Sets
' References    None

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 01.20         2017-05-04              Eerste Release
' 02.52         2019-05-24              Refactoring
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' Sub UserForm_Initialize()
' Sub cmdExit_Click()
' Sub cmdZoom_Click()
' DrillDown(sTarget As String)
' Sub cmdDump_Click()
' Sub optNaarBestand_Change()
' Sub cmdKleiner_Click()
' Sub cmdGroter_Click()
' Sub cmdStandaard_Click()
' Sub optJoker_Change()
' Sub cmdZoekFragment_Click()
' -

' Functions:
' ~~~~~~~~~~
' -
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Base 1


' **********************************************************************************************
' * Procedure:      Private Sub UserForm_Activate()
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   UserForm initialiseren
' * Aanroep van:    OOXML
' * Argumenten:     -                   -
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        OOXML
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-17      Eerste Release
' * 2017-04-26      aanpassing naar CombiModus
' *
Private Sub UserForm_Activate()
    Me.Caption = gsAPP & " - Zoom"
    Me.Height = 474 + IIf([cfgDevModus], 50, 0)
    Me.txtZoom.Font.Size = [cfgZoomPuntgrootte]
    Me.txtZoom.Font.Name = [cfgZoomLetterType]
    Me.optNaarKlembord = [cfgDumpNaarKlembord]
    Me.optNaarBestand = Not (Me.optNaarKlembord)
    Me.chkCombiModus = [cfgZoomCombiModus]
    Me.chkVariaties = Not ([cfgZoomNoteerAlles])
    Me.optSelectie = [cfgZoomSelectie]
    Me.optVolledig = Not [cfgZoomSelectie]
    Me.txtBestand = Format(Now, [cfgZoomDumpNaam])
    Me.optBegint = True
    Me.lblLettertype = [cfgZoomLetterType]
    
    DoEvents
    ' Plaats alle namen van werkbladen in ZOOMIN
    VerzamelSets ("ZOOMIN")
    ZoomTarget
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdExit_Click()                                   *
' * ---------------------------------------------------------------------------
' * doel:       UserForm verlaten                                             *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdExit_Click()
    Me.lblInfo = "Even geduld..."
    Me.lblInfo.ForeColor = vbRed
    DoEvents
    Application.Calculation = xlCalculationAutomatic
    Unload Me
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdZoom_Click()                                 *
' * ---------------------------------------------------------------------------
' * doel:       DrillDown met nieuwe Target                                   *
' * gebruikt:   Sub DrillDown()                                               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdZoom_Click()
    ZoomTarget
End Sub

Private Sub ZoomTarget()
    Dim sTarget As String               ' nummerplaat voor drill down
    Dim sCollage As String              ' resultaat van DrillDown
    
    If Me.optVolledig Then
        ' gebruik alle Sets
        [ZoomIn].Offset(1, 1).Resize([ZoomIn].End(xlDown).Row - 2).Value = 1
    Else
        ' gebruik de Sets die gebruikt werden voor het Schema
        SetsInSchema ("ZOOMIN")
    End If
    
    sTarget = Cells(ActiveCell.Row, 1)
    Me.lblKenteken = sTarget
    Me.txtZoom = sTarget & vbCrLf & vbCrLf
    ' onthoud Target om op te zoeken met joker
    Me.Tag = sTarget
    Me.txtCombiTeller = 0
    sCollage = DrillDown(sTarget)
    Me.txtZoom = Me.lblKenteken & vbCrLf & vbCrLf & sCollage & vbCrLf & "--- einde ---"
    G_Schema.Activate
End Sub

' *****************************************************************************
' * procedure:  Function DrillDown(sTarget As String, sVgl as String)         *
' * ---------------------------------------------------------------------------
' * doel:       DrillDown met bepaalde Target                                 *
' * gebruikt:   DC_journal()                                                     *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' * 26/04/17    Luc S           switch naar functie, ifv CombiModus           *
' *
Function DrillDown(sTarget As String, Optional sVgl As String) As String
    Dim lCombiModus As Boolean          ' exact zoeken om CombiModus
    Dim rWerk As Range                  ' werkrange (schema / import)
    Dim sCollage As String              ' collage met nieuwe resultaten
    
    Dim nSets As Integer                ' aantal aanwezige sets
    Dim nX As Integer                   ' lusteller
    Dim sSheet As String                ' naam van worksheet
    Dim nRijen As Double                ' aantal rijen in worksheet
    Dim sAdres As String                ' adres van volledige range
    Dim rGefilterd As Range             ' range van gefilterde records
    Dim sTekst As String                ' bijkomende info uit zichtbare cellen
    Dim rF As Range                     ' gefilterde rij in rGefilterd
    Dim nRijNr As Double                ' rijnummer van overeenkomst
    Dim sRijNr As String                ' rijnummer als string
    Dim nGevonden As Double             ' aantal gevonden
    Dim lNoteerAlles As Boolean         ' in CombiModus ook gelijken noteren
    
    ' snelste plaats voor xlCalculationManual is hier... Reden???
    Application.Calculation = xlCalculationManual
    ' zoek exact, dus sVgl = sTarget, geen CombiModus
    If IsMissing(sVgl) Or Len(Trim(sVgl)) = 0 Then
        sVgl = sTarget
        lCombiModus = False
    Else
        lCombiModus = True
    End If
    
    sTarget = UCase(sTarget)
    DC_Journal "DrillDown op " & sTarget & " / " & sVgl
    Application.ScreenUpdating = False
    lNoteerAlles = Not Me.chkVariaties
    
    Set rWerk = [ZoomIn]
    nSets = Application.CountA(rWerk.Offset(1, 0).Resize(1000, 1))
    If Me.optVolledig Then
        [ZoomIn].Offset(1, 1).Resize(nSets, 1) = 1
    Else
        [Schema].Offset(1, 1).Resize(nSets, 1).Copy Destination:=[ZoomIn].Offset(1, 1)
    End If
    ' welke sets werden geselecteerd
    For nX = 1 To nSets
        If rWerk.Offset(nX, 1) = 1 Then
            ' ga naar een geselecteerde set en noteer de naam van de worksheet
            sSheet = rWerk.Offset(nX, 0)
            Me.lblInfo = sSheet & "..."

            Sheets(sSheet).Activate
            ' test op inhoud van werkblad
            If [K1] = "Combi" And Left(ActiveSheet.Name, 5) <> "Thema" Then
            
                sCollage = sCollage & sSheet & vbCrLf
                ' stel AutoFilter in op kolom 11 (combi land-nummerplaat)
                ' eerst filter leegmaken
                ActiveSheet.Range("A1:L1").AutoFilter Field:=11
                nRijen = LaatsteRij()
                sAdres = Range("A1:L" & nRijen).Address
                ActiveSheet.Range(sAdres).AutoFilter Field:=11, Criteria1:=sTarget
                
                On Error Resume Next
                Set rGefilterd = Range(sAdres).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow
                If ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count = 1 Then
                    ' tellen van rijen werkt niet: telt alleen rijen in eerste aaneengesloten subrange
                    ' alleen titelrij gevonden
                    ' al dan niet CombiModus, noteer negatief resultaat,
                     sCollage = sCollage & Space(10) & "------------------------" & vbCrLf
                Else
                    For Each rF In rGefilterd
                        If rF.Cells(1, 11) Like sTarget Then
                            nRijNr = rF.Cells(1, 1).Row
                            If [cfgZoomUitlijnen] Then
                                sRijNr = Right("00000" & CStr(nRijNr), 6)
                            Else
                                sRijNr = CStr(nRijNr)
                            End If
                            sTekst = Space(10) & "rij " & sRijNr & " - id " & rF.Cells(1, 1) & " - "
                            sTekst = sTekst & Left(rF.Cells(1, 11) & Space(20), 12) & " - "
                            sTekst = sTekst & Format(rF.Cells(1, 3), [cfgFormatDatum]) & " - "
                            sTekst = sTekst & Format(rF.Cells(1, 4), [cfgFormatTijd]) & " - "
                            sTekst = sTekst & rF.Cells(1, 9) & " - " & rF.Cells(1, 10)
                            If lCombiModus Then
                                If lNoteerAlles Or Trim(rF.Cells(1, 11)) <> sVgl Then
                                    sCollage = sCollage & sTekst & vbCrLf
                                    nGevonden = nGevonden + 1
                                End If
                            Else
                                sCollage = sCollage & sTekst & vbCrLf
                                nGevonden = nGevonden + 1
                            End If
                        End If
                    Next
                End If
                Set rGefilterd = Nothing
                On Error GoTo 0
                Me.txtZoom = sTarget & vbCrLf & vbCrLf & sCollage
                DoEvents
            End If
        End If
    Next nX
    Me.txtCombiTeller = Me.txtCombiTeller + nGevonden
    Me.lblInfo = Me.Caption
    Application.ScreenUpdating = True
    DrillDown = sCollage
    Me.lblInfo = sTarget & " - " & nGevonden & " x gevonden"
End Function

' *****************************************************************************
' * procedure:  Private Sub cmdDump_Click()                                   *
' * ---------------------------------------------------------------------------
' * doel:       kopieer informatie naar KlemBord of naar een bestand          *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdDump_Click()
    Dim obj As New DataObject
    Dim sBestand As String
    If Me.optNaarKlembord Then
        obj.SetText Me.txtZoom
        obj.PutInClipboard
    Else
        sBestand = ActiveWorkbook.path & "\" & Me.txtBestand & ".txt"
        Open sBestand For Append Shared As #1
        'Width #1, 120
        Print #1, Me.txtZoom.Text
        Close #1
    End If
End Sub

' *****************************************************************************
' * procedure:  Private Sub optNaarBestand_Change()                           *
' * ---------------------------------------------------------------------------
' * doel:       txtBestand wordt actief als opNaarBestand actief is           *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub optNaarBestand_Change()
    txtBestand.Enabled = Me.optNaarBestand
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdKleiner_Click()                                *
' *             Private Sub cmdGroter_Click()                                 *
' *             Private Sub cmdStandaard_Click()                              *
' * ---------------------------------------------------------------------------
' * doel:       tools                                                         *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdKleiner_Click()
    txtZoom.Font.Size = Application.WorksheetFunction.Max(txtZoom.Font.Size - 2, 8)
    cmdStandaard.Caption = txtZoom.Font.Size
End Sub

Private Sub cmdGroter_Click()
    txtZoom.Font.Size = Application.WorksheetFunction.Min(txtZoom.Font.Size + 2, 16)
    cmdStandaard.Caption = txtZoom.Font.Size
End Sub

Private Sub cmdStandaard_Click()
    Me.txtZoom.Font.Size = Range("cfgZoomPuntGrootte")
    cmdStandaard.Caption = txtZoom.Font.Size
End Sub

' *****************************************************************************
' * procedure:  Private Sub optJoker_Change()                                 *
' *             Private Sub optCombi_Change()
' * ---------------------------------------------------------------------------
' * doel:       plaatst Tag in txtFragment als optJoker waar is               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub optJoker_Change()
    If Me.optJoker Then
        Me.txtFragment = Me.Tag
    End If
End Sub

Private Sub optCombi_Change()
    If Me.optCombi Then
        Me.txtFragment = Me.Tag
    End If
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdZoekFragment_Click()                           *
' * ---------------------------------------------------------------------------
' * doel:       zoekt txtFragment in alle sets, diverse opties voorzien       *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdZoekFragment_Click()
    Dim sIntro As String                ' introtekst
    Dim sFragment As String             ' nummerplaat voor drill down
    
    If Len(Trim(Me.txtBeginEindeBevat)) = 0 Then
        MsgBox ("zoeksleutel is leeg...")
        Exit Sub
    End If
    
    If Me.optBegint Then
        sIntro = "| Target begint met: " & vbCrLf
        sFragment = "" & Me.txtBeginEindeBevat & "*"
    ElseIf Me.optEindigt Then
        sIntro = "| Target eindigt met: " & vbCrLf
        sFragment = "*" & Me.txtBeginEindeBevat
    ElseIf Me.optBevat Then
        sIntro = "| Target bevat: " & vbCrLf
        sFragment = "*" & Me.txtBeginEindeBevat & "*"
    End If
    
    ' begin van Resultaat toont de vraagstelling
    Me.txtZoom = sIntro & "| " & sFragment & vbCrLf & vbCrLf
    Me.txtZoom = Me.txtZoom & DrillDown(sFragment)
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdZoekJoker_Click()                              *
' * ---------------------------------------------------------------------------
' * doel:       zoekt txtFragment in alle sets, diverse opties voorzien       *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdZoekJoker_Click()
    Dim nType As Integer                ' 1 x of 2 x
    Dim sIntro As String                ' introtekst
    Dim sFragment As String             ' nummerplaat voor drill down

    Dim nCombis As Integer              ' aantal combis
    Dim nX As Integer                   ' lusteller
    Dim sResult As String               ' resultaat van deelopzoeking
    Dim sCombiTekst As String           ' tekst met alle info
    Dim sTest As String                 ' test voor combi
    
    ' cumulatieve teller wordt gebruikt in CombiModus
    Me.txtCombiTeller = 0

    If Len(Trim(Me.txtFragment)) = 0 Then
        MsgBox ("zoeksleutel is leeg...")
        Exit Sub
    End If
    
    If Me.optJoker Then
        nType = 1
        sIntro = "| Target met joker(s): " & vbCrLf
        sFragment = Me.txtFragment
    ElseIf Me.optCombi Then
        nType = 2
        sIntro = "| Target in CombiModus: " & Me.txtFragment & vbCrLf
        sIntro = sIntro & IIf(Me.optVolledig, "| Zoek in alle sets ", "| Zoek alleen in de gekozen sets ") & vbCrLf
        sIntro = sIntro & IIf(Me.chkVariaties, "| Toon alleen variaties", "| Toon overeenkomsten en variaties") & vbCrLf
        sFragment = Me.txtFragment
    End If
    Me.txtZoom = sIntro & "| " & sFragment & vbCrLf
    
    If nType = 1 Then
        Me.txtZoom = Me.txtZoom & DrillDown(sFragment)
    Else
        nCombis = Len(sFragment) - 3
        For nX = 1 To nCombis
            sTest = Left(sFragment, 2 + nX) & "?" & Mid(sFragment, 4 + nX)
            ' bereken deeltekst
            sResult = DrillDown(sTest, sFragment)
            ' voeg deeltekst bij geheel
            sCombiTekst = sCombiTekst & vbCrLf & sTest & " (" & sFragment & ") " & vbCrLf
            sCombiTekst = sCombiTekst & vbCrLf & sResult & vbCrLf & "********************" & vbCrLf
        Next nX
        Me.txtZoom = sIntro & "| " & sFragment & vbCrLf
        Me.txtZoom = Me.txtZoom & sCombiTekst & vbCrLf & "--- einde ---"
        
        Me.lblInfo = Me.txtFragment & " - " & Me.txtCombiTeller & " x gevonden"
    End If
End Sub
'
' EINDE frmZoom **************************************************************
