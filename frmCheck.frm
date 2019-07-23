VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCheck 
   Caption         =   "Check"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   OleObjectBlob   =   "frmCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCheck"
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
' Gerard v. 1.20 - ANPR vs. Rondtrekkende Daders          frmCheck *
' ******************************************************************

' ************************
' * overzicht procedures *
' ************************
' * Private Sub UserForm_Initialize()
' * Private Sub cmdExit_Click()
' * Sub VerzamelChecks()
' * Sub VulListBoxChecks()
' * Private Sub cmdKeuzeOmkeren_Click()
' * Private Sub cmdKiesAlle_Click()
' * Private Sub lstNamen_Change()
' * Function ListCounter(lst As MSForms.ListBox) As Integer
' * Sub Status(sTekst As String)
' * Sub StatusCSV()
' * Private Sub cmdCheck_Click()

Option Explicit

' *****************************************************************************
' * procedure:  Private Sub UserForm_Initialize()                             *
' * ---------------------------------------------------------------------------
' * doel:       UserForm initialiseren                                        *
' *             gewenste instellingen activeren                               *
' * gebruikt:   Sub VerzamelChecks()                                          *
' *             Sub VulListBoxChecks()                                        *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub UserForm_Initialize()
    Me.Caption = gsAPP & " - Check"
    ' verzamel beschikbare Sets voor Schema
    VerzamelChecks
    VulListBoxChecks
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
' * procedure:  Sub VerzamelSets()                                            *
' * ---------------------------------------------------------------------------
' * doel:       Verzamel de namen van de beschikbare Sets             *
' * gebruiker:  Private Sub UserForm_Initialize                               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Sub VerzamelChecks()
    Dim nKolommen As Integer            ' totaal aantal werkbladen
    Dim nX As Integer                   ' lusteller
    Dim nOffset As Integer              ' offset in lijst
    
    nOffset = 1
    nKolommen = Sheets("SCHEMA").Cells(1, 1).End(xlToRight).Column - 2
    Range("CHECK").Offset(1, 0).Resize(1000, 3).ClearContents
    For nX = 1 To nKolommen
        Range("CHECK").Offset(nOffset, 0) = Sheets("SCHEMA").Cells(1, 1 + nX)
        Range("CHECK").Offset(nOffset, 1) = 0
        nOffset = nOffset + 1
    Next nX
End Sub

' *****************************************************************************
' * procedure:  Sub VulListBoxChecks()                                        *
' * ---------------------------------------------------------------------------
' * doel:       laadt de namen van de CSV-bestanden in de ListBox             *
' * gebruiker:  Private Sub UserForm_Initialize                               *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Sub VulListBoxChecks()
    Dim nOffset As Integer              ' offset in lijst
    
    nOffset = 1
    lstNamen.Clear
    While Len(Trim(Range("CHECK").Offset(nOffset, 0))) > 0
        lstNamen.AddItem Range("CHECK").Offset(nOffset, 0)
        nOffset = nOffset + 1
    Wend
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

Sub Status(sTekst As String)
    Me.txtStatus = "==> " & sTekst
    Application.EnableEvents = True
End Sub

Sub StatusCSV()
    Me.txtCSVStatus = ListCounter(Me.lstNamen) & " / " & Me.lstNamen.ListCount
End Sub

' *****************************************************************************
' * procedure:  Private Sub cmdCheck_Click()                                  *
' * ---------------------------------------------------------------------------
' * doel:       activeert AutoFilters naargelang de gekozen Sets              *
' * ---------------------------------------------------------------------------
' * 11/04/17    Luc S           creatie, v 1.00                               *
' *
Private Sub cmdCheck_Click()
    Dim nFilters As Integer             ' aantal te plaatsen Filters
    Dim nX As Integer                   ' lusteller
    Dim nRecs As Double                 ' aantal records
    Dim sAdres As String                ' adres voor Filters
    
    ' zet alle Filters af
    If ActiveSheet.AutoFilterMode Then
        If ActiveSheet.FilterMode Then
            ' AutoFilter blijft actief, maar zonder criteria
            ActiveSheet.ShowAllData
        End If
    End If
    
    nFilters = frmCheck.lstNamen.ListCount
    nRecs = Cells(Rows.count, 1).End(xlUp).Row
    sAdres = Range("A1").Resize(nRecs, frmCheck.lstNamen.ListCount + 2).Address
    
    For nX = 1 To nFilters
        If frmCheck.lstNamen.Selected(nX - 1) Then
            ActiveSheet.Range(sAdres).AutoFilter Field:=nX + 1, Criteria1:="<>"
        Else
            ActiveSheet.Range(sAdres).AutoFilter Field:=nX + 1, Criteria1:="="
        End If
    Next nX
End Sub
'
' EINDE frmSchema *************************************************************
