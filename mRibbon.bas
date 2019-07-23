Attribute VB_Name = "mRibbon"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   ##### ##### *
' * #     #     #   # #   # #   # #   #                         #   #     #  ## *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # # # *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #   # ##  # *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        mRibbon
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Ribbon afhandeling
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.21         2019-03-02              Conform nieuw protocol
' 02.67         2019-07-03              Ribbon nieuwe opbouw, call-functies
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' * Sub Initialize(ribbon As IRibbonUI)
' * Sub G_RCB(control As IRibbonControl)
' * Sub rxInfo_getLabel(control As IRibbonControl, ByRef returnedVal)
' * Sub rxInfo_getSupertip(control As IRibbonControl, ByRef returnedVal)
'-----------------------------------------------------------------------------------------------

Option Explicit
Option Private Module

Private myRibbon As IRibbonUI

' **********************************************************************************************
' * Procedure:      Sub Initialize(ribbon As IRibbonUI)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Initialiseren van Ribbon
' * Aanroep van:    -
' * Argumenten:     ribbon              naam van Ribbon
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        OOXML
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release      creatie, v 1.00
' * 2019-03-02      Volledige revisie
'
Sub Initialize(ribbon As IRibbonUI)
    'Callback for customUI.onLoad
    Set myRibbon = ribbon
End Sub

' **********************************************************************************************
' * Procedure:      Sub G_RCB(control As IRibbonControl)
' * --------------------------------------------------------------------------------------------
' * Beschrijving:   Ribbon onAction
' * Aanroep van:    -
' * Argumenten:     control             control die gekozen wordt
' *                 -                   -
' * Gebruikt:       -
' * Scope:          Option Private Module
' * Aanroep:        OOXML
' * --------------------------------------------------------------------------------------------
' * Revisies overzicht:
' * 2017-04-11      Eerste Release      creatie, v 1.00
' * 2019-03-02      Volledige revisie
' * 2019-07-03      Ribbon nieuwe opbouw, code geherstructureerd, call-functies
'
Sub G_RCB(control As IRibbonControl)
    'Callback for onAction
    Select Case control.ID
        ' BLOK 1
        Case "G_Import":                frmImport.Show
        Case "G_Schema":                frmSchema.Show
            
        ' BLOK 2
        Case "G_Zoom":                  Call_Zoomin
        Case "G_Check":                 frmCheck.Show
        Case "G_Tandem":                frmTandem.Show
            
        ' BLOK 3
        Case "G_Inhoudstafel":          InhoudsTafel
        Case "G_Inventaris":            frmInventaris.Show
        Case "G_Thema":                 frmThema.Show
        Case "G_Puzzel":                Call_Puzzel
            
        ' BLOK 4
        Case "G_FilterClear":           FiltersUit
        Case "G_Sorteer":               Call_Sorteren
        Case "G_Accentueren":           Call_Accentueren
        Case "G_DatumTijd":             DatumTijdKolomSplitser
        
        ' BLOK 5
        Case "G_Config":                frmConfig.Show
        Case "G_Onderhoud":             frmOnderhoud.Show
        Case "G_Info":                  frmInfo.Show
    End Select
End Sub

' ****************************************************************************
' * Groep:          Callback functies voor Ribbon
' * --------------------------------------------------------------------------
' * Procedures:     Sub rxInfo_getLabel(control As IRibbonControl, ByRef returnedVal)
' *                 Sub rxInfo_getSupertip(control As IRibbonControl, ByRef returnedVal)
' * --------------------------------------------------------------------------
' * 2017-04-11      Eerste Release      creatie, v 1.00
'
Sub rxInfo_getLabel(control As IRibbonControl, ByRef returnedVal)
    'Callback for getLabel
    returnedVal = gsAPP & vbCrLf & gsVER & vbCrLf & gsBUILD
End Sub

Sub rxInfo_getSupertip(control As IRibbonControl, ByRef returnedVal)
    'Callback for getSupertip
    returnedVal = "Versie " & gsVER & " - build " & gsBUILD
End Sub

' ****************************************************************************
' * Groep:          Called procedures
' * --------------------------------------------------------------------------
' * Procedures:     Sub Call_Zoomin()
' *                 -
' * --------------------------------------------------------------------------
' * 2017-04-11      Eerste Release      creatie, v 1.00
'
Sub Call_Zoomin()
    If ActiveSheet.Name <> "Schema" Then
        MsgBox "werkt alleen in Schema..."
        Exit Sub
    Else
        If IsEmpty(ActiveCell) Then
            MsgBox ("plaats de cursor in een cel met gegevens...")
            Exit Sub
        End If
    End If
    frmZoom.Show
    ' reset Calculation, mogelijk nog Manual door foutief verlaten van user form
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub Call_Accentueren()
    If IsTandemSheet(ActiveSheet) Then
        ShowFrmAccentueren
    Else
        MsgBox "Accentueren werkt alleen op Tandem en Puzzels", vbInformation
    End If
End Sub

Sub Call_Puzzel()
    If IsTandemSheet(ActiveSheet) Then
        frmPuzzel.Show
    Else
        MsgBox "Puzzel werkt alleen op Tandem en Puzzels", vbInformation
    End If

End Sub

Sub Call_Sorteren()
    If Left(ActiveSheet.CodeName, 2) = "G_" And ActiveSheet.CodeName <> "G_Tandem" Then
        MsgBox "Sorteren kan alleen op ANPR-werkbladen, Thema's en Tandem", vbInformation
        Exit Sub
    End If
    frmSorteren.Show
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - mRibbon                                                                       '
'-----------------------------------------------------------------------------------------------

