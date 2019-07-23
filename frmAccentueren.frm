VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAccentueren 
   Caption         =   "Accentueren"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   OleObjectBlob   =   "frmAccentueren.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAccentueren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   #####   #   *
' * #     #     #   # #   # #   # #   #                         #       #  ##   *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # #   *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #       #   *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Module        frmAccentueren
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Accenten leggen in Tandem
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 02.58         2019-06-11              Nieuw
'
'-----------------------------------------------------------------------------------------------

' Subs:
' ~~~~~
' *
' *

' Functions:
' ~~~~~~~~~~

'-----------------------------------------------------------------------------------------------
Option Explicit

Private Sub UserForm_Activate()
    Me.Caption = gsAPP & " - Accentueren"
    Me.chkKleur = True
    Me.optTekstBlauw = True
    Me.optVerwijderOude = True
    Me.txtPuntGrootte = ActiveCell.Font.Size + 2
    Me.spnPuntGrootte.Min = 6
    Me.spnPuntGrootte.Max = 20
    Me.spnPuntGrootte = 12
    UpdateAccent
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub spnPuntGrootte_Change()
    Me.txtPuntGrootte = spnPuntGrootte
End Sub

Private Sub cmdAccentueer_Click()
    Dim rCel As Range
    Dim nAantalCellen As Double
    Dim sAntwoord As String
    Dim nTeller As Double
    Dim nStart As Integer
    Dim nLengte As Integer
    Dim nGrootte As Integer
    Dim sZoek As String
    Dim nUpdate As Double
    Dim r4 As Range
    Dim r8 As Range
    Dim r14 As Range
    Dim rBereik As Range
    Dim nRijen As Double
    
    sZoek = Trim(UCase(Me.txtZoekterm))
    If Len(sZoek) = 0 Then
        MsgBox "Zoekterm mag niet leeg zijn...", vbCritical
        Exit Sub
    End If
    
    If Me.optVerwijderOude Then
        HerstelOpmaak
    End If
    
    nLengte = Len(Trim(sZoek))
    nGrootte = Me.txtPuntGrootte
    If nGrootte > 0 Then
        If Me.txtPuntGrootte < 6 Or Me.txtPuntGrootte > 20 Then
            MsgBox "Kies een puntgrootte tussen 6 en 20!", vbCritical
            Exit Sub
        End If
    End If
            
    nRijen = DC_LaatsteRij()
    Set r4 = Range("D2").Resize(nRijen - 1, 1)
    Set r8 = Range("H2").Resize(nRijen - 1, 1)
    Set r14 = Range("N2").Resize(nRijen - 1, 1)
    Set rBereik = Union(r4, r8, r14)
    
    nAantalCellen = rBereik.Cells.count
    Me.cmdAccentueer.Caption = "Bezig..."
    Application.ScreenUpdating = False
    DoEvents
    For Each rCel In rBereik.SpecialCells(xlCellTypeVisible)
        nTeller = nTeller + 1
        If rCel <> "" Then
            Application.StatusBar = nTeller & " / " & nAantalCellen
            nStart = InStr(UCase(rCel), sZoek)
            If nStart > 0 Then
                nUpdate = nUpdate + 1
                With rCel.Characters(Start:=nStart, Length:=nLengte)
                    If Me.chkVet Then
                        .Font.Bold = True
                    End If
                    If Me.chkCursief Then
                        .Font.Italic = True
                    End If
                    If Me.chkOnderstreept Then
                        .Font.Underline = True
                    End If
                    If Me.chkDoorstreept Then
                        .Font.Strikethrough = True
                    End If
                    If Me.chkDubbelOnderstreept Then
                        .Font.Underline = xlUnderlineStyleDouble
                    End If
                    If Me.chkKleur Then
                        If Me.optTekstBlauw Then
                            .Font.Color = vbBlue
                        ElseIf Me.optTekstGroen Then
                            .Font.Color = vbGreen
                        ElseIf Me.optTekstRood Then
                            .Font.Color = vbRed
                        ElseIf Me.optTekstPaars Then
                            .Font.Color = -6279056
                        ElseIf Me.optTekstFuchsia Then
                            .Font.Color = -65281
                        End If
                    End If
                .Font.Size = Me.txtPuntGrootte
                End With
            End If
        End If
    Next rCel
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Me.cmdAccentueer.Caption = "Accentueer"
    Range("D:D,H:H,N:N").EntireColumn.AutoFit
    MsgBox nUpdate & " cellen aangepast...", vbInformation, gsAPP & " - Accentueren"
    Journal nUpdate & " cellen geaccentueerd voor zoekterm: " & Me.txtZoekterm
    
End Sub

Private Sub cmdHerstel_Click()
    HerstelOpmaak
End Sub

Private Sub HerstelOpmaak()
    Dim nRijen As Double                ' aantal rijen in Tandem
    Dim r4 As Range
    Dim r8 As Range
    Dim r14 As Range
    Dim rBereik As Range
    
    Me.cmdHerstel.Caption = "Bezig"
    DoEvents
    
    nRijen = DC_LaatsteRij()
    Set r4 = Range("D2").Resize(nRijen - 1, 1)
    Set r8 = Range("H2").Resize(nRijen - 1, 1)
    Set r14 = Range("N2").Resize(nRijen - 1, 1)
    Set rBereik = Union(r4, r8, r14)
    
    Application.ScreenUpdating = False
    With rBereik.Font
        .Bold = False
        .Italic = False
        .Underline = False
        .Strikethrough = False
        .Color = vbBlack
        .Size = Range("A2").Font.Size
    End With
    Range("D:D,H:H,N:N").EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Me.cmdHerstel.Caption = "Herstel"
    Journal "Tandem - opmaak van cellen hersteld"
End Sub
'
'-----------------------------------------------------------------------------------------------
' Einde module - frmAccentueren                                                                '
'-----------------------------------------------------------------------------------------------
