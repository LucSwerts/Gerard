Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
End Sub

Sub KleurenPalet()
    Dim nAantal As Double
    Dim nX As Double
    Dim sHex As String
    nAantal = DC_LaatsteRij
    For nX = 1 To nAantal
        sHex = Cells(nX, 2)
        Cells(nX, 4) = HexToLongRGB(sHex)
        Cells(nX, 5).Interior.Color = Cells(nX, 4)
    Next nX
End Sub

Function HexToLongRGB(sHexVal As String) As Long
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long

    lRed = CLng("&H" & Left$(sHexVal, 2))
    lGreen = CLng("&H" & Mid$(sHexVal, 3, 2))
    lBlue = CLng("&H" & Right$(sHexVal, 2))

    HexToLongRGB = RGB(lRed, lGreen, lBlue)

End Function

Sub testkleur()
    Dim s As String
    s = "800000"
    Debug.Print HexToLongRGB(s)
End Sub

Sub SheetsVisible()
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.Sheets
        sh.Visible = True
    Next sh
End Sub
