Attribute VB_Name = "Module1"
Option Explicit

Public Sub SetCellBackgroundColorFromHex(ByVal inputRange As Range)

Dim vaRGB As Variant: vaRGB = HexToRGB(inputRange.Value)

Dim R As Byte: R = vaRGB(0)
Dim G As Byte: G = vaRGB(1)
Dim B As Byte: B = vaRGB(2)

inputRange.Interior.Color = RGB(R, G, B)

End Sub

Public Sub SetCellFontColorFromHex(ByVal inputRange As Range)

Dim vaRGB As Variant: vaRGB = HexToRGB(inputRange.Value)

Dim R As Byte: R = vaRGB(0)
Dim G As Byte: G = vaRGB(1)
Dim B As Byte: B = vaRGB(2)

'this is a comment
inputRange.Font.Color = RGB(R, G, B)

End Sub

Public Sub SetSelectionColorFromHexInCell()

Dim element As Variant

If IsArray(Selection) Then
    For Each element In Selection
        SetCellBackgroundColorFromHex element
        SetCellFontColorFromHex element
    Next element
Else
    SetCellBackgroundColorFromHex Selection
End If

End Sub

Public Function RGBToHex(R As Byte, G As Byte, B As Byte) As String

Dim output As String

output = Format(Hex(R), "00") & _
         Format(Hex(G), "00") & _
         Format(Hex(B), "00")

RGBToHex = output

End Function

Public Function HexToRGB(inputHex As String) As Variant

Dim R As Byte: R = Val("&H" & Mid(inputHex, 1, 2) & "&")
Dim G As Byte: G = Val("&H" & Mid(inputHex, 3, 2) & "&")
Dim B As Byte: B = Val("&H" & Mid(inputHex, 5, 2) & "&")

HexToRGB = Array(R, G, B)

End Function

