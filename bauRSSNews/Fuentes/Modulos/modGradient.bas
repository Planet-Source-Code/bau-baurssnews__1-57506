Attribute VB_Name = "modGradient"
Option Explicit

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Global rRed As Long, rBlue As Long, rGreen As Long

Dim R1 As Integer, R2 As Integer
Dim B1 As Integer, B2 As Integer
Dim G1 As Integer, G2 As Integer

'Public Enum GradientType
'     Horizontal = 0
'     Vertical = 1
'End Enum

Public Sub Gradient(ByRef picBox, _
                    ByVal StartColor As Long, _
                    ByVal EndColor As Long, _
                    ByVal MaxFill As Integer, _
                    ByVal GradType As GradStyle)
                    
     Dim i     As Integer
     Dim Color As Long
     
     Dim Size  As Long
     
     Dim GradR As Integer, GradB As Integer, GradG As Integer

     Call SetupColors(StartColor, EndColor)
     DoEvents
     
     If MaxFill > 100 Then MaxFill = 100

     picBox.AutoRedraw = True
     picBox.ScaleMode = vbPixels
     
     If GradType = GradientHorizontal Then Size = (picBox.ScaleWidth / 100) * MaxFill
     If GradType = GradientVertical Then Size = (picBox.ScaleHeight / 100) * MaxFill
     
     For i = 0 To Size
          GradR = ((R2 - R1) / Size * i) + R1
          GradG = ((G2 - G1) / Size * i) + G1
          GradB = ((B2 - B1) / Size * i) + B1
          
          Color = RGB(GradR, GradG, GradB)
          
          If GradType = GradientHorizontal Then
               picBox.Line (i, 0)-(i, picBox.ScaleHeight), Color, BF
          ElseIf GradType = GradientVertical Then
               picBox.Line (0, i)-(picBox.ScaleWidth, i), Color, BF
          End If
     Next i
     
     picBox.ScaleMode = vbTwips
End Sub

Sub SetupColors(ByVal StartColor, EndColor)
     ExtractRGBValues StartColor
     B1 = rBlue
     G1 = rGreen
     R1 = rRed

     ExtractRGBValues EndColor
     B2 = rBlue
     G2 = rGreen
     R2 = rRed
End Sub

Public Function ConvertRGBFormat(ByVal Color As OLE_COLOR) As Long
     TranslateColor Color, 0, ConvertRGBFormat
End Function

Function ExtractRGBValues(ByVal vColor As Long)
     rRed = (vColor And &HFF&)
     rGreen = (vColor And &HFF00&) / &H100
     rBlue = (vColor And &HFF0000) / &H10000
End Function
