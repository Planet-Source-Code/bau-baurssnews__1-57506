VERSION 5.00
Begin VB.UserControl InfoHeader 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "InfoHeader.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   4050
   Begin VB.PictureBox picGrad 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   -15
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   0
      Top             =   0
      Width           =   1785
      Begin VB.Image imgIco 
         Height          =   240
         Left            =   120
         Picture         =   "InfoHeader.ctx":0028
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   165
         Width           =   465
      End
   End
End
Attribute VB_Name = "InfoHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' Information Header v1.0
' Coded By: Jayson Ragasa
' Copyright© 2004 Baguio City, Philippines
' ----------------------------------------
'
' any comments or suggestion is appriciated
' if you like this, please vote for it! tnx!


Option Explicit

Public Enum CaptionAlignment
     AlignLeft = 0
     AlignCenter = 1
     AlignRight = 2
     AlignIconBottom = 3
End Enum

Public Enum GradStyle
     GradientHorizontal = 0
     GradientVertical = 1
End Enum

Dim CaptionAlign    As CaptionAlignment
Dim MltiLyn         As Boolean

Dim b_HasIcon       As Boolean
Dim pictIcon        As StdPicture

Dim GradStyl        As GradStyle
Dim Color_LEFT      As OLE_COLOR
Dim Color_RIGHT     As OLE_COLOR
Dim iMaxFill        As Integer

Public Property Let GradientStyle(ByVal newVal As GradStyle)
     GradStyl = newVal
     PropertyChanged "GradientStyle"
     
     Call reDraw(True)
End Property
Public Property Get GradientStyle() As GradStyle
     GradientStyle = GradStyl
End Property

Public Property Let MaxFill(ByVal newVal As Integer)
     If newVal > 100 Then newVal = 100
     
     iMaxFill = newVal
     PropertyChanged "MaxFill"
     
     Call reDraw(True)
End Property
Public Property Get MaxFill() As Integer
Attribute MaxFill.VB_ProcData.VB_Invoke_Property = "GradientStyle"
     MaxFill = iMaxFill
End Property

Public Property Let LeftColor(ByVal newVal As OLE_COLOR)
     Color_LEFT = ConvertRGBFormat(newVal)
     PropertyChanged "LeftColor"
     
     Call reDraw(True)
End Property
Public Property Get LeftColor() As OLE_COLOR
     LeftColor = Color_LEFT
End Property

Public Property Let RightColor(ByVal newVal As OLE_COLOR)
     Color_RIGHT = ConvertRGBFormat(newVal)
     PropertyChanged "RightColor"
     
     Call reDraw(True)
End Property
Public Property Get RightColor() As OLE_COLOR
     RightColor = Color_RIGHT
End Property

Public Property Set FontName(ByVal newVal As StdFont)
     Set lblCaption.Font = newVal
     PropertyChanged "FontName"
     
     Call reDraw(False)
End Property
Public Property Get FontName() As StdFont
     Set FontName = lblCaption.Font
End Property

Public Property Let ForeColor(ByVal newVal As OLE_COLOR)
     lblCaption.ForeColor = newVal
     PropertyChanged "ForeColor"

     Call reDraw(False)
End Property
Public Property Get ForeColor() As OLE_COLOR
     ForeColor = lblCaption.ForeColor
End Property

Public Property Let Caption(ByVal newVal As String)
     If Not MltiLyn Then
          lblCaption.Caption = newVal
     ElseIf MltiLyn Then
          lblCaption.Caption = DoMultiLining(newVal)
     End If
     
     PropertyChanged "Caption"
     
     Call reDraw(False)
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "TextStyle"
     Caption = lblCaption.Caption
End Property

Public Property Let Alignment(ByVal newVal As CaptionAlignment)
     If HasIcon = False And newVal = AlignIconBottom Then
          MsgBox "Cannot set alignment because 'HasIcon' is false"
          
          Exit Property
     End If
     
     CaptionAlign = newVal
     PropertyChanged "Alignment"
     
     Call reDraw(False)
End Property
Public Property Get Alignment() As CaptionAlignment
     Alignment = CaptionAlign
End Property

Public Property Let MultiLine(ByVal newVal As Boolean)
     MltiLyn = newVal
     
     If newVal = True Then
          lblCaption.Caption = DoMultiLining(Caption)
     ElseIf newVal = False Then
          lblCaption.Caption = RemoveCRLF(Caption)
     End If
     
     PropertyChanged "MultiLine"
     
     Call reDraw(False)
End Property
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_ProcData.VB_Invoke_Property = "TextStyle"
     MultiLine = MltiLyn
End Property

Public Property Let HasIcon(ByVal newVal As Boolean)
     If newVal = False Then
          If Alignment = AlignIconBottom Then
               Alignment = AlignLeft
               CaptionAlign = AlignLeft
          End If
     End If
     b_HasIcon = newVal
     PropertyChanged "HasIcon"
     
     Call reDraw(False)
End Property
Public Property Get HasIcon() As Boolean
     HasIcon = b_HasIcon
End Property

Public Property Set Picture(ByVal newVal As StdPicture)
     Set imgIco.Picture = newVal
     PropertyChanged "Picture"
     
     Call reDraw(False)
End Property
Public Property Get Picture() As StdPicture
     Set Picture = imgIco.Picture
End Property

Private Sub UserControl_Initialize()
     GradStyl = GradientHorizontal
     Color_LEFT = ConvertRGBFormat(SystemColorConstants.vb3DShadow)
     Color_RIGHT = ConvertRGBFormat(SystemColorConstants.vb3DFace)
     iMaxFill = 100
     
     lblCaption.Font.Name = "Tahoma"
     lblCaption.ForeColor = ConvertRGBFormat(SystemColorConstants.vb3DDKShadow)
     lblCaption.FontBold = True
     lblCaption.FontSize = 10
     lblCaption.Caption = "Information Header v1.0"
     
     CaptionAlign = AlignLeft
     MltiLyn = False
     
     b_HasIcon = True
     
     Call reDraw(True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     GradStyl = PropBag.ReadProperty("GradientStyle", GradStyle.GradientHorizontal)
     Color_LEFT = PropBag.ReadProperty("LeftColor", ConvertRGBFormat(SystemColorConstants.vb3DShadow))
     Color_RIGHT = PropBag.ReadProperty("RightColor", ConvertRGBFormat(SystemColorConstants.vb3DFace))
     iMaxFill = PropBag.ReadProperty("MaxFill", iMaxFill)
     
     lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", ConvertRGBFormat(SystemColorConstants.vb3DDKShadow))
     lblCaption.Font.Name = PropBag.ReadProperty("FontName", "Tahoma")
     lblCaption.Caption = PropBag.ReadProperty("Caption", "Information Header v1.0")
     MltiLyn = PropBag.ReadProperty("MultiLine", False)
     CaptionAlign = PropBag.ReadProperty("Alignment", CaptionAlignment.AlignLeft)
     
     b_HasIcon = PropBag.ReadProperty("HasIcon", True)
     Set imgIco.Picture = PropBag.ReadProperty("Picture", imgIco.Picture)
     
     Call reDraw(True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     PropBag.WriteProperty "GradientStyle", GradStyl
     PropBag.WriteProperty "LeftColor", Color_LEFT
     PropBag.WriteProperty "RightColor", Color_RIGHT
     PropBag.WriteProperty "MaxFill", iMaxFill
     
     PropBag.WriteProperty "FontName", lblCaption.Font.Name
     PropBag.WriteProperty "ForeColor", lblCaption.ForeColor
     PropBag.WriteProperty "Caption", lblCaption.Caption
     PropBag.WriteProperty "MultiLine", MltiLyn
     PropBag.WriteProperty "Alignment", CaptionAlign
     
     PropBag.WriteProperty "HasIcon", b_HasIcon
     PropBag.WriteProperty "Picture", imgIco.Picture
End Sub

Sub reDraw(ByVal reDrawGrad As Boolean)
     Dim tmpTop     As Long
     Dim tmpHeight  As Long
     
     If reDrawGrad = True Then
          picGrad.Cls
          Gradient picGrad, Color_LEFT, Color_RIGHT, iMaxFill, GradStyl
     End If
     
     If HasIcon Then
          imgIco.Visible = True
          
          
     
          If MltiLyn = False Then
               lblCaption.top = (picGrad.Height - lblCaption.Height) \ 2
               
               imgIco.Move 4 * Screen.TwipsPerPixelX, (picGrad.Height - imgIco.Height) \ 2
          ElseIf MltiLyn = True Then
               If imgIco.Height > lblCaption.Height Then
                    imgIco.top = (picGrad.Height - imgIco.Height) \ 2
                    
                    lblCaption.top = imgIco.top
               Else
                    lblCaption.top = (picGrad.Height - lblCaption.Height) \ 2
                    
                    imgIco.top = lblCaption.top
               End If
          End If
          
          If CaptionAlign = AlignLeft Then
               lblCaption.left = imgIco.left + imgIco.Width + (7 * Screen.TwipsPerPixelX)
               
               lblCaption.Alignment = AlignmentConstants.vbLeftJustify
          ElseIf CaptionAlign = AlignCenter Then
               lblCaption.Alignment = AlignmentConstants.vbCenter
               
               lblCaption.left = (((picGrad.Width - (imgIco.left + imgIco.Width)) - lblCaption.Width) \ 2) + (imgIco.left + imgIco.Width)
          ElseIf CaptionAlign = AlignRight Then
               lblCaption.left = picGrad.ScaleWidth - lblCaption.Width - (4 * Screen.TwipsPerPixelX)
               
               lblCaption.Alignment = AlignmentConstants.vbRightJustify
          ElseIf CaptionAlign = AlignIconBottom Then
               lblCaption.Alignment = AlignmentConstants.vbLeftJustify
               tmpHeight = imgIco.Height + lblCaption.Height + (2 * Screen.TwipsPerPixelY)
               tmpTop = (picGrad.ScaleHeight - tmpHeight) \ 2
               
               imgIco.top = tmpTop
               
               lblCaption.top = imgIco.top + imgIco.Height + (2 * Screen.TwipsPerPixelY)
               lblCaption.left = imgIco.left
          End If
     ElseIf Not HasIcon Then
          imgIco.Visible = False
          
          lblCaption.top = (picGrad.Height - lblCaption.Height) \ 2
          
          If CaptionAlign = AlignLeft Then
               lblCaption.left = 4 * Screen.TwipsPerPixelX
               
               lblCaption.Alignment = AlignmentConstants.vbLeftJustify
          ElseIf CaptionAlign = AlignCenter Then
               lblCaption.left = ((picGrad.ScaleWidth - lblCaption.Width) \ 2)
               
               lblCaption.Alignment = AlignmentConstants.vbCenter
          ElseIf CaptionAlign = AlignRight Then
               lblCaption.left = picGrad.ScaleWidth - lblCaption.Width - (4 * Screen.TwipsPerPixelX)
               
               lblCaption.Alignment = AlignmentConstants.vbRightJustify
          End If
     End If
End Sub

Private Sub UserControl_Resize()
     'If Height < 510 Then Height = 510
     
     picGrad.Move 0, 0, ScaleWidth, ScaleHeight
     
     Call reDraw(True)
End Sub

' DoMultiLining v1.0
Function DoMultiLining(ByVal sText As String) As String
     Dim exLines    As Variant
     Dim i          As Integer
     Dim tmp        As String
     
     If InStr(1, sText, "|") <> 0 Then
          exLines = Split(sText, "|")
               
          For i = 0 To UBound(exLines)
               tmp = tmp + exLines(i) + vbCrLf
          Next i
          
          DoMultiLining = left$(tmp, Len(tmp) - 2)
     Else
          DoMultiLining = sText
     End If
End Function

Function RemoveCRLF(ByVal sText As String) As String
     Dim i     As Integer
     Dim c     As String
     Dim ret   As String
     
     For i = 1 To Len(sText)
          c = Mid$(sText, i, 1)
          
          If Asc(c) = 10 Then
          ElseIf Asc(c) = 13 Then
               ret = ret + "|"
          Else
               ret = ret + c
          End If
     Next i
     
     RemoveCRLF = ret
End Function
