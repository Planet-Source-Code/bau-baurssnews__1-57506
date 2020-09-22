VERSION 5.00
Begin VB.UserControl ButtonXP 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ButtonXP.ctx":0000
End
Attribute VB_Name = "ButtonXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" _
  (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function OleTranslateColor Lib "oleaut32.dll" _
  (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
  (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, _
   ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" _
  (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" _
  (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function Ellipse Lib "gdi32" _
  (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" _
  (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" _
  (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" _
  (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0

Private Declare Function CreateRectRgn Lib "gdi32" _
  (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" _
  (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" _
  (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
  (ByVal hwnd As Long, ByVal hRGN As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" _
  (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" _
  (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function WindowFromPoint Lib "user32" _
  (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" _
  (lpPoint As POINTAPI) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type RGBTRIPLE
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBTRIPLE
End Type

Public Enum ButtonTypes
  [Windows 16-bit] = 1    'the old-fashioned Win16 button
  [Windows 32-bit] = 2    'the classic windows button
  [Windows XP] = 3        'the new brand XP button totally owner-drawn
  [Mac] = 4               'i suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
  [Java metal] = 5        'there are also other styles but not so different from windows one
  [Netscape 6] = 6        'this is the button displayed in web-pages, it also appears in some java apps
  [Simple Flat] = 7       'the standard flat button seen on toolbars
  [Flat Highlight] = 8    'again the flat button but this one has no border until the mouse is over it
  [Office XP] = 9         'the new Office XP button
  '[MacOS-X] = 10         'this is a plan for the future...
  [Transparent] = 11      'suggested from a user...
  [3D Hover] = 12         'took this one from "Noteworthy Composer" toolbal
  [Oval Flat] = 13        'a simple Oval Button
  [KDE 2] = 14            'the great standard KDE2 button!
End Enum

Public Enum ColorTypes
  [Use Windows] = 1
  [Custom] = 2
  [Force Standard] = 3
  [Use Container] = 4
End Enum

Public Enum PicPositions
  cbLeft = 0
  cbRight = 1
  cbTop = 2
  cbBottom = 3
  cbBackground = 4
End Enum

Public Enum fx
  cbNone = 0
  cbEmbossed = 1
  cbEngraved = 2
  cbShadowed = 3
End Enum

Private Const FXDEPTH As Long = &H28

'events
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()

'variables
Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes
Private PicPosition As PicPositions
Private SFX As fx 'font and picture effects

Private He As Long  'the height of the button
Private Wi As Long  'the width of the button

Private BackC As Long 'back color
Private BackO As Long 'back color when mouse is over
Private ForeC As Long 'fore color
Private ForeO As Long 'fore color when mouse is over
Private MaskC As Long 'mask color
Private OXPb As Long, OXPf As Long
Private useMask As Boolean, useGrey As Boolean
Private useHand As Boolean

Private picNormal As StdPicture, picHover As StdPicture
Private pDC As Long, pBM As Long, oBM As Long 'used for the treansparent button

Private elTex As String     'current text

Private rc As RECT, rc2 As RECT, rc3 As RECT, fc As POINTAPI 'text and focus rect locations
Private picPT As POINTAPI, picSZ As POINTAPI  'picture Position & Size
Private rgnNorm As Long

Private LastButton As Byte, LastKeyDown As Byte, lastStat As Byte
Private blnEnabled As Boolean, blnSoft As Boolean
Private blnHasFocus As Boolean, blnShowFocusRectangle As Boolean

Private cFace As Long, cLight As Long, cHighLight As Long
Private cShadow As Long, cDarkShadow As Long, cText As Long
Private cTextO As Long, cFaceO As Long, cMask As Long, XPFace As Long

Private TE As String

Private blnIsOver As Boolean, blnInLoop As Boolean, blnIsShown As Boolean  'used to avoid unnecessary repaints

Private Locked As Boolean

Private captOpt As Long
Private blnIsCheckbox As Boolean, blnCheckBoxValue As Boolean

Private Sub moveMouse()
  If Not IsMouseOver() Then
    blnIsOver = False
    Redraw 0, True
    RaiseEvent MouseOut
    ReleaseCapture
  Else
    SetCapture UserControl.hwnd
  End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  LastButton = 1
  UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  If MyColorType <> [Custom] Then
    SetColors
    Redraw lastStat, True
  End If
End Sub

Private Sub UserControl_Click()
  If LastButton = vbLeftButton And blnEnabled Then
    If blnIsCheckbox Then
      blnCheckBoxValue = Not blnCheckBoxValue
    End If
    Redraw 0, True 'be sure that the normal status is drawn
    UserControl.Refresh
    RaiseEvent Click
  End If
End Sub

Private Sub UserControl_DblClick()
  If LastButton = vbLeftButton Then
    UserControl_MouseDown 1, 0, 0, 0
    SetCapture hwnd
  End If
End Sub

Private Sub UserControl_GotFocus()
  blnHasFocus = True
  Redraw lastStat, True
End Sub

Private Sub UserControl_Hide()
  blnIsShown = False
End Sub

Private Sub UserControl_Initialize()
  'this makes the control to be slow, remark this line if the "not redrawing" problem is not important for you: ie, you intercept the Load_Event (with breakpoint or messageBox) and the button does not repaint...
  blnIsShown = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
  LastKeyDown = KeyCode
  Select Case KeyCode
    Case 32 'spacebar pressed
      Redraw 2, False
    Case 39, 40 'right and down arrows
      SendKeys "{Tab}"
    Case 37, 38 'left and up arrows
      SendKeys "+{Tab}"
  End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
  If KeyCode = 32 And LastKeyDown = 32 Then 'spacebar pressed, and not cancelled by the user
    If blnIsCheckbox Then
      blnCheckBoxValue = Not blnCheckBoxValue
    End If
    Redraw 0, False
    UserControl.Refresh
    RaiseEvent Click
  End If
End Sub

Private Sub UserControl_LostFocus()
  blnHasFocus = False
  Redraw lastStat, True
End Sub

Private Sub UserControl_InitProperties()
    blnEnabled = True
    blnShowFocusRectangle = True
    useMask = True
    elTex = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    MyButtonType = [Windows 32-bit]
    MyColorType = [Use Windows]
    SetColors
    BackC = cFace
    BackO = BackC
    ForeC = cText
    ForeO = ForeC
    MaskC = &HC0C0C0
    CalcTextRects
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
  LastButton = Button
  If Button <> vbRightButton Then
    Redraw 2, False
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
  moveMouse
  If Button < vbRightButton Then
    If Not IsMouseOver Then
      'we are outside the button
      Redraw 0, False
    Else
      'we are inside the button
      If Button = 0 And Not blnIsOver Then
        blnIsOver = True
        Redraw 0, True
        RaiseEvent MouseOver
      ElseIf Button = 1 Then
        blnIsOver = True
        Redraw 2, False
        blnIsOver = False
      End If
    End If
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
  If Button <> vbRightButton Then
    Redraw 0, False
  End If
End Sub

'########## BUTTON PROPERTIES ##########
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
  BackColor = BackC
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
  BackC = theCol
  If Not Ambient.UserMode Then
    BackO = theCol
  End If
  SetColors
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get BackOver() As OLE_COLOR
Attribute BackOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
  BackOver = BackO
End Property

Public Property Let BackOver(ByVal theCol As OLE_COLOR)
  BackO = theCol
  SetColors
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = ForeC
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
  ForeC = theCol
  If Not Ambient.UserMode Then
    ForeO = theCol
  End If
  SetColors
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get ForeOver() As OLE_COLOR
Attribute ForeOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ForeOver = ForeO
End Property

Public Property Let ForeOver(ByVal theCol As OLE_COLOR)
  ForeO = theCol
  SetColors
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  MaskColor = MaskC
End Property

Public Property Let MaskColor(ByVal theCol As OLE_COLOR)
  MaskC = theCol
  SetColors
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get ButtonType() As ButtonTypes
Attribute ButtonType.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal NewValue As ButtonTypes)
  MyButtonType = NewValue
  If MyButtonType = [Java metal] And Not Ambient.UserMode Then
    UserControl.FontBold = True
  ElseIf MyButtonType = Transparent And blnIsShown Then
    GetParentPic
  End If
  UserControl_Resize
  PropertyChanged
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = 0
  Caption = elTex
End Property

Public Property Let Caption(ByVal NewValue As String)
  elTex = NewValue
  SetAccessKeys
  CalcTextRects
  Redraw 0, True
  PropertyChanged
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
  Enabled = blnEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  blnEnabled = NewValue
  Redraw 0, True
  UserControl.Enabled = blnEnabled
  PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByRef NewFont As Font)
  Set UserControl.Font = NewFont
  CalcTextRects
  Redraw 0, True
  PropertyChanged
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
  FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal NewValue As Boolean)
  UserControl.FontBold = NewValue
  CalcTextRects
  Redraw 0, True
  PropertyChanged
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
  FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal NewValue As Boolean)
  UserControl.FontItalic = NewValue
  CalcTextRects
  Redraw 0, True
  PropertyChanged
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
  FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal NewValue As Boolean)
  UserControl.FontUnderline = NewValue
  CalcTextRects
  Redraw 0, True
  PropertyChanged
End Property

Public Property Get FontSize() As Integer
Attribute FontSize.VB_MemberFlags = "400"
  FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal NewValue As Integer)
  UserControl.FontSize = NewValue
  CalcTextRects
  Redraw 0, True
  PropertyChanged
End Property

Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
  FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal NewValue As String)
  UserControl.FontName = NewValue
  CalcTextRects
  Redraw 0, True
  PropertyChanged
End Property

'it is very common that a windows user uses custom color
'schemes to view his/her desktop, and is also very
'common that this color scheme has weird colors that
'would alter the nice look of my buttons.
'So if you want to force the button to use the windows
'standard colors you may change this property to "Force Standard"
Public Property Get ColorScheme() As ColorTypes
Attribute ColorScheme.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal NewValue As ColorTypes)
  MyColorType = NewValue
  SetColors
  Redraw 0, True
  PropertyChanged
End Property

Public Property Get ShowFocusRectangle() As Boolean
Attribute ShowFocusRectangle.VB_ProcData.VB_Invoke_Property = ";Appearance"
  ShowFocusRectangle = blnShowFocusRectangle
End Property

Public Property Let ShowFocusRectangle(ByVal NewValue As Boolean)
  blnShowFocusRectangle = NewValue
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Appearance"
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)
  UserControl.MousePointer = newPointer
  PropertyChanged
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal newIcon As StdPicture)
  On Local Error Resume Next
    Set UserControl.MouseIcon = newIcon
    PropertyChanged
End Property

Public Property Get HandPointer() As Boolean
  HandPointer = useHand
End Property

Public Property Let HandPointer(ByVal newVal As Boolean)
  useHand = newVal
  If useHand Then
    Set UserControl.MouseIcon = LoadResPicture(101, 2)
    UserControl.MousePointer = 99
  Else
    Set UserControl.MouseIcon = Nothing
    UserControl.MousePointer = 1
  End If
  PropertyChanged
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
  hwnd = UserControl.hwnd
End Property

Public Property Get SoftBevel() As Boolean
Attribute SoftBevel.VB_ProcData.VB_Invoke_Property = ";Appearance"
  SoftBevel = blnSoft
End Property

Public Property Let SoftBevel(ByVal NewValue As Boolean)
  blnSoft = NewValue
  SetColors
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get PictureNormal() As StdPicture
Attribute PictureNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Set PictureNormal = picNormal
End Property

Public Property Set PictureNormal(ByVal newPic As StdPicture)
  Set picNormal = newPic
  CalcPicSize
  CalcTextRects
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get PictureOver() As StdPicture
Attribute PictureOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
  Set PictureOver = picHover
End Property

Public Property Set PictureOver(ByVal newPic As StdPicture)
  Set picHover = newPic
  If blnIsOver Then
    Redraw lastStat, True 'only redraw i we need to see this picture immediately
  End If
  PropertyChanged
End Property

Public Property Get PicturePosition() As PicPositions
Attribute PicturePosition.VB_ProcData.VB_Invoke_Property = ";Position"
  PicturePosition = PicPosition
End Property

Public Property Let PicturePosition(ByVal newPicPos As PicPositions)
  PicPosition = newPicPos
  PropertyChanged
  CalcTextRects
  Redraw lastStat, True
End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
  UseMaskColor = useMask
End Property

Public Property Let UseMaskColor(ByVal NewValue As Boolean)
  useMask = NewValue
  If Not picNormal Is Nothing Then
    Redraw lastStat, True
  End If
  PropertyChanged
End Property

Public Property Get UseGreyscale() As Boolean
Attribute UseGreyscale.VB_ProcData.VB_Invoke_Property = ";Appearance"
  UseGreyscale = useGrey
End Property

Public Property Let UseGreyscale(ByVal NewValue As Boolean)
  useGrey = NewValue
  If Not picNormal Is Nothing Then
    Redraw lastStat, True
  End If
  PropertyChanged
End Property

Public Property Get SpecialEffect() As fx
Attribute SpecialEffect.VB_ProcData.VB_Invoke_Property = ";Appearance"
  SpecialEffect = SFX
End Property

Public Property Let SpecialEffect(ByVal NewValue As fx)
  SFX = NewValue
  Redraw lastStat, True
  PropertyChanged
End Property

Public Property Get CheckBoxBehaviour() As Boolean
  CheckBoxBehaviour = blnIsCheckbox
End Property

Public Property Let CheckBoxBehaviour(ByVal NewValue As Boolean)
  blnIsCheckbox = NewValue
  Redraw lastStat, True
End Property

Public Property Get Value() As Boolean
  Value = blnCheckBoxValue
End Property

Public Property Let Value(ByVal NewValue As Boolean)
  blnCheckBoxValue = NewValue
  If blnIsCheckbox Then
    Redraw 0, True
  End If
  PropertyChanged
End Property
'########## END OF PROPERTIES ##########

Private Sub UserControl_Resize()
  If Not blnInLoop Then
    'get button size
    GetClientRect UserControl.hwnd, rc3
    'assign these values to He and Wi
    He = rc3.Bottom
    Wi = rc3.Right
    'build the FocusRect size and position depending on the button type
    If MyButtonType >= [Simple Flat] And MyButtonType <= [Oval Flat] Then
      InflateRect rc3, -3, -3
    ElseIf MyButtonType = [KDE 2] Then
      InflateRect rc3, -5, -5
      OffsetRect rc3, 1, 1
    Else
      InflateRect rc3, -4, -4
    End If
    CalcTextRects
    If rgnNorm Then
      DeleteObject rgnNorm
    End If
    MakeRegion
    SetWindowRgn UserControl.hwnd, rgnNorm, True
    If He Then
      Redraw 0, True
    End If
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    BackColor = .ReadProperty("Backcolor", GetSysColor(COLOR_BTNFACE))
    BackOver = .ReadProperty("Backover", GetSysColor(COLOR_BTNFACE))
    ForeColor = .ReadProperty("ForeColor", GetSysColor(COLOR_BTNTEXT))
    ForeOver = .ReadProperty("ForeOver", &HC00000)
    MaskColor = .ReadProperty("MaskColor", &HC0C0C0)
    ButtonType = .ReadProperty("ButtonType", 2)
    Caption = .ReadProperty("Caption", "")
    Enabled = .ReadProperty("Enabled", True)
    Font = .ReadProperty("Font", UserControl.Font)
    FontBold = .ReadProperty("FontBold", False)
    FontItalic = .ReadProperty("FontItalic", False)
    FontSize = .ReadProperty("FontSize", UserControl.FontSize)
    FontName = .ReadProperty("FontName", UserControl.FontName)
    ColorScheme = .ReadProperty("ColorScheme", 1)
    ShowFocusRectangle = .ReadProperty("ShowFocusRectangle", True)
    MousePointer = .ReadProperty("MousePointer", 0)
    Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
    HandPointer = .ReadProperty("HandPointer", False)
    SoftBevel = .ReadProperty("SoftBevel", False)
    Set PictureNormal = .ReadProperty("PictureNormal", Nothing)
    Set PictureOver = .ReadProperty("PictureOver", Nothing)
    PicturePosition = .ReadProperty("PicturePosition", 0)
    UseMaskColor = .ReadProperty("UseMaskColor", True)
    UseGreyscale = .ReadProperty("UseGreyScale", False)
    SpecialEffect = .ReadProperty("SpecialEffect", 0)
    CheckBoxBehaviour = .ReadProperty("CheckBoxBehaviour", False)
    Value = .ReadProperty("Value", False)
  End With
  UserControl.Enabled = blnEnabled
  CalcPicSize
  CalcTextRects
  SetAccessKeys
End Sub

Private Sub UserControl_Show()
  If MyButtonType = 11 Then
    If pDC = 0 Then
      pDC = CreateCompatibleDC(UserControl.hdc)
      pBM = CreateBitmap(Wi, He, 1, GetDeviceCaps(hdc, 12), ByVal 0&)
      oBM = SelectObject(pDC, pBM)
    End If
    GetParentPic
  End If
  blnIsShown = True
  SetColors
  Redraw 0, True
End Sub

Private Sub UserControl_Terminate()
  blnIsShown = False
  DeleteObject rgnNorm
  If pDC Then
    DeleteObject SelectObject(pDC, oBM)
    DeleteDC pDC
  End If
  ReleaseCapture
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "BackColor", BackColor, GetSysColor(COLOR_BTNFACE)
    .WriteProperty "BackOver", BackOver, GetSysColor(COLOR_BTNFACE)
    .WriteProperty "ForeColor", ForeColor, GetSysColor(COLOR_BTNTEXT)
    .WriteProperty "ForeOver", ForeOver, &HC00000
    .WriteProperty "MaskColor", MaskColor, &HC0C0C0
    .WriteProperty "ButtonType", ButtonType, 2
    .WriteProperty "Caption", Caption, ""
    .WriteProperty "Enabled", Enabled, True
    .WriteProperty "Font", Font, UserControl.Font
    .WriteProperty "FontBold", FontBold, False
    .WriteProperty "FontItalic", FontItalic, False
    .WriteProperty "FontSize", FontSize, UserControl.FontSize
    .WriteProperty "FontName", FontName, UserControl.FontName
    .WriteProperty "ColorScheme", ColorScheme, 1
    .WriteProperty "ShowFocusRectangle", ShowFocusRectangle, True
    .WriteProperty "MousePointer", MousePointer, 0
    .WriteProperty "MouseIcon", MouseIcon, Nothing
    .WriteProperty "HandPointer", HandPointer, False
    .WriteProperty "SoftBevel", SoftBevel, False
    .WriteProperty "PictureNormal", PictureNormal, Nothing
    .WriteProperty "PictureOver", PictureOver, Nothing
    .WriteProperty "PicturePosition", PicturePosition, 0
    .WriteProperty "UseMaskColor", UseMaskColor, True
    .WriteProperty "UseGreyScale", UseGreyscale, False
    .WriteProperty "SpecialEffect", SpecialEffect, 0
    .WriteProperty "CheckBoxBehaviour", CheckBoxBehaviour, False
    .WriteProperty "Value", Value, False
  End With
End Sub

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
'here is the CORE of the button, everything is drawn here
'it's not well commented but i think that everything is
'pretty self explanatory...
If blnIsCheckbox And blnCheckBoxValue Then
  curStat = 2
End If

If Not Force Then  'check drawing redundancy
  If curStat = lastStat And TE = elTex Then
    Exit Sub
  End If
End If
If He = 0 Or Not blnIsShown Then
  Exit Sub   'we don't want errors
End If

lastStat = curStat
TE = elTex

Dim i As Long, stepXP1 As Single, XPFace2 As Long, tempCol As Long

With UserControl
.Cls
If blnIsOver And MyColorType = Custom Then
  tempCol = BackC
  BackC = BackO
  SetColors
End If

DrawRectangle 0, 0, Wi, He, cFace

If blnEnabled Then
    If curStat = 0 Then
'#@#@#@#@#@# BUTTON NORMAL STATE #@#@#@#@#@#
        Select Case MyButtonType
            Case ButtonTypes.[Windows 16-bit] 'Windows 16-bit
                DrawCaption Abs(blnIsOver)
                DrawFrame cHighLight, cShadow, cHighLight, cShadow, True
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                Call DrawFocusR
            Case ButtonTypes.[Windows 32-bit] 'Windows 32-bit
                DrawCaption Abs(blnIsOver)
                If Ambient.DisplayAsDefault And blnShowFocusRectangle Then
                    DrawFrame cHighLight, cDarkShadow, cLight, cShadow, True
                    Call DrawFocusR
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                Else
                    DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
                End If
            Case ButtonTypes.[Windows XP] 'Windows XP
                stepXP1 = 25 / He
                For i = 1 To He
                  DrawLine 0, i, Wi, i, ShiftColor(XPFace, -stepXP1 * i, True)
                Next i
                DrawCaption Abs(blnIsOver)
                DrawRectangle 0, 0, Wi, He, &H733C00, True
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                If blnIsOver Then
                  DrawRectangle 1, 2, Wi - 2, He - 4, &H31B2FF, True
                  DrawLine 2, He - 2, Wi - 2, He - 2, &H96E7&
                  DrawLine 2, 1, Wi - 2, 1, &HCEF3FF
                  DrawLine 1, 2, Wi - 1, 2, &H8CDBFF
                  DrawLine 2, 3, 2, He - 3, &H6BCBFF
                  DrawLine Wi - 3, 3, Wi - 3, He - 3, &H6BCBFF
                ElseIf ((blnHasFocus Or Ambient.DisplayAsDefault) And blnShowFocusRectangle) Then
                  DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                  DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                  DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                  DrawLine 1, 2, Wi - 1, 2, &HF7D7BD
                  DrawLine 2, 3, 2, He - 3, &HF0D1B5
                  DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                Else 'we do not draw the bevel always because the above code would repaint over it
                  DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H30, True)
                  DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPFace, -&H20, True)
                  DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H24, True)
                  DrawLine Wi - 3, 3, Wi - 3, He - 3, ShiftColor(XPFace, -&H18, True)
                  DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPFace, &H10, True)
                  DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPFace, &HA, True)
                  DrawLine 1, 2, 1, He - 2, ShiftColor(XPFace, -&H5, True)
                  DrawLine 2, 3, 2, He - 3, ShiftColor(XPFace, -&HA, True)
                End If
            Case ButtonTypes.Mac 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                DrawCaption Abs(blnIsOver)
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                DrawLine 1, 2, 2, 0, cFace
                DrawLine 3, 2, Wi - 3, 2, cHighLight
                DrawLine 2, 2, 2, He - 3, cHighLight
                mSetPixel 3, 3, cHighLight
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
                DrawLine 1, He - 3, Wi - 3, He - 3, cFace
                mSetPixel Wi - 4, He - 4, cFace
                DrawLine Wi - 2, 2, Wi - 2, He - 2, cShadow
                DrawLine 2, He - 2, Wi - 2, He - 2, cShadow
                mSetPixel Wi - 3, He - 3, cShadow
            Case ButtonTypes.[Java metal] 'Java
                DrawRectangle 1, 1, Wi - 1, He - 1, ShiftColor(cFace, &HC)
                DrawCaption Abs(blnIsOver)
                DrawRectangle 1, 1, Wi - 1, He - 1, cHighLight, True
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                mSetPixel 1, He - 2, ShiftColor(cShadow, &H1A)
                mSetPixel Wi - 2, 1, ShiftColor(cShadow, &H1A)
                If blnHasFocus And blnShowFocusRectangle Then
                  DrawRectangle rc.Left - 2, rc.Top - 1, fc.x + 4, fc.y + 2, &HCC9999, True
                End If
            Case ButtonTypes.[Netscape 6] 'Netscape
                DrawCaption Abs(blnIsOver)
                DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False
                DrawFocusR
            Case ButtonTypes.[Simple Flat], ButtonTypes.[Flat Highlight], ButtonTypes.[3D Hover] 'Flat buttons
                DrawCaption Abs(blnIsOver)
                If (MyButtonType = [Simple Flat]) Then
                  DrawFrame cHighLight, cShadow, 0, 0, False, True
                ElseIf blnIsOver Then
                  If MyButtonType = [Flat Highlight] Then
                    DrawFrame cHighLight, cShadow, 0, 0, False, True
                  Else
                    DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False, False
                  End If
                End If
                DrawFocusR
            Case ButtonTypes.[Office XP] 'Office XP
                If blnIsOver Then
                  DrawRectangle 1, 1, Wi, He, OXPf
                End If
                DrawCaption Abs(blnIsOver)
                If blnIsOver Then
                  DrawRectangle 0, 0, Wi, He, OXPb, True
                End If
                DrawFocusR
            Case ButtonTypes.Transparent 'transparent
                BitBlt hdc, 0, 0, Wi, He, pDC, 0, 0, vbSrcCopy
                DrawCaption Abs(blnIsOver)
                DrawFocusR
            Case ButtonTypes.[Oval Flat] 'Oval
                DrawEllipse 0, 0, Wi, He, Abs(blnIsOver) * cShadow + Abs(Not blnIsOver) * cFace, cFace
                DrawCaption Abs(blnIsOver)
            Case ButtonTypes.[KDE 2] 'KDE 2
                Dim prevBold As Boolean
                If Not blnIsOver Then
                  stepXP1 = 58 / He
                  For i = 1 To He
                    DrawLine 0, i, Wi, i, ShiftColor(cHighLight, -stepXP1 * i)
                  Next i
                Else
                  DrawRectangle 0, 0, Wi, He, cLight
                End If
                If Ambient.DisplayAsDefault Then
                  blnIsShown = False
                  prevBold = Me.FontBold
                  Me.FontBold = True
                End If
                DrawCaption Abs(blnIsOver)
                If Ambient.DisplayAsDefault Then
                  Me.FontBold = prevBold
                  blnIsShown = True
                End If
                DrawRectangle 0, 0, Wi, He, ShiftColor(cShadow, -&H32), True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cFace, -&H9), True
                DrawRectangle 2, 2, Wi - 4, 2, cHighLight
                DrawRectangle 2, 4, 2, He - 6, cHighLight
                DrawFocusR
        End Select
        DrawPictures (0)
    ElseIf curStat = 2 Then
'#@#@#@#@#@# BUTTON IS DOWN #@#@#@#@#@#
        Select Case MyButtonType
            Case ButtonTypes.[Windows 16-bit] 'Windows 16-bit
                DrawCaption 2
                DrawFrame cShadow, cHighLight, cShadow, cHighLight, True
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                DrawFocusR
            Case ButtonTypes.[Windows 32-bit] 'Windows 32-bit
                DrawCaption 2
                If blnShowFocusRectangle And Ambient.DisplayAsDefault Then
                  DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                  DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                  DrawFocusR
                Else
                  DrawFrame cDarkShadow, cHighLight, cShadow, cLight, False
                End If
            Case ButtonTypes.[Windows XP] 'Windows XP
                stepXP1 = 25 / He
                XPFace2 = ShiftColor(XPFace, -32, True)
                For i = 1 To He
                    DrawLine 0, He - i, Wi, He - i, ShiftColor(XPFace2, -stepXP1 * i, True)
                Next i
                DrawCaption 2
                DrawRectangle 0, 0, Wi, He, &H733C00, True
                mSetPixel 1, 1, &H7B4D10
                mSetPixel 1, He - 2, &H7B4D10
                mSetPixel Wi - 2, 1, &H7B4D10
                mSetPixel Wi - 2, He - 2, &H7B4D10
                DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPFace2, &H10, True)
                DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPFace2, &HA, True)
                DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPFace2, &H5, True)
                DrawLine Wi - 3, 3, Wi - 3, He - 3, XPFace
                DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPFace2, -&H20, True)
                DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPFace2, -&H18, True)
                DrawLine 1, 2, 1, He - 2, ShiftColor(XPFace2, -&H20, True)
                DrawLine 2, 2, 2, He - 2, ShiftColor(XPFace2, -&H16, True)
            Case ButtonTypes.Mac 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                XPFace = ShiftColor(cShadow, -&H10)
                DrawCaption 2
                XPFace = ShiftColor(cFace, &H30)
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                mSetPixel 2, 2, ShiftColor(cShadow, -&H40)
                mSetPixel 3, 3, ShiftColor(cShadow, -&H20)
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cShadow
                DrawLine 1, He - 3, Wi - 2, He - 3, cShadow
                mSetPixel Wi - 4, He - 4, cShadow
                DrawLine Wi - 2, 3, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                DrawLine 3, He - 2, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                DrawLine Wi - 2, He - 3, Wi - 4, He - 1, ShiftColor(cShadow, -&H20)
                mSetPixel 2, He - 2, ShiftColor(cShadow, -&H20)
                mSetPixel Wi - 2, 2, ShiftColor(cShadow, -&H20)
            Case ButtonTypes.[Java metal] 'Java
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                DrawLine Wi - 1, 1, Wi - 1, He, cHighLight
                DrawLine 1, He - 1, Wi - 1, He - 1, cHighLight
                SetTextColor .hdc, cTextO
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
                If blnHasFocus And blnShowFocusRectangle Then
                  DrawRectangle rc.Left - 2, rc.Top - 1, fc.x + 4, fc.y + 2, &HCC9999, True
                End If
            Case ButtonTypes.[Netscape 6] 'Netscape
                DrawCaption 2
                DrawFrame cShadow, ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), False
                DrawFocusR
             Case ButtonTypes.[Simple Flat], ButtonTypes.[Flat Highlight], ButtonTypes.[3D Hover] 'Flat buttons
                DrawCaption 2
                If MyButtonType = [3D Hover] Then
                  DrawFrame cDarkShadow, cHighLight, cShadow, cLight, False, False
                Else
                  DrawFrame cShadow, cHighLight, 0, 0, False, True
                End If
                DrawFocusR
            Case ButtonTypes.[Office XP] 'Office XP
                If blnIsOver Then
                  DrawRectangle 0, 0, Wi, He, Abs(MyColorType = 2) * ShiftColor(OXPf, -&H20) + Abs(MyColorType <> 2) * ShiftColorOXP(OXPb, &H80)
                End If
                DrawCaption 2
                DrawRectangle 0, 0, Wi, He, OXPb, True
                DrawFocusR
            Case ButtonTypes.Transparent 'transparent
                BitBlt hdc, 0, 0, Wi, He, pDC, 0, 0, vbSrcCopy
                DrawCaption 2
                DrawFocusR
            Case ButtonTypes.[Oval Flat] 'Oval
                DrawEllipse 0, 0, Wi, He, cDarkShadow, ShiftColor(cFace, -&H20)
                DrawCaption 2
            Case ButtonTypes.[KDE 2] 'KDE 2
                DrawRectangle 1, 1, Wi, He, ShiftColor(cFace, -&H9)
                DrawRectangle 0, 0, Wi, He, ShiftColor(cShadow, -&H30), True
                DrawLine 2, He - 2, Wi - 2, He - 2, cHighLight
                DrawLine Wi - 2, 2, Wi - 2, He - 1, cHighLight
                DrawCaption 7
                DrawFocusR
        End Select
        DrawPictures 1
    End If
Else
'#~#~#~#~#~# DISABLED STATUS #~#~#~#~#~#
    Select Case MyButtonType
        Case ButtonTypes.[Windows 16-bit] 'Windows 16-bit
            DrawCaption 3
            DrawFrame cHighLight, cShadow, cHighLight, cShadow, True
            DrawRectangle 0, 0, Wi, He, cDarkShadow, True
        Case ButtonTypes.[Windows 32-bit] 'Windows 32-bit
            DrawCaption 3
            DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
        Case ButtonTypes.[Windows XP] 'Windows XP
            DrawRectangle 0, 0, Wi, He, ShiftColor(XPFace, -&H18, True)
            DrawCaption 5
            DrawRectangle 0, 0, Wi, He, ShiftColor(XPFace, -&H54, True), True
            mSetPixel 1, 1, ShiftColor(XPFace, -&H48, True)
            mSetPixel 1, He - 2, ShiftColor(XPFace, -&H48, True)
            mSetPixel Wi - 2, 1, ShiftColor(XPFace, -&H48, True)
            mSetPixel Wi - 2, He - 2, ShiftColor(XPFace, -&H48, True)
        Case ButtonTypes.Mac 'Mac
            DrawRectangle 1, 1, Wi - 2, He - 2, cLight
            DrawCaption 3
            DrawRectangle 0, 0, Wi, He, cDarkShadow, True
            mSetPixel 1, 1, cDarkShadow
            mSetPixel 1, He - 2, cDarkShadow
            mSetPixel Wi - 2, 1, cDarkShadow
            mSetPixel Wi - 2, He - 2, cDarkShadow
            DrawLine 1, 2, 2, 0, cFace
            DrawLine 3, 2, Wi - 3, 2, cHighLight
            DrawLine 2, 2, 2, He - 3, cHighLight
            mSetPixel 3, 3, cHighLight
            DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
            DrawLine 1, He - 3, Wi - 3, He - 3, cFace
            mSetPixel Wi - 4, He - 4, cFace
            DrawLine Wi - 2, 2, Wi - 2, He - 2, cShadow
            DrawLine 2, He - 2, Wi - 2, He - 2, cShadow
            mSetPixel Wi - 3, He - 3, cShadow
        Case ButtonTypes.[Java metal] 'Java
            DrawCaption 4
            DrawRectangle 0, 0, Wi, He, cShadow, True
        Case ButtonTypes.[Netscape 6] 'Netscape
            DrawCaption 4
            DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False
        Case ButtonTypes.[Simple Flat], ButtonTypes.[Flat Highlight], ButtonTypes.[3D Hover], ButtonTypes.[Oval Flat] 'Flat buttons
            DrawCaption 3
            If MyButtonType = [Simple Flat] Then
              DrawFrame cHighLight, cShadow, 0, 0, False, True
            End If
        Case ButtonTypes.[Office XP] 'Office XP
            DrawCaption 4
        Case ButtonTypes.Transparent 'transparent
            BitBlt hdc, 0, 0, Wi, He, pDC, 0, 0, vbSrcCopy
            DrawCaption 3
        Case ButtonTypes.[KDE 2] 'KDE 2
            stepXP1 = 58 / He
            For i = 1 To He
                DrawLine 0, i, Wi, i, ShiftColor(cHighLight, -stepXP1 * i)
            Next i
            DrawRectangle 0, 0, Wi, He, ShiftColor(cShadow, -&H32), True
            DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cFace, -&H9), True
            DrawRectangle 2, 2, Wi - 4, 2, cHighLight
            DrawRectangle 2, 4, 2, He - 6, cHighLight
            DrawCaption 6
    End Select
    DrawPictures 2
End If
End With
If blnIsOver And MyColorType = Custom Then
  BackC = tempCol
  SetColors
End If
End Sub

Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
'this is my custom function to draw rectangles and frames
'it's faster and smoother than using the line method
Dim bRECT As RECT
Dim hBrush As Long

  bRECT.Left = x
  bRECT.Top = y
  bRECT.Right = x + Width
  bRECT.Bottom = y + Height
  hBrush = CreateSolidBrush(Color)
  If OnlyBorder Then
    FrameRect UserControl.hdc, bRECT, hBrush
  Else
    FillRect UserControl.hdc, bRECT, hBrush
  End If
  DeleteObject hBrush
End Sub

Private Sub DrawEllipse(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal BorderColor As Long, ByVal FillColor As Long)
Dim pBrush As Long, pPen As Long

  pBrush = SelectObject(hdc, CreateSolidBrush(FillColor))
  pPen = SelectObject(hdc, CreatePen(PS_SOLID, 2, BorderColor))
  Ellipse hdc, x, y, x + Width, y + Height
  DeleteObject SelectObject(hdc, pBrush)
  DeleteObject SelectObject(hdc, pPen)
End Sub

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'a fast way to draw lines
Dim pt As POINTAPI
Dim oldPen As Long, hPen As Long

  With UserControl
    hPen = CreatePen(PS_SOLID, 1, Color)
    oldPen = SelectObject(.hdc, hPen)
    MoveToEx .hdc, X1, Y1, pt
    LineTo .hdc, X2, Y2
    SelectObject .hdc, oldPen
    DeleteObject hPen
  End With
End Sub

Private Sub DrawFrame(ByVal ColHigh As Long, ByVal ColDark As Long, ByVal ColLight As Long, ByVal ColShadow As Long, ByVal ExtraOffset As Boolean, Optional ByVal Flat As Boolean = False)
'a very fast way to draw windows-like frames
Dim pt As POINTAPI
Dim frHe As Long, frWi As Long, frXtra As Long

  frHe = He - 1 + ExtraOffset: frWi = Wi - 1 + ExtraOffset: frXtra = Abs(ExtraOffset)
  With UserControl
    DeleteObject SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColHigh))
    '=============================
    MoveToEx .hdc, frXtra, frHe, pt
    LineTo .hdc, frXtra, frXtra
    LineTo .hdc, frWi, frXtra
    '=============================
    DeleteObject SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColDark))
    '=============================
    LineTo .hdc, frWi, frHe
    LineTo .hdc, frXtra - 1, frHe
    MoveToEx .hdc, frXtra + 1, frHe - 1, pt
    If Flat Then
      Exit Sub
    End If
    '=============================
    DeleteObject SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColLight))
    '=============================
    LineTo .hdc, frXtra + 1, frXtra + 1
    LineTo .hdc, frWi - 1, frXtra + 1
    '=============================
    DeleteObject SelectObject(.hdc, CreatePen(PS_SOLID, 1, ColShadow))
    '=============================
    LineTo .hdc, frWi - 1, frHe - 1
    LineTo .hdc, frXtra, frHe - 1
  End With
End Sub

Private Sub mSetPixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
  SetPixel UserControl.hdc, x, y, Color
End Sub

Private Sub DrawFocusR()
  If blnShowFocusRectangle And blnHasFocus Then
    SetTextColor UserControl.hdc, cText
    DrawFocusRect UserControl.hdc, rc3
  End If
End Sub

Private Sub SetColors()
'this function sets the colors taken as a base to build
'all the other colors and styles.
If MyColorType = Custom Then
    cFace = ConvertFromSystemColor(BackC)
    cFaceO = ConvertFromSystemColor(BackO)
    cText = ConvertFromSystemColor(ForeC)
    cTextO = ConvertFromSystemColor(ForeO)
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
    cDarkShadow = ShiftColor(cFace, -&HC0)
    OXPb = ShiftColor(cFace, -&H80)
    OXPf = cFace
ElseIf MyColorType = [Force Standard] Then
    cFace = &HC0C0C0
    cFaceO = cFace
    cShadow = &H808080
    cLight = &HDFDFDF
    cDarkShadow = &H0
    cHighLight = &HFFFFFF
    cText = &H0
    cTextO = cText
    OXPb = &H800000
    OXPf = &HD1ADAD
ElseIf MyColorType = [Use Container] Then
    cFace = GetBkColor(GetDC(GetParent(hwnd)))
    cFaceO = cFace
    cText = GetTextColor(GetDC(GetParent(hwnd)))
    cTextO = cText
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F)
    cDarkShadow = ShiftColor(cFace, -&HC0)
    OXPb = GetSysColor(COLOR_HIGHLIGHT)
    OXPf = ShiftColorOXP(OXPb)
Else
'if MyColorType is 1 or has not been set then use windows colors
    cFace = GetSysColor(COLOR_BTNFACE)
    cFaceO = cFace
    cShadow = GetSysColor(COLOR_BTNSHADOW)
    cLight = GetSysColor(COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
    cText = GetSysColor(COLOR_BTNTEXT)
    cTextO = cText
    OXPb = GetSysColor(COLOR_HIGHLIGHT)
    OXPf = ShiftColorOXP(OXPb)
End If
cMask = ConvertFromSystemColor(MaskC)
XPFace = ShiftColor(cFace, &H30, MyButtonType = [Windows XP])
End Sub

Private Sub MakeRegion()
'this function creates the regions to "cut" the UserControl
'so it will be transparent in certain areas
Dim rgn1 As Long, rgn2 As Long
    
DeleteObject rgnNorm
rgnNorm = CreateRectRgn(0, 0, Wi, He)
rgn2 = CreateRectRgn(0, 0, 0, 0)
Select Case MyButtonType
    Case ButtonTypes.[Windows 16-bit], ButtonTypes.[Java metal], ButtonTypes.[KDE 2] 'Windows 16-bit, Java & KDE 2
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        If MyButtonType <> ButtonTypes.[Java metal] Then  'the above was common code
          rgn1 = CreateRectRgn(0, 0, 1, 1)
          CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
          DeleteObject rgn1
          rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 1)
          CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
          DeleteObject rgn1
        End If
    Case ButtonTypes.[Windows XP], ButtonTypes.Mac 'Windows XP and Mac
        rgn1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
    Case ButtonTypes.[Oval Flat]
        DeleteObject rgnNorm
        rgnNorm = CreateEllipticRgn(0, 0, Wi, He)
End Select

DeleteObject rgn2
End Sub

Private Sub SetAccessKeys()
'this is a TRUE access keys parser
'the basic rule is that if an ampersand is followed by another,
'  a single ampersand is drawn and this is not the access key.
'  So we continue searching for another possible access key.
'   I only do a second pass because no one writes text like "Me & them & everyone"
'   so the caption prop should be "Me && them && &everyone", this is rubbish and a
'   search like this would only waste time
Dim ampersandPos As Long

'we first Clear the AccessKeys property, and will be filled if one is found
UserControl.AccessKeys = ""

If Len(elTex) > 1 Then
  ampersandPos = InStr(1, elTex, "&", vbTextCompare)
  If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
    If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
      UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
    Else 'do only a second pass to find another ampersand character
      ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
      If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
        UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
      End If
    End If
  End If
End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
'this function will add or remove a certain color
'quantity and return the result
Dim Red As Long, Blue As Long, Green As Long

'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
If blnSoft Then
  Value = Value \ 2
End If
If Not isXP Then 'for XP button i use a work-aroud that works fine
  Blue = ((Color \ &H10000) Mod &H100) + Value
Else
  Blue = ((Color \ &H10000) Mod &H100)
  Blue = Blue + ((Blue * Value) \ &HC0)
End If
Green = ((Color \ &H100) Mod &H100) + Value
Red = (Color And &HFF) + Value
'a bit of optimization done here, values will overflow a
' byte only in one direction... eg: if we added 32 to our
' color, then only a > 255 overflow can occurr.
If Value > 0 Then
  If Red > 255 Then
    Red = 255
  End If
  If Green > 255 Then
    Green = 255
  End If
  If Blue > 255 Then
    Blue = 255
  End If
ElseIf Value < 0 Then
  If Red < 0 Then
    Red = 0
  End If
  If Green < 0 Then
    Green = 0
  End If
  If Blue < 0 Then
    Blue = 0
  End If
End If
'more optimization by replacing the RGB function by its correspondent calculation
ShiftColor = Red + 256& * Green + 65536 * Blue
End Function

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
Dim Red As Long, Blue As Long, Green As Long
Dim Delta As Long

  Blue = ((theColor \ &H10000) Mod &H100)
  Green = ((theColor \ &H100) Mod &H100)
  Red = (theColor And &HFF)
  Delta = &HFF - Base
  Blue = Base + Blue * Delta \ &HFF
  Green = Base + Green * Delta \ &HFF
  Red = Base + Red * Delta \ &HFF
  If Red > 255 Then
    Red = 255
  End If
  If Green > 255 Then
    Green = 255
  End If
  If Blue > 255 Then
    Blue = 255
  End If
  ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

Private Sub CalcTextRects()
'this sub will calculate the rects required to draw the text
  Select Case PicPosition
    Case PicPositions.cbLeft
      rc2.Left = 1 + picSZ.x
      rc2.Right = Wi - 2
      rc2.Top = 1
      rc2.Bottom = He - 2
    Case PicPositions.cbRight
      rc2.Left = 1
      rc2.Right = Wi - 2 - picSZ.x
      rc2.Top = 1
      rc2.Bottom = He - 2
    Case PicPositions.cbTop
      rc2.Left = 1
      rc2.Right = Wi - 2
      rc2.Top = 1 + picSZ.y
      rc2.Bottom = He - 2
    Case PicPositions.cbBottom
      rc2.Left = 1
      rc2.Right = Wi - 2
      rc2.Top = 1
      rc2.Bottom = He - 2 - picSZ.y
    Case PicPositions.cbBackground
      rc2.Left = 1
      rc2.Right = Wi - 2
      rc2.Top = 1
      rc2.Bottom = He - 2
  End Select
  DrawText UserControl.hdc, elTex, Len(elTex), rc2, DT_CALCRECT Or DT_WORDBREAK
  CopyRect rc, rc2: fc.x = rc.Right - rc.Left: fc.y = rc.Bottom - rc.Top
  Select Case PicPosition
      Case PicPositions.cbLeft, PicPositions.cbTop
        OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
      Case PicPositions.cbRight
        OffsetRect rc, (Wi - rc.Right - picSZ.x - 4) \ 2, (He - rc.Bottom) \ 2
      Case PicPositions.cbBottom
        OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom - picSZ.y - 4) \ 2
      Case PicPositions.cbBackground
        OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
  End Select
  CopyRect rc2, rc: OffsetRect rc2, 1, 1
  CalcPicPos 'once we have the text position we are able to calculate the pic position
End Sub

Public Sub DisableRefresh()
'this is for fast button editing, once you disable the refresh,
' you can change every prop without triggering the drawing methods.
' once you are done, you call Refresh.
  blnIsShown = False
End Sub

Public Sub Refresh()
  If MyButtonType = ButtonTypes.Transparent Then
    GetParentPic
  End If
  SetColors
  CalcTextRects
  blnIsShown = True
  Redraw lastStat, True
End Sub

Private Function ConvertFromSystemColor(ByVal theColor As Long) As Long
  OleTranslateColor theColor, 0, ConvertFromSystemColor
End Function

Private Sub DrawCaption(ByVal State As Byte)
'this code is commonly shared through all the buttons so
' i took it and put it toghether here for easier readability
' of the code, and to cut-down disk size.
  captOpt = State
  With UserControl
    Select Case State 'in this select case, we only change the text color and draw only text that needs rc2, at the end, text that uses rc will be drawn
        Case 0 'normal caption
            txtFX rc
            SetTextColor .hdc, cText
        Case 1 'hover caption
            txtFX rc
            SetTextColor .hdc, cTextO
        Case 2 'down caption
            txtFX rc2
            If MyButtonType = Mac Then
              SetTextColor .hdc, cLight
            Else
              SetTextColor .hdc, cTextO
            End If
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
        Case 3 'disabled embossed caption
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cShadow
        Case 4 'disabled grey caption
            SetTextColor .hdc, cShadow
        Case 5 'WinXP disabled caption
            SetTextColor .hdc, ShiftColor(XPFace, -&H68, True)
        Case 6 'KDE 2 disabled
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cFace
        Case 7 'KDE 2 down
            SetTextColor .hdc, ShiftColor(cShadow, -&H32)
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hdc, cHighLight
    End Select
    'we now draw the text that is common in all the captions
    If State <> 2 Then
      DrawText .hdc, elTex, Len(elTex), rc, DT_CENTER
    End If
  End With
End Sub

Private Sub DrawPictures(ByVal State As Byte)
If picNormal Is Nothing Then
  Exit Sub 'check if there is a main picture, if not then exit
End If
With UserControl
Select Case State
    Case 0 'normal & hover
        If Not blnIsOver Then
          DoFX 0, picNormal
          TransBlt .hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picNormal, cMask, , , useGrey, (MyButtonType = [Office XP])
        Else
            If MyButtonType = [Office XP] Then
                DoFX -1, picNormal
                TransBlt .hdc, picPT.x + 1, picPT.y + 1, picSZ.x, picSZ.y, picNormal, cMask, cShadow
                TransBlt .hdc, picPT.x - 1, picPT.y - 1, picSZ.x, picSZ.y, picNormal, cMask
            Else
                If Not picHover Is Nothing Then
                    DoFX 0, picHover
                    TransBlt .hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picHover, cMask
                Else
                    DoFX 0, picNormal
                    TransBlt .hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picNormal, cMask
                End If
            End If
        End If
    Case 1 'down
        If picHover Is Nothing Or MyButtonType = [Office XP] Then
            Select Case MyButtonType
            Case 5, 9
                DoFX 0, picNormal
                TransBlt .hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picNormal, cMask
            Case Else
                DoFX 1, picNormal
                TransBlt .hdc, picPT.x + 1, picPT.y + 1, picSZ.x, picSZ.y, picNormal, cMask
            End Select
        Else
            TransBlt .hdc, picPT.x + Abs(MyButtonType <> [Java metal]), picPT.y + Abs(MyButtonType <> [Java metal]), picSZ.x, picSZ.y, picHover, cMask
        End If
    Case 2 'disabled
        Select Case MyButtonType
        Case ButtonTypes.[Java metal], ButtonTypes.[Netscape 6], ButtonTypes.[Office XP] 'draw flat grey pictures
            TransBlt .hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picNormal, cMask, Abs(MyButtonType = [Office XP]) * ShiftColor(cShadow, &HD) + Abs(MyButtonType <> [Office XP]) * cShadow, True
        Case ButtonTypes.[Windows XP]          'for WinXP draw a greyscaled image
            TransBlt .hdc, picPT.x + 1, picPT.y + 1, picSZ.x, picSZ.y, picNormal, cMask, , , True
        Case Else       'draw classic embossed pictures
            TransBlt .hdc, picPT.x + 1, picPT.y + 1, picSZ.x, picSZ.y, picNormal, cMask, cHighLight, True
            TransBlt .hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picNormal, cMask, cShadow, True
        End Select
End Select
End With
If PicPosition = cbBackground Then
  DrawCaption captOpt
End If
End Sub

Private Sub DoFX(ByVal offset As Long, ByVal thePic As StdPicture)
  If SFX > cbNone Then
    Dim curFace As Long
    If MyButtonType = [Windows XP] Then
      curFace = XPFace
    Else
      If offset = -1 And MyColorType <> Custom Then
        curFace = OXPf
      Else
        curFace = cFace
      End If
    End If
    TransBlt UserControl.hdc, picPT.x + 1 + offset, picPT.y + 1 + offset, picSZ.x, picSZ.y, thePic, cMask, ShiftColor(curFace, Abs(SFX = cbEngraved) * FXDEPTH + (SFX <> cbEngraved) * FXDEPTH)
    If SFX < cbShadowed Then
      TransBlt UserControl.hdc, picPT.x - 1 + offset, picPT.y - 1 + offset, picSZ.x, picSZ.y, thePic, cMask, ShiftColor(curFace, Abs(SFX <> cbEngraved) * FXDEPTH + (SFX = cbEngraved) * FXDEPTH)
    End If
  End If
End Sub

Private Sub txtFX(ByRef theRect As RECT)
Dim curFace As Long
Dim tempR As RECT

  If SFX > cbNone Then
    With UserControl
      CopyRect tempR, theRect
      OffsetRect tempR, 1, 1
      Select Case MyButtonType
          Case ButtonTypes.[Windows 16-bit], ButtonTypes.Mac, ButtonTypes.[KDE 2]
            curFace = XPFace
          Case Else
            If lastStat = 0 And blnIsOver And MyColorType <> Custom And MyButtonType = [Office XP] Then
              curFace = OXPf
            Else
              curFace = cFace
            End If
      End Select
      SetTextColor .hdc, ShiftColor(curFace, Abs(SFX = cbEngraved) * FXDEPTH + (SFX <> cbEngraved) * FXDEPTH)
      DrawText .hdc, elTex, Len(elTex), tempR, DT_CENTER
      If SFX < cbShadowed Then
        OffsetRect tempR, -2, -2
        SetTextColor .hdc, ShiftColor(curFace, Abs(SFX <> cbEngraved) * FXDEPTH + (SFX = cbEngraved) * FXDEPTH)
        DrawText .hdc, elTex, Len(elTex), tempR, DT_CENTER
      End If
    End With
  End If
End Sub

Private Sub CalcPicSize()
  If Not picNormal Is Nothing Then
    picSZ.x = UserControl.ScaleX(picNormal.Width, 8, UserControl.ScaleMode)
    picSZ.y = UserControl.ScaleY(picNormal.Height, 8, UserControl.ScaleMode)
  Else
    picSZ.x = 0
    picSZ.y = 0
  End If
End Sub

Private Sub CalcPicPos()
'exit if there's no picture
If picNormal Is Nothing And picHover Is Nothing Then
  Exit Sub
End If

  If (Trim(elTex) <> "") And (PicPosition <> 4) Then 'if there is no caption, or we have the picture as background, then we put the picture at the center of the button
    Select Case PicPosition
      Case PicPositions.cbLeft 'left
        picPT.x = rc.Left - picSZ.x - 4
        picPT.y = (He - picSZ.y) \ 2
      Case PicPositions.cbRight 'right
        picPT.x = rc.Right + 4
        picPT.y = (He - picSZ.y) \ 2
      Case PicPositions.cbTop 'top
        picPT.x = (Wi - picSZ.x) \ 2
        picPT.y = rc.Top - picSZ.y - 2
      Case PicPositions.cbBottom 'bottom
        picPT.x = (Wi - picSZ.x) \ 2
        picPT.y = rc.Bottom + 2
    End Select
  Else 'center the picture
    picPT.x = (Wi - picSZ.x) \ 2
    picPT.y = (He - picSZ.y) \ 2
  End If
End Sub

Private Sub TransBlt(ByVal dstDC As Long, ByVal DstX As Long, ByVal DstY As Long, _
                     ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, _
                     Optional ByVal TransColor As Long = -1, _
                     Optional ByVal BrushColor As Long = -1, _
                     Optional ByVal MonoMask As Boolean = False, _
                     Optional ByVal isGreyscale As Boolean = False, _
                     Optional ByVal XPBlend As Boolean = False)
    If DstW = 0 Or DstH = 0 Then
      Exit Sub
    End If
        
    Dim B As Long, H As Long, F As Long, i As Long, newW As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
    
    Dim SrcDC As Long, tObj As Long, ttt As Long

    SrcDC = CreateCompatibleDC(hdc)
    
    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
    
    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else
        Dim br As RECT, hBrush As Long: br.Right = DstW: br.Bottom = DstH
        ttt = CreateCompatibleBitmap(dstDC, DstW, DstH): tObj = SelectObject(SrcDC, ttt)
        hBrush = CreateSolidBrush(MaskColor): FillRect SrcDC, br, hBrush
        DeleteObject hBrush
        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, 0, 0, 0, 0, &H1 Or &H2
    End If
 
    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(dstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(dstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With
    
    BitBlt TmpDC, 0, 0, DstW, DstH, dstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0
    
    If BrushColor > 0 Then
        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
        BrushRGB.rgbRed = BrushColor And &HFF
    End If
    
    If Not useMask Then TransColor = -1
    
    newW = DstW - 1
    
    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To newW
            i = F + B
            If (CLng(Data2(i).rgbRed) + 256& * Data2(i).rgbGreen + 65536 * Data2(i).rgbBlue) <> TransColor Then
            With Data1(i)
                If BrushColor > -1 Then
                    If MonoMask Then
                        If (CLng(Data2(i).rgbRed) + Data2(i).rgbGreen + Data2(i).rgbBlue) <= 384 Then Data1(i) = BrushRGB
                    Else
                        Data1(i) = BrushRGB
                    End If
                Else
                    If isGreyscale Then
                        gCol = CLng(Data2(i).rgbRed * 0.3) + Data2(i).rgbGreen * 0.59 + Data2(i).rgbBlue * 0.11
                        .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                    Else
                        If XPBlend Then
                            .rgbRed = (CLng(.rgbRed) + Data2(i).rgbRed * 2) \ 3
                            .rgbGreen = (CLng(.rgbGreen) + Data2(i).rgbGreen * 2) \ 3
                            .rgbBlue = (CLng(.rgbBlue) + Data2(i).rgbBlue * 2) \ 3
                        Else
                            Data1(i) = Data2(i)
                        End If
                    End If
                End If
            End With
            End If
        Next
    Next

    SetDIBitsToDevice dstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = 3 Then
      DeleteObject SelectObject(SrcDC, tObj)
    End If
    DeleteDC TmpDC
    DeleteDC Sr2DC
    DeleteObject tObj
    DeleteObject ttt
    DeleteDC SrcDC
End Sub

Private Function IsMouseOver() As Boolean
Dim pt As POINTAPI

  GetCursorPos pt
  IsMouseOver = (WindowFromPoint(pt.x, pt.y) = hwnd)
End Function

Private Sub GetParentPic()
  On Local Error Resume Next
    blnInLoop = True
    UserControl.Height = 0
    DoEvents
    BitBlt pDC, 0, 0, Wi, He, GetDC(GetParent(hwnd)), Extender.Left, Extender.Top, vbSrcCopy
    UserControl.Height = ScaleY(He, vbPixels, vbTwips)
    blnInLoop = False
End Sub
