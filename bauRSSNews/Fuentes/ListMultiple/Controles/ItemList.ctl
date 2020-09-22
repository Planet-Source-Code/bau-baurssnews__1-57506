VERSION 5.00
Begin VB.UserControl ItemList 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   ScaleHeight     =   1425
   ScaleWidth      =   6555
   Begin VB.Label lblSubCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subcaption"
      Height          =   195
      Left            =   4260
      TabIndex        =   2
      Top             =   90
      Width           =   810
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   6345
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   225
      Left            =   630
      TabIndex        =   0
      Top             =   150
      Width           =   4695
   End
   Begin VB.Image imgChecked 
      Height          =   285
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   315
   End
   Begin VB.Image imgMain 
      Height          =   285
      Left            =   60
      Stretch         =   -1  'True
      Top             =   60
      Width           =   315
   End
End
Attribute VB_Name = "ItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Elemento de un ListMultiple
Option Explicit

'Eventos
Public Event Check(ByVal strKey As String)
Public Event DblClick(ByVal strKey As String)
Public Event KeyDown(ByVal strKey As String, KeyCode As Integer, Shift As Integer)
Public Event MouseDown(ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)

'Variables privadas con los valores de las propiedades
Private strKey As String
Private strTag As String
Private strCaption As String
Private strSubCaption As String
Private strDescription As String
Private dtmDate As Date
Private blnChecked As Boolean
Private blnSelected As Boolean
Private picMain As StdPicture

'Variables privadas
Private blnIsOver As Boolean

Private Sub Init()
'--> Inicializa las variables del control
End Sub

Public Sub Redraw()
'--> Redibuja el control
  'Cambia el color de fondo
    If blnSelected Then
      BackColor = modListMultiple.colBackColorSelected
    ElseIf blnIsOver Then
      BackColor = modListMultiple.colBackColorOver
    Else
      BackColor = modListMultiple.colBackColor
    End If
  'Cambia la imagen principal
    Set imgMain.Picture = picMain
  'Cambia la imagen seleccionada
    If blnChecked Then
      Set imgChecked.Picture = modListMultiple.picChecked
    Else
      Set imgChecked.Picture = modListMultiple.picNoChecked
    End If
  'Cambia el título
    With lblCaption
      .Caption = strCaption
      If Not modListMultiple.fntList Is Nothing Then
        Set .Font = modListMultiple.fntList
      End If
      If blnSelected Then
        .ForeColor = vbWhite ' modListMultiple.colForeColorCaptionSelected
      ElseIf blnIsOver Then
        .ForeColor = vbBlue ' modListMultiple.colForeColorCaptionOver
      Else
        .ForeColor = vbBlue ' modListMultiple.colForeColorCaption
      End If
    End With
    With lblSubCaption
      .Caption = strSubCaption
      If Not modListMultiple.fntList Is Nothing Then
        Set .Font = modListMultiple.fntList
      End If
      If blnSelected Then
        .ForeColor = vbWhite ' modListMultiple.colForeColorCaptionSelected
      ElseIf blnIsOver Then
        .ForeColor = vbRed ' modListMultiple.colForeColorCaptionOver
      Else
        .ForeColor = vbRed ' modListMultiple.colForeColorCaption
      End If
      .Font.Bold = False
    End With
  'Cambia la descripción
    With lblDescription
      .Caption = strDescription
      If Not modListMultiple.fntList Is Nothing Then
        Set .Font = modListMultiple.fntList
      End If
      If blnSelected Then
        .ForeColor = modListMultiple.colForeColorSelected
      ElseIf blnIsOver Then
        .ForeColor = modListMultiple.colForeColorOver
      Else
        .ForeColor = modListMultiple.colForeColor
      End If
    End With
End Sub

Private Sub Resize()
'--> Cambia el tamaño y posición de los controles
  On Error Resume Next
    'Imagen principal
      With imgMain
        .Top = ScaleTop + 50
        .Left = ScaleLeft + 50
      End With
    'Imagen checked
      With imgChecked
        .Top = imgMain.Top
        .Left = ScaleWidth - .Width - 50
      End With
    'Label de título
      With lblCaption
        .Top = imgMain.Top + (imgMain.Height - .Height) / 2
        .Left = imgMain.Left + imgMain.Width + 50
        .Width = ScaleWidth - imgMain.Width - imgChecked.Width - lblSubCaption.Width - 100
      End With
    'Label del subtítulo
      With lblSubCaption
        .Top = lblCaption.Top
        .Left = ScaleWidth - .Width - imgChecked.Width - 100
      End With
    'Label de descripción
      With lblDescription
        .Top = imgMain.Top + imgMain.Height
        .Left = lblCaption.Left
        .Width = ScaleWidth - .Left - 60
        .Height = ScaleHeight - .Top - 30
      End With
End Sub

'------------------------------------------------------------------------------------------
'---------- Propiedades
'------------------------------------------------------------------------------------------
Public Property Get Key() As String
  Key = strKey
End Property

Public Property Let Key(ByVal strNewKey As String)
  strKey = strNewKey
  PropertyChanged
End Property

Public Property Get Caption() As String
  Caption = strCaption
End Property

Public Property Let Caption(ByVal strNewCaption As String)
  strCaption = strNewCaption
  PropertyChanged
End Property

Public Property Get SubCaption() As String
  SubCaption = strSubCaption
End Property

Public Property Let SubCaption(ByVal strNewSubCaption As String)
  strSubCaption = strNewSubCaption
  PropertyChanged
End Property

Public Property Get Description() As String
  Description = strDescription
End Property

Public Property Let Description(ByVal strNewDescription As String)
  strDescription = strNewDescription
  PropertyChanged
End Property

Public Property Get Tag() As String
  Tag = strTag
End Property

Public Property Let Tag(ByVal strNewTag As String)
  strTag = strNewTag
  PropertyChanged
End Property

Public Property Get Checked() As Boolean
  Checked = blnChecked
End Property

Public Property Let Checked(ByVal blnNewChecked As Boolean)
  blnChecked = blnNewChecked
  Redraw
  PropertyChanged
End Property

Public Property Get Selected() As Boolean
  Selected = blnSelected
End Property

Public Property Let Selected(ByVal blnNewSelected As Boolean)
  blnSelected = blnNewSelected
  PropertyChanged
End Property

Public Property Get IconMain() As StdPicture
  Set IconMain = picMain
End Property

Public Property Set IconMain(ByVal picNewMain As StdPicture)
  Set picMain = picNewMain
  Redraw
End Property

Public Property Get FontCaption() As StdFont
  Set FontCaption = lblCaption.Font
End Property

Public Property Set FontCaption(ByRef fntNewFontCaption As StdFont)
  Set lblCaption.Font = fntNewFontCaption
End Property

Public Property Get ForeColorCaption() As OLE_COLOR
  ForeColorCaption = lblCaption.ForeColor
End Property

Public Property Let ForeColorCaption(ByVal colNewForeColor As OLE_COLOR)
  lblCaption.ForeColor = colNewForeColor
End Property

Public Property Let IsOver(ByVal blnNewOver As Boolean)
  blnIsOver = blnNewOver
  Redraw
End Property

Public Property Get hwnd()
  hwnd = UserControl.hwnd
End Property

Private Sub imgChecked_DblClick()
  RaiseEvent DblClick(Key)
End Sub

Private Sub imgChecked_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Checked = Not Checked
  Redraw
  Refresh
  RaiseEvent Check(Key)
End Sub

Private Sub imgChecked_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Key, Button, Shift, x, y)
End Sub

Private Sub imgMain_DblClick()
  RaiseEvent DblClick(Key)
End Sub

Private Sub imgMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Key, Button, Shift, x, y)
End Sub

Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Key, Button, Shift, x, y)
End Sub

Private Sub imgMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Key, Button, Shift, x, y)
End Sub

Private Sub lblCaption_DblClick()
  RaiseEvent DblClick(Key)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Key, Button, Shift, x, y)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Key, Button, Shift, x, y)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Key, Button, Shift, x, y)
End Sub

Private Sub lblDescription_DblClick()
  RaiseEvent DblClick(Key)
End Sub

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Key, Button, Shift, x, y)
End Sub

Private Sub lblDescription_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Key, Button, Shift, x, y)
End Sub

Private Sub lblDescription_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Key, Button, Shift, x, y)
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick(Key)
End Sub

Private Sub UserControl_Initialize()
  Init
End Sub

Private Sub UserControl_InitProperties()
  strKey = ""
  strCaption = Name
  strSubCaption = ""
  blnChecked = False
  blnSelected = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(Key, KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Key, Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Key, Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Key, Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
  Resize
End Sub
