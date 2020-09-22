VERSION 5.00
Begin VB.UserControl SpliterVertical 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   ScaleHeight     =   4365
   ScaleWidth      =   6420
   ToolboxBitmap   =   "SpliterVertical.ctx":0000
   Begin VB.PictureBox picSpliter 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   4155
      Left            =   1320
      MouseIcon       =   "SpliterVertical.ctx":0182
      MousePointer    =   99  'Custom
      ScaleHeight     =   4155
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   72
      Width           =   105
   End
   Begin VB.PictureBox picSpliterHide 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      Enabled         =   0   'False
      FillColor       =   &H00C0C0C0&
      Height          =   3396
      Left            =   2940
      ScaleHeight     =   3390
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "SpliterVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control de usuario para el manejo de spliter vertical que permiten establecer el tamaño de determinados elementos de la ventana
Option Explicit

Private Const cnstColBackColorBar As Long = &HFFC0C0

Public Enum SpliterBorder 'Constantes con el borde del spliter
  SplBorderNone = 0 'Sin borde
  SplBorderSolid = 1 'Con borde sólido
End Enum

Public Event Resize(ByVal SpliterLeft As Integer) 'Evento lanzado al formulario padre para que redimensione los controles

Private blnResizing As Boolean 'Variable que indica que el usuario está moviendo el spliter
Private intSpliterMinimo As Integer 'Mínima posición a la izquierda del spliter

Public Sub Resize(Optional ByVal intWidth As Long = -1, Optional ByVal intHeight As Long = -1)
'--> Redimensiona el tamaño del control
  'Tener en cuenta que si está dentro de un control de usuario no tiene parent
  On Error Resume Next
  If intHeight = -1 Then
    Height = Parent.ScaleHeight
  Else
    Height = intHeight
  End If
  If intWidth = -1 Then
    Width = Parent.ScaleWidth
  Else
    Width = intWidth
  End If
  With picSpliter
    .Top = 0
    .Height = ScaleHeight
  End With
  With picSpliterHide
    .Top = 0
    .Height = ScaleHeight
  End With
End Sub

Property Let SpliterLeft(ByVal lngLeft As Long)
'--> Consigue la posición inicial de la barra de Spliter
  picSpliter.Left = lngLeft
End Property

Property Get SpliterLeft() As Long
'--> Consigue la posición actual de la barra de Spliter
  SpliterLeft = picSpliter.Left
End Property

Property Let Minimo(ByVal SplMinimo As Integer)
'--> Asigna la posición mínima de la barra de Spliter
  intSpliterMinimo = SplMinimo
  PropertyChanged
End Property

Property Get Minimo() As Integer
'--> Devuelve la posición mínima de la barra de Spliter
  Minimo = intSpliterMinimo
End Property

Property Let BorderStyle(ByVal uBorderStyle As SpliterBorder)
'--> Cambia el estilo del borde
  picSpliter.BorderStyle = uBorderStyle
End Property

Property Get BorderStyle() As SpliterBorder
'--> Devuelve el estilo del borde
  BorderStyle = picSpliter.BorderStyle
End Property

Property Get BarBackColor() As OLE_COLOR
  BarBackColor = picSpliter.BackColor
End Property

Property Let BarBackColor(ByVal colNewBackColor As OLE_COLOR)
  picSpliter.BackColor = colNewBackColor
End Property

Property Let SpliterPictureWidth(ByVal intWidth As Integer)
'--> Cambia el ancho del spliter
  If intWidth < 120 Then intWidth = 120
  If intWidth > 500 Then intWidth = 500
  picSpliter.Width = intWidth
End Property

Property Get SpliterPictureWidth() As Integer
'--> Obtiene el ancho del spliter
  SpliterPictureWidth = picSpliter.Width
End Property

Private Sub picSpliter_DblClick()
'--> Al pulsar dos veces sobre el spliter debe ponerse a la izquierda o a la derecha del todo
  If SpliterLeft + 60 >= Minimo And SpliterLeft - 60 <= Minimo Then
    SpliterLeft = Width - Minimo
  Else
    SpliterLeft = Minimo
  End If
  RaiseEvent Resize(SpliterLeft)
End Sub

Private Sub picSpliter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Si se pulsa el botón sobre el Spliter comienza a redimensionar
  blnResizing = (Button = vbLeftButton)
  If blnResizing Then
    With picSpliter
      picSpliterHide.Move .Left, .Top + 20, .Width / 2, .Height - 40
    End With
    picSpliterHide.Visible = True
    picSpliterHide.ZOrder ZOrderConstants.vbBringToFront
  End If
End Sub

Private Sub picSpliter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Si se mueve el Spliter mientras está pulsado se realiza el movimiento
Dim sngPosicion As Single, sngPosicionAnterior As Single

  If blnResizing And Button = vbLeftButton Then
    sngPosicionAnterior = picSpliterHide.Left
    sngPosicion = x + picSpliter.Left
    If sngPosicion < intSpliterMinimo Then
      picSpliterHide.Left = intSpliterMinimo
    ElseIf sngPosicion > Width - intSpliterMinimo Then
      picSpliterHide.Left = Width - intSpliterMinimo
    Else
      picSpliterHide.Left = sngPosicion
    End If
    If sngPosicionAnterior <> picSpliterHide.Left Then
      On Error Resume Next
      Parent.Refresh
    End If
  Else
    blnResizing = False
  End If
End Sub

Private Sub picSpliter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Cuando se suelta el ratón se lanza un evento <B> Redimensionar </B> al propietario del control
  If blnResizing Then
    picSpliter.Left = picSpliterHide.Left
    blnResizing = False
    picSpliterHide.Visible = False
    RaiseEvent Resize(picSpliterHide.Left)
  End If
End Sub

Private Sub UserControl_Initialize()
  blnResizing = False
  SpliterLeft = Width / 2
End Sub

Private Sub UserControl_InitProperties()
  Minimo = 1500
  BorderStyle = SplBorderNone
  SpliterPictureWidth = 120
  BarBackColor = cnstColBackColorBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Minimo = PropBag.ReadProperty("Minimo", 1500)
  SpliterPictureWidth = PropBag.ReadProperty("SpliterPictureWidth", 120)
  BorderStyle = PropBag.ReadProperty("BorderStyle", SplBorderNone)
  BarBackColor = PropBag.ReadProperty("BarBackColor", cnstColBackColorBar)
End Sub

Private Sub UserControl_Show()
  Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Minimo", intSpliterMinimo, 1500
  PropBag.WriteProperty "SpliterPictureWidth", SpliterPictureWidth, 120
  PropBag.WriteProperty "BorderStyle", BorderStyle, SplBorderNone
  PropBag.WriteProperty "BarBackColor", BarBackColor, cnstColBackColorBar
End Sub
