VERSION 5.00
Begin VB.UserControl SpliterHorizontal 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   ScaleHeight     =   4305
   ScaleWidth      =   6420
   ToolboxBitmap   =   "SpliterHorizontal.ctx":0000
   Begin VB.PictureBox picSpliter 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   450
      MouseIcon       =   "SpliterHorizontal.ctx":0182
      MousePointer    =   99  'Custom
      ScaleHeight     =   135
      ScaleWidth      =   5685
      TabIndex        =   1
      Top             =   2550
      Width           =   5685
   End
   Begin VB.PictureBox picSpliterHide 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   5  'Transparent
      Enabled         =   0   'False
      FillColor       =   &H00C0C0C0&
      Height          =   165
      Left            =   495
      ScaleHeight     =   165
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   3450
      Visible         =   0   'False
      Width           =   5355
   End
End
Attribute VB_Name = "SpliterHorizontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control de usuario para el manejo de las barras que permiten establecer el tamaño de determinados elementos de la ventana (Spliter Horizontal)
Option Explicit

Private Const cnstColBackColorBar As Long = &HFFC0C0

Public Event Resize(ByVal SpliterTop As Integer) 'Evento lanzado al formulario padre para que redimensione los controles

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
    .Left = 0
    .Width = ScaleWidth
  End With
  With picSpliterHide
    .Left = 0
    .Width = ScaleWidth
  End With
End Sub

Property Let SpliterTop(ByVal intTop As Long)
'--> Consigue la posición inicial de la barra de Spliter
  picSpliter.Top = intTop
End Property

Property Get SpliterTop() As Long
'--> Consigue la posición actual de la barra de Spliter
  SpliterTop = picSpliter.Top
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

Property Let SpliterPictureHeight(ByVal intHeight As Integer)
'--> Cambia el ancho del spliter
  If intHeight < 120 Then intHeight = 120
  If intHeight > 500 Then intHeight = 500
  picSpliter.Height = intHeight
End Property

Property Get SpliterPictureHeight() As Integer
'--> Obtiene el ancho del spliter
  SpliterPictureHeight = picSpliter.Height
End Property

Private Sub picSpliter_DblClick()
'--> Al hacer doble click sobre el spliter debe cambiarse arriba o abajo del todo
  If SpliterTop + 60 >= Minimo And SpliterTop - 60 <= Minimo Then
    SpliterTop = Height - Minimo
  Else
    SpliterTop = Minimo
  End If
  RaiseEvent Resize(SpliterTop)
End Sub

Private Sub picSpliter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Si se pulsa el botón sobre el Spliter comienza a redimensionar
  blnResizing = (Button = vbLeftButton)
  If blnResizing Then
    With picSpliter
      picSpliterHide.Move .Left, .Top, .Width, .Height - 40
    End With
    picSpliterHide.Visible = True
    picSpliterHide.ZOrder ZOrderConstants.vbBringToFront
  End If
End Sub

Private Sub picSpliter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Si se mueve el Spliter mientras está pulsado se realiza el movimiento
Dim sngPosicion As Single, sngPosicionAnterior As Single

  If blnResizing And Button = vbLeftButton Then
    sngPosicionAnterior = picSpliterHide.Top
    sngPosicion = y + picSpliter.Top
    If sngPosicion < intSpliterMinimo Then
      picSpliterHide.Top = intSpliterMinimo
    ElseIf sngPosicion > Height - intSpliterMinimo Then
      picSpliterHide.Top = Height - intSpliterMinimo
    Else
      picSpliterHide.Top = sngPosicion
    End If
    If sngPosicionAnterior <> picSpliterHide.Top Then
      'Parent.Refresh
      'Refresh
    End If
  Else
    blnResizing = False
  End If
End Sub

Private Sub picSpliter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'--> Cuando se suelta el ratón se lanza un evento <B> Redimensionar </B> al propietario del control
  If blnResizing Then
    SpliterTop = picSpliterHide.Top
    blnResizing = False
    picSpliterHide.Visible = False
    RaiseEvent Resize(SpliterTop)
  End If
End Sub

Private Sub UserControl_Initialize()
  blnResizing = False
  SpliterTop = Height / 2
End Sub

Private Sub UserControl_InitProperties()
  Minimo = 1500
  BorderStyle = SplBorderNone
  SpliterPictureHeight = 120
  BarBackColor = cnstColBackColorBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Minimo = PropBag.ReadProperty("Minimo", 1500)
  SpliterPictureHeight = PropBag.ReadProperty("SpliterPictureHeight", 120)
  BorderStyle = PropBag.ReadProperty("BorderStyle", SplBorderNone)
  BarBackColor = PropBag.ReadProperty("BarBackColor", cnstColBackColorBar)
End Sub

Private Sub UserControl_Show()
  Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Minimo", intSpliterMinimo, 1500
  PropBag.WriteProperty "SpliterPictureHeight", SpliterPictureHeight, 120
  PropBag.WriteProperty "BorderStyle", BorderStyle, SplBorderNone
  PropBag.WriteProperty "BarBackColor", BarBackColor, cnstColBackColorBar
End Sub
