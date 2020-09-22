VERSION 5.00
Begin VB.UserControl ListMultiple 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ScaleHeight     =   3600
   ScaleWidth      =   5925
   Begin VB.VScrollBar scrVertical 
      Height          =   2955
      Left            =   5370
      TabIndex        =   1
      Top             =   210
      Width           =   255
   End
   Begin bauRSSNews.ItemList lsiItem 
      Height          =   825
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   1455
   End
End
Attribute VB_Name = "ListMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Lista de elementos
Option Explicit

'Declaraciones de API
Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'Constantes
Private Const cnstColBackColor As Long = &HFFFFFF
Private Const cnstColBackColorOver As Long = &HFFC0C0
Private Const cnstColBackColorSelected As Long = &HFF0000
Private Const cnstColForeColor As Long = 0
Private Const cnstColForeColorOver As Long = &HFF&
Private Const cnstColForeColorSelected As Long = &HFFFFFF
Private Const cnstColForeColorCaption As Long = 0
Private Const cnstColForeColorCaptionOver As Long = &HFF&
Private Const cnstColForeColorCaptionSelected As Long = &HFFFFFF

'Eventos
Public Event Check(ByVal strKey As String)
Public Event Click(ByVal strKey As String)
Public Event DblClick(ByVal strKey As String)
Public Event KeyDown(ByVal strKey As String, KeyCode As Integer, Shift As Integer)
Public Event MouseDown(ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)

'Variables privadas
Private blnEnabled As Boolean
Private blnFirst As Boolean
Private blnMultiSelect As Boolean

Private Sub Init()
'--> Inicializa los valores del control
  'Indica que está habilitado
    blnEnabled = True
  'Indica que no es una lista multiselección
    blnMultiSelect = False
  'Indica que es la primera vez
    blnFirst = True
  'Cambia el título y descripción del primer elemento
    lsiItem(0).Caption = Name
    lsiItem(0).Description = "Description " & Name
End Sub

Public Function Count() As Long
'--> Obtiene el número de elementos de la lista
  'Obtiene el número de elementos
    Count = lsiItem.UBound
  'Quita el primer elemento si no está visible
    If Not lsiItem(0).Visible Then
      Count = Count - 1
    End If
End Function

Private Function getIndexFromKey(ByVal strKey As String) As Long
'--> Obtiene el índice a partir de la clave
Dim lngIndex As Long

  'Supone que no encuentra el elemento
    getIndexFromKey = -1
  'Recorre los elementos
    For lngIndex = lsiItem.LBound To lsiItem.UBound
      If lsiItem(lngIndex).Key = strKey Then
        getIndexFromKey = lngIndex
      End If
    Next lngIndex
End Function

Private Sub moveMouse()
'--> Rutina que controla los movimientos de ratón
'Dim pt As POINTAPI
'Dim lngIndex As Long
'
'  'Libera las capturas anteriores
'    ReleaseCapture
'  'Obtiene la posición del cursor
'    GetCursorPos pt
'  'Comprueba si está sobre alguno de los elementos
'    For lngIndex = lsiItem.LBound To lsiItem.UBound
'      If WindowFromPoint(pt.x, pt.y) = lsiItem(lngIndex).hwnd Then
'        lsiItem(lngIndex).IsOver = True
'        SetCapture lsiItem(lngIndex).hwnd
'      Else
'        lsiItem(lngIndex).IsOver = False
'      End If
'    Next lngIndex
'  'Actualiza el control
'    Refresh
End Sub

Private Function getNextKey() As String
'--> Obtiene la siguiente clave
Dim blnFound As Boolean
Dim lngKey As Long, lngIndex As Long

  'Inicializa el contador de clave
    lngKey = lsiItem.Count
  'Recorre los controles comprobando si ya existe la clave
    Do
      'Indica que aún no se ha encontrado una clave válida
        blnFound = False
        getNextKey = "LSI" & lngKey
      'Comprueba si existe la clave
        For lngIndex = lsiItem.LBound To lsiItem.UBound
          If lsiItem(lngIndex).Key = getNextKey Then
            blnFound = True
          End If
        Next lngIndex
      'Incrementa el contador de clave
        lngKey = lngKey + 1
    Loop While blnFound
End Function

Public Function Add(ByVal strCaption As String, ByVal strSubCaption As String, _
                    ByVal strDescription As String, _
                    ByVal blnChecked As Boolean, _
                    Optional ByVal strTag As String = "", _
                    Optional ByVal picMain As StdPicture = 0, _
                    Optional ByVal strKey As String = "") As Boolean
'--> Añade un elemento a la lista
  On Error GoTo errorAdd
    'Supone que no se puede añadir el control
      Add = False
    'Carga otro nuevo control (si es la primera vez se hace sobre lsiItem(0))
      If Not blnFirst Then
        Load lsiItem(lsiItem.UBound + 1)
      End If
    'Modifica los parámetros del control
      With lsiItem(lsiItem.UBound)
        'Posición y visible
          .Left = ScaleLeft
          If blnFirst Then
            .Top = ScaleTop
          Else
            .Top = lsiItem(lsiItem.UBound - 1).Top + lsiItem(lsiItem.UBound - 1).Height
          End If
          .Width = ScaleWidth - .Left
          .Visible = True
        'Propiedades
          .Caption = strCaption
          .SubCaption = strSubCaption
          .Description = strDescription
          .Checked = blnChecked
          .Tag = strTag
          Set .IconMain = picMain
        'Clave
          If strKey = "" Then
            strKey = getNextKey()
          End If
          .Key = strKey
      End With
    'Ajusta la barra de scroll y redibuja
      adjustScrollBar
      Redraw
    'Indica que ya no es la primera vez que se añade un parámetro
      blnFirst = False
    'Indica que todo es correcto
      Add = True
  Exit Function
  
errorAdd:
End Function

Public Sub Clear()
'--> Limpia los elementos
Dim lngIndex As Long

  'Indica que está vacío
    blnFirst = True
  'Descarga los controles
    For lngIndex = lsiItem.UBound To 1 Step -1
      Unload lsiItem(lngIndex)
    Next lngIndex
  'Oculta el primer control
    lsiItem(0).Visible = False
  'Ajusta la barra de scroll
    adjustScrollBar
  'Actualiza
    Refresh
End Sub

Public Sub adjustScrollBar()
'--> Ajusta los parámetros de la barra de scroll
  On Error Resume Next
    'Scroll vertical
      With scrVertical
        .Value = 0
        .Min = 0
        .Max = lsiItem.UBound
        .SmallChange = 1
        .LargeChange = ScaleHeight \ lsiItem(0).Height
        .Visible = (.Max > .LargeChange)
        .Refresh
      End With
End Sub

Private Sub Redraw()
'--> Mueve los controles por el scroll y redibuja
Dim lngIndex As Long
Dim lngTop As Long

  'Cambia el color de fondo
    UserControl.BackColor = BackColor
  'Mueve los controles
    lngTop = ScaleTop
    For lngIndex = lsiItem.LBound To lsiItem.UBound
      With lsiItem(lngIndex)
        If lngIndex < scrVertical.Value Then
          .Top = -5000
        Else
          .Top = lngTop
          lngTop = lngTop + lsiItem(0).Height
          .Redraw
        End If
      End With
    Next lngIndex
  'Actualiza la visualización
    Refresh
End Sub

Private Sub selectItems(ByVal strKeySelected As String)
'--> Selecciona el elemento o elementos
Dim lngIndex As Long

  'Siempre y cuando esté habilitado
    If blnEnabled Then
      'Recorre los controles
        For lngIndex = lsiItem.LBound To lsiItem.UBound
          With lsiItem(lngIndex)
            If .Key = strKeySelected Then
              .Selected = Not .Selected
            ElseIf Not MultiSelect Then
              .Selected = False
            End If
          End With
        Next lngIndex
      'Redibuja
        Redraw
    End If
End Sub

Public Sub Resize()
'--> Redimensiona los controles
Dim intIndex As Integer
Dim lngTop As Long

  On Error Resume Next
    'Controles de parámetros
      lngTop = 100
      For intIndex = lsiItem.LBound To lsiItem.UBound
        If lsiItem(intIndex).Visible Then
          lsiItem(intIndex).Width = ScaleWidth - lsiItem(intIndex).Left - scrVertical.Width
          lsiItem(intIndex).Top = lngTop
          lngTop = lngTop + lsiItem(intIndex).Height
        End If
      Next intIndex
    'Ajusta el scroll
      adjustScrollBar
    'Scroll vertical
      With scrVertical
        .Top = ScaleTop
        .Left = ScaleWidth - .Width
        .Height = ScaleHeight - .Top
      End With
End Sub

Public Property Get Item(ByVal lngIndex As Long) As ItemList
  On Error GoTo errorItem
    Set Item = lsiItem(lngIndex)
  Exit Property
  
errorItem:
  Set Item = Nothing
End Property

Public Property Get IconChecked() As StdPicture
  Set IconChecked = modListMultiple.picChecked
End Property

Public Property Set IconChecked(ByVal picNewChecked As StdPicture)
  Set modListMultiple.picChecked = picNewChecked
  Redraw
  PropertyChanged
End Property

Public Property Get IconUnChecked() As StdPicture
  Set IconUnChecked = modListMultiple.picNoChecked
End Property

Public Property Set IconUnChecked(ByVal picNewNoChecked As StdPicture)
  Set modListMultiple.picNoChecked = picNewNoChecked
  Redraw
  PropertyChanged
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = modListMultiple.colBackColor
End Property

Public Property Let BackColor(ByVal colNewBackColor As OLE_COLOR)
  modListMultiple.colBackColor = colNewBackColor
  Redraw
  PropertyChanged
End Property

Public Property Get BackColorOver() As OLE_COLOR
  BackColorOver = modListMultiple.colBackColorOver
End Property

Public Property Let BackColorOver(ByVal colNewBackColorOver As OLE_COLOR)
  modListMultiple.colBackColorOver = colNewBackColorOver
  Redraw
  PropertyChanged
End Property

Public Property Get BackColorSelected() As OLE_COLOR
  BackColorSelected = modListMultiple.colBackColorSelected
End Property

Public Property Let BackColorSelected(ByVal colNewBackColorSelected As OLE_COLOR)
  modListMultiple.colBackColorSelected = colNewBackColorSelected
  Redraw
  PropertyChanged
End Property

Public Property Get Selected(ByVal strKey As String) As Boolean
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Selected = lsiItem(lngIndex).Selected
  End If
End Property

Public Property Let Selected(ByVal strKey As String, ByVal blnNewSelected As Boolean)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    lsiItem(lngIndex).Selected = blnNewSelected
  End If
End Property

Public Property Get SelectedIndex() As Long
Dim lngIndex As Long

  SelectedIndex = -1
  For lngIndex = lsiItem.LBound To lsiItem.UBound
    If lsiItem(lngIndex).Selected Then
      SelectedIndex = lngIndex
    End If
  Next lngIndex
End Property

Public Property Let SelectedIndex(ByVal lngNewSelected As Long)
Dim lngIndex As Long

  'Busca el elemento a seleccionar
    For lngIndex = lsiItem.LBound To lsiItem.UBound
      lsiItem(lngIndex).Selected = (lngIndex = lngNewSelected)
    Next lngIndex
  'Redibuja
    Redraw
End Property

Public Property Get Icon(ByVal strKey As String) As StdPicture
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Set Icon = lsiItem(lngIndex).IconMain
  End If
End Property

Public Property Set Icon(ByVal strKey As String, ByRef imgNewIcon As StdPicture)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Set lsiItem(lngIndex).IconMain = imgNewIcon
  End If
End Property

Public Property Get FontCaption(ByVal strKey As String) As StdFont
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Set FontCaption = lsiItem(lngIndex).FontCaption
  End If
End Property

Public Property Set FontCaption(ByVal strKey As String, ByRef fntNewFontCaption As StdFont)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Set lsiItem(lngIndex).FontCaption = fntNewFontCaption
  End If
End Property

Public Property Get ForeColorCaptionItem(ByVal strKey As String) As OLE_COLOR
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    ForeColorCaption = lsiItem(lngIndex).ForeColorCaption
  End If
End Property

Public Property Let ForeColorCaptionItem(ByVal strKey As String, ByVal colNewForeColor As OLE_COLOR)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    lsiItem(lngIndex).ForeColorCaption = colNewForeColor
  End If
End Property

Public Property Get Checked(ByVal strKey As String) As Boolean
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Checked = lsiItem(lngIndex).Checked
  End If
End Property

Public Property Let Checked(ByVal strKey As String, ByVal blnNewChecked As Boolean)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    lsiItem(lngIndex).Checked = blnNewChecked
  End If
End Property

Public Property Get Caption(ByVal strKey As String) As String
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Caption = lsiItem(lngIndex).Caption
  End If
End Property

Public Property Let Caption(ByVal strKey As String, ByVal strNewCaption As String)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    lsiItem(lngIndex).Caption = strNewCaption
  End If
End Property

Public Property Get Description(ByVal strKey As String) As String
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Description = lsiItem(lngIndex).Description
  End If
End Property

Public Property Let Description(ByVal strKey As String, ByVal strNewDescription As String)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    lsiItem(lngIndex).Description = strNewDescription
  End If
End Property

Public Property Get Tag(ByVal strKey As String) As String
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    Tag = lsiItem(lngIndex).Tag
  End If
End Property

Public Property Let Tag(ByVal strKey As String, ByVal strNewTag As String)
Dim lngIndex As Long

  lngIndex = getIndexFromKey(strKey)
  If lngIndex <> -1 Then
    lsiItem(lngIndex).Tag = strNewTag
  End If
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = modListMultiple.colForeColor
End Property

Public Property Let ForeColor(ByVal colNewForeColor As OLE_COLOR)
  modListMultiple.colForeColor = colNewForeColor
  Redraw
  PropertyChanged
End Property

Public Property Get ForeColorOver() As OLE_COLOR
  ForeColorOver = modListMultiple.colForeColorOver
End Property

Public Property Let ForeColorOver(ByVal colNewForeColorOver As OLE_COLOR)
  modListMultiple.colForeColorOver = colNewForeColorOver
  Redraw
  PropertyChanged
End Property

Public Property Get ForeColorSelected() As OLE_COLOR
  ForeColorSelected = modListMultiple.colForeColorSelected
End Property

Public Property Let ForeColorSelected(ByVal colNewForeColorSelected As OLE_COLOR)
  modListMultiple.colForeColorSelected = colNewForeColorSelected
  Redraw
  PropertyChanged
End Property

Public Property Get ForeColorCaption() As OLE_COLOR
  ForeColorCaption = modListMultiple.colForeColorCaption
End Property

Public Property Let ForeColorCaption(ByVal colNewForeColorCaption As OLE_COLOR)
  modListMultiple.colForeColorCaption = colNewForeColorCaption
  Redraw
  PropertyChanged
End Property

Public Property Get ForeColorCaptionOver() As OLE_COLOR
  ForeColorCaptionOver = modListMultiple.colForeColorCaptionOver
End Property

Public Property Let ForeColorCaptionOver(ByVal colNewForeColorCaptionOver As OLE_COLOR)
  modListMultiple.colForeColorCaptionOver = colNewForeColorCaptionOver
  Redraw
  PropertyChanged
End Property

Public Property Get ForeColorCaptionSelected() As OLE_COLOR
  ForeColorCaptionSelected = modListMultiple.colForeColorCaptionSelected
End Property

Public Property Let ForeColorCaptionSelected(ByVal colNewForeColorCaptionSelected As OLE_COLOR)
  modListMultiple.colForeColorCaptionSelected = colNewForeColorCaptionSelected
  Redraw
  PropertyChanged
End Property

Public Property Get MultiSelect() As Boolean
  MultiSelect = blnMultiSelect
End Property

Public Property Let MultiSelect(ByVal blnNewMultiSelect As Boolean)
  blnMultiSelect = blnNewMultiSelect
  PropertyChanged
End Property

Public Property Get Enabled() As Boolean
  Enabled = blnEnabled
End Property

Public Property Let Enabled(ByVal blnNewEnabled As Boolean)
  blnEnabled = blnNewEnabled
  Redraw
  PropertyChanged
End Property

Private Sub lsiItem_Check(Index As Integer, ByVal strKey As String)
  RaiseEvent Check(strKey)
End Sub

Private Sub lsiItem_DblClick(Index As Integer, ByVal strKey As String)
  RaiseEvent DblClick(strKey)
End Sub

Private Sub lsiItem_KeyDown(Index As Integer, ByVal strKey As String, KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(strKey, KeyCode, Shift)
End Sub

Private Sub lsiItem_MouseDown(Index As Integer, ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
  selectItems strKey
  RaiseEvent MouseDown(strKey, Button, Shift, x, y)
  RaiseEvent Click(strKey)
End Sub

Private Sub lsiItem_MouseMove(Index As Integer, ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
  moveMouse
  RaiseEvent MouseMove(strKey, Button, Shift, x, y)
End Sub

Private Sub lsiItem_MouseUp(Index As Integer, ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(strKey, Button, Shift, x, y)
End Sub

Private Sub scrVertical_Change()
  Redraw
End Sub

Private Sub UserControl_Initialize()
  Init
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  moveMouse
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set IconChecked = PropBag.ReadProperty("IconChecked", Nothing)
  Set IconUnChecked = PropBag.ReadProperty("IconUnChecked", Nothing)
  BackColor = PropBag.ReadProperty("BackColor", cnstColBackColor)
  BackColorOver = PropBag.ReadProperty("BackColorOver", cnstColBackColorOver)
  BackColorSelected = PropBag.ReadProperty("BackColorSelected", cnstColBackColorSelected)
  ForeColor = PropBag.ReadProperty("ForeColor", cnstColForeColor)
  ForeColorOver = PropBag.ReadProperty("ForeColorOver", cnstColForeColorOver)
  ForeColorSelected = PropBag.ReadProperty("ForeColorSelected", cnstColForeColorSelected)
  ForeColorCaption = PropBag.ReadProperty("ForeColorCaption", cnstColForeColorCaption)
  ForeColorCaptionOver = PropBag.ReadProperty("ForeColorCaptionOver", cnstColForeColorCaptionOver)
  ForeColorCaptionSelected = PropBag.ReadProperty("ForeColorCaptionSelected", cnstColForeColorCaptionSelected)
  MultiSelect = PropBag.ReadProperty("MultiSelect", False)
  Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
  Resize
End Sub

Private Sub UserControl_Show()
  Refresh
End Sub

Private Sub UserControl_Terminate()
  ReleaseCapture
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "IconChecked", IconChecked, 0
  PropBag.WriteProperty "IconUnChecked", IconUnChecked, 0
  PropBag.WriteProperty "BackColor", BackColor, cnstColBackColor
  PropBag.WriteProperty "BackColorOver", BackColorOver, cnstColBackColorOver
  PropBag.WriteProperty "BackColorSelected", BackColorSelected, cnstColBackColorSelected
  PropBag.WriteProperty "ForeColor", ForeColor, cnstColForeColor
  PropBag.WriteProperty "ForeColorOver", ForeColorOver, cnstColForeColorOver
  PropBag.WriteProperty "ForeColorSelected", ForeColorSelected, cnstColForeColorSelected
  PropBag.WriteProperty "ForeColorCaption", ForeColorCaption, cnstColForeColorCaption
  PropBag.WriteProperty "ForeColorCaptionOver", ForeColorCaptionOver, cnstColForeColorCaptionOver
  PropBag.WriteProperty "ForeColorCaptionSelected", ForeColorCaptionSelected, cnstColForeColorCaptionSelected
  PropBag.WriteProperty "MultiSelect", MultiSelect, False
  PropBag.WriteProperty "Enabled", Enabled, True
End Sub
