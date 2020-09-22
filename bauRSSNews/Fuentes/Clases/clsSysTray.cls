VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para mantener iconos en la barra de tareas
Option Explicit

'Constantes
Private Const NIM_ADD As Long = &H0
Private Const NIM_MODIFY As Long = &H1
Private Const NIM_DELETE As Long = &H2
Private Const WM_MOUSEMOVE As Long = &H200
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_ICON As Long = &H2
Private Const NIF_TIP As Long = &H4

Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205

'Declaraciones de API
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    
'Tipo de API
Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

'Eventos
Public Event DblClick(ByVal intButton As Integer, ByVal intShift As Integer, ByVal sngX As Single, ByVal sngY As Single)
Public Event MouseDown(ByVal intButton As Integer, ByVal intShift As Integer, ByVal sngX As Single, ByVal sngY As Single)
Public Event MouseUp(ByVal intButton As Integer, ByVal intShift As Integer, ByVal sngX As Single, ByVal sngY As Single)

'Variables locales
Private notifyIcon As NOTIFYICONDATA
Private WithEvents picHook As PictureBox
Attribute picHook.VB_VarHelpID = -1
Private strToolTip As String

'Variables locales con los valores de las propiedades
Private frmSourceWindow As Form
Private blnDefaultDblClick As Boolean

Public Sub RemoveFromSysTray()
'--> Elimina el icono de la barra de tareas
  notifyIcon.cbSize = Len(notifyIcon)
  notifyIcon.hwnd = picHook.hwnd
  notifyIcon.uId = 1&
  Shell_NotifyIcon NIM_DELETE, notifyIcon
End Sub

Public Sub IconInSysTray()
'--> Coloca el icono en la barra de tareas
  Shell_NotifyIcon NIM_ADD, notifyIcon
End Sub

Public Sub MinToSysTray()
'--> Minimiza la aplicación en la barra de tareas
  IconInSysTray
  frmSourceWindow.Hide
  App.TaskVisible = False
End Sub

Private Sub SetPicHook()
'--> Configura la imagen
  On Error GoTo AlreadyAdded
    Set picHook = frmSourceWindow.Controls.Add("VB.PictureBox", "pichook")
    picHook.Visible = False
    picHook.Picture = frmSourceWindow.Icon
    notifyIcon.hwnd = picHook.hwnd
  Exit Sub

AlreadyAdded:
End Sub

Public Sub ChangeIcon(ByRef picNewIcon As Variant)
'--> Modifica el icono en la barra de sistema
  Set picHook.Picture = picNewIcon
  notifyIcon.hIcon = picHook.Picture
  Shell_NotifyIcon NIM_MODIFY, notifyIcon
End Sub

Public Sub ChangeToolTip(ByVal strNewTip As String)
'--> Cambia el tooltip en la barra de sistema
  strToolTip = strNewTip
  notifyIcon.szTip = strNewTip & Chr(0)
  Shell_NotifyIcon NIM_MODIFY, notifyIcon
  If frmSourceWindow.WindowState = vbMinimized Then
    frmSourceWindow.Caption = strNewTip
  End If
End Sub

Public Property Let ToolTip(ByVal strData As String)
  ChangeToolTip strData
End Property

Public Property Get ToolTip() As String
  ToolTip = strToolTip
End Property

Public Property Let Icon(ByVal varData As Variant)
  ChangeIcon varData
End Property

Public Property Get Icon() As Variant
  Icon = notifyIcon.hIcon
End Property

Public Property Let DefaultDblClk(ByVal blnData As Boolean)
  blnDefaultDblClick = blnData
End Property

Public Property Get DefaultDblClk() As Boolean
  DefaultDblClk = blnDefaultDblClick
End Property

Public Property Set SourceWindow(ByVal frmData As Form)
  Set frmSourceWindow = frmData
  SetPicHook
End Property

Public Property Get SourceWindow() As Form
  Set SourceWindow = frmSourceWindow
End Property

Private Sub Class_Initialize()
  blnDefaultDblClick = True
  notifyIcon.cbSize = Len(notifyIcon)
  notifyIcon.uId = 1&
  notifyIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  notifyIcon.ucallbackMessage = WM_MOUSEMOVE
  notifyIcon.hIcon = Me.Icon
  notifyIcon.szTip = Chr(0)
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static rec As Boolean, msg As Long, oldmsg As Long
    
  oldmsg = msg
  msg = x / Screen.TwipsPerPixelX
  If Not rec Then
    rec = True
    Select Case msg
      Case WM_LBUTTONDBLCLK
        If blnDefaultDblClick Then
          frmSourceWindow.WindowState = vbMaximized
          frmSourceWindow.Show
          App.TaskVisible = True
          RemoveFromSysTray
        End If
        RaiseEvent DblClick(vbLeftButton, Shift, x, y)
      Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(vbLeftButton, Shift, x, y)
      Case WM_LBUTTONUP
        RaiseEvent MouseUp(vbLeftButton, Shift, x, y)
      Case WM_RBUTTONDBLCLK
        RaiseEvent DblClick(vbRightButton, Shift, x, y)
      Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(vbRightButton, Shift, x, y)
      Case WM_RBUTTONUP
        RaiseEvent MouseUp(vbRightButton, Shift, x, y)
    End Select
    rec = False
  End If
End Sub
