VERSION 5.00
Begin VB.Form frmUpdateRSS 
   BackColor       =   &H00E0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Canal RSS"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3915
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   2580
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fraProperties 
      BackColor       =   &H00E0F0F0&
      Caption         =   " Propiedades "
      ForeColor       =   &H00000080&
      Height          =   1485
      Left            =   150
      TabIndex        =   10
      Top             =   90
      Width           =   4965
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   675
         Width           =   3675
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   3540
         TabIndex        =   7
         Top             =   1020
         Width           =   1275
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1140
         TabIndex        =   5
         Top             =   1020
         Width           =   1245
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1140
         TabIndex        =   1
         Top             =   330
         Width           =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   2460
         TabIndex        =   6
         Top             =   1065
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   4
         Top             =   1065
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   375
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmUpdateRSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario para la modificación de los RSS del proyecto
Option Explicit

'Variables públicas
Public strName As String
Public strURL As String
Public strUser As String
Public strPassword As String
Public blnCancel As Boolean

Private Sub Init()
'--> Inicializa el formulario
  'Muestra los valores
    txtName.Text = strName
    txtURL.Text = strURL
    txtUser.Text = strUser
    txtPassword.Text = strPassword
  'Cambia el idioma de la ventana
    changeLanguage
  'Supone que se cancelarán las modificaciones
    blnCancel = True
End Sub

Private Sub changeLanguage()
'--> Cambia el idioma de la ventana
  'Título
    Me.Caption = objColLanguage.searchItem(Me.Name, 1, Me.Caption)
  'Frame
    fraProperties.Caption = objColLanguage.searchItem(Me.Name, 2, fraProperties.Caption)
  'Labels
    Label1(0).Caption = objColLanguage.searchItem(Me.Name, 3, Label1(0).Caption)
    Label1(1).Caption = objColLanguage.searchItem(Me.Name, 4, Label1(1).Caption)
    Label1(2).Caption = objColLanguage.searchItem(Me.Name, 5, Label1(2).Caption)
    Label1(3).Caption = objColLanguage.searchItem(Me.Name, 6, Label1(3).Caption)
  'Botones
    cmdAccept.Caption = objColLanguage.searchItem(Me.Name, 7, cmdAccept.Caption)
    cmdCancel.Caption = objColLanguage.searchItem(Me.Name, 8, cmdCancel.Caption)
End Sub

Private Sub acceptData()
'--> Comprueba los datos y si son correctos los devuelve al programa principal
  'Quita los espacios
    txtName.Text = Trim(txtName.Text)
    txtURL.Text = Trim(txtURL.Text)
    txtUser.Text = Trim(txtUser.Text)
    txtPassword.Text = Trim(txtPassword.Text)
  'Comprueba los datos antes de devolverlos
    If txtName.Text = "" Then
      MsgBox objColLanguage.searchItem(Me.Name, 9, "Introduzca el nombre del canal")
    ElseIf txtURL.Text = "" Then
      MsgBox objColLanguage.searchItem(Me.Name, 10, "Introduzca el nombre del canal")
    Else
      'Graba los datos
        strName = txtName.Text
        strURL = txtURL.Text
        strUser = txtUser.Text
        strPassword = txtPassword.Text
      'Añade "http://" a la URL si es necesario
        If UCase(left(strURL, 7)) <> "HTTP://" Then
          strURL = "http://" & strURL
        End If
      'Indica que se han aceptado los datos
        blnCancel = False
      'Descarga la ventana
        Unload Me
    End If
End Sub

Private Sub cmdAccept_Click()
  acceptData
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Init
End Sub
