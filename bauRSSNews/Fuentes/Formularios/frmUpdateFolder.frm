VERSION 5.00
Begin VB.Form frmUpdateFolder 
   BackColor       =   &H00E0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Propiedades de carpeta"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraProperties 
      BackColor       =   &H00E0F0F0&
      Caption         =   " Propiedades "
      ForeColor       =   &H00000080&
      Height          =   765
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   4965
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   990
         TabIndex        =   3
         Top             =   330
         Width           =   3825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   2430
      TabIndex        =   1
      Top             =   870
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3765
      TabIndex        =   0
      Top             =   870
      Width           =   1215
   End
End
Attribute VB_Name = "frmUpdateFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario para la modificación de las carpetas del proyecto
Option Explicit

'Variables públicas
Public strName As String
Public blnCancel As Boolean

Private Sub Init()
'--> Inicializa el formulario
  'Muestra los valores
    txtName.Text = strName
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
  'Botones
    cmdAccept.Caption = objColLanguage.searchItem(Me.Name, 4, cmdAccept.Caption)
    cmdCancel.Caption = objColLanguage.searchItem(Me.Name, 5, cmdCancel.Caption)
End Sub

Private Sub acceptData()
'--> Comprueba los datos y si son correctos los devuelve al programa principal
  'Quita los espacios
    txtName.Text = Trim(txtName.Text)
  'Comprueba los datos antes de devolverlos
    If txtName.Text = "" Then
      MsgBox objColLanguage.searchItem(Me.Name, 6, "Introduzca el nombre de la carpeta")
    Else
      'Graba los datos
        strName = txtName.Text
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

