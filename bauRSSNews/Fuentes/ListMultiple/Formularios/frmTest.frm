VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin bauRSSNews.ListMultiple ListMultiple1 
      Height          =   2055
      Left            =   390
      TabIndex        =   0
      Top             =   720
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   3625
      IconChecked     =   "frmTest.frx":0000
      IconUnChecked   =   "frmTest.frx":059A
      ForeColorCaptionOver=   8421631
      ForeColorCaptionSelected=   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim picMain As StdPicture, picMain2 As StdPicture
Dim lngIndex As Long

  Set picMain = LoadPicture("C:\Documents and Settings\bautistaja\Mis documentos\Mis imágenes\Iconos\MyDocuments.ico")
  Set picMain2 = LoadPicture("C:\Documents and Settings\bautistaja\Mis documentos\Mis imágenes\Iconos\New 16x16.ico")
  With ListMultiple1
    For lngIndex = 0 To 30
      .Add "Elemento " & lngIndex, "Descripción elemento " & lngIndex, (lngIndex Mod 2 = 0), , _
           IIf(lngIndex Mod 2 = 0, picMain, picMain2)
    Next lngIndex
    .MultiSelect = True
  End With
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    With ListMultiple1
      .Top = ScaleTop
      .Left = ScaleLeft
      .Width = ScaleWidth - .Left
      .Height = ScaleHeight - .Top
    End With
End Sub
