VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSSChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase de mantenimiento de canales RSS
Option Explicit

'Variables con las propiedades del canal
Public strKey As String
Public strTitle As String
Public strURL As String
Public strDescription As String
Public strLanguage As String
Public strCopyright As String
Public strLastBuildDate As String
Public strPubDate As String
Public strGenerator As String
Public strTtl As String

'Objeto de imagen
Public objImage As clsRSSImage
Public objColItems As colRSSItems

Private Sub Init()
'--> Inicializa los valores de la clase
  'Crea la colección de elementos
    Set objColItems = New colRSSItems
  'Crea una nueva imagen
    Set objImage = New clsRSSImage
End Sub

Private Sub Destroy()
  'Limpia los elementos
    If Not objColItems Is Nothing Then
      objColItems.Clear
      Set objColItems = Nothing
    End If
  'Libera la memoria
    Set objImage = Nothing
End Sub

Public Function countNewItems() As Long
'--> Cuenta los elementos no leídos
Dim objItem As clsRSSItem

  'Inicializa la variable de salida
    countNewItems = 0
  'Recorre los elementos contando los no leídos
    For Each objItem In objColItems
      If objItem.intStatus <> rssItemStatusRead Then
        countNewItems = countNewItems + 1
      End If
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
End Function

Public Property Get RSSImage() As clsRSSImage
  Set RSSImage = objImage
End Property

Private Sub Class_Initialize()
  Init
End Sub

Private Sub Class_Terminate()
  Destroy
End Sub

