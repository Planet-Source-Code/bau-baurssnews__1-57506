VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSSImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase de mantenimiento de imágenes
Option Explicit

'Variables públicas
Public strTitle As String
Public strWidth As String
Public strHeight As String
Public strLink As String
Public strURL As String

Public Function Load(ByRef objDocument As MSXML2.IXMLDOMNode) As enumError
'--> Carga los datos de la imagen de un canal
'    <image>
'      <title>Yahoo! News</title>
'      <width>142</width>
'      <height>18</height>
'      <link>http://news.yahoo.com/</link>
'      <url>http://us.i1.yimg.com/us.yimg.com/i/us/nws/th/main_142.gif</url>
'    </image>
Dim objNode As MSXML2.IXMLDOMNode

  On Error GoTo errorLoad
    'Supone que puede cargar el archivo
      Load = errorNoError
    'Busca los datos del canal
      For Each objNode In objDocument.childNodes
        If Load = errorNoError Then '... para qué vamos a continuar si hay algún error
          If objNode.baseName = "title" Then
            strTitle = objNode.Text
          ElseIf objNode.baseName = "width" Then
            strWidth = objNode.Text
          ElseIf objNode.baseName = "height" Then
            strHeight = objNode.Text
          ElseIf objNode.baseName = "link" Then
            strLink = objNode.Text
          ElseIf objNode.baseName = "url" Then
            strURL = objNode.Text
          End If
        End If
      Next objNode
    'Libera la memoria
      Set objNode = Nothing
  Exit Function
  
errorLoad:
  Load = errorUnknown
End Function
