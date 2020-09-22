VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOPML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con los datos de un archivo OPML
'Formato de un archivo OPML
'<?xml version="1.0" encoding="ISO-8859-1"?>
'<opml version="1.0">
'  <head>
'    <title>PapaScott Favorites</title>
'    <dateCreated>Thu, 09 Mar 2000 19:13:50 GMT</dateCreated>
'    <dateModified>Sun, 01 Jul 2001 08:25:47 GMT</dateModified>
'    <ownerName>Scott Hanson</ownerName>
'    <ownerEmail>shanson@shcon.com</ownerEmail>
'    <expansionState></expansionState>
'    <vertScrollState>30</vertScrollState>
'    <windowTop>288</windowTop>
'    <windowLeft>285</windowLeft>
'    <windowBottom>488</windowBottom>
'    <windowRight>667</windowRight>
'  </head>
'  <body>
'    <outline text="2020 Hindsight"
'             htmlUrl = "http://www.2020Hindsight.org/"
'             type="weblogsComFavorites"
'             url="http://static.userland.com/weblogMonitor/favorites/2020Hindsight.opml"/>
'  </body>
'</opml>
Option Explicit

'Variables públicas
Public strTitle As String
Public strDateCreated As String
Public strDateModified As String
Public strOwnerName As String
Public strOwnerEMail As String

'Objetos públicos
Public objColItems As colItems

Private Sub Init()
'--> Inicializa las variables
  Set objColItems = New colItems
End Sub

Private Sub Destroy()
'--> Limpia la memoria
  Clear
  Set objColItems = Nothing
End Sub

Public Sub Clear()
'--> Limpia los datos del archivo OPML
  If Not objColItems Is Nothing Then
    objColItems.Clear
  End If
End Sub

Public Function parseXML(ByVal objDocument As MSXML2.DOMDocument30) As enumError
'--> Interpreta el archivo OPML
Dim objNode As MSXML2.IXMLDOMNode
Dim objItem As MSXML2.IXMLDOMNode
Dim objHead As MSXML2.IXMLDOMNode
Dim objOutline As MSXML2.IXMLDOMNode
Dim objAttribute As MSXML2.IXMLDOMAttribute

  'Busca la raíz
    For Each objNode In objDocument.childNodes
      If objNode.baseName = "opml" Then '... si se trata de un archivo OPML
        'Recorre los elementos del archivo OPML
          For Each objItem In objNode.childNodes
            If objItem.baseName = "head" Then '... lee la cabecera
              'Recorre los elementos
                For Each objHead In objItem.childNodes
                  If objHead.baseName = "title" Then
                    strTitle = objHead.Text
                  ElseIf objHead.baseName = "dateCreated" Then
                    strDateCreated = objHead.Text
                  ElseIf objHead.baseName = "dateModified" Then
                    strDateModified = objHead.Text
                  ElseIf objHead.baseName = "ownerName" Then
                    strOwnerName = objHead.Text
                  ElseIf objHead.baseName = "ownerEmail" Then
                    strOwnerEMail = objHead.Text
                  End If
                Next objHead
              'Libera la memoria
                Set objHead = Nothing
            ElseIf objItem.baseName = "body" Then '... lee el cuerpo
              'Recorre los elementos
                For Each objOutline In objItem.childNodes
                  'Añade un elemento a la colección de elementos del OPML
                    With objColItems.Add(itemRSS, "", "", "", "")
                      If objOutline.baseName = "outline" Then
                        'Recorre los atributos
                          For Each objAttribute In objOutline.Attributes
                            If objAttribute.baseName = "text" Then
                              .strTitle = objAttribute.Text
                            ElseIf objAttribute.baseName = "htmlURL" Then
                              .strWeb = objAttribute.Text
                            ElseIf objAttribute.baseName = "type" Then
                              .strType = objAttribute.Text
                            ElseIf objAttribute.baseName = "url" Then
                              .strURL = objAttribute.Text
                              .strKey = .strURL '... crea una clave única
                            End If
                          Next objAttribute
                        'Libera la memoria
                          Set objAttribute = Nothing
                      End If
                    End With
                Next objOutline
              'Libera la memoria
                Set objOutline = Nothing
            End If
          Next objItem
        'Libera la memoria
          Set objItem = Nothing
      End If
    Next objNode
  'Libera la memoria
    Set objNode = Nothing
End Function

Public Function Merge() As Boolean
'--> Mezcla el contenido del archivo actual con el de la Web
  MsgBox "Merge OPML"
End Function

Public Function writeHTML(ByVal lngFile As Long) As Boolean
'--> Crea el archivo HTML con el resumen de los canales
Dim objItem As clsItem

  'Recorre los elementos del canal
    For Each objItem In objColItems
      If objItem.intType = itemRSS Then '... realmente esto no es necesario
        'Carga el elemento
          If objItem.Load() Then
            'Escribe el HTML
              objItem.objRSS.writeHTML lngFile, ""
          End If
      End If
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
End Function

Public Function getXML() As String
'--> Obtiene el XML de este objeto
Dim objRSS As clsRSS

  'Cabecera
    getXML = getXML & " <opml version='1.0'>" & vbCrLf
    getXML = getXML & "   <head>"
    getXML = getXML & "     <title>" & strTitle & "</title>"
    getXML = getXML & "     <dateCreated>" & strDateCreated & "</dateCreated>"
    getXML = getXML & "     <dateModified>" & strDateModified & "</dateModified>"
    getXML = getXML & "     <ownerName>" & strOwnerName & "</ownerName>"
    getXML = getXML & "     <ownerEmail>" & strOwnerEMail & "</ownerEmail>"
    getXML = getXML & "   </head>"
  'Contenido
    getXML = getXML & "   <body>"
    For Each objRSS In objColItems
      getXML = getXML & "<ouline text = '" & objRSS & "'"
      getXML = getXML & " htmlUrl = '" & objRSS.strWeb & "'"
      getXML = getXML & " type = '" & objOutline.strType & "'"
      getXML = getXML & " url = '" & objOutline.strURL & "'"
      getXML = getXML & "/>"
    Next objOutline
    Set objOutline = Nothing
    getXML = getXML & "   </body>"
  'Cierre
    strXML = strXML & "</opml>"
End Function

Public Function loadList(ByVal strKeyItem As String, ByRef lsmRSS As ListMultiple) As Boolean
'--> Carga una lista con los parámetros asociados a un OPML
Dim objItem As clsItem

  'Recorre los elementos del OPML
    For Each objItem In objColItems
      'Añade el elemento a la lista
        With objItem
          lsmRSS.Add .strTitle, .strURL, False, , , .strKey
        End With
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
  'Indica que todo ha ido bien
    loadList = True
End Function

Private Sub Class_Initialize()
  Init
End Sub

Private Sub Class_Terminate()
  Destroy
End Sub
