VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con los datos del proyecto de noticias
'<?xml version="1.0" encoding="iso-8859-1" ?>
'<NewsProject>
'  <rss>
'    <title>Awasu</title>
'    <url>http://www.awasu.com/weblog/wp-rss2.php</url>
'    <user></user>
'    <password></password>
'  </rss>
'  <opml>
'    <url>http://www.almirot.com/almirot_15_12_03.opml></url>
'    <user></user>
'    <password></password>
'  </opml>
'</NewsProject>
Option Explicit

'Objetos públicos
Public objColItems As colItems

'Variables públicas
Public lngLastKey As Long

Public Sub Init()
'--> Inicializa los objetos
  Set objColItems = New colItems
  lngLastKey = 0
End Sub

Public Sub Clear()
'--> Limpia la memoria
  objColItems.Clear
  lngLastKey = 0
End Sub

Public Function Load(ByVal strFileName As String) As Boolean
'--> Carga un archivo de proyecto
Dim objDocument As New MSXML2.DOMDocument30
Dim objNode As MSXML2.IXMLDOMNode

  On Error GoTo errorLoad
    'Supone que no puede cargar el archivo
      Load = False
    'Carga el archivo
      If objDocument.Load(strFileName) Then
        'Busca la raíz
          For Each objNode In objDocument.childNodes
            If objNode.baseName = "NewsProject" Then
              Load = objColItems.Load(objNode)
            End If
          Next objNode
        'Libera la memoria
          Set objNode = Nothing
      End If
    'Libera la memoria
      Set objDocument = Nothing
  Exit Function
  
errorLoad:
End Function

Public Sub markReadRSS(ByVal strURL As String)
'--> Marca una noticia como leída cuando se pulsa sobre la URL
  objColItems.markReadRSS UCase(Trim(strURL))
End Sub

Public Function Save(ByVal strFileName As String) As Boolean
'--> Graba los datos
Dim lngFile As Long
Dim objRSS As clsRSS

  On Error GoTo errorSave
    'Supone que no puede grabar el archivo
      Save = False
    'Abre el archivo
      lngFile = FreeFile()
      Open strFileName For Output As #lngFile
    'Escribe los datos XML
      'Cabecera
        Print #lngFile, "<?xml version='1.0' encoding='iso-8859-1' ?>"
        Print #lngFile, "<NewsProject>"
      'Escribe los datos de los elementos
        objColItems.Save lngFile
      'Cierre
        Print #lngFile, "</NewsProject>"
    'Cierra el archivo
      Close #lngFile
    'Si ha llegado hasta aquí es porque todo ha ido bien
      Save = True
  Exit Function

errorSave:
End Function

Public Function loadList(ByVal strKey As String, ByRef lsmRSS As ListMultiple, _
                         ByRef imgFolder As StdPicture, ByRef imgRSS As StdPicture, _
                         ByRef imgNew As StdPicture, ByRef imgRead As StdPicture, _
                         ByRef imgPageWeb As StdPicture) As Boolean
'--> Carga una lista con los parámetros asociados a un elemento
Dim objItem As clsItem

  'Supone que no puede cargar la lista
    loadList = False
  'Obtiene el objeto
    Set objItem = objColItems.Search(strKey)
  'Si realmente hemos encontrado algo cargamos la lista
    If Not objItem Is Nothing Then
      loadList = objItem.loadList(lsmRSS, imgFolder, imgRSS, imgNew, imgRead, imgPageWeb)
    End If
  'Libera la memoria
    Set objItem = Nothing
End Function

Public Function loadHTML(ByVal strKey As String, ByRef brwRSS As WebBrowser) As Boolean
'--> Crea el HTML de un elemento
Dim strFileName As String
Dim lngFile As Long
Dim objItem As clsItem

  On Error GoTo errorLoad
    'Supone que no puede crear el HTML
      loadHTML = False
    'Obtiene el objeto seleccionado
      Set objItem = objColItems.Search(strKey)
    'Si realmente hemos encontrado algo creamos el HTML
      If Not objItem Is Nothing Then
        If objItem.intType = itemWebPage Then
          'Simplemente se pasa a la página
            brwRSS.Navigate2 objItem.strURL
          'Y se indica que es correcto
            loadHTML = True
        Else
          'Obtiene el nombre del archivo temporal
            strFileName = App.Path & "\tempRSS.html"
          'Elimina el archivo temporal
            killFile strFileName
          'Obtiene el handle al archivo
            lngFile = FreeFile()
          'Abre el archivo temporal para escritura
            Open strFileName For Output As #lngFile
          'Escribe el archivo HTML
            loadHTML = objItem.writeHTML(lngFile, strKey)
          'Cierra el archivo
            Close #lngFile
          'y si lo ha podido escribir, lo carga en el navegador
            If loadHTML Then
              brwRSS.Navigate2 strFileName
            End If
        End If
      End If
  Exit Function
  
errorLoad:
End Function

Public Function getURL(ByVal strKey As String) As String
'--> Obtiene la URL de una noticia
Dim objItem As clsItem
Dim arrStrKey() As String

  On Error GoTo errorGet
    'Supone que no encuentra la noticia
      getURL = ""
    'Busca el objeto
      Set objItem = objColItems.Search(strKey)
    'Si hemos encontrado el objeto
      If Not objItem Is Nothing Then
        '... y es un RSS
          If objItem.intType = itemRSS Then
            'Separa la clave en partes
              arrStrKey = Split(strKey, "¬")
            'Si tenemos toda la información
              If UBound(arrStrKey) = 2 Then
                'Obtenemos la URL
                  getURL = objItem.objRSS.objColChannels(arrStrKey(1)).objColItems(arrStrKey(2)).strLink
              End If
          End If
      End If
    'Libera la memoria
      Set objItem = Nothing
  Exit Function
  
errorGet:
End Function

Public Sub removeItem(ByVal strKey As String)
'--> Elimina un elemento
Dim objItem As clsItem

  'Busca el objeto
    Set objItem = objColItems.Search(strKey)
  'Elimina el elemento
    If Not objItem Is Nothing Then
      On Error Resume Next
        objItem.Clear
        objItem.strKey = ""
    End If
End Sub

Public Sub removeRSSItem(ByVal strKey As String, ByVal strChannel As String, _
                         ByVal strNew As String)
'--> Cambia el estado de un elemento
Dim objItem As clsItem

  'Busca el objeto
    Set objItem = objColItems.Search(strKey)
  'Si lo ha encontrado busca el canal
    If Not objItem Is Nothing Then
      On Error Resume Next
        objItem.objRSS.objColChannels(strChannel).objColItems(strNew).intStatus = rssItemStatusDeleted
        If Err.Number = 0 Then
          objItem.writeXML
        End If
    End If
  'Libera la memoria
    Set objItem = Nothing
End Sub

Private Sub Class_Initialize()
  Init
End Sub

Private Sub Class_Terminate()
  Clear
  Set objColItems = Nothing
End Sub
