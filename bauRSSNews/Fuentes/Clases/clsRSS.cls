VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con los métodos de carga y manejo de RSS
Option Explicit

'Variables públicas
'Public strKey As String
'Public strTitle As String

'Objetos públicos
Public objColChannels As colRSSChannels

Private Sub Init()
'--> Inicializa las variables
  'Inicializa la colección de canales
    Set objColChannels = New colRSSChannels
End Sub

Private Sub Destroy()
'--> Limpia la memoria
  Clear
  Set objColChannels = Nothing
End Sub

Public Sub Clear()
'--> Limpia los datos del archivo RSS
  If Not objColChannels Is Nothing Then
    objColChannels.Clear
  End If
End Sub

Public Function countNewItems() As Long
'--> Cuenta los elementos no leídos
Dim objChannel As clsRSSChannel

  'Inicializa la variable de salida
    countNewItems = 0
  'Recorre los canales contando los elementos no leídos
    For Each objChannel In objColChannels
      countNewItems = countNewItems + objChannel.countNewItems
    Next objChannel
  'Libera la memoria
    Set objChannel = Nothing
End Function

Public Function parseXML(ByVal objDocument As MSXML2.DOMDocument30) As enumError
'--> Interpreta la cadena XML
Dim objNode As MSXML2.IXMLDOMNode, objNodeChannel As MSXML2.IXMLDOMNode

  'Busca la raíz
    For Each objNode In objDocument.childNodes
      If objNode.baseName = "rss" Then
        'Busca los canales
          For Each objNodeChannel In objNode.childNodes
            If parseXML = errorNoError Then '... para qué vamos a continuar si hay algún error
              If objNodeChannel.baseName = "channel" Then
                parseXML = objColChannels.Load(objNodeChannel)
              End If
            End If
          Next objNodeChannel
        'Libera la memoria
          Set objNodeChannel = Nothing
      ElseIf (objNode.prefix = "rdf" And objNode.baseName = "RDF") Then 'RSS 0.9
        'Busca los canales, imágenes y elementos
          For Each objNodeChannel In objNode.childNodes
            If parseXML = errorNoError Then '... para qué vamos a continuar si hay algún error
              If objNodeChannel.baseName = "channel" Then
                parseXML = objColChannels.Load(objNodeChannel)
              ElseIf objNodeChannel.baseName = "image" Then
                If objColChannels.Count > 0 Then
                  If Not objColChannels.Item(objColChannels.Count) Is Nothing Then
                    parseXML = objColChannels.Item(objColChannels.Count).objImage.Load(objNodeChannel)
                  End If
                End If
              ElseIf objNodeChannel.baseName = "item" Then
                If objColChannels.Count > 0 Then
                  If Not objColChannels.Item(objColChannels.Count) Is Nothing Then
                    parseXML = objColChannels.Item(objColChannels.Count).objColItems.Load(objNodeChannel)
                  End If
                End If
              End If
            End If
          Next objNodeChannel
        'Libera la memoria
          Set objNodeChannel = Nothing
      End If
    Next objNode
  'Libera la memoria
    Set objNode = Nothing
End Function

Public Function Merge(ByVal strURL As String, ByVal strKey As String) As Boolean
'--> Mezcla el archivo local con el archivo de la Web
Dim objRSSWeb As New clsItem
Dim objChannel As clsRSSChannel, objChannelWeb As clsRSSChannel
Dim objItemWeb As clsRSSItem, objItem As clsRSSItem
Dim blnFoundChannel As Boolean, blnFoundItem As Boolean, blnUpdate As Boolean, blnUpdateOnlyHeader As Boolean
Dim strXML As String

  'Supone que no hay nada nuevo
    blnUpdate = False
  'Supone que sólo se modifica la cabecera
    blnUpdateOnlyHeader = True
  'Carga el nuevo objeto
    objRSSWeb.strURL = strURL
    objRSSWeb.strKey = strKey
    objRSSWeb.intType = itemRSS
    If objRSSWeb.Load(True) Then
      'Recorre los canales que se han bajado de la web
        For Each objChannelWeb In objRSSWeb.objRSS.objColChannels
          'Supone que no se encuentra el canal
            blnFoundChannel = False
          'Recorre los canales del objeto actual
            For Each objChannel In objColChannels
              If UCase(Trim(objChannel.strTitle)) = UCase(Trim(objChannelWeb.strTitle)) Then
                'Indica que ha encontrado una coincidencia
                  blnFoundChannel = True
                'Si es de una fecha distinta ...
                  'If UCase(Trim(objChannel.strLastBuildDate)) <> UCase(Trim(objChannelWeb.strLastBuildDate)) Or _
                     UCase(Trim(objChannel.strPubDate)) <> UCase(Trim(objChannelWeb.strPubDate)) Then
                    'Cambia las propiedades del canal
                      With objChannel
                        .strTitle = objChannelWeb.strTitle
                        .strURL = objChannelWeb.strURL
                        .strDescription = objChannelWeb.strDescription
                        .strLanguage = objChannelWeb.strLanguage
                        .strLastBuildDate = objChannelWeb.strLastBuildDate
                        .strPubDate = objChannelWeb.strPubDate
                        .strGenerator = objChannelWeb.strGenerator
                        .strTtl = objChannelWeb.strTtl
                        .objImage.strTitle = objChannelWeb.objImage.strTitle
                        .objImage.strLink = objChannelWeb.objImage.strLink
                        .objImage.strURL = objChannelWeb.objImage.strURL
                        .objImage.strWidth = objChannelWeb.objImage.strWidth
                        .objImage.strHeight = objChannelWeb.objImage.strHeight
                      End With
                    'Indica que se han realizado modificaciones
                    '  blnUpdate = True
                    'Busca entre los artículos
                      For Each objItemWeb In objChannelWeb.objColItems
                        'Indica que no se ha encontrado el elemento
                          blnFoundItem = False
                        'Busca el elemento
                          For Each objItem In objChannel.objColItems
                            If objItemWeb.strLink = objItem.strLink Then '... si estamos en el mismo elemento
                              'Indica que se ha encontrado
                                blnFoundItem = True
                              'Si se ha publicado una modificación y no se ha borrado anteriormente el elemento
                                If UCase(Trim(objItemWeb.strPubDate)) <> UCase(Trim(objItem.strPubDate)) And _
                                   objItem.intStatus <> rssItemStatusDeleted Then
                                  'Cambia los propiedades del elemento
                                    With objItem
                                      .strTitle = objItemWeb.strTitle
                                      .strLink = objItemWeb.strLink
                                      .strComments = objItemWeb.strComments
                                      .strAuthor = objItemWeb.strAuthor
                                      .strPubDate = objItemWeb.strPubDate
                                      .strCategory = objItemWeb.strCategory
                                      .strSubject = objItemWeb.strSubject
                                      .strGuid = objItemWeb.strGuid
                                      .strDescription = objItemWeb.strDescription
                                      .strContent = objItemWeb.strContent
                                    End With
                                  'Indica que se han realizado modificaciones
                                    blnUpdate = True
                                    blnUpdateOnlyHeader = False
                                End If
                            End If
                          Next objItem
                        'Si no se ha encontrado el elemento, lo añade
                          If Not blnFoundItem Then
                            'Añade el elemento
                              With objChannel.objColItems.Add()
                                .strTitle = objItemWeb.strTitle
                                .strLink = objItemWeb.strLink
                                .strComments = objItemWeb.strComments
                                .strAuthor = objItemWeb.strAuthor
                                .strPubDate = objItemWeb.strPubDate
                                .strCategory = objItemWeb.strCategory
                                .strSubject = objItemWeb.strSubject
                                .strGuid = objItemWeb.strGuid
                                .strDescription = objItemWeb.strDescription
                                .strContent = objItemWeb.strContent
                              End With
                            'Indica que se han realizado modificaciones
                              blnUpdate = True
                              blnUpdateOnlyHeader = False
                          End If
                        'Libera la memoria
                          Set objItem = Nothing
                      Next objItemWeb
                    'Libera la memoria
                      Set objItemWeb = Nothing
                  'End If
              End If
            Next objChannel
          'Libera la memoria
            Set objChannel = Nothing
          'Si no ha encontrado el canal, lo crea
            If Not blnFoundChannel Then
              'Añade el canal y le pasa las propiedades obtenidas del Web
                With objColChannels.Add()
                  'Cambia las propiedades
                    .strTitle = objChannelWeb.strTitle
                    .strURL = objChannelWeb.strURL
                    .strDescription = objChannelWeb.strDescription
                    .strLanguage = objChannelWeb.strLanguage
                    .strLastBuildDate = objChannelWeb.strLastBuildDate
                    .strPubDate = objChannelWeb.strPubDate
                    .strGenerator = objChannelWeb.strGenerator
                    .strTtl = objChannelWeb.strTtl
                    .objImage.strTitle = objChannelWeb.objImage.strTitle
                    .objImage.strLink = objChannelWeb.objImage.strLink
                    .objImage.strURL = objChannelWeb.objImage.strURL
                    .objImage.strWidth = objChannelWeb.objImage.strWidth
                    .objImage.strHeight = objChannelWeb.objImage.strHeight
                  'Añade los elementos del canal obtenido del web
                    For Each objItemWeb In objChannelWeb.objColItems
                      With .objColItems.Add()
                        .strTitle = objItemWeb.strTitle
                        .strLink = objItemWeb.strLink
                        .strComments = objItemWeb.strComments
                        .strAuthor = objItemWeb.strAuthor
                        .strPubDate = objItemWeb.strPubDate
                        .strCategory = objItemWeb.strCategory
                        .strSubject = objItemWeb.strSubject
                        .strGuid = objItemWeb.strGuid
                        .strDescription = objItemWeb.strDescription
                        .strContent = objItemWeb.strContent
                      End With
                    Next objItemWeb
                  'Libera la memoria
                    Set objItemWeb = Nothing
                  'Indica que se han realizado modificaciones
                    blnUpdate = True
                    blnUpdateOnlyHeader = False
                End With
            End If
        Next objChannelWeb
      'Libera la memoria
        Set objChannelWeb = Nothing
    End If
  'Si se han realizado modificaciones, se graba el XML
    Merge = blnUpdate And Not blnUpdateOnlyHeader
  'Libera la memoria
    Set objRSSWeb = Nothing
End Function

Public Function writeHTML(ByVal lngFile As Long, ByVal strKey As String) As Boolean
'--> Crea el archivo HTML con el resumen de las noticias
  objColChannels.writeHTML lngFile, strKey
End Function

Public Function getXML() As String
'--> Obtiene el XML de este objeto
  'Cabecera
    getXML = " <rss version='2.0' xmlns:content='http://purl.org/rss/1.0/modules/content/'>" & vbCrLf
  'Contenido
    getXML = getXML & objColChannels.getXML()
  'Cierre
    getXML = getXML & "</rss>"
End Function

Public Function markReadRSS(ByVal strURL As String) As Boolean
'--> Marca una noticia como leída cuando se pulsa sobre la URL
Dim objChannel As clsRSSChannel
Dim objRSS As clsRSSItem

  'Indica que no ha marcado ningún elemento
    markReadRSS = False
  'Recorre los canales marcando los elementos como leídos
    For Each objChannel In objColChannels
      'Recorre los elementos del canal
        For Each objRSS In objChannel.objColItems
          If objRSS.intStatus <> rssItemStatusDeleted Then
            If strURL = UCase(Trim(objRSS.strLink)) Then
              'Cambia el estado
                objRSS.intStatus = rssItemStatusRead
              'Indica que se ha marcado como leído
                markReadRSS = True
            End If
          End If
        Next objRSS
      'Libera la memoria
        Set objRSS = Nothing
    Next objChannel
  'Libera la memoria
    Set objChannel = Nothing
End Function

Public Function loadList(ByVal strKeyItem As String, ByRef lsmRSS As ListMultiple, _
                         ByRef imgNew As StdPicture, ByRef imgRead As StdPicture) As Boolean
'--> Carga una lista con los parámetros asociados a un RSS
Dim objChannel As clsRSSChannel
Dim objRSSItem As clsRSSItem
Dim imgIcon As StdPicture
Dim strKey As String
  
  'Recorre los canales del RSS
    For Each objChannel In objColChannels
      'Recorre los elementos del canal
        For Each objRSSItem In objChannel.objColItems
          With objRSSItem
            If .intStatus <> rssItemStatusDeleted Then '... se salta los borrados
              'Crea la clave
                strKey = strKeyItem & "¬" & objChannel.strKey & "¬" & .strKey
              'Obtiene la imagen a utilizar
                If .intStatus = rssItemStatusNew Then
                  Set imgIcon = imgNew
                Else
                  Set imgIcon = imgRead
                End If
              'Añade el elemento a la lista
                lsmRSS.Add .strTitle, .strPubDate, .strDescription, False, .strLink, imgIcon, strKey
              'Cambia la fuente del título
                lsmRSS.FontCaption(strKey).Bold = True
                lsmRSS.FontCaption(strKey).Size = 9
              'Pone el título en negrita si es una nueva noticia
                If .intStatus = rssItemStatusNew Then
                  lsmRSS.FontCaption(strKey).Bold = True
                  lsmRSS.ForeColorCaptionItem(strKey) = vbRed
                End If
            End If
          End With
        Next objRSSItem
      'Libera la memoria
        Set objRSSItem = Nothing
    Next objChannel
  'Libera la memoria
    Set objChannel = Nothing
  'Indica que todo ha ido bien
    loadList = True
End Function

Private Sub Class_Initialize()
  Init
End Sub

Private Sub Class_Terminate()
  Destroy
End Sub
