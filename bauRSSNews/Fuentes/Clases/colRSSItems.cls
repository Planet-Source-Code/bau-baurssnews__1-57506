VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colRSSItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsRSSItem"
Attribute VB_Ext_KEY = "Member0" ,"clsRSSItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección de elementos de un canal RSS
Option Explicit

'Constantes locales
Private Const cnstStrStatusNew = "New"
Private Const cnstStrStatusRead = "Read"
Private Const cnstStrStatusDeleted = "Deleted"

'Variable local con los datos de la colección
Private objColItems As Collection

Public Sub Clear()
'--> Limpia la colección de elementos
Dim lngIndex As Long

  For lngIndex = Count To 1 Step -1
    Remove lngIndex
  Next lngIndex
End Sub

Public Function Add(Optional ByVal strKey As String = "") As clsRSSItem
'--> Añade un objeto a la colección
Dim objNewMember As New clsRSSItem
  
  'Añade el elemento a la colección
    If Trim(strKey) = "" Then
      strKey = getNextKey()
    End If
    objColItems.Add objNewMember, strKey
  'Inicializa la clave
    objNewMember.strKey = strKey
  'Devuelve el objeto creado y libera la memoria
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Private Function getNextKey() As String
'--> Obtiene la siguiente clave
Dim lngLastKey As Long

  'Obtiene la siguiente clave
    lngLastKey = Count + 1
  'Mientras que exista esa clave, intenta pasar a la siguiente
    While Not Item("RSSItem" & lngLastKey) Is Nothing
      'Pasa a la siguiente clave
        lngLastKey = lngLastKey + 1
    Wend
  'Devuelve la clave
    getNextKey = "RSSItem" & lngLastKey
End Function

Public Function Load(ByRef objDocument As MSXML2.IXMLDOMNode) As enumError
'--> Carga los datos de un elemento
'<item>
'  <title>&#039;Let&#039;s Just Cuddle&#039; (Reuters)</title>
'  <link>http://us.rd.yahoo.com/dailynews/rss/oddlyenough/*http://story.news.yahoo.com/news?tmpl=story2&amp;u=/nm/20040825/od_nm/odd_germany_dictionary_dc</link>
'  <guid isPermaLink="false">nm/20040825/odd_germany_dictionary_dc</guid>
'  <pubDate>Wed, 25 Aug 2004 12:26:25 GMT</pubDate>
'  <description>
'    Reuters - A leading German dictionary publisher plans to launch a guide it says will help men translate the
'    subtext of female conversation.
'  </description>
'</item>
Dim objNode As MSXML2.IXMLDOMNode
Dim objItem As clsRSSItem

  On Error GoTo errorLoad
    'Supone que puede cargar el archivo
      Load = errorNoError
    'Añade un objeto a la colección
      Set objItem = Add(getNextKey())
    'Busca los datos del elemento
      For Each objNode In objDocument.childNodes
        If Load = errorNoError Then '... para qué vamos a continuar si hay algún error
          With objItem
            If objNode.baseName = "title" Then
              .strTitle = objNode.Text
            ElseIf objNode.baseName = "link" Then
              .strLink = objNode.Text
            ElseIf objNode.baseName = "author" Then
              .strAuthor = objNode.Text
            ElseIf objNode.baseName = "comments" Then
              .strComments = objNode.Text
            ElseIf objNode.baseName = "pubDate" Then
              .strPubDate = objNode.Text
            ElseIf objNode.baseName = "category" Then
              .strCategory = objNode.Text
            ElseIf objNode.baseName = "subject" Then
              .strSubject = objNode.Text
            ElseIf objNode.baseName = "guid" Then
              .strGuid = objNode.Text
            ElseIf objNode.baseName = "description" Then
              .strDescription = objNode.Text
            ElseIf objNode.prefix = "content" And objNode.baseName = "encoded" Then
              .strContent = objNode.Text
            ElseIf objNode.baseName = "status" Then
              Select Case UCase(Trim(objNode.Text))
                Case UCase(Trim(cnstStrStatusRead))
                  .intStatus = rssItemStatusRead
                Case UCase(Trim(cnstStrStatusDeleted))
                  .intStatus = rssItemStatusDeleted
                Case Else
                  .intStatus = rssItemStatusNew
              End Select
            End If
          End With
        End If
      Next objNode
    'Libera la memoria
      Set objNode = Nothing
      Set objItem = Nothing
  Exit Function
  
errorLoad:
  Load = errorUnknown
End Function

Public Sub writeHTML(ByVal lngFile As Long, ByVal strKey As String)
'--> Escribe los elementos del RSS en el archivo HTML
Dim objItem As clsRSSItem
Dim arrStrKey() As String
Dim blnContinue As Boolean

  'Separa la clave en partes
    arrStrKey = Split(strKey, "¬")
  'Recorre los elementos
    For Each objItem In objColItems
      'Supone que debe escribir la información de este canal
        blnContinue = True
      'Comprueba si realmente debe escribir información de este canal
        If UBound(arrStrKey) > 0 Then
          If arrStrKey(2) <> objItem.strKey Then
            blnContinue = False
          End If
        End If
      'Si debe escribir información de este elemento
        If blnContinue Then
          With objItem
            If .intStatus <> rssItemStatusDeleted Then '... se salta los elementos eliminados
              'Cabecera del Div
                Print #lngFile, "<DIV class='itemList'>"
                Print #lngFile, "<p></p>"
                Print #lngFile, "<DIV class='itemTitle'>"
                If .intStatus = rssItemStatusNew Then
                  Print #lngFile, "<IMG src='Images/newFeedItem.gif'>"
                Else
                  Print #lngFile, "<IMG src='Images/readFeedItem.gif'>"
                End If
                Print #lngFile, "<A href='" & .strLink & "'>" & .strTitle & "</A>"
                Print #lngFile, "</DIV>"
                Print #lngFile, "<p></p>"
                Print #lngFile, "<DIV class='itemDescription'>"
                If .strContent <> "" Then
                  Print #lngFile, .strContent
                Else
                  Print #lngFile, .strDescription
                End If
                Print #lngFile, "</DIV>"
                Print #lngFile, "<p></p>"
                Print #lngFile, "<DIV class='itemDetails'>"
                Print #lngFile, objColLanguage.searchItem("colRSSItems", 1, "Categoría:") & "&nbsp;" & .strCategory & _
                                "&nbsp;&nbsp;&nbsp;" & _
                                objColLanguage.searchItem("colRSSItems", 2, "Asunto:") & "&nbsp;" & .strSubject & "<BR>"
                If .strAuthor <> "" Then
                  Print #lngFile, objColLanguage.searchItem("colRSSItems", 3, "Autor:") & "&nbsp;" & .strAuthor & "<BR>"
                End If
                If .strPubDate <> "" Then
                  Print #lngFile, objColLanguage.searchItem("colRSSItems", 4, "Fecha publicación:") & "&nbsp;" & _
                                  .strPubDate & "<BR>"
                End If
              'Fin del DIV del elemento
                Print #lngFile, "</DIV>"
                Print #lngFile, "<p></p>"
                Print #lngFile, "<HR>"
            End If
          End With
        End If
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
End Sub

Public Function getXML() As String
'--> Obtiene el XML de los canales
'<item>
'    <title>Laziness as a virtue</title>
'    <link>http://www.awasu.com/weblog/index.php?p=179</link>
'    <comments>http://www.awasu.com/weblog/index.php?p=179#comments</comments>
'    <author>azamsharp</author>
'    <pubDate>Tue, 24 Aug 2004 13:33:23 +0100</pubDate>
'    <category>General</category>
'    <subject>.NET</subject>
'    <guid isPermaLink='false'>179@http://www.awasu.com/weblog/index.php</guid>
'    <description><![CDATA[Continuing a trend,]]>
'    </description>
'    <content:encoded><![CDATA[aa]]></content:encoded>
'</item>
Dim objItem As clsRSSItem

  'Inicializa la variable de salida
    getXML = ""
  'Recorre los canales creando el XML
    For Each objItem In objColItems
      With objItem
        getXML = getXML & "<item>" & vbCrLf
        getXML = getXML & "    <title>" & getCData(.strTitle) & "</title>" & vbCrLf
        getXML = getXML & "    <link>" & getCData(.strLink) & "</link>" & vbCrLf
        getXML = getXML & "    <comments>" & getCData(.strComments) & "</comments>" & vbCrLf
        getXML = getXML & "    <author>" & getCData(.strAuthor) & "</author>" & vbCrLf
        getXML = getXML & "    <pubDate>" & getCData(.strPubDate) & "</pubDate>" & vbCrLf
        getXML = getXML & "    <category>" & getCData(.strCategory) & "</category>" & vbCrLf
        getXML = getXML & "    <subject>" & getCData(.strSubject) & "</subject>" & vbCrLf
        getXML = getXML & "    <guid isPermaLink='false'>" & getCData(.strGuid) & "</guid>" & vbCrLf
        getXML = getXML & "    <description>" & getCData(.strDescription) & "</description>" & vbCrLf
        If .strContent <> "" Then
          getXML = getXML & "    <content:encoded>" & getCData(.strContent) & "</content:encoded>" & vbCrLf
        End If
        getXML = getXML & "    <status>"
        Select Case .intStatus
          Case enumRSSItemStatus.rssItemStatusRead
            getXML = getXML & cnstStrStatusRead
          Case enumRSSItemStatus.rssItemStatusDeleted
            getXML = getXML & cnstStrStatusDeleted
          Case Else
            getXML = getXML & cnstStrStatusNew
        End Select
        getXML = getXML & "</status>" & vbCrLf
        getXML = getXML & "</item>" & vbCrLf
      End With
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
End Function

Public Property Get Item(varIndexKey As Variant) As clsRSSItem
Attribute Item.VB_UserMemId = 0
'--> Obtiene un elemento de la colección
  On Error GoTo errorItem
    Set Item = objColItems(varIndexKey)
  Exit Property
  
errorItem:
  Set Item = Nothing
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos de la colección
  If objColItems Is Nothing Then
    Count = 0
  Else
    Count = objColItems.Count
  End If
End Property

Public Sub Remove(ByVal varIndexKey As Variant)
'--> Elimina un elemento de la colección
  On Error Resume Next
    objColItems.Remove varIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Enumera los elementos de la colección
  Set NewEnum = objColItems.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Crea la colección al crear la clase
    Set objColItems = New Collection
End Sub

Private Sub Class_Terminate()
'--> Destruye la colección cuando se termina la clase
  Clear
  Set objColItems = Nothing
End Sub
