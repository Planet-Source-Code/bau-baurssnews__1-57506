VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colRSSChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsRSSChannel"
Attribute VB_Ext_KEY = "Member0" ,"clsRSSChannel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección de canales RSS
Option Explicit

'Variable local con los datos de la colección
Private objColChannels As Collection

Public Sub Clear()
'--> Limpia la colección de canales
Dim lngIndex As Long

  For lngIndex = Count To 1 Step -1
    Remove lngIndex
  Next lngIndex
End Sub

Public Function Add(Optional ByVal strKey As String = "") As clsRSSChannel
'--> Añade un canal a la colección
Dim objNewMember As New clsRSSChannel
  
  'Añade el objeto a la colección
    If Trim(strKey) = "" Then
      strKey = getNextKey()
    End If
    objColChannels.Add objNewMember, strKey
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
    While Not Item("RSSChannel" & lngLastKey) Is Nothing
      'Pasa a la siguiente clave
        lngLastKey = lngLastKey + 1
    Wend
  'Devuelve la clave
    getNextKey = "RSSChannel" & lngLastKey
End Function

Public Function Load(ByRef objDocument As MSXML2.IXMLDOMNode) As enumError
'--> Carga los datos de un canal
'<channel>
'   <title>Yahoo! News - Oddly Enough</title>
'   <link>http://news.yahoo.com/news?tmpl=index&amp;cid=757</link>
'   <description>Yahoo! News - Oddly Enough</description>
'   <language>en-us</language>
'   <lastBuildDate>Thu, 26 Aug 2004 09:47:30 GMT</lastBuildDate>
'   <ttl>5</ttl>
'   <image>
'      ...
'   </image>
'   <item>
'      ...
'   </item>
'</channel>
Dim objNode As MSXML2.IXMLDOMNode, objNodeItem As MSXML2.IXMLDOMNode
Dim objChannel As clsRSSChannel

  On Error GoTo errorLoad
    'Supone que puede cargar el archivo
      Load = errorNoError
    'Añade un objeto a la colección
      Set objChannel = Add(getNextKey())
    'Busca los datos del canal
      For Each objNode In objDocument.childNodes
        If Load = errorNoError Then '... para qué vamos a continuar si hay algún error
          If objNode.baseName = "title" Then
            objChannel.strTitle = objNode.Text
          ElseIf objNode.baseName = "link" Then
            objChannel.strURL = objNode.Text
          ElseIf objNode.baseName = "description" Then
            objChannel.strDescription = objNode.Text
          ElseIf objNode.baseName = "language" Then
            objChannel.strLanguage = objNode.Text
          ElseIf objNode.baseName = "copyright" Then
            objChannel.strCopyright = objNode.Text
          ElseIf objNode.baseName = "lastBuildDate" Then
            objChannel.strLastBuildDate = objNode.Text
          ElseIf objNode.baseName = "pubDate" Then
            objChannel.strPubDate = objNode.Text
          ElseIf objNode.baseName = "generator" Then
            objChannel.strGenerator = objNode.Text
          ElseIf objNode.baseName = "ttl" Then
            objChannel.strTtl = objNode.Text
          ElseIf objNode.baseName = "image" Then
            Load = objChannel.objImage.Load(objNode)
          ElseIf objNode.baseName = "item" Then
            Load = objChannel.objColItems.Load(objNode)
          End If
        End If
      Next objNode
    'Libera la memoria
      Set objNode = Nothing
      Set objChannel = Nothing
  Exit Function
  
errorLoad:
  Load = errorUnknown
End Function

Public Sub writeHTML(ByVal lngFile As Long, ByVal strKey As String)
'--> Escribe el archivo HTML
Dim objChannel As clsRSSChannel
Dim strTemp As String
Dim arrStrKey() As String
Dim blnContinue As Boolean

  'Separa la clave en partes
    arrStrKey = Split(strKey, "¬")
  'Recorre los canales
    For Each objChannel In objColChannels
      'Supone que debe escribir la información de este canal
        blnContinue = True
      'Comprueba si realmente debe escribir información de este canal
        If UBound(arrStrKey) > 0 Then
          If arrStrKey(1) <> objChannel.strKey Then
            blnContinue = False
          End If
        End If
      'Si debe escribir la información de este canal
        If blnContinue Then
          With objChannel
            'Div de cabecera
              Print #lngFile, "<DIV class='banner'>"
            'Imagen
              If .objImage.strURL <> "" Then
                strTemp = "<A href='" & .objImage.strLink & "'>"
                strTemp = strTemp & "<IMG src='" & .objImage.strURL & "'"
                If IsNumeric(.objImage.strWidth) Then
                  strTemp = strTemp & " width='" & .objImage.strWidth & "'"
                End If
                If IsNumeric(.objImage.strHeight) Then
                  strTemp = strTemp & " height='" & .objImage.strHeight & "'"
                End If
                strTemp = strTemp & " border='0' alt='" & .objImage.strTitle & "'>"
                strTemp = strTemp & "</A>"
                Print #lngFile, strTemp
              End If
            'Título y fecha de última modificación
              Print #lngFile, "<SPAN class='channelName'>"
              Print #lngFile, "<A class='channelName' href='" & .strURL & "'>" & .strTitle & "</A>"
              Print #lngFile, "</SPAN>"
              Print #lngFile, "&nbsp;"
              Print #lngFile, "<SPAN class='lastUpdateTime'>(" & _
                              objColLanguage.searchItem("colRSSChannels", 1, "Ultima modificación:") & _
                              "&nbsp;" & IIf(.strLastBuildDate = "", .strPubDate, .strPubDate) & ")</SPAN><BR>"
              Print #lngFile, "<SPAN class='feedItemCountInfo'></SPAN>"
            'Descripción del canal
              Print #lngFile, "<p></p>"
              Print #lngFile, "<DIV class='channelDescription'>"
              Print #lngFile, .strDescription
              Print #lngFile, "</DIV>"
            'Copyright
              Print #lngFile, "<p></p>"
              Print #lngFile, "<DIV class='channelDescription'>"
              Print #lngFile, objColLanguage.searchItem("colRSSChannels", 2, "Copyright:") & "&nbsp;" & .strCopyright
              Print #lngFile, "</DIV>"
            'Fin del div de cabecera
              Print #lngFile, "</DIV>"
            'Noticias
              If Not .objColItems Is Nothing Then
                .objColItems.writeHTML lngFile, strKey
              End If
          End With
        End If
    Next objChannel
  'Libera la memoria
    Set objChannel = Nothing
End Sub

Public Function getXML() As String
'--> Obtiene el XML de los canales
'<channel>
'  <title>Awasu</title>
'  <link>http://www.awasu.com/weblog/index.php</link>
'  <description>The Wonderful World of Awasu</description>
'  <language>en</language>
'  <copyright>Copyright 2004</copyright>
'  <lastBuildDate>Thu, 26 Aug 2004 10:02:40 +0000</lastBuildDate>
'  <pubDate>Thu, 26 Aug 2004 10:02:40 +0000</pubDate>
'  <generator>http://wordpress.org/?v=1.0.1</generator>
'  <image>
'    <title>Awasu</title>
'    <url>http://www.awasu.com/images/logo-small.jpeg</url>
'    <link>http://www.awasu.com/weblog/index.php</link>
'  </image>
'  <item>
'  </item>
'</channel>
Dim objChannel As clsRSSChannel

  'Inicializa la variable de salida
    getXML = ""
  'Recorre los canales creando el XML
    For Each objChannel In objColChannels
      With objChannel
        getXML = getXML & "<channel>" & vbCrLf
        getXML = getXML & "  <title>" & getCData(.strTitle) & "</title>" & vbCrLf
        getXML = getXML & "  <link>" & getCData(.strURL) & "</link>" & vbCrLf
        getXML = getXML & "  <description>" & getCData(.strDescription) & "</description>" & vbCrLf
        getXML = getXML & "  <language>" & getCData(.strLanguage) & "</language>" & vbCrLf
        getXML = getXML & "  <copyright>" & getCData(.strCopyright) & "</copyright>" & vbCrLf
        getXML = getXML & "  <lastBuildDate>" & getCData(.strLastBuildDate) & "</lastBuildDate>" & vbCrLf
        getXML = getXML & "  <pubDate>" & getCData(.strPubDate) & "</pubDate>" & vbCrLf
        getXML = getXML & "  <generator>" & getCData(.strGenerator) & "</generator>" & vbCrLf
        getXML = getXML & "  <ttl>" & getCData(.strTtl) & "</ttl>" & vbCrLf
        If Not .objImage Is Nothing Then
          getXML = getXML & "  <image>" & vbCrLf
          getXML = getXML & "    <title>" & getCData(.objImage.strTitle) & "</title>" & vbCrLf
          getXML = getXML & "    <url>" & getCData(.objImage.strURL) & "</url>" & vbCrLf
          getXML = getXML & "    <link>" & getCData(.objImage.strLink) & "</link>" & vbCrLf
          getXML = getXML & "    <width>" & getCData(.objImage.strWidth) & "</width>" & vbCrLf
          getXML = getXML & "    <height>" & getCData(.objImage.strHeight) & "</height>" & vbCrLf
          getXML = getXML & "  </image>" & vbCrLf
        End If
        If Not .objColItems Is Nothing Then
          getXML = getXML & .objColItems.getXML()
        End If
        getXML = getXML & "</channel>" & vbCrLf
      End With
    Next objChannel
  'Libera la memoria
    Set objChannel = Nothing
End Function

Public Property Get Item(ByVal varIndexKey As Variant) As clsRSSChannel
Attribute Item.VB_UserMemId = 0
'--> Obtiene un elemento de la colección
  On Error GoTo errorItem
    Set Item = objColChannels(varIndexKey)
  Exit Property
  
errorItem:
  Set Item = Nothing
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos de la colección
  If objColChannels Is Nothing Then
    Count = 0
  Else
    Count = objColChannels.Count
  End If
End Property

Public Sub Remove(ByVal varIndexKey As Variant)
'--> Elimina un elemento de la coleeción
  On Error Resume Next
    objColChannels.Remove varIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Enumera los elementos de la colección
  Set NewEnum = objColChannels.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Al crea la clase crea la colección
  Set objColChannels = New Collection
End Sub

Private Sub Class_Terminate()
'--> Al destruir la clase destruye la colección
  Clear
  Set objColChannels = Nothing
End Sub
