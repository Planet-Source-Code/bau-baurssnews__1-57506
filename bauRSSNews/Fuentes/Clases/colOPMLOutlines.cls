VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colOPMLOutlines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsOPMLOutline"
Attribute VB_Ext_KEY = "Member0" ,"clsOPMLOutline"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección de objeto Outline de OPML
'  <body>
'    <outline text="2020 Hindsight"
'             htmlUrl = "http://www.2020Hindsight.org/"
'             type="weblogsComFavorites"
'             url="http://static.userland.com/weblogMonitor/favorites/2020Hindsight.opml"/>
'  </body>
Option Explicit

'Variable local con los datos de la colección
Private objColOutlines As Collection

Public Sub Clear()
'--> Limpia la colección de elementos
Dim lngIndex As Long

  For lngIndex = Count To 1 Step -1
    Remove lngIndex
  Next lngIndex
End Sub

Public Function Add(Optional ByVal strKey As String = "") As clsOPMLOutline
'--> Añade un objeto a la colección
Dim objNewMember As New clsOPMLOutline
  
  'Añade el elemento a la colección
    If Trim(strKey) = "" Then
      strKey = getNextKey()
    End If
    objColOutlines.Add objNewMember, strKey
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
    While Not Item("OPMLOutline" & lngLastKey) Is Nothing
      'Pasa a la siguiente clave
        lngLastKey = lngLastKey + 1
    Wend
  'Devuelve la clave
    getNextKey = "OPMLOutline" & lngLastKey
End Function

Public Function Load(ByRef objDocument As MSXML2.IXMLDOMNode) As enumError
'--> Carga los datos de un elemento
'    <outline text="2020 Hindsight"
'             htmlUrl = "http://www.2020Hindsight.org/"
'             type="weblogsComFavorites"
'             url="http://static.userland.com/weblogMonitor/favorites/2020Hindsight.opml"/>
Dim objNode As MSXML2.IXMLDOMNode
Dim objAttribute As MSXML2.IXMLDOMAttribute
Dim objItem As clsOPMLOutline

  On Error GoTo errorLoad
    'Supone que puede cargar el archivo
      Load = errorOPMLNoError
    'Busca los datos del elemento
      For Each objNode In objDocument.childNodes
        If Load = errorOPMLNoError Then '... para qué vamos a continuar si hay algún error
          If objNode.baseName = "outline" Then
            'Añade un objeto a la colección
              Set objItem = Add(getNextKey())
            'Pasa los valores de los atributos al objeto OPML Outline
              With objItem
                'Recorre los atributos
                  For Each objAttribute In objNode.Attributes
                    If objAttribute.baseName = "text" Then
                      .strText = objAttribute.Text
                    ElseIf objAttribute.baseName = "htmlUrl" Then
                      .strLink = objAttribute.Text
                    ElseIf objAttribute.baseName = "type" Then
                      .strType = objAttribute.Text
                    ElseIf objAttribute.baseName = "url" Then
                      .strURL = objAttribute.Text
                    End If
                  Next objAttribute
                'Libera la memoria
                  Set objAttribute = Nothing
              End With
          End If
        End If
      Next objNode
    'Libera la memoria
      Set objNode = Nothing
      Set objItem = Nothing
  Exit Function
  
errorLoad:
  Load = errorOPMLUnknown
End Function

Public Sub writeHTML(ByVal lngFile As Long)
'--> Escribe los elementos del RSS en el archivo HTML
'Dim objItem As clsOPMLOutline
'
'  'Recorre los elementos
'    For Each objItem In objColOutlines
'      With objItem
'        If .intStatus <> OPMLOutlineStatusDeleted Then '... se salta los elementos eliminados
'          'Cabecera del Div
'            Print #lngFile, "<DIV class='itemList'>"
'            Print #lngFile, "<p></p>"
'            Print #lngFile, "<DIV class='itemTitle'>"
'            If .intStatus = OPMLOutlineStatusNew Then
'              Print #lngFile, "<IMG src='Images/newFeedItem.gif'>"
'            Else
'              Print #lngFile, "<IMG src='Images/readFeedItem.gif'>"
'            End If
'            Print #lngFile, "<A href='" & .strLink & "'>" & .strTitle & "</A>"
'            Print #lngFile, "</DIV>"
'            Print #lngFile, "<p></p>"
'            Print #lngFile, "<DIV class='itemDescription'>"
'            If .strContent <> "" Then
'              Print #lngFile, .strContent
'            Else
'              Print #lngFile, .strDescription
'            End If
'            Print #lngFile, "</DIV>"
'            Print #lngFile, "<p></p>"
'            Print #lngFile, "<DIV class='itemDetails'>"
'            Print #lngFile, objColLanguage.searchItem("colOPMLOutlines", 1, "Categoría:") & "&nbsp;" & .strCategory & _
'                            "&nbsp;&nbsp;&nbsp;" & _
'                            objColLanguage.searchItem("colOPMLOutlines", 2, "Asunto:") & "&nbsp;" & .strSubject & "<BR>"
'            If .strAuthor <> "" Then
'              Print #lngFile, objColLanguage.searchItem("colOPMLOutlines", 3, "Autor:") & "&nbsp;" & .strAuthor & "<BR>"
'            End If
'            If .strPubDate <> "" Then
'              Print #lngFile, objColLanguage.searchItem("colOPMLOutlines", 4, "Fecha publicación:") & "&nbsp;" & _
'                              .strPubDate & "<BR>"
'            End If
'          'Fin del DIV del elemento
'            Print #lngFile, "</DIV>"
'            Print #lngFile, "<p></p>"
'            Print #lngFile, "<HR>"
'        End If
'      End With
'    Next objItem
'  'Libera la memoria
'    Set objItem = Nothing
End Sub

Public Property Get Item(varIndexKey As Variant) As clsOPMLOutline
'--> Obtiene un elemento de la colección
  On Error GoTo errorItem
    Set Item = objColOutlines(varIndexKey)
  Exit Property
  
errorItem:
  Set Item = Nothing
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos de la colección
  If objColOutlines Is Nothing Then
    Count = 0
  Else
    Count = objColOutlines.Count
  End If
End Property

Public Sub Remove(ByVal varIndexKey As Variant)
'--> Elimina un elemento de la colección
  On Error Resume Next
    objColOutlines.Remove varIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Enumera los elementos de la colección
  Set NewEnum = objColOutlines.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Crea la colección al crear la clase
    Set objColOutlines = New Collection
End Sub

Private Sub Class_Terminate()
'--> Destruye la colección cuando se termina la clase
  Clear
  Set objColOutlines = Nothing
End Sub



