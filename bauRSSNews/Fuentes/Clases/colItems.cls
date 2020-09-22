VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsItem"
Attribute VB_Ext_KEY = "Member0" ,"clsItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección de elementos del proyecto
Option Explicit

'Variable local con los datos de la colección
Private objColItems As Collection

Public Function Add(ByVal intType As enumTypeItem, ByVal strTitle As String, ByVal strURL As String, _
                    ByVal strUser As String, ByVal strPassword As String, _
                    Optional ByVal strKey As String = "") As clsItem
'--> Añade un objeto a la colección
Dim objNewMember As clsItem

  'Crea un nuevo objeto
    Set objNewMember = New clsItem
  'Obtiene la clave
    If strKey = "" Then
      strKey = getNextKey()
    End If
  'Asigna las propiedades al objeto (debe estar antes porque la clave depende de la URL)
    With objNewMember
      .strKey = strKey
      .strTitle = strTitle
      .strURL = strURL
      .strUser = strUser
      .strPassword = strPassword
      .intType = intType
    End With
  'Añade el objeto a la colección
    objColItems.Add objNewMember, strKey
  'Devuelve el objeto creado y libera la memoria
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Private Function getNextKey() As String
'--> Obtiene la siguiente clave
  'Obtiene la siguiente clave
    lngProjectLastKey = lngProjectLastKey + 1
  'Devuelve la clave
    getNextKey = "RSS" & lngProjectLastKey
    Debug.Print getNextKey
End Function

Public Function Load(ByRef objDocument As MSXML2.IXMLDOMNode) As Boolean
'--> Carga los datos de un elemento RSS del proyecto
Dim objNode As MSXML2.IXMLDOMNode
Dim objDefinition As MSXML2.IXMLDOMNode

  On Error GoTo errorLoad
    'Supone que puede cargar el archivo
      Load = True
    'Busca los datos del elemento
      For Each objNode In objDocument.childNodes
        If Load Then
          If objNode.baseName = "opml" Or objNode.baseName = "folder" Or _
             objNode.baseName = "rss" Or objNode.baseName = "webPage" Then
              'Añade un objeto a la colección
                With Add(itemFolder, "", "", "", "")
                  'Cambia el tipo del elemento
                    Select Case objNode.baseName
                      Case "rss"
                        .intType = itemRSS
                      Case "folder"
                        .intType = itemFolder
                      Case "webPage"
                        .intType = itemWebPage
                    End Select
                  'Recorre los nodos buscando los datos
                    For Each objDefinition In objNode.childNodes
                      If objDefinition.baseName = "title" Then
                        .strTitle = objDefinition.Text
                      ElseIf objDefinition.baseName = "url" Then
                        .strURL = objDefinition.Text
                      ElseIf objDefinition.baseName = "user" Then
                        .strUser = objDefinition.Text
                      ElseIf objDefinition.baseName = "password" Then
                        .strPassword = objDefinition.Text
                      ElseIf objDefinition.baseName = "items" Then '... elementos de carpeta
                        If .intType = itemFolder Then '... carga los elementos
                          Load = .objColItems.Load(objDefinition)
                        End If
                      End If
                    Next objDefinition
                  'Libera la memoria
                    Set objDefinition = Nothing
                End With
          End If
        End If
      Next objNode
    'Libera la memoria
      Set objNode = Nothing
  Exit Function
  
errorLoad:
  Load = False
End Function

Public Sub Save(ByVal lngFile As Long)
'--> Graba los datos del archivo del proyecto
Dim objItem As clsItem
Dim strHeader As String

  'Recorre los elementos de la colección
    For Each objItem In objColItems
      With objItem
        If .strKey <> "" Then
          Select Case .intType
            Case enumTypeItem.itemFolder
              Print #lngFile, "  <folder>" & vbCrLf
              Print #lngFile, "    <title>" & getCData(.strTitle) & "</title>" & vbCrLf
              Print #lngFile, "    <items>" & vbCrLf
              .objColItems.Save lngFile
              Print #lngFile, "    </items>" & vbCrLf
              Print #lngFile, "  </folder>" & vbCrLf
            Case enumTypeItem.itemWebPage
              Print #lngFile, "  <webPage>" & vbCrLf
              Print #lngFile, "    <title>" & getCData(.strTitle) & "</title>" & vbCrLf
              Print #lngFile, "    <url>" & getCData(.strURL) & "</url>" & vbCrLf
              Print #lngFile, "  </webPage>" & vbCrLf
            Case Else 'OPML, RSS, ...
              'Obtiene la cabecera
                Select Case .intType
                  Case enumTypeItem.itemRSS
                    strHeader = "rss"
                End Select
              'Graba los datos
                Print #lngFile, "  <" & strHeader & ">" & vbCrLf
                Print #lngFile, "    <title>" & getCData(.strTitle) & "</title>" & vbCrLf
                Print #lngFile, "    <url>" & getCData(.strURL) & "</url>" & vbCrLf
                Print #lngFile, "    <user>" & getCData(.strUser) & "</user>" & vbCrLf
                Print #lngFile, "    <password>" & getCData(.strPassword) & "</password>" & vbCrLf
                Print #lngFile, "    <localFile>" & getCData(.LocalFile) & "</localFile>" & vbCrLf
                Print #lngFile, "  </" & strHeader & ">" & vbCrLf
          End Select
        End If
      End With
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
End Sub

Public Sub Clear()
'--> Limpia la colección
Dim lngIndex As Long

  For lngIndex = Count To 1 Step -1
    Remove lngIndex
  Next lngIndex
End Sub

Public Function Search(ByVal strKey As String) As clsItem
'--> Busca un elemento en la colección
Dim arrStrKey() As String
Dim objItem As clsItem
  
  'Si realmente hay algo que buscar ...
    If strKey <> "" Then
      'Inicializa el valor de salida
        Set Search = Nothing
      'Separa la clave en partes
        arrStrKey = Split(strKey, "¬")
      'Recorre la colección buscando el elemento
        For Each objItem In objColItems
          If Search Is Nothing Then '¿Para qué seguir buscando?
            If objItem.strKey = arrStrKey(0) Then
              Set Search = objItem
            ElseIf objItem.intType = itemFolder Then
              Set Search = objItem.objColItems.Search(strKey)
            End If
          End If
        Next objItem
    End If
  'Libera la memoria
    Set objItem = Nothing
End Function

Public Sub markReadRSS(ByVal strURL As String)
'--> Marca una noticia como leída cuando se pulsa sobre la URL
Dim objItem As clsItem

  'Recorre los elementos
    For Each objItem In objColItems
      Select Case objItem.intType
        Case enumTypeItem.itemFolder
          'Recorre la carpeta
            objItem.objColItems.markReadRSS strURL
        Case enumTypeItem.itemRSS
          'Marca los elementos como leídos
            If objItem.objRSS.markReadRSS(strURL) Then
              objItem.writeXML
            End If
      End Select
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
End Sub

Public Property Get Item(ByVal varIndexKey As Variant) As clsItem
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
