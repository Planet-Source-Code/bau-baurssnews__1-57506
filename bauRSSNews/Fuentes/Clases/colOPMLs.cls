VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colOPMLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsOPML"
Attribute VB_Ext_KEY = "Member0" ,"clsOPML"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección de archivos OPML asociados al proyecto
Option Explicit

'Variable local con los elementos de la colección
Private objColOPMLs As Collection

Public Function Add(ByVal strTitle As String, ByVal strURL As String, _
                    ByVal strUser As String, ByVal strPassword As String, _
                    Optional ByVal strKey As String = "") As clsOPML
'--> Añade un nuevo objeto a la colección
Dim objNewMember As New clsOPML
  
  'Pasa las propiedades al objeto
    With objNewMember
      .strTitle = strTitle
      .strURL = strURL
      .strUser = strUser
      .strPassword = strPassword
    End With
  'Añade el objeto a la colección
    If strKey = "" Then
      strKey = getNextKey()
    End If
    objColOPMLs.Add objNewMember, strKey
  'Asigna la clave al objeto
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
    While Not Item("OPML" & lngLastKey) Is Nothing
      'Pasa a la siguiente clave
        lngLastKey = lngLastKey + 1
    Wend
  'Devuelve la clave
    getNextKey = "OPML" & lngLastKey
End Function

Public Sub Clear()
'--> Limpia la memoria
Dim lngIndex As Long

  For lngIndex = Count To 1 Step -1
    Remove lngIndex
  Next lngIndex
End Sub

Public Property Get Item(ByVal varIndexKey As Variant) As clsOPML
Attribute Item.VB_UserMemId = 0
'--> Obtiene un elemento de la colección
  On Error GoTo errorItem
    Set Item = objColOPMLs(varIndexKey)
  Exit Property
  
errorItem:
  Set Item = Nothing
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos de la colección
  If objColOPMLs Is Nothing Then
    Count = 0
  Else
    Count = objColOPMLs.Count
  End If
End Property

Public Sub Remove(ByVal varIndexKey As Variant)
'--> Elimina un elemento de la colección
  On Error Resume Next
    objColOPMLs.Remove varIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Enumera los elementos de la colección
  Set NewEnum = objColOPMLs.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Crea la colección cuando se crea la clase
  Set objColOPMLs = New Collection
End Sub

Private Sub Class_Terminate()
'--> Destruye la colección cuando se termina la clase
  Clear
  Set objColOPMLs = Nothing
End Sub
