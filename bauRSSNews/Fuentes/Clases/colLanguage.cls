VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsLanguage"
Attribute VB_Ext_KEY = "Member0" ,"clsLanguage"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección con los datos del lenguaje
Option Explicit

Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
  (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private objColLanguage As Collection 'Variable local con la colección
Private strLastError As String 'Cadena con el último error
Private strFileName As String 'Nombre del archivo de idioma

Private Function loadIni(ByVal strFileName As String, ByVal strTopic As String, ByVal strSection As String, _
                         Optional ByVal strDefault As String = "@@", _
                         Optional ByVal strLen As Integer = 5000) As String
'--> Recubre a la rutina de la API GetPrivateProfileString
Dim strSalida As String

  strSalida = Space(strLen)
  GetPrivateProfileString strTopic, strSection, strDefault, strSalida, strLen, strFileName
  strSalida = Trim(strSalida)
  If strSalida <> "" Then
    strSalida = Left(strSalida, Len(strSalida) - 1)
  End If
  loadIni = strSalida
End Function

Public Function Add(ByVal strType As String, ByVal strName As String, ByVal strValue As String, _
                    Optional ByVal strKey As String = "") As clsLanguage
'--> Agrega un nuevo objeto a la colección
Dim objNewMember As clsLanguage

  'Crea un nuevo objeto
    Set objNewMember = New clsLanguage
  'Establece las propiedades que se transfieren al método
    With objNewMember
      .strType = strType
      .strName = strName
      .strValue = strValue
    End With
  'Agrega el objeto
    If Len(strKey) = 0 Then
      objColLanguage.Add objNewMember
    Else
      objColLanguage.Add objNewMember, strKey
    End If
  'Devuelve el objeto creado y libera la memoria
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Function initLanguage(ByVal strNewFileName As String, ByVal strKeyApp As String, _
                             ByVal strVersion As String) As Boolean
'--> Inicializa la colección con los datos del lenguaje
  On Error GoTo errorInit
    'Supone que no puede cargar
      initLanguage = False
    'Comprueba si es un archivo de lenguaje correcto
      If loadIni(strNewFileName, strKeyApp, "Version", "") <> strVersion Then
        strLastError = "No es un archivo de lenguaje correcto"
      Else
        strFileName = strNewFileName
        initLanguage = True
      End If
  Exit Function
  
errorInit:
  strLastError = Err.Description
End Function

Private Function loadSection(ByVal strSection As String)
'--> Carga la sección del lenguaje determinado
Dim intIndex As Integer
Dim strValue As String

  intIndex = 1
  Do
    strValue = loadIni(strFileName, strSection, "Title" & intIndex, "@@")
    If strValue <> "@@" Then
      Add UCase(Trim(strSection)), UCase(Trim("Title" & intIndex)), strValue
      intIndex = intIndex + 1
    End If
  Loop While strValue <> "@@"
End Function

Private Function existSection(ByVal strSection As String) As Boolean
'--> Comprueba si existe una sección
Dim objLanguage As clsLanguage

  'Supone que no existe
    existSection = False
  'Comprueba si existe
    For Each objLanguage In objColLanguage
      If objLanguage.strType = strSection Then
        existSection = True
      End If
    Next objLanguage
  'Libera la memoria
    Set objLanguage = Nothing
End Function

Public Function searchItem(ByVal strSection As String, ByVal intIndex As Integer, ByVal strDefault As String) As String
'--> Obtiene una cadena en el lenguage determinado
Dim objLanguage As clsLanguage
Dim strIndex As String

  'Supone que no lo puede encontrar
    searchItem = ""
  'Para ahorrar tiempo
    strIndex = UCase(Trim("Title" & intIndex))
    strSection = UCase(Trim(strSection))
  'Carga la sección si no existe
    If Not existSection(strSection) Then
      loadSection strSection
    End If
  'Busca el elemento
    For Each objLanguage In objColLanguage
      If objLanguage.strType = strSection And objLanguage.strName = strIndex Then
        searchItem = objLanguage.strValue
      End If
    Next objLanguage
  'Si no ha encontrado ningún valor, devuelve el valor por defecto
    If searchItem = "" Then
      searchItem = strDefault
    End If
  'Añade los saltos de línea
    searchItem = Replace(searchItem, "\n", vbCrLf)
  'Libera la memoria
    Set objLanguage = Nothing
End Function

Public Function searchItemValue(ByVal strSection As String, ByVal intIndex As Integer, ByVal strDefault As String, _
                                ParamArray arrStrParam()) As String
'--> Obtiene una cadena con parámetros
Dim intIndexArr As Integer

  'Obtiene la cadena
    searchItemValue = searchItem(strSection, intIndex, strDefault)
  'Cambia los parámetros
    For intIndexArr = LBound(arrStrParam) To UBound(arrStrParam)
      searchItemValue = Replace(searchItemValue, "%" & intIndexArr + 1, arrStrParam(intIndexArr))
    Next intIndexArr
End Function

Public Sub Clear()
'--> Elimina los elementos de la colección
Dim intIndex As Integer

  For intIndex = objColLanguage.Count To 1 Step -1
    objColLanguage.Remove intIndex
  Next intIndex
End Sub

Public Property Get Item(ByVal varIndexKey As Variant) As clsLanguage
Attribute Item.VB_UserMemId = 0
'--> Obtiene un elemento de la colección
  Set Item = objColLanguage(varIndexKey)
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos de la colección
  Count = objColLanguage.Count
End Property

Public Sub Remove(ByVal varIndexKey As Variant)
'--> Elimina un elemento de la colección
  objColLanguage.Remove varIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Permite enumera los elementos de la colección
  Set NewEnum = objColLanguage.[_NewEnum]
End Property

Public Property Get ErrorMessage() As String
'--> Obtiene el último mensaje de error (y lo limpia)
  ErrorMessage = strLastError
  strLastError = ""
End Property

Private Sub Class_Initialize()
'--> Al crea la colección crea la clase
  Set objColLanguage = New Collection
End Sub

Private Sub Class_Terminate()
'--> Elimina la colección al terminar con la clase
  Set objColLanguage = Nothing
End Sub
