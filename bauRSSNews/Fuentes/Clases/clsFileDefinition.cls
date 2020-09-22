VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Clase con la definición de un fichero OPML, RSS
Option Explicit

'Enumerados
Public Enum enumError 'Enumerados
  errorNoError = 0
  errorDownloadURL
  errorParse
  errorUnknown
End Enum

Public Enum enumTypeFile 'Tipo de archivo
  typeFileOPML = 0
  typeFileRSS
End Enum

'Variables públicas comunes a OPML y RSS
Public strKey As String
Public strTitle As String
Public strURL As String
Public strUser As String
Public strPassword As String

'Variables privadas
Private intType As enumTypeFile
Private strLocalFile As String
Private lngLastError As enumError

'Objetos privados
Private objOPML As clsOPML
Private objRSS As clsRSS

Private Sub Init()
'--> Inicializa las variables
  'Inicializa las variables
    strURL = ""
    strUser = ""
    strPassword = ""
  'Inicializa los objetos
    Set objOPML = New clsOPML
    Set objRSS = New clsRSS
  'Limpia los errores
    clearError
End Sub

Private Sub Destroy()
'--> Limpia la memoria
  Clear
  Set objOPML = Nothing
  Set objRSS = Nothing
End Sub

Public Sub Clear()
'--> Limpia los datos del archivo
  objOPML.Clear
  objRSS.Clear
End Sub

Private Sub clearError()
'--> Limpia los errores
  lngLastError = errorNoError
End Sub

Private Function getFileNameLocal() As String
'--> Obtiene el nombre del archivo local
  getFileNameLocal = App.Path & "\RSS\" & LocalFile
End Function

Public Function Load(Optional ByVal blnFromWeb As Boolean = False) As Boolean
'--> Carga un archivo
Dim strXML As String

  On Error GoTo errorLoad
    'Supone que no puede cargar el archivo
      Load = False
    'Limpia los errores
      clearError
    'Limpia los datos
      Clear
    'Carga el archivo
      If existFile(getFileNameLocal()) And Not blnFromWeb Then
        strXML = loadStringFromURL("file://" & getFileNameLocal(), strUser, strPassword)
      Else
        'Obtiene el archivo
          strXML = loadStringFromURL(strURL, strUser, strPassword)
        'Graba el RSS local la primera vez
          If Not existFile(getFileNameLocal()) Then '... cuando blnFromWeb = true, se carga de la web, pero aún no se debe grabar
            'Interpreta la cadena XML
              If parseXML(strXML) = errorNoError Then
                 '... Y guarda el archivo local
                  writeXML
                  Clear
                'Carga desde el archivo local (en ocasiones, no interpreta correctamente el archivo leído directamente de la web)
                  strXML = loadStringFromURL("file://" & getFileNameLocal(), strUser, strPassword)
              End If
          End If
      End If
    'Interpreta el documento XML
      lngLastError = parseXML(strXML)
    'Indica si se ha realizado una carga correcta
      Load = (lngLastError = errorNoError)
    'Si no se ha podido cargar y existe un archivo, lo elimina
      If Not Load Then
        If existFile(getFileNameLocal()) Then
          Kill getFileNameLocal()
        End If
      End If
  Exit Function
  
errorLoad:
  lngLastError = errorUnknown
End Function

Private Function parseXML(ByVal strXML As String) As enumError
'--> Interpreta la cadena XML
Dim objDocument As New MSXML2.DOMDocument30

  'Interpreta la cadena XML
    If strXML = "" Then
      parseXML = errorDownloadURL
    ElseIf Not objDocument.loadXML(strXML) Then
      parseXML = errorParse
    Else
      'Dependiendo del tipo de archivo
      Select Case intType
        Case enumTypeFile.typeFileOPML
          parseXML = objOPML.parseXML(objDocument)
        Case enumTypeFile.typeFileRSS
          parseXML = objRSS.parseXML(objDocument)
      End Select
  'Libera la memoria
    Set objDocument = Nothing
End Function

Private Function loadStringFromURL(ByVal strURL As String, ByVal strUser As String, ByVal strPassword As String) As String
Dim objDocument As New MSXML2.DOMDocument30
Dim objXMLHTTP As New MSXML2.XMLHTTP

  On Error GoTo errorLoad
    'Supone que no puede descargar el archivo a partir de la URL
      loadStringFromURL = ""
    'Si es un archivo local ...
      If UCase(Left(strURL, 7)) = "FILE://" Then
        'Lee el archivo
          If objDocument.Load(strURL) Then
            loadStringFromURL = objDocument.xml
          End If
      Else
        'Inicializa los datos para cargar la URL (Async = False)
          objXMLHTTP.open "GET", Trim(strURL), False, strUser, strPassword
        'Envía la llamada
          objXMLHTTP.send
        'Obtiene la cadena leída
          loadStringFromURL = objXMLHTTP.responseText
      End If
    'Libera la memoria
      Set objXMLHTTP = Nothing
      Set objDocument = Nothing
  Exit Function
  
errorLoad:
End Function

Public Function Merge() As Boolean
'--> Mezcla el archivo local con el archivo de la Web
  'Si no existe el archivo, carga y graba
    If Not existFile(getFileNameLocal()) Then
      Load
    Else '... si existe
      'Dependiendo del tipo de archivo
        Select Case intType
          Case enumTypeFile.typeFileOPML
            Merge = objOPML.Merge()
          Case enumTypeFile.typeFileRSS
            Merge = objRSS.Merge()
        End Select
      'Si se han realizado modificaciones, se graba el XML
        If Merge Then
          writeXML
          Merge = Not blnUpdateOnlyHeader
        End If
      'Limpia el objeto actual
        Clear
    End If
End Function

Public Function writeHTML(ByVal lngFile As Long) As Boolean
'--> Crea el archivo HTML
  On Error GoTo errorWrite
    'Supone que no puede escribir el archivo
      writeHTML = False
    'Cabecera HTML
      Print #lngFile, "<html>"
      Print #lngFile, "<head>"
      Print #lngFile, "<title>" & strTitle & "</title>"
      'Print #lngFile, "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=utf-8'>"
    'Escribe el estilo
      writeHTMLStyle lngFile
      Print #lngFile, "</head>"
      Print #lngFile, "<body>"
    'Cuerpo dependiendo del tipo
      Select Case intType
        Case enumTypeFile.typeFileOPML
          writeHTML = objOPML.writeHTML(lngifle)
        Case enumTypeFile.typeFileRSS
          writeHTML = objOPML.writeHTML(lngFile)
    'Cierre HTML
      Print #lngFile, "</body>"
      Print #lngFile, "</html>"
    'Si ha llegado hasta aquí es porque todo ha ido bien
      writeHTML = True
  Exit Function
  
errorWrite:
End Function

Private Sub writeHTMLStyle(ByVal lngFile As Long)
'--> Escribe el estilo de la página (se saca de la rutina principal para que no sea demasiado grande)
  Print #lngFile, "<STYLE>"
  Print #lngFile, "    /* --- STYLESHEET BEGINS HERE --- */"
  Print #lngFile, "    /* Define the general appearance of the page */"
  Print #lngFile, "    BODY"
  Print #lngFile, "    {"
  Print #lngFile, "        background-attachment: fixed;"
  Print #lngFile, "        background-position: 97% bottom;" 'top;"
  Print #lngFile, "        background-image: url(Images/logoBackGround.gif);"
  Print #lngFile, "        background-repeat: no-repeat;"
  Print #lngFile, "        padding: 5px;"
  Print #lngFile, "        font-family: Georgia;"
  Print #lngFile, "        font-size: 80%;"
  Print #lngFile, "    }"
  Print #lngFile, "    HR"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #993300;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the general appearance of the banner */"
  Print #lngFile, "    .banner"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #FFF;"
  Print #lngFile, "        background: #993300;"
  Print #lngFile, "        border: 3px double;"
  Print #lngFile, "        padding: 10px;"
  Print #lngFile, "    }"
  Print #lngFile, "    .banner A:link,"
  Print #lngFile, "    .banner A: visited"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #FFF;"
  Print #lngFile, "        text-decoration: none;"
  Print #lngFile, "    }"
  Print #lngFile, "    .banner :hover,"
  Print #lngFile, "    .banner A: active"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #FFF;"
  Print #lngFile, "        text-decoration: underline;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the channel name in the banner */"
  Print #lngFile, "    .channelName"
  Print #lngFile, "    {"
  Print #lngFile, "        font-variant: small-caps;"
  Print #lngFile, "        font-size: 150%;"
  Print #lngFile, "    }"
  Print #lngFile, "    .channelName A"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #FFF;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the channel description in the banner */"
  Print #lngFile, "    .channelDescription"
  Print #lngFile, "    {"
  Print #lngFile, "        font-size: 125%;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the channel details in the banner."
  Print #lngFile, "     * NOTE: This information will only appear if it has been supplied by"
  Print #lngFile, "     *       the channel publisher."
  Print #lngFile, "     */"
  Print #lngFile, "    .channelDetails"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #FFF;"
  Print #lngFile, "        margin-left: 20px;"
  Print #lngFile, "        margin-right: 20px;"
  Print #lngFile, "        margin-bottom: 10px;"
  Print #lngFile, "        font-style: italic;"
  Print #lngFile, "    }"
  Print #lngFile, "    .channelDetails A"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #FFF;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the info in the banner */"
  Print #lngFile, "    .feedItemCountInfo"
  Print #lngFile, "    {"
  Print #lngFile, "        font-style: italic;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the Last Updated time in the banner */"
  Print #lngFile, "    .lastUpdateTime"
  Print #lngFile, "    {"
  Print #lngFile, "        font-style: italic;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the general appearance of the feed item list */"
  Print #lngFile, "    .itemList"
  Print #lngFile, "    {"
  Print #lngFile, "        font-family: Bookman;"
  Print #lngFile, "        color: #330033;"
  Print #lngFile, "    }"
  Print #lngFile, "    .itemList A: link"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #CC3300;"
  Print #lngFile, "        font-weight: bold;"
  Print #lngFile, "        text-decoration: none;"
  Print #lngFile, "    }"
  Print #lngFile, "    .itemList A: visited"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #CC3300;"
  Print #lngFile, "        font-weight: normal;"
  Print #lngFile, "        text-decoration: none;"
  Print #lngFile, "    }"
  Print #lngFile, "    .itemList A:active,"
  Print #lngFile, "    .itemList A: hover"
  Print #lngFile, "    {"
  Print #lngFile, "       color: #FFF;"
  Print #lngFile, "       background: #CC3300;"
  Print #lngFile, "       padding: 2px;"
  Print #lngFile, "       text-decoration: none;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the feed item titles */"
  Print #lngFile, "    .itemTitle"
  Print #lngFile, "    {"
  Print #lngFile, "        font-size: 125%;"
  Print #lngFile, "    }"
  Print #lngFile, "    .itemTitle A"
  Print #lngFile, "    {"
  Print #lngFile, "        text-decoration: none;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the body of each feed item */"
  Print #lngFile, "    .itemDescription"
  Print #lngFile, "    {"
  Print #lngFile, "        margin-left: 20px;"
  Print #lngFile, "        margin-right: 20px;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of any other information attached to each feed item */"
  Print #lngFile, "    .itemDetails"
  Print #lngFile, "    {"
  Print #lngFile, "        font-style: italic;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* --- STYLESHEET ENDS HERE --- */"
  Print #lngFile, "</STYLE>"
End Sub

Public Sub writeXML()
'--> Graba el XML de este objeto
Dim strXML As String

  'Obtiene el XML
    'Cabecera
      strXML = "<?xml version='1.0'  encoding='iso-8859-1' ?>" & vbCrLf
      strXML = strXML & " <!-- generator='" & App.Title & "/" & App.Major & "." & _
                        App.Minor & "." & App.Revision & "' -->" & vbCrLf
    'Contenido
      Select Case intType
        Case enumTypeFile.typeFileOPML
          strXML = strXML & objOPML.getXML()
        Case enumTypeFile.typeFileRSS
      End Select
  'Graba el XML
    saveLocalFile strXML
End Sub

Private Sub saveLocalFile(ByVal strXML As String)
'--> Graba el archivo local
Dim lngFile As Long

  On Error GoTo errorSave
    'Abre el archivo
      lngFile = FreeFile()
      Open getFileNameLocal() For Output As #lngFile
    'Graba el XML
      Print #lngFile, strXML
    'Cierra el archivo
      Close #lngFile
  Exit Sub
  
errorSave:
End Sub

Private Function normalizeFileName(ByVal strFileName As String) As String
'--> Quita los valores extraños a un nombre de archivo
Dim lngIndex As Long
Dim strChar As String

  'Inicializa la variable de salida
    normalizeFileName = ""
  'Recorre el nombre de archivo quitando los caracteres extraños
    For lngIndex = 1 To Len(strFileName)
      'Coge el carácter 'n'
        strChar = Mid(strFileName, lngIndex, 1)
      'Si es un carácter 'normal' lo añade a la cadena de salida
        If (strChar >= "A" And strChar <= "Z") Or (strChar >= "a" And strChar <= "z") Or _
           (strChar >= "0" And strChar <= "9") Or strChar = "." Then
          normalizeFileName = normalizeFileName & strChar
        End If
    Next lngIndex
End Function

Private Function getExtension() As String
'--> Obtiene la extensión del archivo
  Select Case intType
    Case enumTypeFile.typeFileOPML
      getExtension = "OPML"
    Case enumTypeFile.typeFileRSS
      getExtension = "RSS"
  End Select
End Function

Public Property Get LocalFile() As String
  If strLocalFile = "" Then
    LocalFile = strKey & "_" & normalizeFileName(strURL) & "." & getExtension()
  Else
    LocalFile = strLocalFile
  End If
End Property

Public Property Get LastError() As enumError
  LastError = lngLastError
End Property

Private Sub Class_Initialize()
  Init
End Sub

Private Sub Class_Terminate()
  Destroy
End Sub
