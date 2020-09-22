VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con la definición de un fichero OPML, RSS
Option Explicit

'Enumerados
Public Enum enumError 'Enumerados
  errorNoError = 0
  errorDownloadURL
  errorParse
  errorUnknown
End Enum

Public Enum enumTypeItem 'Tipo de archivo
  itemFolder = 0
  itemRSS
  itemWebPage
End Enum

'Variables públicas comunes a OPML y RSS
Public strKey As String
Public intType As enumTypeItem
Public strTitle As String
Public strURL As String
Public strUser As String
Public strPassword As String

'Variables públicas relacionadas con OPML
Public strWeb As String
Public strType As String

'Objetos públicos
Public objColItems As colItems
Public objRSS As clsRSS

'Variables privadas
Private strLocalFile As String
Private strLastError As String
Private lngLastError As enumError

Private Sub Init()
'--> Inicializa las variables
  'Inicializa las variables
    strURL = ""
    strUser = ""
    strPassword = ""
  'Inicializa los objetos
    Set objColItems = New colItems
    Set objRSS = New clsRSS
  'Limpia los errores
    clearError
End Sub

Private Sub Destroy()
'--> Limpia la memoria
  Clear
  Set objRSS = Nothing
  Set objColItems = Nothing
End Sub

Public Sub Clear()
'--> Limpia los datos
  objColItems.Clear
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

Public Function countNewItems() As Long
'--> Cuenta los elementos no leídos
  'Inicializa la variable de salida
    countNewItems = 0
  'Si es un RSS
    If intType = itemRSS Then
      'If Load(False) Then
        countNewItems = objRSS.countNewItems()
      'End If
    End If
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
        Case enumTypeItem.itemRSS
          parseXML = objRSS.parseXML(objDocument)
      End Select
    End If
  'Libera la memoria
    Set objDocument = Nothing
End Function

Private Function downloadFileHTTP(ByVal strURL As String, ByVal strDestination As String, _
                                  ByVal strUser As String, ByVal strPassword As String, _
                                  ByRef dtmLastUpdate As Date) As Boolean
'--> Descarga un archivo desde un servidor HTTP
Const cnstLngChunkSize As Long = 1024 'Tamaño del trozo de imagen se descarga de Internet
Const cnstLngRollBack As Long = 4096 'Bytes para deshacer en caso de error. Se puede ser menos conservador y deshacer menos bytes pero no lo recomiendo
Dim bytArrData() As Byte 'Variable con los datos
Dim blnResume As Boolean 'True si se deshace la descarga
Dim blnCancel As Boolean 'Para evitar un goto, se cancela la descarga cuando se reciben ciertas cabeceras
Dim lngFile As Integer 'Identificador de archivo
Dim lngBytesReceived As Long 'Bytes recibidos hasta ahora
Dim lngFileLength As Long 'Longitud total del archivo en bytes
Dim strHeader As String 'Cabecera HTTP
Dim intRetry As Integer 'Número de reintentos hasta el momento

  On Error GoTo errorDownload
    'Supone que no se puede descargar el archivo
      downloadFileHTTP = False
    'Indica que no se ha reintentado nunca
      intRetry = 0
    'Cancela la petición actual
      While frmMain.inetTransfer.StillExecuting
        frmMain.inetTransfer.Cancel
        DoEvents
      Wend
    'Entra en un bucle para intentar descargar la página 10 veces
      Do
        'Actualiza el número de bytes recibidos
          If blnResume Then
            lngBytesReceived = lngBytesReceived - cnstLngRollBack
            If lngBytesReceived < 0 Then
              lngBytesReceived = 0
            End If
          End If
        'Da tiempo a Windows
          DoEvents
        'Descarga el archivo
          With frmMain.inetTransfer
              .URL = strURL
              .UserName = strUser
              .Password = strPassword
              If blnResume Then
                'GET del archivo, envía la cabecera mágica de continuar
                  .Execute , "GET", , "Range: bytes=" & lngBytesReceived & "-" & vbCrLf
              Else
                'GET estándar
                  .Execute , "GET"
              End If
          End With
        'Mientras se abre la conexión, pasa los mensajes a Windows
          While frmMain.inetTransfer.StillExecuting
            DoEvents
          Wend
        'Comprueba si se debe transferir
      '    If dtmLastUpdate < CDate(inetTransfer.GetHeader("Last-modified")) Then
      '      inetTransfer.Cancel
      '      GoTo ExitDownload
      '    Else
      '      dtmLastUpdate = inetTransfer.GetHeader("Last-modified")
      '    End If
        'Obtiene la primera cabecera ("HTTP/X.X XXX ...")
          strHeader = frmMain.inetTransfer.GetHeader
        'Atrapa los códigos de respuesta HTTP habituales
          blnCancel = False 'Supone que todo irá bien
          Select Case Mid(strHeader, 10, 3)
            Case "200" 'OK
              'Si se está corrigiendo un error anterior, sin embargo, esto es un fallo
                If blnResume Then
                  'Elimina el archivo parcialmente descargado
                    killFile strDestination
                  'Indica que se empieza de nuevo
                    blnResume = False
                End If
            Case "206" '206=Contenido parcial (cuando estamos corrigiendo errores)
            Case "204" 'Sin contenido
              strLastError = "Sin contenido"
              blnCancel = True
            Case "401" 'No autorizado
              strLastError = "No autorizado"
              blnCancel = True
            Case "404" 'No se encuentra el archivo
              strLastError = "Archivo desconocido"
              blnCancel = True
            Case vbCrLf 'Sin cabecera
              strLastError = "No se encuentra la cabecera de archivo"
              blnCancel = True
            Case Else 'Otros errores
              strLastError = Left(strHeader, InStr(strHeader, vbCr))
              blnCancel = True
          End Select
        'Obtiene la longitud de archivo de "Content-Length"
        'Al ser un archivo XML no tiene content-length
'          If Not blnResume And Not blnCancel Then
'            'Comprueba el tamaño del archivo
'              strHeader = frmMain.inetTransfer.GetHeader("Content-Length")
'              lngFileLength = Val(strHeader)
'              If lngFileLength = 0 Then
'                strLastError = "Archivo vacío"
'                blnCancel = True
'              End If
'          End If
        'Si en la cabecera no nos ha devuelto ningún error
          If Not blnCancel Then
            'Actualiza el sistema
              DoEvents
            'Inicializa los bytes recibidos si no se está corrigiendo ningún error
              If Not blnResume Then
                lngBytesReceived = 0
              End If
            'Si no ha ocurrido ningún error, pasa el archivo a disco
              lngFile = FreeFile() 'Obtiene un handle a archivo libre
              Open strDestination For Binary Access Write As #lngFile
            'Si se está corrigiendo un error, busca el byte en el archivo descargado donde se quedó la última vez
              If blnResume Then
                Seek #lngFile, lngBytesReceived + 1
              End If
            'Obtiene los trozos de archivo
              Do
                'Obtiene el trozo
                  bytArrData = frmMain.inetTransfer.GetChunk(cnstLngChunkSize, icByteArray)
                'Lo guarda en el archivo
                  Put #lngFile, , bytArrData
                'Ajusta el número de bytes recibidos
                  lngBytesReceived = lngBytesReceived + UBound(bytArrData, 1) + 1
                'Ajusta la barra de progreso
                  DoEvents
              Loop While UBound(bytArrData, 1) > 0 ' ... mientras queden datos
            'Cierra el archivo
              Close #lngFile
          End If
        'Es correcto si el número de bytes recibidos = content length
'          If lngBytesReceived = lngFileLength Then
'            downloadFileHTTP = True
'          Else ' ... no ha descargado
'            blnResume = (intRetry < 10) '... sólo lo reintenta 10 veces
'            intRetry = intRetry + 1 '... incrementa el número de reintentos (siempre, aunque blnResume = false)
'          End If
        'Al ser un archivo XML no tiene content-lengt
          downloadFileHTTP = True
      Loop While Not downloadFileHTTP And blnResume
  'No tiene exit Function porque siempre debe cerrar la conexión a Internet

errorDownload:
  'Cierra la conexión a Internet
    frmMain.inetTransfer.Cancel
End Function

Private Function loadStringFromURL(ByVal strURL As String, ByVal strUser As String, ByVal strPassword As String) As String
'--> Carga una cadena a partir de una URL local o Internet
Dim strFileName As String
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
        'Obtiene el nombre del archivo local
          strFileName = getFileNameLocal()
        'Descarga el archivo HTTP
          If downloadFileHTTP(strURL, strFileName, strUser, strPassword, Now()) Then
            'Lee el archivo
              If objDocument.Load("FILE://" & strFileName) Then
                loadStringFromURL = objDocument.xml
              End If
          End If
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
      'Espera que termine la última transmisión
        While frmMain.inetTransfer.StillExecuting
          DoEvents
        Wend
      'Carga los datos
        Load
      'Dependiendo del tipo de archivo
        Select Case intType
          Case enumTypeItem.itemRSS
            Merge = objRSS.Merge(strURL, strKey)
        End Select
      'Si se han realizado modificaciones, se graba el XML
        If Merge Then
          writeXML
        End If
      'Limpia el objeto actual
        Clear
    End If
End Function

Public Function writeHTML(ByVal lngFile As Long, ByVal strKey As String) As Boolean
'--> Crea el archivo HTML
Dim objItem As clsItem

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
        Case enumTypeItem.itemFolder
          'Recorre los elementos de la carpeta
            For Each objItem In objColItems
              Select Case objItem.intType
                Case enumTypeItem.itemRSS
                  objItem.Load False
                  objItem.objRSS.writeHTML lngFile, strKey
              End Select
            Next objItem
          'Libera la memoria
            Set objItem = Nothing
        Case enumTypeItem.itemRSS
          writeHTML = objRSS.writeHTML(lngFile, strKey)
      End Select
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
  Print #lngFile, "        font-size: 125%;"
  Print #lngFile, "    }"
  Print #lngFile, "    .channelName A"
  Print #lngFile, "    {"
  Print #lngFile, "        color: #FFF;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of the channel description in the banner */"
  Print #lngFile, "    .channelDescription"
  Print #lngFile, "    {"
  Print #lngFile, "        font-size: 110%;"
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
  Print #lngFile, "        font-size: 100%;"
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
  Print #lngFile, "        font-size: 90%;"
  Print #lngFile, "    }"
  Print #lngFile, "    /* Define the appearance of any other information attached to each feed item */"
  Print #lngFile, "    .itemDetails"
  Print #lngFile, "    {"
  Print #lngFile, "        font-style: italic;"
  Print #lngFile, "        font-size: 80%;"
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
        Case enumTypeItem.itemRSS
          strXML = strXML & objRSS.getXML()
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
    Case enumTypeItem.itemRSS
      getExtension = "RSS"
  End Select
End Function

Public Function loadList(ByRef lsmRSS As ListMultiple, ByRef imgFolder As StdPicture, _
                         ByRef imgRSS As StdPicture, ByRef imgNew As StdPicture, _
                         ByRef imgRead As StdPicture, ByRef imgPageWeb As StdPicture) As Boolean
'--> Carga una lista con los parámetros asociados a un elemento
Dim objItem As clsItem

  'Limpia la lista
    lsmRSS.Clear
  'Dependiendo del tipo de elemento
    Select Case intType
      Case enumTypeItem.itemFolder
        'Recorre los elementos de la carpeta
          For Each objItem In objColItems
            With objItem
              If .strKey <> "" Then
                If .intType = itemFolder Then
                  lsmRSS.Add .strTitle, "", .strURL, False, "", imgFolder, .strKey
                ElseIf .intType = itemWebPage Then
                  lsmRSS.Add .strTitle, "", .strURL, False, "", imgPageWeb, .strKey
                Else
                  lsmRSS.Add .strTitle, "", .strURL, False, "", imgRSS, .strKey
                End If
              End If
            End With
          Next objItem
        'Libera la memoria
          Set objItem = Nothing
      Case enumTypeItem.itemWebPage
        lsmRSS.Add strTitle, "", strURL, False, "", imgPageWeb, strKey
      Case enumTypeItem.itemRSS
        If strKey <> "" Then
          'Carga el archivo RSS
            Load
          'Carga la lista
            objRSS.loadList strKey, lsmRSS, imgNew, imgRead
        End If
    End Select
  'Indica que todo ha ido bien
    loadList = True
End Function

Public Property Get LocalFile() As String
  If strLocalFile = "" Then
    LocalFile = normalizeFileName(strURL) & "." & getExtension()
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
