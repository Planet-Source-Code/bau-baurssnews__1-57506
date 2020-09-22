Attribute VB_Name = "modTypes"
'--> Tipos, enumerados y variables globales
Option Explicit

'Enumerados públicos
Public Enum enumIcons
  iconNew = 1
  iconUpdate
  iconDrop
  iconArrowFirst
  iconArrowLast
  iconArrowPrevious
  iconArrowNext
End Enum

Public Enum enumIconsRSS
  iconRSSNew = 1
  iconRSSRead
  iconRSS
  iconRSSFolder
  iconRSSWithNew
  iconRSSPageWeb
End Enum

'Variables públicas
Public frmNewAlertWindow As frmAlert
Public objColLanguage As colLanguage
Public lngProjectLastKey As Long

Public Sub killFile(ByVal strFileName As String)
'--> Elimina un archivo sin tener en cuenta los errores
  On Local Error Resume Next
    Kill strFileName
End Sub

Public Function getCData(ByVal strValue As String) As String
'--> Obtiene una cadena CData de XML
  getCData = "<![CDATA[" & strValue & "]]>"
End Function

Public Function existFile(ByVal strFileName As String) As Boolean
'--> Comprueba si existe un fichero
Dim lngLen As Long

  On Error Resume Next
    'Obtiene la longitud del archivo (si no existe devuelve un error)
      lngLen = FileLen(strFileName)
    'Comprueba si existe el archivo
      existFile = (Err.Number = 0)
End Function
