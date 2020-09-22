Attribute VB_Name = "modListMultiple"
'--> Módulo con variables, enumerados y tipos globales del control ListMultiple
Option Explicit

'Colores de los elementos de la lista
Public colBackColor As OLE_COLOR
Public colBackColorSelected As OLE_COLOR
Public colBackColorOver As OLE_COLOR
Public colForeColor As OLE_COLOR
Public colForeColorSelected As OLE_COLOR
Public colForeColorOver As OLE_COLOR
Public colForeColorCaption As OLE_COLOR
Public colForeColorCaptionSelected As OLE_COLOR
Public colForeColorCaptionOver As OLE_COLOR


'Imágenes generales de los elementos de la lista
Public picChecked As StdPicture
Public picNoChecked As StdPicture

'Fuente utilizada en el control
Public fntList As StdFont
