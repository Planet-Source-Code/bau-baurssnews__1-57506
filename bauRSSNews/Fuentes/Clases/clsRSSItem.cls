VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRSSItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con los datos de un elemento Item de RSS
Option Explicit

'Enumerado con los tipos de estado de un elemento
Public Enum enumRSSItemStatus
  rssItemStatusNew = 0
  rssItemStatusRead
  rssItemStatusDeleted
End Enum

'Variables públicas
Public strKey As String
Public strTitle As String
Public strLink As String
Public strComments As String
Public strAuthor As String
Public strPubDate As String
Public strCategory As String
Public strSubject As String
Public strGuid As String
Public strDescription As String
Public strContent As String
Public intStatus As enumRSSItemStatus
