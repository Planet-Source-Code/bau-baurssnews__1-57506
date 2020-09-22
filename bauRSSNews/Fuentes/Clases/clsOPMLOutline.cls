VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOPMLOutline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con los datos de un elemento Outline de OPML
'Formato de un elemento Outline de OPML:
'    <outline text="2020 Hindsight"
'             htmlUrl = "http://www.2020Hindsight.org/"
'             type="weblogsComFavorites"
'             url="http://static.userland.com/weblogMonitor/favorites/2020Hindsight.opml"/>
Option Explicit

'Variables públicas
Public strKey As String
Public strText As String
Public strLink As String
Public strType As String
Public strURL As String

