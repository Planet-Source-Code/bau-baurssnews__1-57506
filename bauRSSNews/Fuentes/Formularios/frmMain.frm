VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0F0F0&
   Caption         =   "bauRSSNews"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet inetTransfer 
      Left            =   2940
      Top             =   6540
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlRSS 
      Index           =   1
      Left            =   2190
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView trvProjects 
      Height          =   2565
      Left            =   60
      TabIndex        =   7
      Top             =   780
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4524
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin bauRSSNews.ListMultiple lsmRSS 
      Height          =   1665
      Left            =   4140
      TabIndex        =   6
      Top             =   1560
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   2937
      ForeColorCaption=   12582912
      ForeColorCaptionOver=   8421631
      ForeColorCaptionSelected=   0
   End
   Begin bauRSSNews.InfoHeader lblHeader 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   8490
      _ExtentX        =   18865
      _ExtentY        =   661
      GradientStyle   =   0
      LeftColor       =   16711680
      RightColor      =   16761024
      MaxFill         =   100
      FontName        =   "Tahoma"
      ForeColor       =   16777215
      Caption         =   "Seleccione el canal en la lista de la izquierda"
      MultiLine       =   0   'False
      Alignment       =   0
      HasIcon         =   -1  'True
      Picture         =   "frmMain.frx":0CCA
   End
   Begin MSComctlLib.ImageList imlRSS 
      Index           =   0
      Left            =   1800
      Top             =   4260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1064
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":193E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2218
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3CA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrMergeRSS 
      Enabled         =   0   'False
      Left            =   150
      Top             =   5700
   End
   Begin MSComctlLib.ImageList imlButtons 
      Left            =   1140
      Top             =   4230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4980
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":584E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6260
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7684
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "New folder"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "New RSS"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "New web page"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin bauRSSNews.SpliterVertical splVertical 
      Height          =   7155
      Left            =   3600
      TabIndex        =   2
      Top             =   2580
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   12621
      BarBackColor    =   16744576
   End
   Begin bauRSSNews.SpliterHorizontal splHorizontal 
      Height          =   7155
      Left            =   3060
      TabIndex        =   1
      Top             =   1800
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   12621
      BarBackColor    =   16744576
   End
   Begin SHDocVwCtl.WebBrowser brwRSS 
      Height          =   3165
      Left            =   3630
      TabIndex        =   0
      Top             =   3000
      Width           =   7125
      ExtentX         =   12568
      ExtentY         =   5583
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblStatus 
      Caption         =   "Preparado"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   6750
      Width           =   750
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "Menú systray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "Menu modificación proyecto"
      Visible         =   0   'False
      Begin VB.Menu mnuProjectNew 
         Caption         =   "&Nuevo"
         Begin VB.Menu mnuEditNewFolder 
            Caption         =   "&Carpeta"
         End
         Begin VB.Menu mnuEditNewRSS 
            Caption         =   "&RSS"
         End
         Begin VB.Menu mnuEditNewPageWeb 
            Caption         =   "&Página Web"
         End
      End
      Begin VB.Menu mnuProjectUpdate 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnuProjectDrop 
         Caption         =   "&Borrar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario principal del lector de RSS
Option Explicit

'Enumerados
Private Enum enumButtons 'Botones de la barra de herramientas
  buttonNew = 1
  buttonUpdate
  buttonDrop
  buttonSeparator
  buttonFirst
  buttonPrevious
  buttonNext
  buttonLast
End Enum

Private Enum enumListView
  listProject = 0
  listRSS
  listOPML
End Enum

Private Enum enumIconsSize
  iconsBig = 0
  iconsSmall
End Enum

'Variables privadas
Private intSizeIcons As enumIconsSize
Private intListFocus As enumListView
Private blnTreatingDblClick As Boolean
Private strActualKey As String
Private blnLoading As Boolean

'Objetos privados
Private objProject As clsProject
Private WithEvents objSysTray As clsSysTray
Attribute objSysTray.VB_VarHelpID = -1

Private Sub Init()
'--> Inicializa los objetos y variables
  'Desactiva el temporizador
    enableTimer False
  'Objeto de bandeja de sistema
    Set objSysTray = New clsSysTray
  'Inicializa el objeto de SysTray
    With objSysTray
      Set .SourceWindow = Me
      .Icon = Me.Icon
      .ToolTip = Me.Caption
    End With
  'Inicializa la colección de idioma
    loadLanguage App.Path & "\Data\English.lng"
  'Cambia el idioma del formulario
    changeLanguage
  'Copia los iconos
    copyIcons
  'Inicializa el tamaño de los iconos
    intSizeIcons = iconsSmall
  'Inicializa la barra de herramientas
    initToolBar
  'Carga los datos
    Load
  'Activa el temporizador
    enableTimer True
  'Inicializa la barra de estado
    setStatus ""
End Sub

Private Sub copyIcons()
'--> Copia las imágenes de la lista de iconos grandes a la lista de iconos pequeños
Dim lngIndex As Long

  For lngIndex = 1 To imlRSS(iconsBig).ListImages.Count
    imlRSS(iconsSmall).ListImages.Add , , imlRSS(iconsBig).ListImages(lngIndex).Picture
  Next lngIndex
End Sub

Private Sub initToolBar()
'--> Inicializa la barra de herramientas
  With tlbMain
    'Inicializa la lista de imágenes
      Set .ImageList = imlButtons
    'Inicialia los botones
      .Buttons(enumButtons.buttonNew).Image = enumIcons.iconNew
      .Buttons(enumButtons.buttonUpdate).Image = enumIcons.iconUpdate
      .Buttons(enumButtons.buttonDrop).Image = enumIcons.iconDrop
      .Buttons(enumButtons.buttonFirst).Image = enumIcons.iconArrowFirst
      .Buttons(enumButtons.buttonPrevious).Image = enumIcons.iconArrowPrevious
      .Buttons(enumButtons.buttonNext).Image = enumIcons.iconArrowNext
      .Buttons(enumButtons.buttonLast).Image = enumIcons.iconArrowLast
  End With
End Sub

Private Sub enableTimer(ByVal blnEnabled As Boolean)
'--> Inicializa el temporizador
  'Inicializa el temporizador
    With tmrMergeRSS
      .Enabled = blnEnabled
      .Interval = 45000
    End With
End Sub

Private Sub enableMoveButtons()
'--> Activa los botones de movimiento
Dim lngSelected As Long

  'Obtiene el elemento seleccionado
    lngSelected = lsmRSS.SelectedIndex()
  'Activa / desactiva los botones
    tlbMain.Buttons(enumButtons.buttonFirst).Enabled = (lngSelected > 0)
    tlbMain.Buttons(enumButtons.buttonPrevious).Enabled = (lngSelected > 0)
    tlbMain.Buttons(enumButtons.buttonNext).Enabled = (lngSelected < lsmRSS.Count)
    tlbMain.Buttons(enumButtons.buttonLast).Enabled = (lngSelected < lsmRSS.Count)
End Sub

Private Sub moveSelectedNew(ByVal lngIndexNew As Long)
'--> Selecciona una noticia en la lista y la carga en el navegador
  'Ajusta el índice
    If lngIndexNew < 0 Then
      lngIndexNew = 0
    ElseIf lngIndexNew > lsmRSS.Count - 1 Then
      lngIndexNew = lsmRSS.Count - 1
    End If
  'Si realmente tenemos alguna noticia
    If lsmRSS.Count > 0 Then
      'Selecciona la noticia
        lsmRSS.SelectedIndex = lngIndexNew
      'Carga la noticia seleccionada
        loadContentRSS lsmRSS.Item(lngIndexNew).Key, lsmRSS.Item(lngIndexNew).Caption
    End If
End Sub

Private Sub setFocusList(ByVal intNewListFocus As enumListView)
'--> Cambia la lista que tiene el foco
  'Cambia el valor de la variable global
    intListFocus = intNewListFocus
  'Habilita / inhabilita los botones
    tlbMain.Buttons(enumButtons.buttonNew).Enabled = (intListFocus = listProject)
    tlbMain.Buttons(enumButtons.buttonUpdate).Enabled = (intListFocus = listProject)
  'Muestra / oculta los menús
    mnuProjectNew.Visible = (intListFocus = listProject)
    mnuProjectUpdate.Visible = (intListFocus = listProject)
End Sub

Private Sub loadLanguage(ByVal strFileName As String)
'--> Carga las cadenas de un idioma
  'Crea la colección si no existe
    If objColLanguage Is Nothing Then
      Set objColLanguage = New colLanguage
    End If
  'Carga el idioma
    objColLanguage.initLanguage strFileName, "bauRSSNews", "1.0.0"
End Sub

Private Sub changeLanguage()
'--> Cambia el idioma actual de la pantalla
  'Etiqueta con el nombre del canal
    lblHeader.Caption = objColLanguage.searchItem(Me.Name, 27, lblHeader.Caption)
  'Menú SysTray
    mnuSysTrayExit.Caption = objColLanguage.searchItem(Me.Name, 1, mnuSysTrayExit.Caption)
  'Menú de modificación de proyecto
    mnuProjectNew.Caption = objColLanguage.searchItem(Me.Name, 2, mnuProjectNew.Caption)
    mnuEditNewFolder.Caption = objColLanguage.searchItem(Me.Name, 35, mnuEditNewFolder.Caption)
    mnuEditNewRSS.Caption = objColLanguage.searchItem(Me.Name, 36, mnuEditNewRSS.Caption)
    mnuEditNewPageWeb.Caption = objColLanguage.searchItem(Me.Name, 37, mnuEditNewPageWeb.Caption)
    mnuProjectUpdate.Caption = objColLanguage.searchItem(Me.Name, 3, mnuProjectUpdate.Caption)
    mnuProjectDrop.Caption = objColLanguage.searchItem(Me.Name, 4, mnuProjectDrop.Caption)
  'Tooltip de los botones de la barra de herramientas
    With tlbMain.Buttons(enumButtons.buttonNew)
      .ToolTipText = objColLanguage.searchItem(Me.Name, 5, .ToolTipText)
    End With
    With tlbMain.Buttons(enumButtons.buttonUpdate)
      .ToolTipText = objColLanguage.searchItem(Me.Name, 6, .ToolTipText)
    End With
    With tlbMain.Buttons(enumButtons.buttonDrop)
      .ToolTipText = objColLanguage.searchItem(Me.Name, 7, .ToolTipText)
    End With
End Sub

Private Sub Load()
'--> Carga el proyecto
  'Indica que se está cargando
    blnLoading = True
  'Limpia el navegador
    brwRSS.Navigate "about:blank"
  'Limpia el objeto de proyecto si existía alguno
    If Not objProject Is Nothing Then
      objProject.Clear
      Set objProject = Nothing
    End If
  'Crea el nuevo proyecto
    Set objProject = New clsProject
  'Carga los datos
    If Not objProject.Load(App.Path & "\Data\bauRSSNews.xml") Then
      MsgBox objColLanguage.searchItem(Me.Name, 9, "Error al cargar el archivo de definición")
    End If
  'Carga el árbol de proyecto
    loadTreeProject
  'Indica que se ha terminado la carga
    blnLoading = False
End Sub

Private Sub loadTreeProject()
'--> Carga el árbol de proyecto
  'Inicializa el árbol
    With trvProjects
      'Lista de imágenes
        Set .ImageList = imlRSS(intSizeIcons)
      'Limpia el árbol
        .Nodes.Clear
      'Configura las propiedades
        .HideSelection = False
        .Indentation = 0
        .LabelEdit = tvwManual
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
    End With
  'Carga los nodos
    If Not objProject Is Nothing Then
      loadTreeItems Nothing, objProject.objColItems
    End If
  'Carga el primer archivo si existe
    If trvProjects.Nodes.Count > 0 Then
      loadItem trvProjects.Nodes(1).Key, trvProjects.Nodes(1).Text
    End If
End Sub

Private Function addNodeTree(ByVal tnParent As Node, _
                             ByVal strKey As String, ByVal strTitle As String, _
                             ByVal intIcon As enumIconsRSS) As Node
'--> Añade un nodo al árbol
  If tnParent Is Nothing Then
    Set addNodeTree = trvProjects.Nodes.Add(, tvwLast, strKey, strTitle, intIcon)
  Else
    Set addNodeTree = trvProjects.Nodes.Add(tnParent, tvwChild, strKey, strTitle, intIcon)
  End If
End Function

Private Sub loadTreeItems(ByRef tnParent As Node, ByRef objColItems As colItems)
'--> Carga el árbol de elementos
Dim tnNode As Node
Dim objItem As clsItem
Dim lngNew As Long
Dim intIcon As enumIconsRSS

  'Recorre los elementos de la colección
    For Each objItem In objColItems
      'Añade el elemento
        With objItem
          Select Case objItem.intType 'Dependiendo del tipo de elemento
            Case enumTypeItem.itemFolder '...carpeta
              'Añade el nodo con la carpeta
                Set tnNode = addNodeTree(tnParent, .strKey, .strTitle, iconRSSFolder)
                tnNode.ForeColor = vbRed
              '... y los hijos de la carpeta
                loadTreeItems tnNode, .objColItems
            Case enumTypeItem.itemWebPage '... página web
              Set tnNode = addNodeTree(tnParent, .strKey, .strTitle, iconRSSPageWeb)
              tnNode.ForeColor = vbBlue
            Case enumTypeItem.itemRSS '... RSS
              'Obtiene el número de elementos nuevos
                lngNew = objItem.countNewItems()
              'Añade la cabecera de RSS
                If lngNew > 0 Then
                  Set tnNode = addNodeTree(tnParent, .strKey, _
                                           .strTitle & " (" & Format(lngNew, "#,##0") & ")", iconRSSWithNew)
                  tnNode.Bold = True
                  tnNode.ForeColor = vbBlue
                Else
                  Set tnNode = addNodeTree(tnParent, .strKey, .strTitle, iconRSS)
                End If
          End Select
        End With
    Next objItem
  'Libera la memoria
    Set objItem = Nothing
End Sub

Private Sub loadItem(ByVal strKey As String, ByVal strCaption As String)
'--> Carga un elemento en el listMultiple y el webBrowser
Static strLastKey As String

  'Carga el listMultiple con el elemento seleccionado
    If strLastKey <> strKey And strKey <> "" Then '... si realmente ha cambiado
      If Not objProject.loadList(strKey, lsmRSS, imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSSFolder).Picture, _
                                 imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSS).Picture, _
                                 imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSSNew).Picture, _
                                 imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSSRead).Picture, _
                                 imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSSPageWeb).Picture) Then
        MsgBox objColLanguage.searchItem(Me.Name, 31, "No se pueden encontrar los datos de este elemento")
      Else 'Carga el browser
        loadContentRSS strKey, strCaption
      End If
    End If
  'Activa los botones
    enableMoveButtons
  'Almacena la clave seleccionada
    strLastKey = strKey
End Sub

Private Sub loadContentRSS(ByVal strKey As String, ByVal strCaption As String)
'--> Carga el contenido de un RSS
Static strLastKey As String
Dim strURL As String

  On Error GoTo errorLoad
    'Desactiva el temporizador
      enableTimer False
    'Si realmente hemos cambiado de noticia
    ' If strLastKey <> strKey Then
        'Si estamos en una noticia
          If InStr(1, strKey, "¬") > 0 Then
            'Obtiene la URL de la noticia
              strURL = objProject.getURL(strKey)
            'Carga la noticia
              If strURL <> "" Then
                'Carga el navegador
                  brwRSS.Navigate strURL
                'Marca la noticia como leída
                  markReadRSS strURL
              End If
          Else '... si no estamos en una noticia
            'Carga el HTML de esa elemento
              If Not objProject.loadHTML(strKey, brwRSS) Then
                MsgBox objColLanguage.searchItemValue(Me.Name, 11, "Error al cargar la noticia %1", strCaption)
              End If
            'Cambia el título
              lblHeader.Caption = strCaption
          End If
    '  End If
    'Cambia el mensaje
      setStatus ""
    'Activa los botones
      enableMoveButtons
    'Actualiza la noticia actual
      strLastKey = strKey
    'Activa el temporizador
      enableTimer True
  Exit Sub

errorLoad:
  MsgBox objColLanguage.searchItemValue(Me.Name, 11, "Error al cargar la noticia %1", strKey)
End Sub

Private Sub saveProject()
'--> Graba el proyecto
  If Not objProject.Save(App.Path & "\Data\bauRSSNews.xml") Then
    MsgBox objColLanguage.searchItem(Me.Name, 16, "Error al grabar el archivo de proyecto")
'  Else '... si no ha habido errores carga de nuevo el proyecto
'    Load
  End If
End Sub

Private Function openWindowFolder(ByRef objFolder As clsItem) As Boolean
'--> Abre la ventana de modificación de una carpeta y recoge sus parámetros
Dim frmNewUpdateFolder As New frmUpdateFolder

  'Supone que se cancela la modificación
    openWindowFolder = False
  'Desactiva el temporizador
    enableTimer False
  'Abre la ventana y le pasa los parámetros
    With frmNewUpdateFolder
      'Pasa los valores a la ventana
        .strName = objFolder.strTitle
      'Muestra la ventana
        .Show vbModal, Me
      'Modifica los datos del elemento si no se ha cancelado y graba el proyecto
        If Not .blnCancel Then
          'Modifica los datos del elemento
            objFolder.strTitle = .strName
          'E indica que se han modificado los datos
            openWindowFolder = True
        End If
    End With
  'Activa el temporizador
    enableTimer True
End Function

Private Sub newFolder()
'--> Nueva carpeta
Dim objFolder As New clsItem
Dim objParent As clsItem

  'Obtiene los datos de la carpeta
    If openWindowFolder(objFolder) Then
      'Obtiene el elemento padre ...
        If trvProjects.SelectedItem Is Nothing Then '... sobre todos los demás
          Set objParent = Nothing
        Else '... obtiene el elemento que se corresponde con el nodo del árbol
          Set objParent = objProject.objColItems.Search(trvProjects.SelectedItem.Key)
        End If
      '... si no ha encontrado nada, lo añade directamente al proyecto
        If objParent Is Nothing Then
          Set objFolder = objProject.objColItems.Add(itemFolder, objFolder.strTitle, "", "", "")
        ElseIf objParent.intType = itemFolder Then
          Set objFolder = objParent.objColItems.Add(itemFolder, objFolder.strTitle, "", "", "")
        ElseIf objParent.intType <> itemFolder Then '... ¿quién es su padre?
          MsgBox objColLanguage.searchItem(Me.Name, 32, "No se puede añadir una carpeta a un archivo RSS")
          objFolder.strKey = ""
        End If
      '... si se puede añadir la carpeta
        If Not objFolder Is Nothing Then
          If objFolder.strKey <> "" Then
            'Graba el proyecto
              saveProject
            'Añade el nodo al árbol
              addNodeTree trvProjects.SelectedItem, objFolder.strKey, objFolder.strTitle, iconRSSFolder
          End If
        End If
    End If
  'Libera la memoria
    Set objFolder = Nothing
End Sub

Private Sub updateItem()
'--> Modifica los datos de un elemento
Dim objItem As clsItem
Dim blnSave As Boolean

  'Obtiene el elemento si se ha seleccionado alguno
    If trvProjects.SelectedItem Is Nothing Then
      MsgBox objColLanguage.searchItem(Me.Name, 33, "Seleccione el elemento a modificar")
    Else
      'Supone que no se debe grabar nada
        blnSave = False
      'Obtiene el elemento
        Set objItem = objProject.objColItems.Search(trvProjects.SelectedItem.Key)
      'Si realmente se ha encontrado un elemento se modifican sus datos
        If Not objItem Is Nothing Then
          If objItem.intType = itemFolder Then
            blnSave = openWindowFolder(objItem) '... indica que se debe realizar la grabación
          ElseIf objItem.intType = itemWebPage Then
            blnSave = openWindowPageWeb(objItem)
          ElseIf objItem.intType = itemRSS Then
            blnSave = openWindowRSS(objItem)
          End If
          'Si realmente se han modificado los datos
            If blnSave Then
              'Graba el proyecto
                saveProject
              'Modifica el nombre del nodo
                trvProjects.SelectedItem.Text = objItem.strTitle
            End If
        End If
    End If
  'Libera la memoria
    Set objItem = Nothing
End Sub

Private Function openWindowRSS(ByRef objRSS As clsItem) As Boolean
'--> Abre la ventana de modificación de un elemento RSS y recoge sus parámetros
Dim frmNewUpdateRSS As New frmUpdateRSS

  'Supone que se cancela la modificación
    openWindowRSS = False
  'Desactiva el temporizador
    enableTimer False
  'Abre la ventana y le pasa los parámetros
    With frmNewUpdateRSS
      'Pasa los valores a la ventana
        .strName = objRSS.strTitle
        .strURL = objRSS.strURL
        .strUser = objRSS.strUser
        .strPassword = objRSS.strPassword
      'Muestra la ventana
        .Show vbModal, Me
      'Modifica los datos del elemento si no se ha cancelado y graba el proyecto
        If Not .blnCancel Then
          'Modifica los datos del elemento
            objRSS.strTitle = .strName
            objRSS.strURL = .strURL
            objRSS.strUser = .strUser
            objRSS.strPassword = .strPassword
          'E indica que se han modificado los datos
            openWindowRSS = True
        End If
    End With
  'Activa el temporizador
    enableTimer True
  'Libera la memoria
    Set frmNewUpdateRSS = Nothing
End Function

Private Sub newRSS()
'--> Nuevo RSS
Dim objRSS As New clsItem
Dim objParent As clsItem

  'Obtiene los datos del RSS
    If openWindowRSS(objRSS) Then
      'Obtiene el elemento padre ...
        If trvProjects.SelectedItem Is Nothing Then '... sobre todos los demás
          Set objParent = Nothing
        Else '... obtiene el elemento que se corresponde con el nodo del árbol
          Set objParent = objProject.objColItems.Search(trvProjects.SelectedItem.Key)
        End If
      '... si no ha encontrado nada, lo añade directamente al proyecto
        If objParent Is Nothing Then
          Set objRSS = objProject.objColItems.Add(itemRSS, objRSS.strTitle, objRSS.strURL, objRSS.strUser, objRSS.strPassword)
        ElseIf objParent.intType = itemFolder Then
          Set objRSS = objParent.objColItems.Add(itemRSS, objRSS.strTitle, objRSS.strURL, objRSS.strUser, objRSS.strPassword)
        ElseIf objParent.intType <> itemFolder Then '... ¿quién es su padre?
          MsgBox objColLanguage.searchItem(Me.Name, 34, "No se puede añadir un archivo RSS a otro archivo RSS")
          objRSS.strKey = ""
        End If
      '... si se puede añadir el RSS
        If Not objRSS Is Nothing Then
          If objRSS.strKey <> "" Then
            'Graba el proyecto
              saveProject
            'Añade el nodo al árbol
              addNodeTree trvProjects.SelectedItem, objRSS.strKey, objRSS.strTitle, iconRSS
          End If
        End If
    End If
  'Libera la memoria
    Set objRSS = Nothing
End Sub

Private Sub dropRSS()
'--> Elimina un RSS
  If trvProjects.SelectedItem Is Nothing Then
    MsgBox objColLanguage.searchItem(Me.Name, 17, "Seleccione un canal")
  Else
    If MsgBox(objColLanguage.searchItem(Me.Name, 19, "¿Realmente desea eliminar este elemento?"), vbYesNo) = vbYes Then
      'Elimina el elemento
        objProject.removeItem trvProjects.SelectedItem.Key
      'Elimina el nodo
        trvProjects.Nodes.Remove trvProjects.SelectedItem.Key
      'Y graba el proyecto
        saveProject
    End If
  End If
End Sub

Private Sub newPageWeb()
'--> Nueva página Web
Dim objItem As New clsItem
Dim objParent As clsItem

  'Obtiene los datos de la página
    If openWindowPageWeb(objItem) Then
      'Obtiene el elemento padre ...
        If trvProjects.SelectedItem Is Nothing Then '... sobre todos los demás
          Set objParent = Nothing
        Else '... obtiene el elemento que se corresponde con el nodo del árbol
          Set objParent = objProject.objColItems.Search(trvProjects.SelectedItem.Key)
        End If
      '... si no ha encontrado nada, lo añade directamente al proyecto
        If objParent Is Nothing Then
          Set objItem = objProject.objColItems.Add(itemWebPage, objItem.strTitle, objItem.strURL, objItem.strUser, objItem.strPassword)
        ElseIf objParent.intType = itemFolder Then
          Set objItem = objParent.objColItems.Add(itemWebPage, objItem.strTitle, objItem.strURL, objItem.strUser, objItem.strPassword)
        ElseIf objParent.intType <> itemFolder Then '... ¿quién es su padre?
          MsgBox objColLanguage.searchItem(Me.Name, 38, "Sólo se pueden añadir páginas a carpetas")
          objItem.strKey = ""
        End If
      '... si se puede añadir el RSS
        If Not objItem Is Nothing Then
          If objItem.strKey <> "" Then
            'Graba el proyecto
              saveProject
            'Añade el nodo al árbol
              addNodeTree trvProjects.SelectedItem, objItem.strKey, objItem.strTitle, iconRSSPageWeb
          End If
        End If
    End If
  'Libera la memoria
    Set objItem = Nothing
End Sub

Private Function openWindowPageWeb(ByRef objItem As clsItem) As Boolean
'--> Abre la ventana de modificación de una página Web
Dim frmNewUpdatePageWeb As New frmUpdatePageWeb

  'Supone que se cancela la modificación
    openWindowPageWeb = False
  'Desactiva el temporizador
    enableTimer False
  'Abre la ventana y le pasa los parámetros
    With frmNewUpdatePageWeb
      'Pasa los valores a la ventana
        .strName = objItem.strTitle
        .strURL = objItem.strURL
        .strUser = objItem.strUser
        .strPassword = objItem.strPassword
      'Muestra la ventana
        .Show vbModal, Me
      'Modifica los datos del elemento si no se ha cancelado y graba el proyecto
        If Not .blnCancel Then
          'Modifica los datos del elemento
            objItem.strTitle = .strName
            objItem.strURL = .strURL
            objItem.strUser = .strUser
            objItem.strPassword = .strPassword
          'E indica que se han modificado los datos
            openWindowPageWeb = True
        End If
    End With
  'Activa el temporizador
    enableTimer True
  'Libera la memoria
    Set frmNewUpdatePageWeb = Nothing
End Function

Private Sub dropPageWeb()
'--> Elimina una página web
  If trvProjects.SelectedItem Is Nothing Then
    MsgBox objColLanguage.searchItem(Me.Name, 39, "Seleccione una página")
  Else
    If MsgBox(objColLanguage.searchItem(Me.Name, 19, "¿Realmente desea eliminar este elemento?"), vbYesNo) = vbYes Then
      'Elimina el elemento
        objProject.removeItem trvProjects.SelectedItem.Key
      'Elimina el nodo
        trvProjects.Nodes.Remove trvProjects.SelectedItem.Key
      'Y graba el proyecto
        saveProject
    End If
  End If
End Sub

Private Sub dropRSSItem(ByVal strKey As String)
'--> Elimina una noticia de un canal
Dim arrStrKey() As String

    If strKey = "" Then
      MsgBox objColLanguage.searchItem(Me.Name, 12, "No se ha seleccionado ninguna noticia")
    ElseIf MsgBox(objColLanguage.searchItem(Me.Name, 26, "¿Realmente desea eliminar esta noticia?"), vbYesNo) = vbYes Then
      'Separa las partes de la clave (canal - noticia)
        arrStrKey = Split(strKey, "¬")
      'Elimina el elemento
        On Error Resume Next
          objProject.removeRSSItem arrStrKey(0), arrStrKey(1), arrStrKey(2)
          If Err.Number = 0 Then
            'objProject.writeXML
            loadItem "", "" '... para que vacía strLastKey
            loadItem arrStrKey(0), lblHeader.Caption
          End If
    End If
End Sub

Private Sub mergeRSS()
'--> Descarga el RSS de la web y lo mezcla con el archivo local
Static lngLastRSS As Long
Dim objItem As clsItem
Dim blnFound As Boolean

  'Desactiva el temporizador para evitar llamadas reentrantes
    enableTimer False
  'Mezcla el siguiente RSS
    If Not objProject Is Nothing Then
      If Not objProject.objColItems Is Nothing Then
        'Actualiza el contador
          If lngLastRSS < 1 Or lngLastRSS > trvProjects.Nodes.Count Then
            lngLastRSS = 1
          Else
            lngLastRSS = lngLastRSS + 1
          End If
        'Descarga el RSS de la web y lo mezcla
          blnFound = False
          While Not blnFound And lngLastRSS <= trvProjects.Nodes.Count
            'Obtiene el elemento que se corresponde con el nodo del árbol
              Set objItem = objProject.objColItems.Search(trvProjects.Nodes(lngLastRSS).Key)
            'Si es un RSS se marca como encontrado
              If Not objItem Is Nothing Then
                If objItem.intType = itemRSS Then
                  blnFound = True
                End If
              End If
            'Si no se ha encontrado se pasa al siguiente
              If Not blnFound Then
                lngLastRSS = lngLastRSS + 1
              End If
          Wend
        'Si se ha encontrado algo
          If blnFound And Not objItem Is Nothing Then
            'Cambia la barra de estado
              setStatus objColLanguage.searchItemValue(Me.Name, 20, "Obteniendo los últimos datos de '%1'", objItem.strTitle)
            'Mezcla
              If objItem.Merge() Then
                'Muestra la ventana de alerta
                  openWindowAlert objColLanguage.searchItemValue(Me.Name, 21, "Se han obtenido nuevas noticias de '%1'", objItem.strTitle)
                'Si se ha bajado de la web una versión más reciente que la que se está mostrando, actualiza la vista
'                  If Not objActualRSS Is Nothing Then
'                    If .strKey = objActualRSS.strKey Then
'                      loadRSS objActualRSS.strKey
'                    End If
'                  End If
              End If
          End If
        'Cambia la barra de estado
          setStatus ""
      End If
    End If
  'Activa de nuevo el temporizador
    enableTimer True
End Sub

Private Sub setRSSRead(ByVal strKey As String, ByVal blnRead As Boolean)
'--> Marca el estado de lectura de una noticia
  If blnRead Then
    Set lsmRSS.Icon(strKey) = imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSSRead).Picture
    With lsmRSS.FontCaption(strKey)
      .Bold = False
    End With
  Else
    Set lsmRSS.Icon(strKey) = imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSSNew).Picture
    With lsmRSS.FontCaption(strKey)
      .Bold = True
    End With
  End If
End Sub

Private Sub refreshTreeProject()
'--> Actualiza el árbol de proyecto
Dim lngNumberNew As Long
Dim trnNode As Node

  'Actualiza el árbol
    For Each trnNode In trvProjects.Nodes
      If Not objProject.objColItems(trnNode.Key) Is Nothing Then
        With objProject.objColItems(trnNode.Key)
          '... si es un RSS
            If .intType = itemRSS Then
              If .Load(False) Then
                'Cuenta el número de elementos
                  lngNumberNew = .countNewItems()
                'Según el número de elementos actualiza el árbol
                  If .countNewItems = 0 Then
                    trnNode.Text = .strTitle
                    trnNode.Image = enumIconsRSS.iconRSS
                    trnNode.Bold = False
                    trnNode.ForeColor = vbBlack
                  Else
                    trnNode.Text = .strTitle & " (" & Format(lngNumberNew, "#,##0") & ")"
                    trnNode.Image = enumIconsRSS.iconRSSWithNew
                    trnNode.Bold = True
                    trnNode.ForeColor = vbBlue
                  End If
              End If
            End If
        End With
      End If
    Next trnNode
  'Libera la memoria
    Set trnNode = Nothing
End Sub

Private Sub markReadRSS(ByVal strURL As String)
'--> Marca una noticia como leída cuando se pulsa sobre la URL
Dim lngIndex As Long

  If Not blnLoading And Right(strURL, 12) <> "tempRSS.html" Then
    'Marca el RSS como leído
      objProject.markReadRSS strURL
    'Marca el RSS como leído en el listImage
      For lngIndex = 0 To lsmRSS.Count
        With lsmRSS.Item(lngIndex)
          If UCase(Trim(.Tag)) = UCase(Trim(strURL)) Then
            Set .IconMain = imlRSS(intSizeIcons).ListImages(enumIconsRSS.iconRSSRead).Picture
          End If
        End With
      Next lngIndex
    'Actualiza el árbol
      refreshTreeProject
  End If
End Sub

Private Sub openWindowAlert(ByVal strMessage As String)
'--> Muestra la ventana de alerta
  If frmNewAlertWindow Is Nothing Then
    Set frmNewAlertWindow = New frmAlert
  End If
  frmNewAlertWindow.DisplayMessage strMessage, 4, False, True, False, &HC0FFFF, False
End Sub

Private Sub setStatus(ByVal strMessage As String)
'--> Cambia la barra de estado
  If strMessage = "" Then
    strMessage = objColLanguage.searchItem(Me.Name, 8, "Preparado ...")
  End If
  lblStatus.Caption = "  " & strMessage
  lblStatus.Refresh
  lblStatus.ZOrder 0
  DoEvents
End Sub

Private Sub Resize()
'--> Redimensiona los controles
  On Error Resume Next
    If Me.WindowState = vbMinimized Then
      objSysTray.MinToSysTray
    Else
      'Spliter vertical
        With splVertical
          .Top = tlbMain.Top + tlbMain.Height + lblHeader.Height
          .Left = Me.ScaleLeft
          .Resize Me.ScaleWidth - .Left, Me.ScaleHeight - .Top - lblStatus.Height
          .ZOrder 0
        End With
      'Spliter horizontal
        With splHorizontal
          .Top = splVertical.Top
          .Left = splVertical.Left + splVertical.SpliterLeft + splVertical.SpliterPictureWidth
          .Resize Me.ScaleWidth - .Left, splVertical.Height
          .ZOrder 0
        End With
      'TreeView de RSS del proyecto
        With trvProjects
          .Top = splVertical.Top
          .Left = splVertical.Left
          .Width = splVertical.SpliterLeft
          .Height = splVertical.Height
        End With
      'ListView RSS
        With lsmRSS
          .Top = splHorizontal.Top
          .Left = splHorizontal.Left
          .Width = Me.ScaleWidth - .Left
          .Height = splHorizontal.SpliterTop
        End With
      'Browser
        With brwRSS
          .Top = splHorizontal.Top + splHorizontal.SpliterTop + splHorizontal.SpliterPictureHeight
          .Left = lsmRSS.Left
          .Width = lsmRSS.Width
          .Height = Me.ScaleHeight - .Top - lblStatus.Height
        End With
      'Barra de estado
        With lblStatus
          .Left = splVertical.Left
          .Width = splVertical.Width
          .Top = splVertical.Top + splVertical.Height
          .ZOrder 0
        End With
    End If
End Sub

Private Sub exitApp()
'--> Sale del programa
  'Desactiva el temporizador
    enableTimer False
  'Limpia la ventana de alerta
    If Not frmNewAlertWindow Is Nothing Then
      Unload frmNewAlertWindow
    End If
    Set frmNewAlertWindow = Nothing
  'Libera la memoria
    Set objProject = Nothing
    Set objSysTray = Nothing
  'Descarga el programa
    Unload Me
    End
End Sub

Private Sub brwRSS_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
  markReadRSS URL
End Sub

Private Sub Form_Load()
  Init
End Sub

Private Sub Form_Resize()
  Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
  exitApp
End Sub

Private Sub lsmRSS_Click(ByVal strKey As String)
  setFocusList listRSS
  If strKey <> "" Then
    'Carga el listMultiple
      If UBound(Split(strKey, "¬")) = 0 Then
        loadItem strKey, lsmRSS.Caption(strKey)
      End If
    'Carga el HTML
      loadContentRSS strKey, lsmRSS.Caption(strKey)
    'Guarda la clave del elemento
      strActualKey = strKey
  End If
End Sub

Private Sub lsmRSS_DblClick(ByVal strKey As String)
  setFocusList listRSS
  If strKey <> "" Then
    strActualKey = strKey
  End If
End Sub

Private Sub lsmRSS_GotFocus()
  setFocusList listRSS
End Sub

Private Sub lsmRSS_KeyDown(ByVal strKey As String, KeyCode As Integer, Shift As Integer)
  If strKey <> "" Then
    If KeyCode = vbKeyDelete Then
      dropRSSItem strKey
    ElseIf KeyCode = 13 Then 'Enter
      loadContentRSS strKey, lsmRSS.Caption(strKey)
    End If
  End If
End Sub

Private Sub lsmRSS_MouseDown(ByVal strKey As String, Button As Integer, Shift As Integer, x As Single, y As Single)
  setFocusList listRSS
  strActualKey = strKey
  If Button = vbRightButton Then
    PopupMenu mnuProject
  End If
End Sub

Private Sub mnuEditNewFolder_Click()
  newFolder
End Sub

Private Sub mnuEditNewPageWeb_Click()
  newPageWeb
End Sub

Private Sub mnuEditNewRSS_Click()
  newRSS
End Sub

Private Sub mnuProjectDrop_Click()
  If intListFocus = listProject Then
    dropRSS
  ElseIf intListFocus = listRSS Then
    dropRSSItem strActualKey
  End If
End Sub

Private Sub mnuProjectUpdate_Click()
  updateItem
End Sub

Private Sub mnuSysTrayExit_Click()
  exitApp
End Sub

Private Sub objSysTray_DblClick(ByVal intButton As Integer, ByVal intShift As Integer, ByVal sngX As Single, ByVal sngY As Single)
  Me.WindowState = FormWindowStateConstants.vbNormal
End Sub

Private Sub objSysTray_MouseDown(ByVal intButton As Integer, ByVal intShift As Integer, ByVal sngX As Single, ByVal sngY As Single)
  If intButton = vbRightButton Then
    PopupMenu mnuSysTray
  End If
End Sub

Private Sub splHorizontal_Resize(ByVal SpliterTop As Integer)
  Resize
End Sub

Private Sub splVertical_Resize(ByVal SpliterLeft As Integer)
  Resize
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case enumButtons.buttonUpdate
      updateItem
    Case enumButtons.buttonDrop
      If intListFocus = listProject Then
        dropRSS
      Else
        dropRSSItem strActualKey
      End If
    Case enumButtons.buttonFirst
      moveSelectedNew 0
    Case enumButtons.buttonPrevious
      moveSelectedNew lsmRSS.SelectedIndex - 1
    Case enumButtons.buttonNext
      moveSelectedNew lsmRSS.SelectedIndex + 1
    Case enumButtons.buttonLast
      moveSelectedNew lsmRSS.Count - 1
  End Select
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Select Case ButtonMenu.Index
    Case 1
      newFolder
    Case 2
      newRSS
    Case 3
      newPageWeb
  End Select
End Sub

Private Sub tmrMergeRSS_Timer()
  mergeRSS
End Sub

Private Sub trvProjects_Click()
  setFocusList listProject
  If Not trvProjects.SelectedItem Is Nothing Then
    loadItem trvProjects.SelectedItem.Key, trvProjects.SelectedItem.Text
  End If
End Sub

Private Sub trvProjects_GotFocus()
  setFocusList listProject
End Sub

Private Sub trvProjects_KeyDown(KeyCode As Integer, Shift As Integer)
  setFocusList listProject
  If KeyCode = vbKeyDelete Then
    If Not trvProjects.SelectedItem Is Nothing Then
      dropRSS
    End If
  ElseIf KeyCode = 13 Then 'Enter
    If Not trvProjects.SelectedItem Is Nothing Then
      loadItem trvProjects.SelectedItem.Key, trvProjects.SelectedItem.Text
    End If
  End If
End Sub

Private Sub trvProjects_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  setFocusList listProject
  If Button = vbRightButton Then
    If trvProjects.HitTest(x, y) Is Nothing Then
      Set trvProjects.SelectedItem = Nothing
    End If
    mnuProjectUpdate.Visible = Not trvProjects.SelectedItem Is Nothing
    mnuProjectDrop.Visible = mnuProjectUpdate.Visible
    PopupMenu mnuProject
  End If
End Sub
