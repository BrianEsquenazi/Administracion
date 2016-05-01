VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Proyecto1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "En&viar a"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "Nuev&o"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "E&liminar"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "&Cambiar nombre"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Propiedades"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Cerrar"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edición"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cor&tar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Peg&ado especial..."
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "&Seleccionar todo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "&Invertir selección"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Ver"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Barra de herramientas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Ba&rra de estado"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Iconos grandes"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Iconos pe&queños"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Lista"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Det&alles"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Organiz&ar iconos"
         Begin VB.Menu mnuVAIByDate 
            Caption         =   "por f&echa"
         End
         Begin VB.Menu mnuVAIByName 
            Caption         =   "por &nombre"
         End
         Begin VB.Menu mnuVAIByType 
            Caption         =   "por &tipo"
         End
         Begin VB.Menu mnuVAIBySize 
            Caption         =   "por t&amaño"
         End
      End
      Begin VB.Menu mnuViewLineUpIcons 
         Caption         =   "Alinear icono&s"
      End
      Begin VB.Menu mnuViewBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "Ac&tualizar"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Opciones..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Buscar Ayuda acerca de..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de Proyecto1..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)




Private Sub mnuHelpContents_Click()
    

    Dim nRet As Integer


    'Si no hay archivo de Ayuda para este proyecto, muestra un mensaje al usuario
    'puede establecer el archivo de Ayuda para su aplicación en el cuadro de
    'diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "Imposible mostrar los contenidos de la Ayuda. No hay una Ayuda asociada con este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuHelpSearch_Click()
    

    Dim nRet As Integer


    'Si no hay archivo de Ayuda para este proyecto, muestra un mensaje al usuario
    'puede establecer el archivo de Ayuda para su aplicación en el cuadro de
    'diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "Imposible mostrar los contenidos de la Ayuda. No hay una Ayuda asociada con este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub



Private Sub mnuVAIByDate_Click()
    'Para hacer
'  lvListView.SortKey = DATE_COLUMN
End Sub


Private Sub mnuVAIByName_Click()
    'Para hacer
'  lvListView.SortKey = NAME_COLUMN
End Sub


Private Sub mnuVAIBySize_Click()
    'Para hacer
'  lvListView.SortKey = SIZE_COLUMN
End Sub


Private Sub mnuVAIByType_Click()
    'Para hacer
'  lvListView.SortKey = TYPE_COLUMN
End Sub


Private Sub mnuListViewMode_Click(Index As Integer)
    'Desactiva el tipo actual
    mnuListViewMode(lvListView.View).Checked = False
    'Establece el modo listview
    lvListView.View = Index
    'Activa el nuevo tipo
    mnuListViewMode(Index).Checked = True
    'Establece la barra de herramientas al mismo tipo nuevo
    tbToolBar.Buttons(Index + LISTVIEW_BUTTON).Value = tbrPressed
End Sub


Private Sub mnuViewLineUpIcons_Click()
    'Para hacer
    lvListView.Arrange = lvwAutoLeft
End Sub


Private Sub mnuViewRefresh_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para renovar"
End Sub

Private Sub mnuEditCopy_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para copiar"
End Sub


Private Sub mnuEditCut_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para cortar"
End Sub


Private Sub mnuEditDSelectAll_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para seleccionar todo"
End Sub


Private Sub mnuEditInvertSelection_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para invertir selección"
End Sub


Private Sub mnuEditPaste_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para pegar"
End Sub


Private Sub mnuEditPasteSpecial_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de pegado especial"
End Sub


Private Sub mnuEditUndo_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para deshacer"
End Sub

Private Sub mnuFileOpen_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para abrir"
End Sub


Private Sub mnuFileFind_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para buscar"
End Sub


Private Sub mnuFileSendTo_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para enviar a"
End Sub


Private Sub mnuFileNew_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de nuevo archivo"
End Sub


Private Sub mnuFileDelete_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de eliminar"
End Sub


Private Sub mnuFileRename_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código para cambiar nombre"
End Sub


Private Sub mnuFileProperties_Click()
    'Para hacer
    MsgBox "Aquí se sitúa el código de propiedades"
End Sub


Private Sub mnuFileMRU_Click(Index As Integer)
    'Para hacer
    MsgBox "Aquí se sitúa el código de archivos recientes"
End Sub


Private Sub mnuFileClose_Click()
    'Descarga el formulario
    Unload Me
End Sub

