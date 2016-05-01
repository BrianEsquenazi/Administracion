VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaEnsayo 
   Caption         =   "Ingreso de Ensayos"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11640
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   9840
      TabIndex        =   130
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   9720
      TabIndex        =   129
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   9840
      TabIndex        =   128
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "CargaEnsayo.frx":0000
      Left            =   240
      List            =   "CargaEnsayo.frx":0007
      TabIndex        =   7
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton AvisoErrorII 
      Caption         =   "No se puede ejecutar el procedimiento. Sistema sin Conexion con las otras plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1080
      Picture         =   "CargaEnsayo.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Version 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   8
      Text            =   " "
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Orden 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####"
      PromptChar      =   " "
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   6255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   11033
      _Version        =   327680
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Formula"
      TabPicture(0)   =   "CargaEnsayo.frx":0757
      Tab(0).ControlCount=   14
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "WVector1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "WTexto3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Realizado"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "WTexto2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "WCombo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "WTexto1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "GeneraHojaII"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Hoja"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "IngresoDescripcion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "GrabaPlanta"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LeeAnterior"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "AgregaRenglonII"
      Tab(0).Control(13).Enabled=   0   'False
      TabCaption(1)   =   "Proceso"
      TabPicture(1)   =   "CargaEnsayo.frx":0773
      Tab(1).ControlCount=   11
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "AgendaIII"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WVector2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "WTexto32"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "RealizadoII"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "WTexto22"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "WCombo12"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "WTexto12"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "AgregaRenglon"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Imprime"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "LeeAnteriorII"
      Tab(1).Control(10).Enabled=   -1  'True
      TabCaption(2)   =   "Resultados Laboratorio"
      TabPicture(2)   =   "CargaEnsayo.frx":078F
      Tab(2).ControlCount=   10
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "AgendaII"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "WVector3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "WTexto33"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "WTexto23"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "WCombo13"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "WTexto13"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "Visto"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "LeeAnteriorIII"
      Tab(2).Control(9).Enabled=   -1  'True
      TabCaption(3)   =   "Ensayos adicionales"
      TabPicture(3)   =   "CargaEnsayo.frx":07AB
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Agenda"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "Revisiones"
      TabPicture(4)   =   "CargaEnsayo.frx":07C7
      Tab(4).ControlCount=   6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "WVector4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "WTexto34"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "WTexto24"
      Tab(4).Control(2).Enabled=   -1  'True
      Tab(4).Control(3)=   "WCombo14"
      Tab(4).Control(3).Enabled=   -1  'True
      Tab(4).Control(4)=   "WTexto14"
      Tab(4).Control(4).Enabled=   -1  'True
      Tab(4).Control(5)=   "AgregaRenglonIV"
      Tab(4).Control(5).Enabled=   -1  'True
      TabCaption(5)   =   "Costo"
      TabPicture(5)   =   "CargaEnsayo.frx":07E3
      Tab(5).ControlCount=   11
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TipoCosto"
      Tab(5).Control(0).Enabled=   -1  'True
      Tab(5).Control(1)=   "Recalcula"
      Tab(5).Control(1).Enabled=   -1  'True
      Tab(5).Control(2)=   "CostoKilo"
      Tab(5).Control(2).Enabled=   -1  'True
      Tab(5).Control(3)=   "CostoTotal"
      Tab(5).Control(3).Enabled=   -1  'True
      Tab(5).Control(4)=   "WTexto15"
      Tab(5).Control(4).Enabled=   -1  'True
      Tab(5).Control(5)=   "WCombo15"
      Tab(5).Control(5).Enabled=   -1  'True
      Tab(5).Control(6)=   "WTexto25"
      Tab(5).Control(6).Enabled=   -1  'True
      Tab(5).Control(7)=   "WVector5"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "WTexto35"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Label10"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label9"
      Tab(5).Control(10).Enabled=   0   'False
      TabCaption(6)   =   "Documentacion"
      TabPicture(6)   =   "CargaEnsayo.frx":07FF
      Tab(6).ControlCount=   2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "AgendaIV"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label13"
      Tab(6).Control(1).Enabled=   0   'False
      TabCaption(7)   =   "Datos de Entrada"
      TabPicture(7)   =   "CargaEnsayo.frx":081B
      Tab(7).ControlCount=   53
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "LeeAnteriorIV"
      Tab(7).Control(0).Enabled=   -1  'True
      Tab(7).Control(1)=   "AVerificarXII"
      Tab(7).Control(1).Enabled=   -1  'True
      Tab(7).Control(2)=   "InformativoXII"
      Tab(7).Control(2).Enabled=   -1  'True
      Tab(7).Control(3)=   "ComentarioXII"
      Tab(7).Control(3).Enabled=   -1  'True
      Tab(7).Control(4)=   "RequisitoXII"
      Tab(7).Control(4).Enabled=   -1  'True
      Tab(7).Control(5)=   "AVerificarXI"
      Tab(7).Control(5).Enabled=   -1  'True
      Tab(7).Control(6)=   "InformativoXI"
      Tab(7).Control(6).Enabled=   -1  'True
      Tab(7).Control(7)=   "ComentarioXI"
      Tab(7).Control(7).Enabled=   -1  'True
      Tab(7).Control(8)=   "RequisitoXI"
      Tab(7).Control(8).Enabled=   -1  'True
      Tab(7).Control(9)=   "AVerificarX"
      Tab(7).Control(9).Enabled=   -1  'True
      Tab(7).Control(10)=   "InformativoX"
      Tab(7).Control(10).Enabled=   -1  'True
      Tab(7).Control(11)=   "ComentarioX"
      Tab(7).Control(11).Enabled=   -1  'True
      Tab(7).Control(12)=   "RequisitoX"
      Tab(7).Control(12).Enabled=   -1  'True
      Tab(7).Control(13)=   "AVerificarIX"
      Tab(7).Control(13).Enabled=   -1  'True
      Tab(7).Control(14)=   "InformativoIX"
      Tab(7).Control(14).Enabled=   -1  'True
      Tab(7).Control(15)=   "ComentarioIX"
      Tab(7).Control(15).Enabled=   -1  'True
      Tab(7).Control(16)=   "RequisitoIX"
      Tab(7).Control(16).Enabled=   -1  'True
      Tab(7).Control(17)=   "AVerificarVIII"
      Tab(7).Control(17).Enabled=   -1  'True
      Tab(7).Control(18)=   "InformativoVIII"
      Tab(7).Control(18).Enabled=   -1  'True
      Tab(7).Control(19)=   "ComentarioVIII"
      Tab(7).Control(19).Enabled=   -1  'True
      Tab(7).Control(20)=   "RequisitoVIII"
      Tab(7).Control(20).Enabled=   -1  'True
      Tab(7).Control(21)=   "AVerificarVII"
      Tab(7).Control(21).Enabled=   -1  'True
      Tab(7).Control(22)=   "InformativoVII"
      Tab(7).Control(22).Enabled=   -1  'True
      Tab(7).Control(23)=   "ComentarioVII"
      Tab(7).Control(23).Enabled=   -1  'True
      Tab(7).Control(24)=   "RequisitoVII"
      Tab(7).Control(24).Enabled=   -1  'True
      Tab(7).Control(25)=   "AVerificarVI"
      Tab(7).Control(25).Enabled=   -1  'True
      Tab(7).Control(26)=   "InformativoVI"
      Tab(7).Control(26).Enabled=   -1  'True
      Tab(7).Control(27)=   "ComentarioVI"
      Tab(7).Control(27).Enabled=   -1  'True
      Tab(7).Control(28)=   "RequisitoVI"
      Tab(7).Control(28).Enabled=   -1  'True
      Tab(7).Control(29)=   "AVerificarV"
      Tab(7).Control(29).Enabled=   -1  'True
      Tab(7).Control(30)=   "InformativoV"
      Tab(7).Control(30).Enabled=   -1  'True
      Tab(7).Control(31)=   "ComentarioV"
      Tab(7).Control(31).Enabled=   -1  'True
      Tab(7).Control(32)=   "RequisitoV"
      Tab(7).Control(32).Enabled=   -1  'True
      Tab(7).Control(33)=   "AVerificarIV"
      Tab(7).Control(33).Enabled=   -1  'True
      Tab(7).Control(34)=   "InformativoIV"
      Tab(7).Control(34).Enabled=   -1  'True
      Tab(7).Control(35)=   "ComentarioIV"
      Tab(7).Control(35).Enabled=   -1  'True
      Tab(7).Control(36)=   "RequisitoIV"
      Tab(7).Control(36).Enabled=   -1  'True
      Tab(7).Control(37)=   "AVerificarIII"
      Tab(7).Control(37).Enabled=   -1  'True
      Tab(7).Control(38)=   "InformativoIII"
      Tab(7).Control(38).Enabled=   -1  'True
      Tab(7).Control(39)=   "ComentarioIII"
      Tab(7).Control(39).Enabled=   -1  'True
      Tab(7).Control(40)=   "RequisitoIII"
      Tab(7).Control(40).Enabled=   -1  'True
      Tab(7).Control(41)=   "AVerificarII"
      Tab(7).Control(41).Enabled=   -1  'True
      Tab(7).Control(42)=   "InformativoII"
      Tab(7).Control(42).Enabled=   -1  'True
      Tab(7).Control(43)=   "ComentarioII"
      Tab(7).Control(43).Enabled=   -1  'True
      Tab(7).Control(44)=   "RequisitoII"
      Tab(7).Control(44).Enabled=   -1  'True
      Tab(7).Control(45)=   "AVerificarI"
      Tab(7).Control(45).Enabled=   -1  'True
      Tab(7).Control(46)=   "InformativoI"
      Tab(7).Control(46).Enabled=   -1  'True
      Tab(7).Control(47)=   "ComentarioI"
      Tab(7).Control(47).Enabled=   -1  'True
      Tab(7).Control(48)=   "RequisitoI"
      Tab(7).Control(48).Enabled=   -1  'True
      Tab(7).Control(49)=   "Label20"
      Tab(7).Control(49).Enabled=   0   'False
      Tab(7).Control(50)=   "Label18"
      Tab(7).Control(50).Enabled=   0   'False
      Tab(7).Control(51)=   "Label17"
      Tab(7).Control(51).Enabled=   0   'False
      Tab(7).Control(52)=   "Label16"
      Tab(7).Control(52).Enabled=   0   'False
      Begin VB.ComboBox TipoCosto 
         Height          =   315
         Left            =   -70560
         TabIndex        =   131
         Top             =   5520
         Width           =   2895
      End
      Begin VB.CommandButton LeeAnteriorIV 
         Caption         =   "Lee Version Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   127
         Top             =   5760
         Width           =   2055
      End
      Begin VB.CheckBox AVerificarXII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   126
         Top             =   5280
         Width           =   255
      End
      Begin VB.CheckBox InformativoXII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   125
         Top             =   5280
         Width           =   255
      End
      Begin VB.TextBox ComentarioXII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   124
         Text            =   " "
         Top             =   5280
         Width           =   3855
      End
      Begin VB.TextBox RequisitoXII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   123
         Text            =   " "
         Top             =   5280
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarXI 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   122
         Top             =   4920
         Width           =   255
      End
      Begin VB.CheckBox InformativoXI 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   121
         Top             =   4920
         Width           =   255
      End
      Begin VB.TextBox ComentarioXI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   120
         Text            =   " "
         Top             =   4920
         Width           =   3855
      End
      Begin VB.TextBox RequisitoXI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   119
         Text            =   " "
         Top             =   4920
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarX 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   118
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox InformativoX 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   117
         Top             =   4560
         Width           =   255
      End
      Begin VB.TextBox ComentarioX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   116
         Text            =   " "
         Top             =   4560
         Width           =   3855
      End
      Begin VB.TextBox RequisitoX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   115
         Text            =   " "
         Top             =   4560
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarIX 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   114
         Top             =   4200
         Width           =   255
      End
      Begin VB.CheckBox InformativoIX 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   113
         Top             =   4200
         Width           =   255
      End
      Begin VB.TextBox ComentarioIX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   112
         Text            =   " "
         Top             =   4200
         Width           =   3855
      End
      Begin VB.TextBox RequisitoIX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   111
         Text            =   " "
         Top             =   4200
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarVIII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   110
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox InformativoVIII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   109
         Top             =   3840
         Width           =   255
      End
      Begin VB.TextBox ComentarioVIII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   108
         Text            =   " "
         Top             =   3840
         Width           =   3855
      End
      Begin VB.TextBox RequisitoVIII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   107
         Text            =   " "
         Top             =   3840
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarVII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   106
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox InformativoVII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   105
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox ComentarioVII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   104
         Text            =   " "
         Top             =   3480
         Width           =   3855
      End
      Begin VB.TextBox RequisitoVII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   103
         Text            =   " "
         Top             =   3480
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarVI 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   102
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox InformativoVI 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   101
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox ComentarioVI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   100
         Text            =   " "
         Top             =   3120
         Width           =   3855
      End
      Begin VB.TextBox RequisitoVI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   99
         Text            =   " "
         Top             =   3120
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarV 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   98
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox InformativoV 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   97
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox ComentarioV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   96
         Text            =   " "
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox RequisitoV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   95
         Text            =   " "
         Top             =   2760
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarIV 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   94
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox InformativoIV 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   93
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox ComentarioIV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   92
         Text            =   " "
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox RequisitoIV 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   91
         Text            =   " "
         Top             =   2400
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarIII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   63
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox InformativoIII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   62
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox ComentarioIII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   61
         Text            =   " "
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox RequisitoIII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   60
         Text            =   " "
         Top             =   2040
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   59
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox InformativoII 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   58
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox ComentarioII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   57
         Text            =   " "
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox RequisitoII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   56
         Text            =   " "
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CheckBox AVerificarI 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -69000
         TabIndex        =   55
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox InformativoI 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -70200
         TabIndex        =   54
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox ComentarioI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67800
         MaxLength       =   50
         TabIndex        =   53
         Text            =   " "
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox RequisitoI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   52
         Text            =   " "
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CommandButton AgregaRenglonII 
         Caption         =   "Agrega Renglon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9120
         TabIndex        =   51
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton LeeAnteriorIII 
         Caption         =   "Lee Version Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74760
         TabIndex        =   50
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton LeeAnteriorII 
         Caption         =   "Lee Version Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69120
         TabIndex        =   49
         Top             =   5760
         Width           =   2055
      End
      Begin VB.CommandButton LeeAnterior 
         Caption         =   "Lee Version Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         TabIndex        =   48
         Top             =   4800
         Width           =   2055
      End
      Begin VB.ComboBox GrabaPlanta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   47
         Top             =   5280
         Width           =   2775
      End
      Begin VB.CommandButton Imprime 
         Caption         =   "Imprime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -65040
         TabIndex        =   46
         Top             =   5760
         Width           =   855
      End
      Begin VB.Frame IngresoDescripcion 
         Height          =   1215
         Left            =   1800
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   6975
         Begin VB.TextBox Descripcion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   44
            Text            =   " "
            Top             =   720
            Width           =   6615
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Ingrese Descripcion del Producto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   120
            Width           =   6615
         End
      End
      Begin VB.TextBox Hoja 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   42
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton Recalcula 
         Caption         =   "Recalcula"
         Height          =   615
         Left            =   -70560
         TabIndex        =   41
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox CostoKilo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -66360
         TabIndex        =   40
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox CostoTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -66360
         TabIndex        =   39
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton AgregaRenglonIV 
         Caption         =   "Agrega Renglon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -65040
         TabIndex        =   38
         Top             =   5520
         Width           =   1095
      End
      Begin VB.TextBox WTexto15 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -72960
         TabIndex        =   37
         Top             =   2040
         Width           =   375
      End
      Begin VB.ComboBox WCombo15 
         Height          =   315
         Left            =   -73560
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73560
         TabIndex        =   35
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox WTexto14 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -73080
         TabIndex        =   34
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox WCombo14 
         Height          =   315
         Left            =   -73680
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73680
         TabIndex        =   32
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox WTexto2555 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   -360
         Width           =   375
      End
      Begin VB.TextBox WTexto1555 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Top             =   -360
         Width           =   375
      End
      Begin VB.TextBox Visto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72960
         MaxLength       =   50
         TabIndex        =   29
         Text            =   " "
         Top             =   5640
         Width           =   8535
      End
      Begin VB.TextBox WTexto13 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -73320
         TabIndex        =   28
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox WCombo13 
         Height          =   315
         Left            =   -73920
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73920
         TabIndex        =   26
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton AgregaRenglon 
         Caption         =   "Agrega Renglon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66960
         TabIndex        =   25
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton GeneraHojaII 
         Caption         =   "HOJA PILOTO EN PLANTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   4800
         Width           =   2775
      End
      Begin VB.TextBox WTexto12 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -73440
         TabIndex        =   23
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox WCombo12 
         Height          =   315
         Left            =   -74040
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74040
         TabIndex        =   21
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox RealizadoII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73200
         MaxLength       =   50
         TabIndex        =   20
         Text            =   " "
         Top             =   5760
         Width           =   3975
      End
      Begin VB.TextBox WTexto1 
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox Realizado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   5760
         Width           =   5895
      End
      Begin RichTextLib.RichTextBox Agenda 
         Height          =   4455
         Left            =   -74640
         TabIndex        =   64
         Top             =   1440
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7858
         _Version        =   327680
         ScrollBars      =   3
         RightMargin     =   8900
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"CargaEnsayo.frx":0837
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   2280
         TabIndex        =   65
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   3735
         Left            =   240
         TabIndex        =   66
         Top             =   960
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6588
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   -72840
         TabIndex        =   67
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   3255
         Left            =   -74640
         TabIndex        =   68
         Top             =   720
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5741
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto33 
         Height          =   285
         Left            =   -72720
         TabIndex        =   69
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid WVector3 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   70
         Top             =   960
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4683
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin RichTextLib.RichTextBox AgendaII 
         Height          =   1695
         Left            =   -72960
         TabIndex        =   71
         Top             =   3720
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2990
         _Version        =   327680
         ScrollBars      =   3
         RightMargin     =   8900
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"CargaEnsayo.frx":08B3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox WTexto3555 
         Height          =   285
         Left            =   2520
         TabIndex        =   72
         Top             =   -360
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16711935
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox WTexto34 
         Height          =   285
         Left            =   -72480
         TabIndex        =   73
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid WVector4 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   74
         Top             =   720
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8070
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSFlexGridLib.MSFlexGrid WVector5 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   75
         Top             =   840
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6588
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto35 
         Height          =   285
         Left            =   -72360
         TabIndex        =   76
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16776960
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin RichTextLib.RichTextBox AgendaIII 
         Height          =   1575
         Left            =   -74640
         TabIndex        =   77
         Top             =   4080
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2778
         _Version        =   327680
         ScrollBars      =   3
         RightMargin     =   8900
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"CargaEnsayo.frx":092F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox AgendaIV 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   78
         Top             =   1440
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   7858
         _Version        =   327680
         ScrollBars      =   3
         RightMargin     =   8900
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"CargaEnsayo.frx":09AB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label20 
         Caption         =   "Comentarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -67800
         TabIndex        =   90
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "A Verificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -69360
         TabIndex        =   89
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Informativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -70560
         TabIndex        =   88
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Requisitos del Desarrollo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   87
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label13 
         Caption         =   "CONSULTAS VARIAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   86
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label11 
         Caption         =   "Hoja de Produccion "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3480
         TabIndex        =   85
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Costo por Kilo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -68040
         TabIndex        =   84
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Costo Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -68040
         TabIndex        =   83
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Visto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   82
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Otros Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   81
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Realizado por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74640
         TabIndex        =   80
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Label Label19 
         Caption         =   "Realizado por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         TabIndex        =   79
         Top             =   5760
         Width           =   2895
      End
   End
   Begin VB.Label Label15 
      Caption         =   "x Cliente"
      Height          =   255
      Left            =   8880
      TabIndex        =   14
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "x Ensayo"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   7800
      Width           =   855
   End
   Begin VB.Image BusquedaEnsayoII 
      Height          =   480
      Left            =   9000
      Picture         =   "CargaEnsayo.frx":0A27
      ToolTipText     =   "Busqueda de Ensayos por CLiente"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image BusquedaEnsayo 
      Height          =   480
      Left            =   7920
      Picture         =   "CargaEnsayo.frx":0E69
      ToolTipText     =   "Busqueda de Ensayos"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3960
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6720
      MouseIcon       =   "CargaEnsayo.frx":12AB
      MousePointer    =   99  'Custom
      Picture         =   "CargaEnsayo.frx":15B5
      ToolTipText     =   "Salida"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   4800
      MouseIcon       =   "CargaEnsayo.frx":1DF7
      MousePointer    =   99  'Custom
      Picture         =   "CargaEnsayo.frx":2101
      ToolTipText     =   "Elimina el Registro"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "CargaEnsayo.frx":2943
      MousePointer    =   99  'Custom
      Picture         =   "CargaEnsayo.frx":2C4D
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5760
      MouseIcon       =   "CargaEnsayo.frx":348F
      MousePointer    =   99  'Custom
      Picture         =   "CargaEnsayo.frx":3799
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "PrgCargaEnsayo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrdenTrabajo As Recordset
Dim spOrdenTrabajo As String
Dim rstOrdenTrabajoII As Recordset
Dim spOrdenTrabajoII As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstCargaEnsayo As Recordset
Dim spCargaEnsayo As String
Dim rstCargaEnsayoII As Recordset
Dim spCargaEnsayoII As String
Dim rstCargaEnsayoIII As Recordset
Dim spCargaEnsayoIII As String
Dim rstCargaEnsayoIV As Recordset
Dim spCargaEnsayoIV As String
Dim rstCargaEnsayoV As Recordset
Dim spCargaEnsayoV As String
Dim rstCargaEnsayoVI As Recordset
Dim spCargaEnsayoVI As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String

Dim CargaEmpresa(10, 2) As String
Dim ZCarga(10000, 3) As String
Dim ZClienteII As String
Dim ZProceso As Integer

Private ZAuxiliar(100, 7) As String
Dim Producto As String
Dim XCosto1 As Double
Dim XCosto2 As Double
Dim XCosto3 As Double

Dim XParam As String
Dim EmpresaActual As String
Private XEmpresa As String
Dim WVersion As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Rem para el vector II

Dim WBorraII(1000, 20) As String
Dim WParametrosII(10, 20) As Double
Dim WFormatoII(20) As String
Dim WControlII As String

Rem para el vector III

Dim WBoraIII(1000, 20) As String
Dim WParametrosIII(10, 20) As Double
Dim WFormatoIII(20) As String
Dim WControlIII As String

Rem para el vector IV

Dim WBorraIV(1000, 20) As String
Dim WParametrosIV(10, 20) As Double
Dim WFormatoIV(20) As String
Dim WControlIV As String

Rem para el vector V

Dim CargaEnsayoV(1000, 20) As String
Dim WParametrosV(10, 20) As Double
Dim WFormatoV(20) As String
Dim WControlV As String



Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZOrden = Orden.Text
    ZVersion = Version.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Orden.Text = ZOrden
    Version.Text = ZVersion
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayo"
    ZSql = ZSql + " Where CargaEnsayo.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and CargaEnsayo.Version = " + "'" + Version.Text + "'"
    spCargaEnsayo = ZSql
    Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayo.RecordCount > 0 Then
        Fecha.Text = rstCargaEnsayo!Fecha
        Cantidad.Text = Str$(rstCargaEnsayo!Cantidad)
        Realizado.Text = Trim(rstCargaEnsayo!Realizado)
        RealizadoII.Text = Trim(rstCargaEnsayo!RealizadoII)
        Visto.Text = Trim(rstCargaEnsayo!Visto)
        ZExiste = "S"
        rstCargaEnsayo.Close
    End If
    
    Cantidad.Text = Pusing("###,###.###", Val(Cantidad.Text))
    
    Call Limpia_Vector
    Call Limpia_VectorV
    WRenglon = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoII"
    ZSql = ZSql + " Where CargaEnsayoII.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and CargaEnsayoII.Version = " + "'" + Version.Text + "'"
    ZSql = ZSql + " Order by CargaEnsayoII.Clave"
    spCargaEnsayoII = ZSql
    Set rstCargaEnsayoII = db.OpenRecordset(spCargaEnsayoII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoII.RecordCount > 0 Then
        With rstCargaEnsayoII
            .MoveFirst
            Do
                If .EOF = False Then
                    Hoja.Text = IIf(IsNull(rstCargaEnsayoII!Hoja), "0", rstCargaEnsayoII!Hoja)
                    WRenglon = WRenglon + 1
                    
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCargaEnsayoII!Tipo)
                    
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstCargaEnsayoII!Articulo)
            
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCargaEnsayoII!Terminado)
                    
                    WVector1.Col = 4
                    WVector1.Text = Trim(rstCargaEnsayoII!Descripcion)
                    
                    WVector1.Col = 5
                    WVector1.Text = Str$(rstCargaEnsayoII!Cantidad)
                    WVector1.Text = Pusing("###,###.###", Val(WVector1.Text))
                    
                    WVector1.Col = 6
                    WVector1.Text = Trim(rstCargaEnsayoII!Lote)
                    
                    WVector1.Col = 7
                    WVector1.Text = Trim(rstCargaEnsayoII!Stock)
                    
                    
                    
                    WVector5.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector5.Col = 1
                    WVector5.Text = Trim(rstCargaEnsayoII!Tipo)
                    
                    WVector5.Col = 2
                    WVector5.Text = Trim(rstCargaEnsayoII!Articulo)
            
                    WVector5.Col = 3
                    WVector5.Text = Trim(rstCargaEnsayoII!Terminado)
                    
                    WVector5.Col = 4
                    WVector5.Text = Trim(rstCargaEnsayoII!Descripcion)
                    
                    WVector5.Col = 5
                    WVector5.Text = Str$(rstCargaEnsayoII!Cantidad)
                    WVector5.Text = Pusing("###,###.###", Val(WVector5.Text))
                    
                    WVector5.Col = 6
                    WVector5.Text = Str$(rstCargaEnsayoII!Costo)
                    WVector5.Text = Pusing("###,###.##", Val(WVector5.Text))
                    
                    WVector5.Col = 7
                    WVector5.Text = Str$(rstCargaEnsayoII!Cantidad * rstCargaEnsayoII!Costo)
                    WVector5.Text = Pusing("###,###.##", Val(WVector5.Text))
                    
                    WHoja = IIf(IsNull(rstCargaEnsayoII!Hoja), "0", rstCargaEnsayoII!Hoja)
                    Hoja.Text = Str$(WHoja)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoII.Close
    End If
    
    
    Call Limpia_VectorII
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoIII"
    ZSql = ZSql + " Where CargaEnsayoIII.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and CargaEnsayoIII.Version = " + "'" + Version.Text + "'"
    ZSql = ZSql + " Order by CargaEnsayoIII.Clave"
    spCargaEnsayoIII = ZSql
    Set rstCargaEnsayoIII = db.OpenRecordset(spCargaEnsayoIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoIII.RecordCount > 0 Then
        With rstCargaEnsayoIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector2.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector2.Col = 0
                    WVector2.Text = Trim(rstCargaEnsayoIII!Etapa)
                    
                    WVector2.Col = 1
                    WVector2.Text = Trim(rstCargaEnsayoIII!Etapa)
            
                    WVector2.Col = 2
                    WVector2.Text = Trim(rstCargaEnsayoIII!Instrucciones)
            
                    WVector2.Col = 3
                    WVector2.Text = Trim(rstCargaEnsayoIII!Equipo)
                    
                    WVector2.Col = 4
                    WVector2.Text = Trim(rstCargaEnsayoIII!Temperatura)
                    
                    WVector2.Col = 5
                    WVector2.Text = Trim(rstCargaEnsayoIII!Tiempo)
                    
                    WVector2.Col = 6
                    WVector2.Text = Trim(rstCargaEnsayoIII!Control)
            
                    WVector2.Col = 7
                    WVector2.Text = Trim(rstCargaEnsayoIII!Seguridad)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoIII.Close
    End If
    
    
    Call Limpia_VectorIII
    WRenglon = 0
    
    Entra = "S"
    
    If ZExiste = "S" Then
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaEnsayoIV"
        ZSql = ZSql + " Where CargaEnsayoIV.Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and CargaEnsayoIV.Version = " + "'" + Version.Text + "'"
        ZSql = ZSql + " Order by CargaEnsayoIV.Clave"
        spCargaEnsayoIV = ZSql
        Set rstCargaEnsayoIV = db.OpenRecordset(spCargaEnsayoIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaEnsayoIV.RecordCount > 0 Then
            With rstCargaEnsayoIV
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        WRenglon = WRenglon + 1
                        WVector3.Row = WRenglon
                        Renglon = WRenglon
                    
                        WVector3.Col = 1
                        WVector3.Text = Trim(rstCargaEnsayoIV!Ensayo)
                        If Val(WVector3.Text) = 0 Then
                            WVector3.Text = ""
                        End If
                        
                        WVector3.Col = 2
                        WVector3.Text = Trim(rstCargaEnsayoIV!Descripcion)
                
                        WVector3.Col = 3
                        WVector3.Text = Trim(rstCargaEnsayoIV!Esperado)
                        
                        WVector3.Col = 4
                        WVector3.Text = Trim(rstCargaEnsayoIV!Resultado)
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaEnsayoIV.Close
        End If
        
    End If
    
    If Entra = "S" Then
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenTrabajoII"
        ZSql = ZSql + " Where OrdenTrabajoII.Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " Order by OrdenTrabajoII.Clave"
        spOrdenTrabajoII = ZSql
        Set rstOrdenTrabajoII = db.OpenRecordset(spOrdenTrabajoII, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenTrabajoII.RecordCount > 0 Then
            With rstOrdenTrabajoII
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        WRenglon = WRenglon + 1
                        WVector3.Row = WRenglon
                        Renglon = WRenglon
                    
                        WVector3.Col = 1
                        WVector3.Text = Trim(rstOrdenTrabajoII!Ensayo)
                        If Val(WVector3.Text) = 0 Then
                            WVector3.Text = ""
                        End If
                        
                        WVector3.Col = 2
                        WVector3.Text = Trim(rstOrdenTrabajoII!Descripcion)
                
                        WVector3.Col = 3
                        WVector3.Text = Trim(rstOrdenTrabajoII!Resultado)
                        
                        WVector3.Col = 4
                        WVector3.Text = ""
                         
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrdenTrabajoII.Close
        End If
        
    End If
    
    For Ciclo = 1 To WRenglon
        If Val(WVector3.TextMatrix(Ciclo, 1)) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Ensayos"
            Sql3 = " Where Ensayos.Codigo = " + "'" + WVector3.TextMatrix(Ciclo, 1) + "'"
            spEnsayo = Sql1 + Sql2 + Sql3
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                WVector3.TextMatrix(Ciclo, 2) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        End If
    Next Ciclo
    
    
    
    
    
    
    
    Call Limpia_VectorIV
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoV"
    ZSql = ZSql + " Where CargaEnsayoV.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " Order by CargaEnsayoV.orden,CargaEnsayoV.renglon"
    spCargaEnsayoV = ZSql
    Set rstCargaEnsayoV = db.OpenRecordset(spCargaEnsayoV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoV.RecordCount > 0 Then
        With rstCargaEnsayoV
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector4.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector4.Col = 1
                    WVector4.Text = Trim(rstCargaEnsayoV!Version)
                    
                    WVector4.Col = 2
                    WVector4.Text = Trim(rstCargaEnsayoV!Etapa)
            
                    WVector4.Col = 3
                    WVector4.Text = Trim(rstCargaEnsayoV!Fecha)
            
                    WVector4.Col = 4
                    WVector4.Text = Trim(rstCargaEnsayoV!Participantes)
                    
                    WVector4.Col = 5
                    WVector4.Text = Trim(rstCargaEnsayoV!Resultados)
                    
                    WVector4.Col = 6
                    WVector4.Text = Trim(rstCargaEnsayoV!Acciones)
                    
                    WVector4.Col = 7
                    WVector4.Text = Trim(rstCargaEnsayoV!Responsables)
            
                    WVector4.Col = 8
                    WVector4.Text = Trim(rstCargaEnsayoV!Estado)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoV.Close
    End If
    
    
    
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoVI"
    ZSql = ZSql + " Where CargaEnsayoVI.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and CargaEnsayoVI.Version = " + "'" + Version.Text + "'"
    ZSql = ZSql + " Order by CargaEnsayoVI.Clave"
    spCargaEnsayoVI = ZSql
    Set rstCargaEnsayoVI = db.OpenRecordset(spCargaEnsayoVI, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoVI.RecordCount > 0 Then
        With rstCargaEnsayoVI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    Select Case WRenglon
                        Case 1
                            RequisitoI.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioI.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoI.Value = rstCargaEnsayoVI!Informativo
                            AVerificarI.Value = rstCargaEnsayoVI!AVerificar
                        Case 2
                            RequisitoII.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioII.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoII.Value = rstCargaEnsayoVI!Informativo
                            AVerificarII.Value = rstCargaEnsayoVI!AVerificar
                        Case 3
                            RequisitoIII.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioIII.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoIII.Value = rstCargaEnsayoVI!Informativo
                            AVerificarIII.Value = rstCargaEnsayoVI!AVerificar
                        Case 4
                            RequisitoIV.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioIV.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoIV.Value = rstCargaEnsayoVI!Informativo
                            AVerificarIV.Value = rstCargaEnsayoVI!AVerificar
                        Case 5
                            RequisitoV.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioV.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoV.Value = rstCargaEnsayoVI!Informativo
                            AVerificarV.Value = rstCargaEnsayoVI!AVerificar
                        Case 6
                            RequisitoVI.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioVI.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoVI.Value = rstCargaEnsayoVI!Informativo
                            AVerificarVI.Value = rstCargaEnsayoVI!AVerificar
                        Case 7
                            RequisitoVII.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioVII.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoVII.Value = rstCargaEnsayoVI!Informativo
                            AVerificarVII.Value = rstCargaEnsayoVI!AVerificar
                        Case 8
                            RequisitoVIII.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioVIII.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoVIII.Value = rstCargaEnsayoVI!Informativo
                            AVerificarVIII.Value = rstCargaEnsayoVI!AVerificar
                        Case 9
                            RequisitoIX.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioIX.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoIX.Value = rstCargaEnsayoVI!Informativo
                            AVerificarIX.Value = rstCargaEnsayoVI!AVerificar
                        Case 10
                            RequisitoX.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioX.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoX.Value = rstCargaEnsayoVI!Informativo
                            AVerificarX.Value = rstCargaEnsayoVI!AVerificar
                        Case 11
                            RequisitoXI.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioXI.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoXI.Value = rstCargaEnsayoVI!Informativo
                            AVerificarXI.Value = rstCargaEnsayoVI!AVerificar
                        Case Else
                            RequisitoXII.Text = Trim(rstCargaEnsayoVI!Requisito)
                            ComentarioXII.Text = Trim(rstCargaEnsayoVI!Comentario)
                            InformativoXII.Value = rstCargaEnsayoVI!Informativo
                            AVerificarXII.Value = rstCargaEnsayoVI!AVerificar
                    End Select
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoVI.Close
    End If
    
    
    
    
    WVersion = Version.Text
    Call Ceros(WVersion, 4)
    
    Agenda.LoadFile "blanco.rtf", 0
    Agenda.LoadFile "C" + Orden.Text + WVersion + ".rtf", 0
    
    AgendaII.LoadFile "blanco.rtf", 0
    AgendaII.LoadFile "E" + Orden.Text + WVersion + ".rtf", 0
    
    AgendaIII.LoadFile "blanco.rtf", 0
    AgendaIII.LoadFile "P" + Orden.Text + WVersion + ".rtf", 0
    
    AgendaIV.LoadFile "blanco.rtf", 0
    AgendaIV.LoadFile "V" + Orden.Text + WVersion + ".rtf", 0
    
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub



Private Sub AgregaRenglonII_Click()

    Hasta = WVector1.Row

    For iRow = 100 To Hasta Step -1
        WVector1.TextMatrix(iRow, 0) = WVector1.TextMatrix(iRow - 1, 0)
        WVector1.TextMatrix(iRow, 1) = WVector1.TextMatrix(iRow - 1, 1)
        WVector1.TextMatrix(iRow, 2) = WVector1.TextMatrix(iRow - 1, 2)
        WVector1.TextMatrix(iRow, 3) = WVector1.TextMatrix(iRow - 1, 3)
        WVector1.TextMatrix(iRow, 4) = WVector1.TextMatrix(iRow - 1, 4)
        WVector1.TextMatrix(iRow, 5) = WVector1.TextMatrix(iRow - 1, 5)
        WVector1.TextMatrix(iRow, 6) = WVector1.TextMatrix(iRow - 1, 6)
        WVector1.TextMatrix(iRow, 7) = WVector1.TextMatrix(iRow - 1, 7)
    Next iRow

    WVector1.TextMatrix(Hasta, 0) = ""
    WVector1.TextMatrix(Hasta, 1) = ""
    WVector1.TextMatrix(Hasta, 2) = ""
    WVector1.TextMatrix(Hasta, 3) = ""
    WVector1.TextMatrix(Hasta, 4) = ""
    WVector1.TextMatrix(Hasta, 5) = ""
    WVector1.TextMatrix(Hasta, 6) = ""
    WVector1.TextMatrix(Hasta, 7) = ""
    
    WTexto1.Text = ""
    WTexto2.Text = ""

End Sub

Private Sub cmdAdd_Click()

    If Orden.Text <> "" And Version.Text <> "" Then
    
        WVersion = Version.Text
        Call Ceros(WVersion, 4)
        
        WClave = Orden.Text + WVersion
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaEnsayo"
        ZSql = ZSql + " Where CargaEnsayo.Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and CargaEnsayo.Version = " + "'" + Version.Text + "'"
        spCargaEnsayo = ZSql
        Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaEnsayo.RecordCount > 0 Then
        
            rstCargaEnsayo.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaEnsayo SET "
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
            ZSql = ZSql + " Cantidad = " + "'" + Cantidad.Text + "',"
            ZSql = ZSql + " Realizado = " + "'" + Realizado.Text + "',"
            ZSql = ZSql + " RealizadoII = " + "'" + RealizadoII.Text + "',"
            ZSql = ZSql + " Visto = " + "'" + Visto.Text + "'"
            ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
            ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
            spCargaEnsayo = ZSql
            Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaEnsayo ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Realizado ,"
            ZSql = ZSql + "RealizadoII ,"
            ZSql = ZSql + "Visto )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Orden.Text + "',"
            ZSql = ZSql + "'" + Version.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + Cantidad.Text + "',"
            ZSql = ZSql + "'" + Realizado.Text + "',"
            ZSql = ZSql + "'" + RealizadoII.Text + "',"
            ZSql = ZSql + "'" + Visto.Text + "')"
            
            spCargaEnsayo = ZSql
            Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "DELETE CargaEnsayoII"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
        spCargaEnsayoII = ZSql
        Set rstCargaEnsayoII = db.OpenRecordset(spCargaEnsayoII, dbOpenSnapshot, dbSQLPassThrough)
        
        
        WRenglon = 0
        For iRow = 1 To 99
    
            ZTipo = WVector1.TextMatrix(iRow, 1)
            ZArticulo = WVector1.TextMatrix(iRow, 2)
            ZTerminado = WVector1.TextMatrix(iRow, 3)
            ZDescripcion = WVector1.TextMatrix(iRow, 4)
            ZCantidad = WVector1.TextMatrix(iRow, 5)
            ZLote = WVector1.TextMatrix(iRow, 6)
            ZStock = WVector1.TextMatrix(iRow, 7)
            
            ZCosto = WVector5.TextMatrix(iRow, 6)
            
            ZPartiOri = ""
            
            If ZTipo <> "" Or ZArticulo <> "" Or ZTerminado <> "" Or ZDescripcion <> "" Or ZCantidad <> "" Then
            
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                WClave = Orden.Text + WVersion + Auxi
        
                ZSql = ""
                ZSql = ZSql + "INSERT INTO CargaEnsayoII ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Orden ,"
                ZSql = ZSql + "Version ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Costo ,"
                ZSql = ZSql + "Lote ,"
                ZSql = ZSql + "Stock ,"
                ZSql = ZSql + "PartiOri )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Orden.Text + "',"
                ZSql = ZSql + "'" + Version.Text + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + ZTipo + "',"
                ZSql = ZSql + "'" + ZArticulo + "',"
                ZSql = ZSql + "'" + ZTerminado + "',"
                ZSql = ZSql + "'" + ZDescripcion + "',"
                ZSql = ZSql + "'" + ZCantidad + "',"
                ZSql = ZSql + "'" + ZCosto + "',"
                ZSql = ZSql + "'" + ZLote + "',"
                ZSql = ZSql + "'" + ZStock + "',"
                ZSql = ZSql + "'" + ZPartiOri + "')"
            
                spCargaEnsayoII = ZSql
                Set rstCargaEnsayoII = db.OpenRecordset(spCargaEnsayoII, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next iRow
        
        
        
        
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "DELETE CargaEnsayoIII"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
        spCargaEnsayoIII = ZSql
        Set rstCargaEnsayoIII = db.OpenRecordset(spCargaEnsayoIII, dbOpenSnapshot, dbSQLPassThrough)
    
        HastaRenglon = 0
        
        For iRow = 100 To 1 Step -1
        
            Etapa = WVector2.TextMatrix(iRow, 1)
            Instrucciones = WVector2.TextMatrix(iRow, 2)
            Equipo = WVector2.TextMatrix(iRow, 3)
            Temperatura = WVector2.TextMatrix(iRow, 4)
            Tiempo = WVector2.TextMatrix(iRow, 5)
            Control = WVector2.TextMatrix(iRow, 6)
            Seguridad = WVector2.TextMatrix(iRow, 7)
            
            If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
                HastaRenglon = iRow
                Exit For
            End If
            
        Next iRow
    
        WRenglon = 0
        
        For iRow = 1 To HastaRenglon
    
            ZLote = ""
        
            Etapa = WVector2.TextMatrix(iRow, 1)
            Instrucciones = WVector2.TextMatrix(iRow, 2)
            Equipo = WVector2.TextMatrix(iRow, 3)
            Temperatura = WVector2.TextMatrix(iRow, 4)
            Tiempo = WVector2.TextMatrix(iRow, 5)
            Control = WVector2.TextMatrix(iRow, 6)
            Seguridad = WVector2.TextMatrix(iRow, 7)
        
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            WClave = Orden.Text + WVersion + Auxi
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaEnsayoIII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Etapa ,"
            ZSql = ZSql + "Instrucciones ,"
            ZSql = ZSql + "Equipo ,"
            ZSql = ZSql + "Temperatura ,"
            ZSql = ZSql + "Tiempo ,"
            ZSql = ZSql + "Control ,"
            ZSql = ZSql + "Seguridad )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Orden.Text + "',"
            ZSql = ZSql + "'" + Version.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + Etapa + "',"
            ZSql = ZSql + "'" + Instrucciones + "',"
            ZSql = ZSql + "'" + Equipo + "',"
            ZSql = ZSql + "'" + Temperatura + "',"
            ZSql = ZSql + "'" + Tiempo + "',"
            ZSql = ZSql + "'" + Control + "',"
            ZSql = ZSql + "'" + Seguridad + "')"
           
            rsCargaEnsayoIII = ZSql
            Set rstCargaEnsayoIII = db.OpenRecordset(rsCargaEnsayoIII, dbOpenSnapshot, dbSQLPassThrough)
            
        Next iRow
    
    
    
    
    
    
    
    
    
    
    
        
        
        
        
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "DELETE CargaEnsayoIV"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
        spCargaEnsayoIV = ZSql
        Set rstCargaEnsayoIV = db.OpenRecordset(spCargaEnsayoIV, dbOpenSnapshot, dbSQLPassThrough)
    
        WRenglon = 0
        For iRow = 1 To 99
    
            ZEnsayo = WVector3.TextMatrix(iRow, 1)
            ZDescripcion = WVector3.TextMatrix(iRow, 2)
            ZEsperado = WVector3.TextMatrix(iRow, 3)
            ZResultado = WVector3.TextMatrix(iRow, 4)
            
            If ZEnsayo <> "" Or ZDescripcion <> "" Or ZEsperado <> "" Or ZResultado <> "" Then
            
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                WClave = Orden.Text + WVersion + Auxi
        
                ZSql = ""
                ZSql = ZSql + "INSERT INTO CargaEnsayoIV ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Orden ,"
                ZSql = ZSql + "Version ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Ensayo ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Esperado ,"
                ZSql = ZSql + "Resultado )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Orden.Text + "',"
                ZSql = ZSql + "'" + Version.Text + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + ZEnsayo + "',"
                ZSql = ZSql + "'" + ZDescripcion + "',"
                ZSql = ZSql + "'" + ZEsperado + "',"
                ZSql = ZSql + "'" + ZResultado + "')"
            
                spCargaEnsayoIV = ZSql
                Set rstCargaEnsayoIV = db.OpenRecordset(spCargaEnsayoIV, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next iRow
    
    
    
    
    
    
    
    
    
    
    
        
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "DELETE CargaEnsayoV"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        spCargaEnsayoV = ZSql
        Set rstCargaEnsayoV = db.OpenRecordset(spCargaEnsayoV, dbOpenSnapshot, dbSQLPassThrough)
        
        
        HastaRenglon = 0
        
        For iRow = 1000 To 1 Step -1
        
            ZVersion = WVector4.TextMatrix(iRow, 1)
            ZEtapa = WVector4.TextMatrix(iRow, 2)
            ZFecha = WVector4.TextMatrix(iRow, 3)
            ZParticipantes = WVector4.TextMatrix(iRow, 4)
            ZResultados = WVector4.TextMatrix(iRow, 5)
            ZAcciones = WVector4.TextMatrix(iRow, 6)
            ZResponsables = WVector4.TextMatrix(iRow, 7)
            ZEstado = WVector4.TextMatrix(iRow, 8)
            
            If ZVersion <> "" Or ZEtapa <> "" Or ZFecha <> "" Or ZParticipantes <> "" Or ZResultados <> "" Or ZAcciones <> "" Or ZResponsables <> "" Or ZEstado <> "" Then
                HastaRenglon = iRow
                Exit For
            End If
            
        Next iRow
    
        WRenglon = 0
        For iRow = 1 To HastaRenglon
    
            ZVersion = WVector4.TextMatrix(iRow, 1)
            ZEtapa = WVector4.TextMatrix(iRow, 2)
            ZFecha = WVector4.TextMatrix(iRow, 3)
            ZParticipantes = WVector4.TextMatrix(iRow, 4)
            ZResultados = WVector4.TextMatrix(iRow, 5)
            ZAcciones = WVector4.TextMatrix(iRow, 6)
            ZResponsables = WVector4.TextMatrix(iRow, 7)
            ZEstado = WVector4.TextMatrix(iRow, 8)
            
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            WClave = Orden.Text + Auxi
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaEnsayoV ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Etapa ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Participantes ,"
            ZSql = ZSql + "Resultados ,"
            ZSql = ZSql + "Acciones ,"
            ZSql = ZSql + "Responsables ,"
            ZSql = ZSql + "Estado )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Orden.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + ZVersion + "',"
            ZSql = ZSql + "'" + ZEtapa + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZParticipantes + "',"
            ZSql = ZSql + "'" + ZResultados + "',"
            ZSql = ZSql + "'" + ZAcciones + "',"
            ZSql = ZSql + "'" + ZResponsables + "',"
            ZSql = ZSql + "'" + ZEstado + "')"
            
            spCargaEnsayoV = ZSql
            Set rstCargaEnsayoV = db.OpenRecordset(spCargaEnsayoV, dbOpenSnapshot, dbSQLPassThrough)
            
        Next iRow
    
    
    
    
    
        
        
        ZSql = ""
        ZSql = ZSql + "DELETE CargaEnsayoVI"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
        spCargaEnsayoVI = ZSql
        Set rstCargaEnsayoVI = db.OpenRecordset(spCargaEnsayoVI, dbOpenSnapshot, dbSQLPassThrough)
    
        WRenglon = 0
        
        For iRow = 1 To 12
        
            Select Case iRow
                Case 1
                    ZRequisito = RequisitoI.Text
                    ZInformativo = Str$(InformativoI.Value)
                    ZAVerificar = Str$(AVerificarI.Value)
                    ZComentario = ComentarioI.Text
                Case 2
                    ZRequisito = RequisitoII.Text
                    ZInformativo = Str$(InformativoII.Value)
                    ZAVerificar = Str$(AVerificarII.Value)
                    ZComentario = ComentarioII.Text
                Case 3
                    ZRequisito = RequisitoIII.Text
                    ZInformativo = Str$(InformativoIII.Value)
                    ZAVerificar = Str$(AVerificarIII.Value)
                    ZComentario = ComentarioIII.Text
                Case 4
                    ZRequisito = RequisitoIV.Text
                    ZInformativo = Str$(InformativoIV.Value)
                    ZAVerificar = Str$(AVerificarIV.Value)
                    ZComentario = ComentarioIV.Text
                Case 5
                    ZRequisito = RequisitoV.Text
                    ZInformativo = Str$(InformativoV.Value)
                    ZAVerificar = Str$(AVerificarV.Value)
                    ZComentario = ComentarioV.Text
                Case 6
                    ZRequisito = RequisitoVI.Text
                    ZInformativo = Str$(InformativoVI.Value)
                    ZAVerificar = Str$(AVerificarVI.Value)
                    ZComentario = ComentarioVI.Text
                Case 7
                    ZRequisito = RequisitoVII.Text
                    ZInformativo = Str$(InformativoVII.Value)
                    ZAVerificar = Str$(AVerificarVII.Value)
                    ZComentario = ComentarioVII.Text
                Case 8
                    ZRequisito = RequisitoVIII.Text
                    ZInformativo = Str$(InformativoVIII.Value)
                    ZAVerificar = Str$(AVerificarVIII.Value)
                    ZComentario = ComentarioVIII.Text
                Case 9
                    ZRequisito = RequisitoIX.Text
                    ZInformativo = Str$(InformativoIX.Value)
                    ZAVerificar = Str$(AVerificarIX.Value)
                    ZComentario = ComentarioIX.Text
                Case 10
                    ZRequisito = RequisitoX.Text
                    ZInformativo = Str$(InformativoX.Value)
                    ZAVerificar = Str$(AVerificarX.Value)
                    ZComentario = ComentarioX.Text
                Case 11
                    ZRequisito = RequisitoXI.Text
                    ZInformativo = Str$(InformativoXI.Value)
                    ZAVerificar = Str$(AVerificarXI.Value)
                    ZComentario = ComentarioXI.Text
                Case 12
                    ZRequisito = RequisitoXII.Text
                    ZInformativo = Str$(InformativoXII.Value)
                    ZAVerificar = Str$(AVerificarXII.Value)
                    ZComentario = ComentarioXII.Text
                Case Else
            End Select
        
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            WClave = Orden.Text + WVersion + Auxi
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaEnsayoVI ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Requisito ,"
            ZSql = ZSql + "Informativo ,"
            ZSql = ZSql + "AVerificar ,"
            ZSql = ZSql + "Comentario )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Orden.Text + "',"
            ZSql = ZSql + "'" + Version.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + ZRequisito + "',"
            ZSql = ZSql + "'" + ZInformativo + "',"
            ZSql = ZSql + "'" + ZAVerificar + "',"
            ZSql = ZSql + "'" + ZComentario + "')"
           
            rsCargaEnsayoVI = ZSql
            Set rstCargaEnsayoVI = db.OpenRecordset(rsCargaEnsayoVI, dbOpenSnapshot, dbSQLPassThrough)
            
        Next iRow
    
    
    
    
    
    
    
    
        Agenda.SaveFile "C" + Orden.Text + WVersion + ".rtf", 0
        AgendaII.SaveFile "E" + Orden.Text + WVersion + ".rtf", 0
        AgendaIII.SaveFile "P" + Orden.Text + WVersion + ".rtf", 0
        AgendaIV.SaveFile "V" + Orden.Text + WVersion + ".rtf", 0
        
        Call CmdLimpiar_Click
        Orden.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Orden.Text <> "" And Version.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " From CargaEnsayo"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
        spCargaEnsayo = ZSql
        Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaEnsayo.RecordCount > 0 Then
        
            rstCargaEnsayo.Close
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE CargaEnsayo"
                ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
                ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
                spCargaEnsayo = ZSql
                Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "DELETE CargaEnsayoII"
                ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
                ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
                spCargaEnsayoII = ZSql
                Set rstCargaEnsayoII = db.OpenRecordset(spCargaEnsayoII, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "DELETE CargaEnsayoIII"
                ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
                ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
                spCargaEnsayoIII = ZSql
                Set rstCargaEnsayoIII = db.OpenRecordset(spCargaEnsayoIII, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "DELETE CargaEnsayoIV"
                ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
                ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
                spCargaEnsayoIV = ZSql
                Set rstCargaEnsayoIV = db.OpenRecordset(spCargaEnsayoIV, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "DELETE CargaEnsayoV"
                ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
                spCargaEnsayoV = ZSql
                Set rstCargaEnsayoV = db.OpenRecordset(spCargaEnsayoV, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "DELETE CargaEnsayoVI"
                ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
                ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
                spCargaEnsayoVI = ZSql
                Set rstCargaEnsayoVI = db.OpenRecordset(spCargaEnsayoVI, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
                
            End If
        End If
        
    End If
    
    Orden.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Call Limpia_Vector
    Call Limpia_VectorII
    Call Limpia_VectorIII
    Call Limpia_VectorIV
    Call Limpia_VectorV

    Orden.Text = "  -     "
    Version.Text = ""
    Fecha.Text = "  /  /    "
    Cantidad.Text = ""
    
    Realizado.Text = ""
    RealizadoII.Text = ""
    Visto.Text = ""
    
    RequisitoI.Text = ""
    InformativoI.Value = False
    AVerificarI.Value = False
    ComentarioI.Text = ""
    
    RequisitoII.Text = ""
    InformativoII.Value = False
    AVerificarII.Value = False
    ComentarioII.Text = ""
    
    RequisitoIII.Text = ""
    InformativoIII.Value = False
    AVerificarIII.Value = False
    ComentarioIII.Text = ""
    
    RequisitoIV.Text = ""
    InformativoIV.Value = False
    AVerificarIV.Value = False
    ComentarioIV.Text = ""
    
    RequisitoV.Text = ""
    InformativoV.Value = False
    AVerificarV.Value = False
    ComentarioV.Text = ""
    
    RequisitoVI.Text = ""
    InformativoVI.Value = False
    AVerificarVI.Value = False
    ComentarioVI.Text = ""
    
    RequisitoVII.Text = ""
    InformativoVII.Value = False
    AVerificarVII.Value = False
    ComentarioVII.Text = ""
    
    RequisitoVIII.Text = ""
    InformativoVIII.Value = False
    AVerificarVIII.Value = False
    ComentarioVIII.Text = ""
    
    RequisitoIX.Text = ""
    InformativoIX.Value = False
    AVerificarIX.Value = False
    ComentarioIX.Text = ""
    
    RequisitoX.Text = ""
    InformativoX.Value = False
    AVerificarX.Value = False
    ComentarioX.Text = ""
    
    RequisitoXI.Text = ""
    InformativoXI.Value = False
    AVerificarXI.Value = False
    ComentarioXI.Text = ""
    
    RequisitoXII.Text = ""
    InformativoXII.Value = False
    AVerificarXII.Value = False
    ComentarioXII.Text = ""
    
    
    Tablas.Tab = 0
    Agenda.LoadFile "blanco.rtf", 0
    AgendaII.LoadFile "blanco.rtf", 0
    AgendaIII.LoadFile "blanco.rtf", 0
    AgendaIV.LoadFile "blanco.rtf", 0
    
    GrabaPlanta.ListIndex = 0

    Orden.SetFocus
    
End Sub

Private Sub CmdClose_Click()

    Call CmdLimpiar_Click
    PrgCargaEnsayo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Dir2_Change()

End Sub

Private Sub File1_Click()
Agenda = File1.filename

End Sub

Private Sub GeneraHojaII_Click()
    Descripcion.Text = ""
    IngresoDescripcion.Visible = True
    Descripcion.SetFocus
End Sub

Private Sub Descripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Descripcion.Text <> "" Then
            Call GeneraHoja_Click
        End If
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub GeneraHoja_Click()

    If Val(Hoja.Text) <> 0 Then
        
        m$ = "Hoja de Produccion ya ingresada"
        a% = MsgBox(m$, 0, "Ingreso de Ensayos")
        IngresoDescripcion.Visible = False
            
            Else
            
        If GrabaPlanta.ListIndex = 0 Then
            m$ = "NO SE INFORMO PLANTA EN LA QUE SE DEBE DESCONTAR EL STOCK"
            a% = MsgBox(m$, 0, "Ingreso de Ensayos")
            IngresoDescripcion.Visible = False
            Exit Sub
        End If
            
        Rem dada
        Rem dada
        Rem dada
        Rem dada
        
        ZZEmpresa = WEmpresa
            
        If Val(WEmpresa) = 4 Then
            If GrabaPlanta.ListIndex = 2 Then
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
                Else
            If GrabaPlanta.ListIndex = 1 Then
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            If GrabaPlanta.ListIndex = 2 Then
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            If GrabaPlanta.ListIndex = 3 Then
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
        End If
            
        If Val(WEmpresa) = 3 Then
            If Left$(UCase(Orden.Text), 2) = "IF" Then
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
        End If
            
        If Val(WEmpresa) = 3 Then
            If Left$(UCase(Orden.Text), 2) = "IP" Then
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
        End If
            
        IngresoDescripcion.Visible = False
        
        WFechaHoja = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)

        WRenglon = 0
        
        For iRow = 1 To 99
    
            ZTipo = WVector1.TextMatrix(iRow, 1)
            If ZTipo = "M" Then
                ZArticulo = WVector1.TextMatrix(iRow, 2)
                ZTerminado = "  -     -   "
            End If
            If ZTipo = "T" Then
                ZArticulo = "  -   -   "
                ZTerminado = WVector1.TextMatrix(iRow, 3)
            End If
            If ZTipo = "O" Then
                ZArticulo = "  -   -   "
                ZTerminado = "  -     -   "
            End If
            ZCantidad = WVector1.TextMatrix(iRow, 5)
            ZLote = ""
            
            If ZCantidad <> "" Then
        
                Select Case ZTipo
                    Case "M"
                        spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            ZSaldo = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                            rstArticulo.Close
                            If Val(ZCantidad) > ZSaldo Then
                                m$ = ZArticulo + " Producto inexistente o Stock Insufuciente"
                                G% = MsgBox(m$, 0, "Ingreso de Ensayos")
                                If Val(WEmpresa) = 8 Then
                                    WEmpresa = "0004"
                                    txtOdbc = "Empresa04"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                End If
                                Exit Sub
                            End If
                                Else
                            m$ = ZArticulo + " Producto inexistente o Stock Insufuciente"
                            G% = MsgBox(m$, 0, "Ingreso de Ensayos")
                            If Val(WEmpresa) = 8 Then
                                WEmpresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End If
                            Exit Sub
                        End If
        
                    Case "T"
                        spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZSaldo = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                            rstTerminado.Close
                            If Val(ZCantidad) > ZSaldo Then
                                m$ = ZTerminado + " Producto inexistente o Stock Insufuciente"
                                G% = MsgBox(m$, 0, "Ingreso de Ensayos")
                                If Val(WEmpresa) = 8 Then
                                    WEmpresa = "0004"
                                    txtOdbc = "Empresa04"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                End If
                                Exit Sub
                            End If
                                Else
                            m$ = ZTerminado + " Producto inexistente o Stock Insufuciente"
                            G% = MsgBox(m$, 0, "Ingreso de Ensayos")
                            If Val(WEmpresa) = 8 Then
                                WEmpresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End If
                            Exit Sub
                        End If
        
                    Case "O"
                        WENTRA = "S"
        
                    Case Else
                        WENTRA = "N"
                             
                        m$ = ZTerminado + " Producto inexistente "
                       G% = MsgBox(m$, 0, "Ingreso de Ensayos")
                        If Val(WEmpresa) = 8 Then
                            WEmpresa = "0004"
                            txtOdbc = "Empresa04"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End If
                        Exit Sub
        
                End Select
        
            End If
        
        Next iRow
        
        spHoja = "ListaHojaNumero"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveLast
                ZHoja = rstHoja!Hoja + 1
            End With
            rstHoja.Close
                Else
            ZHoja = 1
        End If
        
        Hoja.Text = Str$(ZHoja)
        WCodigo = Orden.Text + "-100"
        
        spTerminado = "ConsultaTerminado " + "'" + WCodigo + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            rstTerminado.Close
                Else
            Call Alta_Producto
        End If
        
        Renglon = 0
        
        For iRow = 1 To 99
    
            Tipo = WVector1.TextMatrix(iRow, 1)
            If Tipo = "M" Then
                Articulo = WVector1.TextMatrix(iRow, 2)
                Terminado = "  -     -   "
            End If
            If Tipo = "T" Then
                Articulo = "  -   -   "
                Terminado = WVector1.TextMatrix(iRow, 3)
            End If
            If Tipo = "O" Then
                Articulo = "  -   -   "
                Terminado = "  -     -   "
            End If
            Canti = WVector1.TextMatrix(iRow, 5)
            Rem Lote = WVector1.TextMatrix(iRow, 6)
            Lote = ""
            
            If Canti <> "" Then
                        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(ZHoja)
                Call Ceros(Auxi1, 6)
                    
                WClave = Auxi1 + Auxi
                WHoja = Str$(ZHoja)
                WRenglon = Str$(Renglon)
                WFecha = WFechaHoja
                WProducto = Orden.Text + "-100"
                WTeorico = Cantidad.Text
                
                Rem WReal = Cantidad.Text
                Rem WFechaing = Fecha.Text
                Rem WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
                WReal = "0"
                WFechaing = "  /  /    "
                WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
                
                WTipo = Tipo
                WArticulo = Articulo
                WTerminado = Terminado
                WCantidad = Canti
                WLote = ""
                WDate = Date$
                WImporte = ""
                WMarca = ""
                WSaldo = Cantidad.Text
                Rem WLote1 = Lote
                WLote1 = ""
                WLote2 = ""
                WLote3 = ""
                Rem WCanti1 = Canti
                WCanti1 = ""
                WCanti2 = ""
                WCanti3 = ""
                WCosto1 = "0"
                WCosto2 = "0"
                WCosto3 = "0"
                
                XParam = "'" + WClave + "','" _
                            + WHoja + "','" _
                            + WRenglon + "','" _
                            + WFecha + "','" _
                            + WProducto + "','" _
                            + WCantidad + "','" _
                            + WTipo + "','" _
                            + WLote + "','" _
                            + WArticulo + "','" _
                            + WTerminado + "','" _
                            + WTeorico + "','" _
                            + WReal + "','" _
                            + WFechaing + "','" _
                            + WFechaingord + "','" _
                            + WDate + "','" _
                            + WImporte + "','" _
                            + WMarca + "','" _
                            + WSaldo + "','" _
                            + WLote1 + "','" + WCanti1 + "','" _
                            + WLote2 + "','" + WCanti2 + "','" _
                            + WLote3 + "','" + WLote3 + "','" _
                            + WCosto1 + "','" _
                            + WCosto2 + "','" _
                            + WCosto3 + "'"
                                           
                spHoja = "AltaHoja " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    
                If Renglon = 1 Then
                
                    spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WCodigo = rstTerminado!Codigo
                        WEntradas = Str$(rstTerminado!Entradas)
                        WProceso = Str$(rstTerminado!Proceso + Val(Cantidad.Text))
                        WDate = Date$
                        rstTerminado.Close
                    
                        XParam = "'" + WCodigo + "','" _
                                     + WEntradas + "','" _
                                     + WProceso + "','" _
                                     + WDate + "'"
                                           
                        spTerminado = "ModificaTerminadoHoja " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                End If
                
                Select Case Tipo
                    Case "M"
                        Rem WControla = 0
                        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            Rem WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                            WCodigo = rstArticulo!Codigo
                            WSalidas = Str$(rstArticulo!Salidas + Val(Canti))
                            WDate = Date$
                            rstArticulo.Close
                            XParam = "'" + WCodigo + "','" _
                                         + WSalidas + "','" _
                                         + WDate + "'"
                                                        
                            spArticulo = "ModificaArticuloSalidas " + XParam
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        Rem If WControla = 0 And Val(Lote) <> 0 Then
                        Rem     XParam = "'" + Lote + "','" _
                        rem                  + Articulo + "'"
                        Rem     spLaudo = "ListaLaudoArticulo " + XParam
                        Rem     Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        Rem     If rstLaudo.RecordCount > 0 Then
                        Rem         WClave = rstLaudo!Clave
                        Rem         WSaldo = Str$(rstLaudo!Saldo - Val(Canti))
                        Rem         WDate = Date$
                        Rem         rstLaudo.Close
                        Rem
                        Rem         XParam = "'" + WClave + "','" _
                        rem                      + WDate + "','" _
                        rem                      + WSaldo + "'"
                        Rem         spLaudo = "ModificaLaudoSaldo " + XParam
                        Rem         Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        Rem
                        Rem             Else
                        Rem
                        Rem         XParam = "'" + Articulo + "','" _
                        rem                      + Lote + "'"
                        Rem         spMovguia = "ListaMovguiaLote " + XParam
                        Rem         Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        Rem         If rstMovguia.RecordCount > 0 Then
                        Rem             WClave = rstMovguia!Clave
                        Rem             WSaldo = Str$(rstMovguia!Saldo - Val(Canti))
                        Rem             WDate = Date$
                        Rem             rstMovguia.Close
                        Rem
                        Rem             XParam = "'" + WClave + "','" _
                        rem                          + WDate + "','" _
                        rem                          + WSaldo + "'"
                        Rem             spMovguia = "ModificaMovguiaSaldo " + XParam
                        Rem             Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        Rem         End If
                        Rem
                        Rem     End If
                        Rem End If
                                            
                    Case "T"
                        Rem WControla = 0
                        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            Rem WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                            WCodigo = rstTerminado!Codigo
                            WSalidas = Str$(rstTerminado!Salidas + Val(Canti))
                            WDate = Date$
                            rstTerminado.Close
                            XParam = "'" + WCodigo + "','" _
                                        + WSalidas + "','" _
                                        + WDate + "'"
                                                    
                            spTerminado = "ModificaTerminadoSalidas " + XParam
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        Rem If WControla = 0 And Val(Lote) <> 0 Then
                        Rem
                        Rem     XParam = "'" + Lote + "','" _
                        rem                  + Terminado + "'"
                        Rem     spHoja = "ListaHojaProducto " + XParam
                        Rem     Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        Rem     If rstHoja.RecordCount > 0 Then
                        Rem         WClave = rstHoja!Clave
                        Rem         WSaldo = Str$(rstHoja!Saldo - Val(Canti))
                        Rem         WDate = Date$
                        Rem         rstHoja.Close
                        Rem         XParam = "'" + WClave + "','" _
                        rem                      + WDate + "','" _
                        rem                      + WSaldo + "'"
                        Rem         spHoja = "ModificaHojaSaldo " + XParam
                        Rem         Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        Rem
                        Rem             Else
                        Rem
                        Rem         XParam = "'" + Terminado + "','" _
                        rem                      + Lote + "'"
                        Rem         spMovguia = "ListaMovguiaLote1 " + XParam
                        Rem         Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        Rem         If rstMovguia.RecordCount > 0 Then
                        Rem             WClave = rstMovguia!Clave
                        Rem             WSaldo = Str$(rstMovguia!Saldo - Val(Canti))
                        Rem             WDate = Date$
                        Rem             rstMovguia.Close
                        Rem             XParam = "'" + WClave + "','" _
                        rem                      + WDate + "','" _
                        rem                      + WSaldo + "'"
                        Rem             spMovguia = "ModificaMovguiaSaldo " + XParam
                        Rem             Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        Rem         End If
                        Rem     End If
                        Rem
                        Rem End If
                        
                    Case Else
                End Select
                
            End If
                        
        Next iRow
            
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WEquipo = ""
        WVersionI = ""
        WVersionII = ""
        WVersionIII = ""
        WMarcaLabora = "S"
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " FechaOrd = " + "'" + WFechaord + "',"
        ZSql = ZSql + " Equipo = " + "'" + WEquipo + "',"
        ZSql = ZSql + " VersionI = " + "'" + WVersionI + "',"
        ZSql = ZSql + " VersionII = " + "'" + WVersionII + "',"
        ZSql = ZSql + " VersionIII = " + "'" + WVersionIII + "',"
        ZSql = ZSql + " MarcaLabora = " + "'" + WMarcaLabora + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
        
    
        Select Case Val(ZZEmpresa)
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                WEmpresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
    
    
        
        Rem If Val(WEmpresa) = 8 Then
        Rem     WEmpresa = "0004"
        Rem     txtOdbc = "Empresa04"
        Rem     strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem     Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem End If
        Rem
        Rem If Val(WEmpresa) = 5 Then
        Rem     WEmpresa = "0003"
        Rem     txtOdbc = "Empresa03"
        Rem     strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem     Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem End If
        Rem
        Rem If Val(WEmpresa) = 1 Then
        Rem     WEmpresa = "0003"
        Rem     txtOdbc = "Empresa03"
        Rem     strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem     Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem End If
        
        
        Call Impresion
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaEnsayoII SET "
        ZSql = ZSql + " Hoja = " + "'" + Hoja.Text + "'"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Version = " + "'" + Version.Text + "'"
        spCargaEnsayoII = ZSql
        Set rstCargaEnsayoII = db.OpenRecordset(spCargaEnsayoII, dbOpenSnapshot, dbSQLPassThrough)
        
        IngresoDescripcion.Visible = False
        
    End If

End Sub

Private Sub Imprime_Click()

    Rem dada
    Rem dada
    Rem dada
    Rem dada
    Rem dada
    Rem dada
    Listado.WindowTitle = ""
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    Listado.WindowTitle = "Planilla de evaluacion Semestral de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaEnsayoIII.Orden, CargaEnsayoIII.Version, CargaEnsayoIII.Renglon, CargaEnsayoIII.Etapa, CargaEnsayoIII.Instrucciones, CargaEnsayoIII.Equipo, CargaEnsayoIII.Temperatura, CargaEnsayoIII.Tiempo, CargaEnsayoIII.Control, CargaEnsayoIII.Seguridad " _
                + "From " _
                + DSQ + ".dbo.CargaEnsayoIII CargaEnsayoIII " _
                + "Where " _
                + "CargaEnsayoIII.Orden >= '" + Orden.Text + "' AND " _
                + "CargaEnsayoIII.Orden <= '" + Orden.Text + "' AND " _
                + "CargaEnsayoIII.Version >= '" + Version.Text + "' AND " _
                + "CargaEnsayoIII.Version <= '" + Version.Text + "'"
                
    Listado.ReportFileName = "ListaProceso.rpt"
                
    Uno = "{CargaEnsayoIII.Orden} in " + Chr$(34) + Orden.Text + Chr$(34) + " to " + Chr$(34) + Orden.Text + Chr$(34)
    Dos = " and {CargaEnsayoIII.Version} in " + Chr$(34) + Version.Text + Chr$(34) + " to " + Chr$(34) + Version.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
                
    Listado.Connect = Connect()
    Listado.Action = 1

End Sub

Private Sub LeeAnterior_Click()

    On Error GoTo WError

    Call Limpia_Vector
    WRenglon = 0
    
    ZZVersion = Trim(Str$(Val(Version.Text) - 1))
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoII"
    ZSql = ZSql + " Where CargaEnsayoII.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and CargaEnsayoII.Version = " + "'" + ZZVersion + "'"
    ZSql = ZSql + " Order by CargaEnsayoII.Clave"
    spCargaEnsayoII = ZSql
    Set rstCargaEnsayoII = db.OpenRecordset(spCargaEnsayoII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoII.RecordCount > 0 Then
        With rstCargaEnsayoII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCargaEnsayoII!Tipo)
                    
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstCargaEnsayoII!Articulo)
            
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCargaEnsayoII!Terminado)
                    
                    WVector1.Col = 4
                    WVector1.Text = Trim(rstCargaEnsayoII!Descripcion)
                    
                    WVector1.Col = 5
                    WVector1.Text = Str$(rstCargaEnsayoII!Cantidad)
                    WVector1.Text = Pusing("###,###.###", Val(WVector1.Text))
                    
                    WVector1.Col = 6
                    WVector1.Text = Trim(rstCargaEnsayoII!Lote)
                    
                    WVector1.Col = 7
                    WVector1.Text = Trim(rstCargaEnsayoII!Stock)
                    
                    WHoja = IIf(IsNull(rstCargaEnsayoII!Hoja), "0", rstCargaEnsayoII!Hoja)
                    Hoja.Text = Str$(WHoja)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoII.Close
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub LeeAnteriorII_Click()

    On Error GoTo WError
    
    Call Limpia_VectorII
    WRenglon = 0
    
    ZZVersion = Trim(Str$(Val(Version.Text) - 1))
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoIII"
    ZSql = ZSql + " Where CargaEnsayoIII.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and CargaEnsayoIII.Version = " + "'" + ZZVersion + "'"
    ZSql = ZSql + " Order by CargaEnsayoIII.Clave"
    spCargaEnsayoIII = ZSql
    Set rstCargaEnsayoIII = db.OpenRecordset(spCargaEnsayoIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoIII.RecordCount > 0 Then
        With rstCargaEnsayoIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector2.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector2.Col = 0
                    WVector2.Text = Trim(rstCargaEnsayoIII!Etapa)
                    
                    WVector2.Col = 1
                    WVector2.Text = Trim(rstCargaEnsayoIII!Etapa)
            
                    WVector2.Col = 2
                    WVector2.Text = Trim(rstCargaEnsayoIII!Instrucciones)
            
                    WVector2.Col = 3
                    WVector2.Text = Trim(rstCargaEnsayoIII!Equipo)
                    
                    WVector2.Col = 4
                    WVector2.Text = Trim(rstCargaEnsayoIII!Temperatura)
                    
                    WVector2.Col = 5
                    WVector2.Text = Trim(rstCargaEnsayoIII!Tiempo)
                    
                    WVector2.Col = 6
                    WVector2.Text = Trim(rstCargaEnsayoIII!Control)
            
                    WVector2.Col = 7
                    WVector2.Text = Trim(rstCargaEnsayoIII!Seguridad)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoIII.Close
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub




Private Sub LeeAnteriorIII_Click()

    If Val(Version.Text) > 1 Then

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaEnsayoIV"
        ZSql = ZSql + " Where CargaEnsayoIV.Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and CargaEnsayoIV.Version = " + "'" + Str$(Val(Version.Text) - 1) + "'"
        ZSql = ZSql + " Order by CargaEnsayoIV.Clave"
        spCargaEnsayoIV = ZSql
        Set rstCargaEnsayoIV = db.OpenRecordset(spCargaEnsayoIV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaEnsayoIV.RecordCount > 0 Then
            With rstCargaEnsayoIV
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        WRenglon = WRenglon + 1
                        Renglon = WRenglon
                    
                        WVector3.TextMatrix(WRenglon, 1) = Trim(rstCargaEnsayoIV!Ensayo)
                        WVector3.TextMatrix(WRenglon, 2) = Trim(rstCargaEnsayoIV!Descripcion)
                        WVector3.TextMatrix(WRenglon, 3) = Trim(rstCargaEnsayoIV!Esperado)
                        WVector3.TextMatrix(WRenglon, 4) = Trim(rstCargaEnsayoIV!Resultado)
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaEnsayoIV.Close
        End If
        
        WVector3.Col = 1
        WVector3.Row = 1
        Call StartEditIII
        
    End If

End Sub

Private Sub LeeAnteriorIV_Click()

    If Val(Version.Text) > 1 Then
    
        WRenglon = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaEnsayoVI"
        ZSql = ZSql + " Where CargaEnsayoVI.Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and CargaEnsayoVI.Version = " + "'" + Trim(Str$(Val(Version.Text) - 1)) + "'"
        ZSql = ZSql + " Order by CargaEnsayoVI.Clave"
        spCargaEnsayoVI = ZSql
        Set rstCargaEnsayoVI = db.OpenRecordset(spCargaEnsayoVI, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaEnsayoVI.RecordCount > 0 Then
            With rstCargaEnsayoVI
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        WRenglon = WRenglon + 1
                        
                        Select Case WRenglon
                            Case 1
                                RequisitoI.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioI.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoI.Value = rstCargaEnsayoVI!Informativo
                                AVerificarI.Value = rstCargaEnsayoVI!AVerificar
                            Case 2
                                RequisitoII.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioII.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoII.Value = rstCargaEnsayoVI!Informativo
                                AVerificarII.Value = rstCargaEnsayoVI!AVerificar
                            Case 3
                                RequisitoIII.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioIII.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoIII.Value = rstCargaEnsayoVI!Informativo
                                AVerificarIII.Value = rstCargaEnsayoVI!AVerificar
                            Case 4
                                RequisitoIV.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioIV.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoIV.Value = rstCargaEnsayoVI!Informativo
                                AVerificarIV.Value = rstCargaEnsayoVI!AVerificar
                            Case 5
                                RequisitoV.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioV.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoV.Value = rstCargaEnsayoVI!Informativo
                                AVerificarV.Value = rstCargaEnsayoVI!AVerificar
                            Case 6
                                RequisitoVI.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioVI.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoVI.Value = rstCargaEnsayoVI!Informativo
                                AVerificarVI.Value = rstCargaEnsayoVI!AVerificar
                            Case 7
                                RequisitoVII.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioVII.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoVII.Value = rstCargaEnsayoVI!Informativo
                                AVerificarVII.Value = rstCargaEnsayoVI!AVerificar
                            Case 8
                                RequisitoVIII.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioVIII.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoVIII.Value = rstCargaEnsayoVI!Informativo
                                AVerificarVIII.Value = rstCargaEnsayoVI!AVerificar
                            Case 9
                                RequisitoIX.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioIX.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoIX.Value = rstCargaEnsayoVI!Informativo
                                AVerificarIX.Value = rstCargaEnsayoVI!AVerificar
                            Case 10
                                RequisitoX.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioX.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoX.Value = rstCargaEnsayoVI!Informativo
                                AVerificarX.Value = rstCargaEnsayoVI!AVerificar
                            Case 11
                                RequisitoXI.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioXI.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoXI.Value = rstCargaEnsayoVI!Informativo
                                AVerificarXI.Value = rstCargaEnsayoVI!AVerificar
                            Case Else
                                RequisitoXII.Text = Trim(rstCargaEnsayoVI!Requisito)
                                ComentarioXII.Text = Trim(rstCargaEnsayoVI!Comentario)
                                InformativoXII.Value = rstCargaEnsayoVI!Informativo
                                AVerificarXII.Value = rstCargaEnsayoVI!AVerificar
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaEnsayoVI.Close
        End If
           
    End If
    
End Sub

Private Sub Realizado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Realizado.SetFocus
    End If
    If KeyAscii = 27 Then
        Realizado.Text = ""
    End If
End Sub

Private Sub RealizadoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RealizadoII.SetFocus
    End If
    If KeyAscii = 27 Then
        RealizadoII.Text = ""
    End If
End Sub


Private Sub TipoCosto_click()
    Call Recalcula_Click
End Sub

Private Sub Visto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Visto.SetFocus
    End If
    If KeyAscii = 27 Then
        Visto.Text = ""
    End If
End Sub



Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Orden.Text <> "" Then
        
            Orden.Text = UCase(Orden.Text)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM OrdenTrabajo"
            ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + Orden.Text + "'"
            spOrdenTrabajo = ZSql
            Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrdenTrabajo.RecordCount > 0 Then
                rstOrdenTrabajo.Close
                Rem Call Imprime_Datos
                Version.SetFocus
                    Else
                WOrden = Orden.Text
                CmdLimpiar_Click
                Orden.Text = WOrden
                Orden.SetFocus
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Orden.Text = "  -     "
    End If
End Sub

Private Sub Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Version.Text <> "" Then
            Call Imprime_Datos
            Fecha.SetFocus
                Else
            Version.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Version.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cantidad.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Cantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cantidad.Text = Pusing("###,###.###", Val(Cantidad.Text))
        Tablas.Tab = 0
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Cantidad.Text = ""
    End If
End Sub



Private Sub RequisitoI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioI.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoI.Text = ""
    End If
End Sub

Private Sub ComentarioI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoII.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioI.Text = ""
    End If
End Sub

Private Sub RequisitoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioII.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoII.Text = ""
    End If
End Sub

Private Sub ComentarioII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioII.Text = ""
    End If
End Sub

Private Sub RequisitoIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioIII.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoIII.Text = ""
    End If
End Sub

Private Sub ComentarioIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoIV.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioIII.Text = ""
    End If
End Sub

Private Sub RequisitoIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioIV.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoIV.Text = ""
    End If
End Sub

Private Sub ComentarioIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoV.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioIV.Text = ""
    End If
End Sub

Private Sub RequisitoV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioV.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoV.Text = ""
    End If
End Sub

Private Sub ComentarioV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoVI.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioV.Text = ""
    End If
End Sub

Private Sub RequisitoVI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioVI.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoVI.Text = ""
    End If
End Sub

Private Sub ComentarioVI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoVII.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioVI.Text = ""
    End If
End Sub

Private Sub RequisitoVII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioVII.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoVII.Text = ""
    End If
End Sub

Private Sub ComentarioVII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoVIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioVII.Text = ""
    End If
End Sub

Private Sub RequisitoVIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioVIII.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoVIII.Text = ""
    End If
End Sub

Private Sub ComentarioVIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoIX.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioVIII.Text = ""
    End If
End Sub

Private Sub RequisitoIX_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioIX.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoIX.Text = ""
    End If
End Sub

Private Sub ComentarioIX_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoX.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioIX.Text = ""
    End If
End Sub

Private Sub RequisitoX_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioX.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoX.Text = ""
    End If
End Sub

Private Sub ComentarioX_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoXI.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioX.Text = ""
    End If
End Sub

Private Sub RequisitoXI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioXI.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoXI.Text = ""
    End If
End Sub

Private Sub ComentarioXI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoXII.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioXI.Text = ""
    End If
End Sub

Private Sub RequisitoXII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ComentarioXII.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoXII.Text = ""
    End If
End Sub

Private Sub ComentarioXII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoI.SetFocus
    End If
    If KeyAscii = 27 Then
        ComentarioXII.Text = ""
    End If
End Sub

Sub Form_Load()

  Rem  MiRuta = "\\193.168.0.4\prueba desarrollo"
    
    
   Rem MiRuta = "c:\"
   Rem ChDir "c:\"
    
   Rem Drive1 = "c:"
  Rem  Dir1.Path = Drive1.Drive
  Rem  File1.Path = Dir1.Path  ' Establece la ruta del archivo.

    Call Limpia_Vector
    Call Limpia_VectorII
    Call Limpia_VectorIII
    Call Limpia_VectorIV
    Call Limpia_VectorV
    WVector1.Col = 1
    WVector1.Row = 1


    Orden.Text = "  -     "
    Version.Text = ""
    Fecha.Text = "  /  /    "
    Cantidad.Text = ""
    Realizado.Text = ""
    RealizadoII.Text = ""
    Visto.Text = ""
    
    RequisitoI.Text = ""
    InformativoI.Value = False
    AVerificarI.Value = False
    ComentarioI.Text = ""
    
    RequisitoII.Text = ""
    InformativoII.Value = False
    AVerificarII.Value = False
    ComentarioII.Text = ""
    
    RequisitoIII.Text = ""
    InformativoIII.Value = False
    AVerificarIII.Value = False
    ComentarioIII.Text = ""
    
    RequisitoIV.Text = ""
    InformativoIV.Value = False
    AVerificarIV.Value = False
    ComentarioIV.Text = ""
    
    RequisitoV.Text = ""
    InformativoV.Value = False
    AVerificarV.Value = False
    ComentarioV.Text = ""
    
    RequisitoVI.Text = ""
    InformativoVI.Value = False
    AVerificarVI.Value = False
    ComentarioVI.Text = ""
    
    RequisitoVII.Text = ""
    InformativoVII.Value = False
    AVerificarVII.Value = False
    ComentarioVII.Text = ""
    
    RequisitoVIII.Text = ""
    InformativoVIII.Value = False
    AVerificarVIII.Value = False
    ComentarioVIII.Text = ""
    
    RequisitoIX.Text = ""
    InformativoIX.Value = False
    AVerificarIX.Value = False
    ComentarioIX.Text = ""
    
    RequisitoX.Text = ""
    InformativoX.Value = False
    AVerificarX.Value = False
    ComentarioX.Text = ""
    
    RequisitoXI.Text = ""
    InformativoXI.Value = False
    AVerificarXI.Value = False
    ComentarioXI.Text = ""
    
    RequisitoXII.Text = ""
    InformativoXII.Value = False
    AVerificarXII.Value = False
    ComentarioXII.Text = ""
    
    Agenda.LoadFile "blanco.rtf", 0
    AgendaII.LoadFile "blanco.rtf", 0
    AgendaIII.LoadFile "blanco.rtf", 0
    AgendaIV.LoadFile "blanco.rtf", 0
    
    
    If Val(WEmpresa) <> 4 Then
    
        GrabaPlanta.Clear
        
        GrabaPlanta.AddItem ""
        GrabaPlanta.AddItem "Planta II"
        GrabaPlanta.AddItem "Planta I"
        GrabaPlanta.AddItem "Planta III"
        
        GrabaPlanta.ListIndex = 0
        GrabaPlanta.Visible = True
        
            Else
            
        GrabaPlanta.Clear
        
        GrabaPlanta.AddItem ""
        GrabaPlanta.AddItem "Planta II"
        GrabaPlanta.AddItem "Planta V"
        
        GrabaPlanta.ListIndex = 0
        GrabaPlanta.Visible = True
        
    End If
    
    TipoCosto.Clear
    
    TipoCosto.AddItem "Costo Standard y Estimado"
    TipoCosto.AddItem "Costo Ultima Compra"
    
    TipoCosto.ListIndex = 0
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 0
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        Case 1
            WVector2.Col = 1
            WVector2.Row = 1
            Call StartEditII
        Case 2
            WVector3.Col = 1
            WVector3.Row = 1
            Call StartEditIII
        Case 3
            Agenda.SetFocus
        Case 4
            WVector4.Col = 1
            WVector4.Row = 1
            Call StartEditIV
        Case 5
            Call Recalcula_Click
        Case 6
            AgendaIV.SetFocus
        Case 7
            RequisitoI.SetFocus
        Case Else
    End Select
End Sub








Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 7
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()

    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            WVector1.Text = UCase(WVector1.Text)
            If WVector1.Text <> "T" And WVector1.Text <> "M" And WVector1.Text <> "O" Then
                WControl = "N"
                    Else
                If WVector1.Text = "M" Then
                    WVector1.TextMatrix(WVector1.Row, 3) = "  -     -   "
                    WVector1.Col = 1
                        Else
                    If WVector1.Text = "T" Then
                        WVector1.TextMatrix(WVector1.Row, 2) = "  -   -   "
                        WVector1.Col = 2
                            Else
                        If WVector1.Text = "O" Then
                            WVector1.TextMatrix(WVector1.Row, 2) = "  -   -   "
                            WVector1.TextMatrix(WVector1.Row, 3) = "  -     -   "
                            WVector1.Col = 3
                        End If
                    End If
                End If
            End If
    
        Case 2
            WVector1.Text = UCase(WVector1.Text)
            If WVector1.Text <> "" Then
                ZCodigo = WVector1.Text
                Sql1 = "Select *"
                Sql2 = " FROM Articulo"
                Sql3 = " Where Articulo.Codigo = " + "'" + ZCodigo + "'"
                spArticulo = Sql1 + Sql2 + Sql3
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 4
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                        Else
                    WControl = "N"
                End If
            End If
            
        Case 3
            WVector1.Text = UCase(WVector1.Text)
            If WVector1.Text <> "" Then
                ZCodigo = WVector1.Text
                Sql1 = "Select *"
                Sql2 = " FROM Terminado"
                Sql3 = " Where Terminado.Codigo = " + "'" + ZCodigo + "'"
                spTerminado = Sql1 + Sql2 + Sql3
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WVector1.Col = 4
                    WVector1.Text = rstTerminado!Descripcion
                    rstTerminado.Close
                        Else
                    WControl = "N"
                End If
            End If
            
        Case 7
            WVector1.Text = UCase(WVector1.Text)
            If WVector1.Text <> "" And WVector1.Text <> "S" Then
                WControl = "N"
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        ZCodigo = WVector1.TextMatrix(iRow, 1)
        If ZCodigo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For DA = 0 To WVector1.Cols - 1
            WVector1.Col = DA
            WVector1.Text = WBorra(Ciclo, DA)
        Next DA
    Next Ciclo
    
    End If
    
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Clientes"
     Opcion.AddItem "Ensayos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
                
            Case 2
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
                
            Case 3
                WVector1.Text = "Terminado"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
                
            Case 4
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
                
            Case 5
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.###"
                
            Case 6
                WVector1.Text = "Lote"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
                
            Case 7
                WVector1.Text = "Stock"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub














Rem
Rem Controles de la WVector2
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector2.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector2.CellLeft + WVector2.Left
            WTexto22.Top = WVector2.CellTop + WVector2.Top
            WTexto22.Width = WVector2.CellWidth
            WTexto22.Height = WVector2.CellHeight
            WTexto22.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector2.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector2.CellLeft + WVector2.Left
            WTexto32.Top = WVector2.CellTop + WVector2.Top
            WTexto32.Width = WVector2.CellWidth
            WTexto32.Height = WVector2.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector2.Text) = 10 Then
                        WTexto32.Text = WVector2.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector2.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector2.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector2.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector2.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormatoII(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector2.CellLeft + WVector2.Left
    WCombo12.Top = WVector2.CellTop + WVector2.Top
    WCombo12.Width = WVector2.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1
        Case 113
            WTexto12.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 123
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Col > 1 Then
                WVector2.Col = WVector2.Col - 1
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1
        Case 113
            WTexto22.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub Wtexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto32.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub Wtexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEditII
End Sub

Private Sub WVector2_LeaveCell()
    EndEditII
End Sub

Private Sub WVector2_GotFocus()
    EndEditII
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector2.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 7
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 1
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControlII = "S"
    Select Case XColumna
        Case 1
            WVector2.TextMatrix(WVector2.Row, 0) = WVector2.TextMatrix(WVector2.Row, 1)
        Case 3, 6, 7
            Rem If Val(WVector2.Text) <> 0 Then
            Rem     ZCodigo = Val(WVector2.Text)
            Rem     Call Ceros(ZCodigo, 4)
            Rem
            Rem     Sql1 = "Select *"
            Rem     Sql2 = " FROM EquipoFabrica"
            Rem     Sql3 = " Where EquipoFabrica.Codigo = " + "'" + ZCodigo + "'"
            Rem     spEquipoFabrica = Sql1 + Sql2 + Sql3
            Rem     Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstEquipoFabrica.RecordCount > 0 Then
            Rem         rstEquipoFabrica.Close
            Rem     End If
            Rem End If
            
        Case Else
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
    
    RenglonAuxiliar = WVector2.Row

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorraII
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        Etapa = WVector2.TextMatrix(iRow, 1)
        Instrucciones = WVector2.TextMatrix(iRow, 2)
        Equipo = WVector2.TextMatrix(iRow, 3)
        Temperatura = WVector2.TextMatrix(iRow, 4)
        Tiempo = WVector2.TextMatrix(iRow, 5)
        Control = WVector2.TextMatrix(iRow, 6)
        Seguridad = WVector2.TextMatrix(iRow, 7)
            
        If Etapa <> "" Or Instrucciones <> "" Or Equipo <> "" Or Temperatura <> "" Or Tiempo <> "" Or Control <> "" Or Seguridad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorraII(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For DA = 0 To WVector2.Cols - 1
            WVector2.Col = DA
            WVector2.Text = WBorraII(Ciclo, DA)
        Next DA
    Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglon_Click()

    Hasta = WVector2.Row

    For iRow = 100 To Hasta Step -1
        WVector2.TextMatrix(iRow, 0) = WVector2.TextMatrix(iRow - 1, 0)
        WVector2.TextMatrix(iRow, 1) = WVector2.TextMatrix(iRow - 1, 1)
        WVector2.TextMatrix(iRow, 2) = WVector2.TextMatrix(iRow - 1, 2)
        WVector2.TextMatrix(iRow, 3) = WVector2.TextMatrix(iRow - 1, 3)
        WVector2.TextMatrix(iRow, 4) = WVector2.TextMatrix(iRow - 1, 4)
        WVector2.TextMatrix(iRow, 5) = WVector2.TextMatrix(iRow - 1, 5)
        WVector2.TextMatrix(iRow, 6) = WVector2.TextMatrix(iRow - 1, 6)
        WVector2.TextMatrix(iRow, 7) = WVector2.TextMatrix(iRow - 1, 7)
    Next iRow

    WVector2.TextMatrix(Hasta, 0) = ""
    WVector2.TextMatrix(Hasta, 1) = ""
    WVector2.TextMatrix(Hasta, 2) = ""
    WVector2.TextMatrix(Hasta, 3) = ""
    WVector2.TextMatrix(Hasta, 4) = ""
    WVector2.TextMatrix(Hasta, 5) = ""
    WVector2.TextMatrix(Hasta, 6) = ""
    WVector2.TextMatrix(Hasta, 7) = ""
    
    WTexto12.Text = ""
    WTexto22.Text = ""

End Sub

Private Sub Limpia_VectorII()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector2.FontName
    WTexto12.FontSize = WVector2.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector2.FontName
    WTexto22.FontSize = WVector2.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector2.FontName
    WTexto32.FontSize = WVector2.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector2.FontName
    WCombo12.FontSize = WVector2.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 8
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 400
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Etapa"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 10
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "Detalle del Trabajo"
                WVector2.ColWidth(Ciclo) = 8900
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 90
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Equipo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 4
                WVector2.Text = "Temperatura"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 20
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 5
                WVector2.Text = "Tiempo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 6
                WVector2.Text = "Control"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 20
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 7
                WVector2.Text = "Seguridad"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 15
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector2.Text
        Rem WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        Rem WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        Rem WTitulo(Ciclo).Width = WVector2.CellWidth
        Rem WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub




















Rem
Rem Controles de la WVector3
Rem

Private Sub GridEditTextIII(ByVal KeyAscii As Integer)

    XColumna = WVector3.Col
    XTipoDato = WParametrosIII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto13.Left = WVector3.CellLeft + WVector3.Left
            WTexto13.Top = WVector3.CellTop + WVector3.Top
            WTexto13.Width = WVector3.CellWidth
            WTexto13.Height = WVector3.CellHeight
            WTexto13.MaxLength = WParametrosIII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto13.Text = WVector3.Text
                    WTexto13.SelStart = Len(WTexto13.Text)
                Case Else
                    WTexto13.Text = Chr$(KeyAscii)
                    WTexto13.SelStart = 1
            End Select
            WTexto13.Visible = True
            WTexto13.SetFocus
        Case 1
            WTexto23.Left = WVector3.CellLeft + WVector3.Left
            WTexto23.Top = WVector3.CellTop + WVector3.Top
            WTexto23.Width = WVector3.CellWidth
            WTexto23.Height = WVector3.CellHeight
            WTexto23.MaxLength = WParametrosIII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto23.Text = WVector3.Text
                    Rem WTexto23.SelStart = Len(WTexto23.Text)
                    WTexto23.SelStart = 0
                Case Else
                    WTexto23.Text = Chr$(KeyAscii)
                    WTexto23.SelStart = 1
            End Select
            WTexto23.Visible = True
            WTexto23.SetFocus
        Case 2
            WTexto33.Left = WVector3.CellLeft + WVector3.Left
            WTexto33.Top = WVector3.CellTop + WVector3.Top
            WTexto33.Width = WVector3.CellWidth
            WTexto33.Height = WVector3.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector3.Text) = 10 Then
                        WTexto33.Text = WVector3.Text
                            Else
                        WTexto33.Text = "  /  /    "
                    End If
                    WTexto33.SelStart = 0
                Case Else
                    WTexto33.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto33.SelStart = 1
            End Select
            WTexto33.Visible = True
            WTexto33.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditIII()
    Pasa = 0
    If WCombo13.Visible Then
        Pasa = 0
        WVector3.Text = WCombo13.Text
        WCombo13.Visible = False
            Else
        If WTexto13.Visible Then
            Pasa = 1
            WVector3.Text = WTexto13.Text
            WTexto13.Visible = False
                Else
            If WTexto23.Visible Then
                Pasa = 1
                WVector3.Text = WTexto23.Text
                WTexto23.Visible = False
                    Else
                If WTexto33.Visible Then
                    Pasa = 1
                    WVector3.Text = WTexto33.Text
                    WTexto33.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoIII(WVector3.Col) <> "" Then
            WVector3.Text = Pusing(WFormatoIII(WVector3.Col), WVector3.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboIII()
    ' Position the ComboBox over the cell.
    WCombo13.Left = WVector3.CellLeft + WVector3.Left
    WCombo13.Top = WVector3.CellTop + WVector3.Top
    WCombo13.Width = WVector3.CellWidth
    WCombo13.Visible = True
    WCombo13.SetFocus
End Sub

Private Sub WTexto13_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto13.Text = ""
            
        Rem F1
        Case 113
            WTexto13.Text = WVector3.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector3.SetFocus
            DoEvents
            Call Control_CampoIII
            If WControlIII = "S" Then
                Call Control_WVectorIII
            End If
            Call StartEditIII

        Case vbKeyDown
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row < WVector3.Rows - 1 Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row + 1
                End If
            End If
            Call StartEditIII

        Case vbKeyUp
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row > WVector3.FixedRows Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row - 1
                End If
            End If
            Call StartEditIII
        Case 34
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow < WVector3.Rows - 12 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow + 12
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 33
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow - 12 > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow - 12
                    WVector3.Row = WVector3.TopRow
                        Else
                    WVector3.TopRow = 1
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 123
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Col > 1 Then
                WVector3.Col = WVector3.Col - 1
            End If
            Call StartEditIII

    End Select
End Sub

Private Sub WTexto23_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto23.Text = ""
            
        Rem F1
        Case 113
            WTexto23.Text = WVector3.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector3.SetFocus
            DoEvents
            Call Control_CampoIII
            If WControlIII = "S" Then
                Call Control_WVectorIII
            End If
            Call StartEditIII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row < WVector3.Rows - 1 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row + 1
                Rem End If
            End If
            Call StartEditIII

        Case vbKeyUp
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row - 1
                Rem End If
            End If
            Call StartEditIII
        Case 34
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow < WVector3.Rows - 12 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow + 12
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 33
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow - 12 > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow - 12
                    WVector3.Row = WVector3.TopRow
                        Else
                    WVector3.TopRow = 1
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII

    End Select
End Sub

Private Sub WTexto33_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto33.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto33.Text = WVector3.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector3.SetFocus
            Call Control_CampoIII
            If WControlIII = "S" Then
                Call Control_WVectorIII
            End If
            Call StartEditIII

        Case vbKeyDown
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row < WVector3.Rows - 1 Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row + 1
                End If
            End If
            Call StartEditIII

        Case vbKeyUp
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row > WVector3.FixedRows Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row - 1
                End If
            End If
            Call StartEditIII
        Case 34
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow < WVector3.Rows - 12 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow + 12
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 33
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow - 12 > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow - 12
                    WVector3.Row = WVector3.TopRow
                        Else
                    WVector3.TopRow = 1
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto13_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto23_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto33_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo13_Click()
    WVector3.SetFocus
End Sub


Private Sub WVector3_Click()
    StartEditIII
End Sub

Private Sub WVector3_LeaveCell()
    EndEditIII
End Sub

Private Sub WVector3_GotFocus()
    EndEditIII
End Sub

Private Sub WVector3_KeyPress(KeyAscii As Integer)
    XColumna = WVector3.Col
    Select Case WParametrosIII(4, WVector3.Col)
        Case 1
        Case Else
            If WParametrosIII(2, XColumna) = 0 Then
                GridEditTextIII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditIII()
    Select Case WParametrosIII(4, WVector3.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo13.Clear
            WCombo13.AddItem "Campo1"
            WCombo13.AddItem "Campo2"
            On Error Resume Next
            WCombo13.Text = WVector3.Text
            On Error GoTo 0
            GridEditComboIII
        Case Else
            If WParametrosIII(2, WVector3.Col) = 0 Then
                GridEditTextIII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVectorIII()
    Select Case WVector3.Col
        Case 4
            If WVector3.Row < WVector3.Rows - 1 Then
                WVector3.Row = WVector3.Row + 1
            End If
            WVector3.Col = 1
        Case Else
            If WVector3.Col < WVector3.Cols - 1 Then
                WVector3.Col = WVector3.Col + 1
            End If
    End Select
    WVector3.SetFocus
    GridEditTextIII KeyAscii
End Sub

Private Sub Control_CampoIII()
    XColumna = WVector3.Col
    XFila = WVector3.Row
    WControlIII = "S"
    Select Case XColumna
        Case 1
            If Val(WVector3.Text) <> 0 Then
            
                ZCodigo = WVector3.Text
                
                Sql1 = "Select *"
                Sql2 = " FROM Ensayos"
                Sql3 = " Where Ensayos.Codigo = " + "'" + ZCodigo + "'"
                spEnsayo = Sql1 + Sql2 + Sql3
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    WVector3.Col = 2
                    WVector3.Text = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
                
            End If
            
        Case Else
            WVector3.Col = XColumna
    End Select
End Sub

Private Sub Limpia_VectorIII()

    WVector3.Clear

    Rem ponga la WVector3 en negritas
    WVector3.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto13.FontName = WVector3.FontName
    WTexto13.FontSize = WVector3.FontSize
    WTexto13.Visible = False
    WTexto23.FontName = WVector3.FontName
    WTexto23.FontSize = WVector3.FontSize
    WTexto23.Visible = False
    WTexto33.FontName = WVector3.FontName
    WTexto33.FontSize = WVector3.FontSize
    WTexto33.Visible = False
    WCombo13.FontName = WVector3.FontName
    WCombo13.FontSize = WVector3.FontSize
    WCombo13.Visible = False

    ' Establesco loa Valores de la WVector3
    
    WVector3.FixedCols = 1
    WVector3.Cols = 5
    WVector3.FixedRows = 1
    WVector3.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector3.Text = "Articulo"
    
    Rem Longitud
    Rem WVector3.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector3.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosIII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosIII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosIII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosIII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector3.ColWidth(0) = 400
    WVector3.Row = 0
    For Ciclo = 1 To WVector3.Cols - 1
        WVector3.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector3.Text = "Ensayo"
                WVector3.ColWidth(Ciclo) = 900
                WVector3.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosIII(1, Ciclo) = 10
                WParametrosIII(2, Ciclo) = 0
                WParametrosIII(3, Ciclo) = 1
                WParametrosIII(4, Ciclo) = 0
                WFormatoIII(Ciclo) = ""
            Case 2
                WVector3.Text = "Descripcion"
                WVector3.ColWidth(Ciclo) = 3000
                WVector3.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIII(1, Ciclo) = 50
                WParametrosIII(2, Ciclo) = 1
                WParametrosIII(3, Ciclo) = 0
                WParametrosIII(4, Ciclo) = 0
                WFormatoIII(Ciclo) = ""
            Case 3
                WVector3.Text = "Requerido"
                WVector3.ColWidth(Ciclo) = 3000
                WVector3.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIII(1, Ciclo) = 50
                WParametrosIII(2, Ciclo) = 0
                WParametrosIII(3, Ciclo) = 0
                WParametrosIII(4, Ciclo) = 0
                WFormatoIII(Ciclo) = ""
            Case 4
                WVector3.Text = "Resultado"
                WVector3.ColWidth(Ciclo) = 3000
                WVector3.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIII(1, Ciclo) = 50
                WParametrosIII(2, Ciclo) = 0
                WParametrosIII(3, Ciclo) = 0
                WParametrosIII(4, Ciclo) = 0
                WFormatoIII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector3.Row = 0
    For Ciclo = 1 To WVector3.Cols - 1
        WVector3.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector3.Text
        Rem WTitulo(Ciclo).Left = WVector3.CellLeft + WVector3.Left
        Rem WTitulo(Ciclo).Top = WVector3.CellTop + WVector3.Top
        Rem WTitulo(Ciclo).Width = WVector3.CellWidth
        Rem WTitulo(Ciclo).Height = WVector3.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector3
    
    WAncho = 400
    For Ciclo = 0 To WVector3.Cols - 1
        WAncho = WAncho + WVector3.ColWidth(Ciclo)
    Next Ciclo
    WVector3.Width = WAncho

    ' Size the columns.
    Font.Name = WVector3.Font.Name
    Font.Size = WVector3.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector3.AllowUserResizing = flexResizeBoth
    
    WVector3.Col = 1
    WVector3.Row = 1
    
End Sub

Private Sub WVector3_Scroll()
    WTexto13.Visible = False
    WTexto23.Visible = False
    WTexto33.Visible = False
End Sub
























Rem
Rem Controles de la WVector4
Rem

Private Sub GridEditTextIV(ByVal KeyAscii As Integer)

    XColumna = WVector4.Col
    XTipoDato = WParametrosIV(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto14.Left = WVector4.CellLeft + WVector4.Left
            WTexto14.Top = WVector4.CellTop + WVector4.Top
            WTexto14.Width = WVector4.CellWidth
            WTexto14.Height = WVector4.CellHeight
            WTexto14.MaxLength = WParametrosIV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto14.Text = WVector4.Text
                    WTexto14.SelStart = Len(WTexto14.Text)
                Case Else
                    WTexto14.Text = Chr$(KeyAscii)
                    WTexto14.SelStart = 1
            End Select
            WTexto14.Visible = True
            WTexto14.SetFocus
        Case 1
            WTexto24.Left = WVector4.CellLeft + WVector4.Left
            WTexto24.Top = WVector4.CellTop + WVector4.Top
            WTexto24.Width = WVector4.CellWidth
            WTexto24.Height = WVector4.CellHeight
            WTexto24.MaxLength = WParametrosIV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto24.Text = WVector4.Text
                    Rem WTexto24.SelStart = Len(WTexto24.Text)
                    WTexto24.SelStart = 0
                Case Else
                    WTexto24.Text = Chr$(KeyAscii)
                    WTexto24.SelStart = 1
            End Select
            WTexto24.Visible = True
            WTexto24.SetFocus
        Case 2
            WTexto34.Left = WVector4.CellLeft + WVector4.Left
            WTexto34.Top = WVector4.CellTop + WVector4.Top
            WTexto34.Width = WVector4.CellWidth
            WTexto34.Height = WVector4.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector4.Text) = 10 Then
                        WTexto34.Text = WVector4.Text
                            Else
                        WTexto34.Text = "  /  /    "
                    End If
                    WTexto34.SelStart = 0
                Case Else
                    WTexto34.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto34.SelStart = 1
            End Select
            WTexto34.Visible = True
            WTexto34.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditIV()
    Pasa = 0
    If WCombo14.Visible Then
        Pasa = 0
        WVector4.Text = WCombo14.Text
        WCombo14.Visible = False
            Else
        If WTexto14.Visible Then
            Pasa = 1
            WVector4.Text = WTexto14.Text
            WTexto14.Visible = False
                Else
            If WTexto24.Visible Then
                Pasa = 1
                WVector4.Text = WTexto24.Text
                WTexto24.Visible = False
                    Else
                If WTexto34.Visible Then
                    Pasa = 1
                    WVector4.Text = WTexto34.Text
                    WTexto34.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoIV(WVector4.Col) <> "" Then
            WVector4.Text = Pusing(WFormatoIV(WVector4.Col), WVector4.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboIV()
    ' Position the ComboBox over the cell.
    WCombo14.Left = WVector4.CellLeft + WVector4.Left
    WCombo14.Top = WVector4.CellTop + WVector4.Top
    WCombo14.Width = WVector4.CellWidth
    WCombo14.Visible = True
    WCombo14.SetFocus
End Sub

Private Sub WTexto14_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto14.Text = ""
            
        Rem F1
        Case 113
            WTexto14.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            DoEvents
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV

        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 123
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Col > 1 Then
                WVector4.Col = WVector4.Col - 1
            End If
            Call StartEditIV

    End Select
End Sub

Private Sub WTexto24_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto24.Text = ""
            
        Rem F1
        Case 113
            WTexto24.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            DoEvents
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                Rem End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                Rem End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV

    End Select
End Sub

Private Sub WTexto34_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto34.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto34.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV

        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto14_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto24_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto34_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo14_Click()
    WVector4.SetFocus
End Sub


Private Sub WVector4_Click()
    StartEditIV
End Sub

Private Sub WVector4_LeaveCell()
    EndEditIV
End Sub

Private Sub WVector4_GotFocus()
    EndEditIV
End Sub

Private Sub WVector4_KeyPress(KeyAscii As Integer)
    XColumna = WVector4.Col
    Select Case WParametrosIV(4, WVector4.Col)
        Case 1
        Case Else
            If WParametrosIV(2, XColumna) = 0 Then
                GridEditTextIV KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditIV()
    Select Case WParametrosIV(4, WVector4.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo14.Clear
            WCombo14.AddItem "Campo1"
            WCombo14.AddItem "Campo2"
            On Error Resume Next
            WCombo14.Text = WVector4.Text
            On Error GoTo 0
            GridEditComboIV
        Case Else
            If WParametrosIV(2, WVector4.Col) = 0 Then
                GridEditTextIV Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVectorIV()
    Select Case WVector4.Col
        Case 8
            If WVector4.Row < WVector4.Rows - 1 Then
                WVector4.Row = WVector4.Row + 1
            End If
            WVector4.Col = 1
        Case Else
            If WVector4.Col < WVector4.Cols - 1 Then
                WVector4.Col = WVector4.Col + 1
            End If
    End Select
    WVector4.SetFocus
    GridEditTextIV KeyAscii
End Sub

Private Sub Control_CampoIV()
    XColumna = WVector4.Col
    XFila = WVector4.Row
    WControlIV = "S"
End Sub

Private Sub WVector4_DblClick()

    If WVector4.Col = 0 Or WVector4.Col = 1 Then
    
    WTexto14.Visible = False
    WTexto24.Visible = False
    WTexto34.Visible = False
    
    RenglonAuxiliar = WVector4.Row

    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        WVector4.Text = ""
    Next Ciclo
    
    Erase WBorraIV
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 99 To 1 Step -1
        
        ZEtapa = WVector4.TextMatrix(iRow, 1)
        ZFecha = WVector4.TextMatrix(iRow, 2)
        ZParticipantes = WVector4.TextMatrix(iRow, 3)
        ZResultados = WVector4.TextMatrix(iRow, 4)
        ZAcciones = WVector4.TextMatrix(iRow, 5)
        ZResponsables = WVector4.TextMatrix(iRow, 6)
        ZEstado = WVector4.TextMatrix(iRow, 7)
            
        If ZEtapa <> "" Or ZFecha <> "" Or ZParticipantes <> "" Or ZResultados <> "" Or ZAcciones <> "" Or ZResponsables <> "" Or ZEstado <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector4.Row = Ciclo
        WVector4.Col = 1
        WAuxi1 = WVector4.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector4.Cols - 1
                WVector4.Col = Ciclo1
                WBorraIV(EntraVector, Ciclo1) = WVector4.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorIV
    
    For Ciclo = 1 To EntraVector
        WVector4.Row = Ciclo
        For DA = 0 To WVector4.Cols - 1
            WVector4.Col = DA
            WVector4.Text = WBorraIV(Ciclo, DA)
        Next DA
    Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglonIV_Click()

    Hasta = WVector4.Row

    For iRow = 1000 To Hasta Step -1
        WVector4.TextMatrix(iRow, 0) = WVector4.TextMatrix(iRow - 1, 0)
        WVector4.TextMatrix(iRow, 1) = WVector4.TextMatrix(iRow - 1, 1)
        WVector4.TextMatrix(iRow, 2) = WVector4.TextMatrix(iRow - 1, 2)
        WVector4.TextMatrix(iRow, 3) = WVector4.TextMatrix(iRow - 1, 3)
        WVector4.TextMatrix(iRow, 4) = WVector4.TextMatrix(iRow - 1, 4)
        WVector4.TextMatrix(iRow, 5) = WVector4.TextMatrix(iRow - 1, 5)
        WVector4.TextMatrix(iRow, 6) = WVector4.TextMatrix(iRow - 1, 6)
        WVector4.TextMatrix(iRow, 7) = WVector4.TextMatrix(iRow - 1, 7)
        WVector4.TextMatrix(iRow, 8) = WVector4.TextMatrix(iRow - 1, 8)
    Next iRow

    WVector4.TextMatrix(Hasta, 0) = ""
    WVector4.TextMatrix(Hasta, 1) = ""
    WVector4.TextMatrix(Hasta, 2) = ""
    WVector4.TextMatrix(Hasta, 3) = ""
    WVector4.TextMatrix(Hasta, 4) = ""
    WVector4.TextMatrix(Hasta, 5) = ""
    WVector4.TextMatrix(Hasta, 6) = ""
    WVector4.TextMatrix(Hasta, 7) = ""
    WVector4.TextMatrix(Hasta, 8) = ""
    
    WTexto14.Text = ""
    WTexto24.Text = ""

End Sub




Private Sub Limpia_VectorIV()

    WVector4.Clear

    Rem ponga la WVector4 en negritas
    WVector4.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto14.FontName = WVector4.FontName
    WTexto14.FontSize = WVector4.FontSize
    WTexto14.Visible = False
    WTexto24.FontName = WVector4.FontName
    WTexto24.FontSize = WVector4.FontSize
    WTexto24.Visible = False
    WTexto34.FontName = WVector4.FontName
    WTexto34.FontSize = WVector4.FontSize
    WTexto34.Visible = False
    WCombo14.FontName = WVector4.FontName
    WCombo14.FontSize = WVector4.FontSize
    WCombo14.Visible = False

    ' Establesco loa Valores de la WVector4
    
    WVector4.FixedCols = 1
    WVector4.Cols = 9
    WVector4.FixedRows = 1
    WVector4.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector4.Text = "Articulo"
    
    Rem Longitud
    Rem WVector4.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector4.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosIV(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosIV(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosIV(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosIV(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector4.ColWidth(0) = 200
    WVector4.Row = 0
    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector4.Text = "Version"
                WVector4.ColWidth(Ciclo) = 800
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 10
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 2
                WVector4.Text = "Etapa"
                WVector4.ColWidth(Ciclo) = 900
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 3
                WVector4.Text = "Fecha"
                WVector4.ColWidth(Ciclo) = 1000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 10
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 4
                WVector4.Text = "Participantes"
                WVector4.ColWidth(Ciclo) = 2000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 5
                WVector4.Text = "Resultados"
                WVector4.ColWidth(Ciclo) = 4000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 6
                WVector4.Text = "Acciones"
                WVector4.ColWidth(Ciclo) = 4000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 7
                WVector4.Text = "Responsables"
                WVector4.ColWidth(Ciclo) = 2000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 8
                WVector4.Text = "Estado"
                WVector4.ColWidth(Ciclo) = 1200
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector4.Row = 0
    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector4.Text
        Rem WTitulo(Ciclo).Left = WVector4.CellLeft + WVector4.Left
        Rem WTitulo(Ciclo).Top = WVector4.CellTop + WVector4.Top
        Rem WTitulo(Ciclo).Width = WVector4.CellWidth
        Rem WTitulo(Ciclo).Height = WVector4.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector4
    
    WAncho = 400
    For Ciclo = 0 To WVector4.Cols - 1
        WAncho = WAncho + WVector4.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector4.Width = WAncho

    ' Size the columns.
    Font.Name = WVector4.Font.Name
    Font.Size = WVector4.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector4.AllowUserResizing = flexResizeBoth
    
    WVector4.Col = 1
    WVector4.Row = 1
    
End Sub

Private Sub WVector4_Scroll()
    WTexto14.Visible = False
    WTexto24.Visible = False
    WTexto34.Visible = False
End Sub










Rem
Rem Controles de la WVector5
Rem

Private Sub GridEditTextV(ByVal KeyAscii As Integer)

    XColumna = WVector5.Col
    XTipoDato = WParametrosV(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto15.Left = WVector5.CellLeft + WVector5.Left
            WTexto15.Top = WVector5.CellTop + WVector5.Top
            WTexto15.Width = WVector5.CellWidth
            WTexto15.Height = WVector5.CellHeight
            WTexto15.MaxLength = WParametrosV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto15.Text = WVector5.Text
                    WTexto15.SelStart = Len(WTexto15.Text)
                Case Else
                    WTexto15.Text = Chr$(KeyAscii)
                    WTexto15.SelStart = 1
            End Select
            WTexto15.Visible = True
            WTexto15.SetFocus
        Case 1
            WTexto25.Left = WVector5.CellLeft + WVector5.Left
            WTexto25.Top = WVector5.CellTop + WVector5.Top
            WTexto25.Width = WVector5.CellWidth
            WTexto25.Height = WVector5.CellHeight
            WTexto25.MaxLength = WParametrosV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto25.Text = WVector5.Text
                    Rem WTexto25.SelStart = Len(WTexto25.Text)
                    WTexto25.SelStart = 0
                Case Else
                    WTexto25.Text = Chr$(KeyAscii)
                    WTexto25.SelStart = 1
            End Select
            WTexto25.Visible = True
            WTexto25.SetFocus
        Case 2
            WTexto35.Left = WVector5.CellLeft + WVector5.Left
            WTexto35.Top = WVector5.CellTop + WVector5.Top
            WTexto35.Width = WVector5.CellWidth
            WTexto35.Height = WVector5.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector5.Text) = 10 Then
                        WTexto35.Text = WVector5.Text
                            Else
                        WTexto35.Text = "  /  /    "
                    End If
                    WTexto35.SelStart = 0
                Case Else
                    WTexto35.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto35.SelStart = 1
            End Select
            WTexto35.Visible = True
            WTexto35.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditV()
    Pasa = 0
    If WCombo15.Visible Then
        Pasa = 0
        WVector5.Text = WCombo15.Text
        WCombo15.Visible = False
            Else
        If WTexto15.Visible Then
            Pasa = 1
            WVector5.Text = WTexto15.Text
            WTexto15.Visible = False
                Else
            If WTexto25.Visible Then
                Pasa = 1
                WVector5.Text = WTexto25.Text
                WTexto25.Visible = False
                    Else
                If WTexto35.Visible Then
                    Pasa = 1
                    WVector5.Text = WTexto35.Text
                    WTexto35.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoV(WVector5.Col) <> "" Then
            WVector5.Text = Pusing(WFormatoV(WVector5.Col), WVector5.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboV()
    ' Position the ComboBox over the cell.
    WCombo15.Left = WVector5.CellLeft + WVector5.Left
    WCombo15.Top = WVector5.CellTop + WVector5.Top
    WCombo15.Width = WVector5.CellWidth
    WCombo15.Visible = True
    WCombo15.SetFocus
End Sub

Private Sub WTexto15_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto15.Text = ""
            
        Rem F1
        Case 113
            WTexto15.Text = WVector5.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector5.SetFocus
            DoEvents
            Call Control_CampoV
            If WControlV = "S" Then
                Call Control_WVectorV
            End If
            Call StartEditV

        Case vbKeyDown
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row < WVector5.Rows - 1 Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row + 1
                End If
            End If
            Call StartEditV

        Case vbKeyUp
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row > WVector5.FixedRows Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row - 1
                End If
            End If
            Call StartEditV
        Case 34
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow < WVector5.Rows - 12 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow + 12
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 33
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow - 12 > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow - 12
                    WVector5.Row = WVector5.TopRow
                        Else
                    WVector5.TopRow = 1
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 123
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Col > 1 Then
                WVector5.Col = WVector5.Col - 1
            End If
            Call StartEditV

    End Select
End Sub

Private Sub WTexto25_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto25.Text = ""
            
        Rem F1
        Case 113
            WTexto25.Text = WVector5.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector5.SetFocus
            DoEvents
            Call Control_CampoV
            If WControlV = "S" Then
                Call Control_WVectorV
            End If
            Call StartEditV
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row < WVector5.Rows - 1 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.Row = WVector5.Row + 1
                Rem End If
            End If
            Call StartEditV

        Case vbKeyUp
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.Row = WVector5.Row - 1
                Rem End If
            End If
            Call StartEditV
        Case 34
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow < WVector5.Rows - 12 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow + 12
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 33
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow - 12 > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow - 12
                    WVector5.Row = WVector5.TopRow
                        Else
                    WVector5.TopRow = 1
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV

    End Select
End Sub

Private Sub WTexto35_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto35.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto35.Text = WVector5.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector5.SetFocus
            Call Control_CampoV
            If WControlV = "S" Then
                Call Control_WVectorV
            End If
            Call StartEditV

        Case vbKeyDown
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row < WVector5.Rows - 1 Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row + 1
                End If
            End If
            Call StartEditV

        Case vbKeyUp
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row > WVector5.FixedRows Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row - 1
                End If
            End If
            Call StartEditV
        Case 34
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow < WVector5.Rows - 12 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow + 12
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 33
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow - 12 > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow - 12
                    WVector5.Row = WVector5.TopRow
                        Else
                    WVector5.TopRow = 1
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto15_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto25_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto35_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo15_Click()
    WVector5.SetFocus
End Sub


Private Sub WVector5_Click()
    StartEditV
End Sub

Private Sub WVector5_LeaveCell()
    EndEditV
End Sub

Private Sub WVector5_GotFocus()
    EndEditV
End Sub

Private Sub WVector5_KeyPress(KeyAscii As Integer)
    XColumna = WVector5.Col
    Select Case WParametrosV(4, WVector5.Col)
        Case 1
        Case Else
            If WParametrosV(2, XColumna) = 0 Then
                GridEditTextV KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditV()
    Select Case WParametrosV(4, WVector5.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo15.Clear
            WCombo15.AddItem "Campo1"
            WCombo15.AddItem "Campo2"
            On Error Resume Next
            WCombo15.Text = WVector5.Text
            On Error GoTo 0
            GridEditComboV
        Case Else
            If WParametrosV(2, WVector5.Col) = 0 Then
                GridEditTextV Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVectorV()
    Select Case WVector5.Col
        Case 6
            If WVector5.Row < WVector5.Rows - 1 Then
                WVector5.Row = WVector5.Row + 1
            End If
            Rem WVector5.Col = 1
        Case Else
            If WVector5.Col < WVector5.Cols - 1 Then
                WVector5.Col = WVector5.Col + 1
            End If
    End Select
    WVector5.SetFocus
    GridEditTextV KeyAscii
End Sub

Private Sub Control_CampoV()
    XColumna = WVector5.Col
    XFila = WVector5.Row
    WControlV = "S"
    Call Recalcula_Costo
End Sub

Private Sub Limpia_VectorV()

    WVector5.Clear

    Rem ponga la WVector5 en negritas
    WVector5.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto15.FontName = WVector5.FontName
    WTexto15.FontSize = WVector5.FontSize
    WTexto15.Visible = False
    WTexto25.FontName = WVector5.FontName
    WTexto25.FontSize = WVector5.FontSize
    WTexto25.Visible = False
    WTexto35.FontName = WVector5.FontName
    WTexto35.FontSize = WVector5.FontSize
    WTexto35.Visible = False
    WCombo15.FontName = WVector5.FontName
    WCombo15.FontSize = WVector5.FontSize
    WCombo15.Visible = False

    ' Establesco loa Valores de la WVector5
    
    WVector5.FixedCols = 1
    WVector5.Cols = 8
    WVector5.FixedRows = 1
    WVector5.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector5.Text = "Articulo"
    
    Rem Longitud
    Rem WVector5.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector5.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosV(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosV(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosV(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosV(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector5.ColWidth(0) = 200
    WVector5.Row = 0
    For Ciclo = 1 To WVector5.Cols - 1
        WVector5.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector5.Text = "Tipo"
                WVector5.ColWidth(Ciclo) = 600
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 1
                WParametrosV(2, Ciclo) = 1
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 2
                WVector5.Text = "Articulo"
                WVector5.ColWidth(Ciclo) = 1300
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 10
                WParametrosV(2, Ciclo) = 1
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 3
                WVector5.Text = "Terminado"
                WVector5.ColWidth(Ciclo) = 1500
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 12
                WParametrosV(2, Ciclo) = 1
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 4
                WVector5.Text = "Descripcion"
                WVector5.ColWidth(Ciclo) = 3500
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 50
                WParametrosV(2, Ciclo) = 1
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 5
                WVector5.Text = "Cantidad"
                WVector5.ColWidth(Ciclo) = 1100
                WVector5.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosV(1, Ciclo) = 10
                WParametrosV(2, Ciclo) = 1
                WParametrosV(3, Ciclo) = 1
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = "###,###.###"
            Case 6
                WVector5.Text = "Costo"
                WVector5.ColWidth(Ciclo) = 1100
                WVector5.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosV(1, Ciclo) = 10
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 1
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = "###,###.##"
            Case 7
                WVector5.Text = "Importe"
                WVector5.ColWidth(Ciclo) = 1100
                WVector5.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosV(1, Ciclo) = 10
                WParametrosV(2, Ciclo) = 1
                WParametrosV(3, Ciclo) = 1
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = "###,###.##"
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector5.Row = 0
    For Ciclo = 1 To WVector5.Cols - 1
        WVector5.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector5.Text
        Rem WTitulo(Ciclo).Left = WVector5.CellLeft + WVector5.Left
        Rem WTitulo(Ciclo).Top = WVector5.CellTop + WVector5.Top
        Rem WTitulo(Ciclo).Width = WVector5.CellWidth
        Rem WTitulo(Ciclo).Height = WVector5.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector5
    
    WAncho = 400
    For Ciclo = 0 To WVector5.Cols - 1
        WAncho = WAncho + WVector5.ColWidth(Ciclo)
    Next Ciclo
    WVector5.Width = WAncho

    ' Size the columns.
    Font.Name = WVector5.Font.Name
    Font.Size = WVector5.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector5.AllowUserResizing = flexResizeBoth
    
    WVector5.Col = 1
    WVector5.Row = 1
    
End Sub

Private Sub WVector5_Scroll()
    WTexto15.Visible = False
    WTexto25.Visible = False
    WTexto35.Visible = False
End Sub

Private Sub Recalcula_Click()

    For Ciclo = 1 To 99
    
        WVector5.TextMatrix(Ciclo, 1) = WVector1.TextMatrix(Ciclo, 1)
        WVector5.TextMatrix(Ciclo, 2) = WVector1.TextMatrix(Ciclo, 2)
        WVector5.TextMatrix(Ciclo, 3) = WVector1.TextMatrix(Ciclo, 3)
        WVector5.TextMatrix(Ciclo, 4) = WVector1.TextMatrix(Ciclo, 4)
        WVector5.TextMatrix(Ciclo, 5) = WVector1.TextMatrix(Ciclo, 5)
        WVector5.TextMatrix(Ciclo, 6) = ""
        WVector5.TextMatrix(Ciclo, 7) = ""
        
        Select Case WVector5.TextMatrix(Ciclo, 1)
            Case "T"
                Producto = WVector5.TextMatrix(Ciclo, 3)
                Call Calcula_Costo(Producto, XCosto2)
                Rem Call Calcula_Costo_Produccion(Producto, XCosto1, XCosto2, XCosto3)
                WVector5.TextMatrix(Ciclo, 6) = Str$(XCosto2)
                WVector5.TextMatrix(Ciclo, 6) = Pusing("###,###.##", Val(WVector5.TextMatrix(Ciclo, 6)))
                
                If XCosto2 = 0 Then
                    m$ = "El costo de la producto terminado " + Producto + "esta en cero"
                    a% = MsgBox(m$, 0, "Calculo de Costos")
                End If
                
                
            Case "M"
                Articulo1 = WVector5.TextMatrix(Ciclo, 2)
                spArticulo = "ConsultaArticulo " + "'" + Articulo1 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DescriArticulo1 = Left$(rstArticulo!Descripcion, 30)
                    Select Case TipoCosto.ListIndex
                        Case 0
                            Costo = rstArticulo!Costo2
                            rstArticulo.Close
                        Case Else
                            Costo = rstArticulo!Costo1
    
                            Costo1 = rstArticulo!Costo1
                            WWCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                            ZCosto1 = IIf(IsNull(rstArticulo!ZCosto1), "0", rstArticulo!ZCosto1)
                            ZZOrdenI = IIf(IsNull(rstArticulo!OrdenI), "0", rstArticulo!OrdenI)
                            ZZOrdenII = IIf(IsNull(rstArticulo!OrdenII), "0", rstArticulo!OrdenII)
                            ZZOrdenIII = IIf(IsNull(rstArticulo!OrdenIII), "0", rstArticulo!OrdenIII)
                            ZZPtaOrdenI = IIf(IsNull(rstArticulo!PtaOrdenI), "0", rstArticulo!PtaOrdenI)
                            ZZPtaOrdenII = IIf(IsNull(rstArticulo!PtaOrdenII), "0", rstArticulo!PtaOrdenII)
                            ZZPtaOrdenIII = IIf(IsNull(rstArticulo!PtaOrdenIII), "0", rstArticulo!PtaOrdenIII)
                                
                            ZZFechaOrdenI = ""
                            ZZFechaOrdenII = ""
                            ZZFechaOrdenIII = ""
                            
                            ZZMoneda = ""
                            
                            rstArticulo.Close
                            
                            XEmpresa = WEmpresa
    
                            If ZZPtaOrdenI <> 0 And ZZOrdenI <> 0 Then
                            
                                Select Case ZZPtaOrdenI
                                    Case 1
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 2
                                        WEmpresa = "0002"
                                        txtOdbc = "Empresa02"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 3
                                        WEmpresa = "0003"
                                        txtOdbc = "Empresa03"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 4
                                        WEmpresa = "0004"
                                        txtOdbc = "Empresa04"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 5
                                        WEmpresa = "0005"
                                        txtOdbc = "Empresa05"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 6
                                        WEmpresa = "0006"
                                        txtOdbc = "Empresa06"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 7
                                        WEmpresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 8
                                        WEmpresa = "0008"
                                        txtOdbc = "Empresa08"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 9
                                        WEmpresa = "0009"
                                        txtOdbc = "Empresa09"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 10
                                        WEmpresa = "0010"
                                        txtOdbc = "Empresa10"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 11
                                        WEmpresa = "0011"
                                        txtOdbc = "Empresa11"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                End Select
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenI) + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZFechaOrdenI = rstOrden!Fecha
                                    Select Case rstOrden!Moneda
                                        Case 0
                                            ZZMoneda = "U$S"
                                        Case 1
                                            ZZMoneda = "$"
                                        Case Else
                                            ZZMoneda = "Eur"
                                    End Select
                                    rstOrden.Close
                                End If
                                
                                Call Conecta_Empresa
                                
                            End If
                            
                            
                            If ZZPtaOrdenII <> 0 And ZZOrdenII <> 0 Then
                                
                                Select Case ZZPtaOrdenII
                                    Case 1
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 2
                                        WEmpresa = "0002"
                                        txtOdbc = "Empresa02"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 3
                                        WEmpresa = "0003"
                                        txtOdbc = "Empresa03"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 4
                                        WEmpresa = "0004"
                                        txtOdbc = "Empresa04"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 5
                                        WEmpresa = "0005"
                                        txtOdbc = "Empresa05"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 6
                                        WEmpresa = "0006"
                                        txtOdbc = "Empresa06"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 7
                                        WEmpresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 8
                                        WEmpresa = "0008"
                                        txtOdbc = "Empresa08"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 9
                                        WEmpresa = "0009"
                                        txtOdbc = "Empresa09"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 10
                                        WEmpresa = "0010"
                                        txtOdbc = "Empresa10"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 11
                                        WEmpresa = "0011"
                                        txtOdbc = "Empresa11"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                End Select
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenII) + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZFechaOrdenII = rstOrden!Fecha
                                    rstOrden.Close
                                End If
                                
                                Call Conecta_Empresa
                                
                            End If
                            
                            If ZZPtaOrdenIII <> 0 And ZZOrdenIII <> 0 Then
                                
                                Select Case ZZPtaOrdenIII
                                    Case 1
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 2
                                        WEmpresa = "0002"
                                        txtOdbc = "Empresa02"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 3
                                        WEmpresa = "0003"
                                        txtOdbc = "Empresa03"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 4
                                        WEmpresa = "0004"
                                        txtOdbc = "Empresa04"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 5
                                        WEmpresa = "0005"
                                        txtOdbc = "Empresa05"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 6
                                        WEmpresa = "0006"
                                        txtOdbc = "Empresa06"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 7
                                        WEmpresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 8
                                        WEmpresa = "0008"
                                        txtOdbc = "Empresa08"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 9
                                        WEmpresa = "0009"
                                        txtOdbc = "Empresa09"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 10
                                        WEmpresa = "0010"
                                        txtOdbc = "Empresa10"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 11
                                        WEmpresa = "0011"
                                        txtOdbc = "Empresa11"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                End Select
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenIII) + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZFechaOrdenIII = rstOrden!Fecha
                                    rstOrden.Close
                                End If
                                
                                Call Conecta_Empresa
                                
                            End If
                            
                            Rem DADA
                            Rem spCambio = "ConsultaCambio " + "'" + ZZFecha + "'"
                            Rem Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                            Rem If rstCambio.RecordCount > 0 Then
                            Rem     ZZParidad = rstCambio!Cambio
                            Rem     rstCambio.Close
                            Rem End If
                            
                            Rem If ZZParidad <> 0 Then
                            Rem     ZZCostoPesos = WPrecio * ZZParidad
                            Rem End If
                                    
                            If ZZFechaOrdenI <> "" Then
                                WFechaOrdI = Right$(ZZFechaOrdenI, 4) + Mid$(ZZFechaOrdenI, 4, 2) + Left$(ZZFechaOrdenI, 2)
                                    Else
                                WFechaOrdI = ""
                            End If
                            If ZZFechaOrdenII <> "" Then
                                WFechaOrdII = Right$(ZZFechaOrdenII, 4) + Mid$(ZZFechaOrdenII, 4, 2) + Left$(ZZFechaOrdenII, 2)
                                    Else
                                WFechaOrdII = ""
                            End If
                            If ZZFechaOrdenIII <> "" Then
                                WFechaOrdIII = Right$(ZZFechaOrdenIII, 4) + Mid$(ZZFechaOrdenIII, 4, 2) + Left$(ZZFechaOrdenIII, 2)
                                    Else
                                WFechaOrdIII = ""
                            End If
                            
                            If WFechaOrdI <> "" And WFechaOrdI > WFechaOrdII And WFechaOrdI > WFechaOrdIII Then
                                Costo = Costo1
                            End If
                            
                            If WFechaOrdII <> "" And WFechaOrdII > WFechaOrdI And WFechaOrdII > WFechaOrdIII Then
                            
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                                 spCambios = "ConsultaCambio  " + "'" + ZZFechaOrdenII + "'"
                                 Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                                 If rstCambios.RecordCount > 0 Then
                                     ZZZZParidad = rstCambios!Cambio
                                     rstCambios.Close
                                     If ZZZZParidad <> 0 Then
                                         ZZCosto1Dol = WWCosto1 / ZZZZParidad
                                     End If
                                 End If
                                Call Conecta_Empresa
                            
                                Costo = ZZCosto1Dol
                            End If
                            
                            If WFechaOrdIII <> "" And WFechaOrdIII > WFechaOrdI And WFechaOrdIII > WFechaOrdII Then
                                Costo = ZCosto1
                            End If
                            WDescriTipo = ""
                                                    
                            Rem WEmpresa = "0001"
                            Rem txtOdbc = "Empresa01"
                            Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                            Rem spCambios = "ConsultaCambio  " + "'" + FechaOrdenII.Text + "'"
                            Rem Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                            Rem If rstCambios.RecordCount > 0 Then
                            Rem     ZZParidad = rstCambios!Cambio
                            Rem     rstCambios.Close
                            Rem     If ZZParidad <> 0 Then
                            Rem         ZZCosto1Dol = Val(WCosto1.Text) / ZZParidad
                            Rem         WCosto1Dol.Text = Str$(ZZCosto1Dol)
                            Rem         WCosto1Dol.Text = Pusing("###,###.##", WCosto1Dol.Text)
                            Rem     End If
                            Rem End If
                            
                    End Select
                    
                    WVector5.TextMatrix(Ciclo, 6) = Str$(Costo)
                    WVector5.TextMatrix(Ciclo, 6) = Pusing("###,###.##", Val(WVector5.TextMatrix(Ciclo, 6)))
                
                    If Costo = 0 Then
                        m$ = "El costo de la materia prima " + Articulo1 + " esta en cero"
                        a% = MsgBox(m$, 0, "Calculo de Costos")
                    End If
                    
                    
                End If
                
            Case Else
        End Select
        
    Next Ciclo
    
    Call Recalcula_Costo

End Sub


Private Sub Recalcula_Costo()

    WCosto = 0

    For Ciclo = 1 To 99
        WImpo = Val(WVector5.TextMatrix(Ciclo, 5)) * Val(WVector5.TextMatrix(Ciclo, 6))
        If WImpo <> 0 Then
            WVector5.TextMatrix(Ciclo, 7) = Str$(WImpo)
            WVector5.TextMatrix(Ciclo, 7) = Pusing("###,###.##", Val(WVector5.TextMatrix(Ciclo, 7)))
            WCosto = WCosto + WImpo
                Else
            WVector5.TextMatrix(Ciclo, 7) = ""
        End If
    Next Ciclo
    
    CostoTotal.Text = Str$(WCosto)
    If Val(Cantidad.Text) <> 0 Then
        CostoKilo.Text = Str$(WCosto / Val(Cantidad.Text))
            Else
        CostoKilo.Text = ""
    End If
        
    CostoTotal.Text = Pusing("###,###.##", Val(CostoTotal.Text))
    CostoKilo.Text = Pusing("###,###.##", Val(CostoKilo.Text))
        
End Sub


Private Sub Calcula_Costo_Produccion(ZProducto As String, ZCosto1 As Double, ZCosto2 As Double, ZCosto3 As Double)

    Dim ZVector(100, 2) As String
    
    Erase ZAuxiliar
    ZRenglon = 0
    
    ZVector(1, 1) = ZProducto
    ZVector(1, 2) = "1"
    ZCosto1 = 0
    ZCosto2 = 0
    ZCosto3 = 0
    ZLugar = 1
    ZCicla = 0
    
    Do
        ZCicla = ZCicla + 1
        If ZVector(ZCicla, 1) <> "" Then
    
            spComposicion = "ConsultaComposicionProducto " + "'" + ZVector(ZCicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZTipo = rstComposicion!Tipo
                        zarticulo1 = rstComposicion!Articulo1
                        ZArticulo2 = rstComposicion!Articulo2
                        ZCantidad = rstComposicion!Cantidad
                        
                        Select Case ZTipo
                            Case "T"
                                If ZProducto <> ZArticulo2 Then
                                    ZLugar = ZLugar + 1
                                    ZVector(ZLugar, 1) = ZArticulo2
                                    ZVector(ZLugar, 2) = Str$(ZCantidad * Val(ZVector(ZCicla, 2)))
                                End If
                            Case "M"
                                ZRenglon = ZRenglon + 1
                                ZAuxiliar(ZRenglon, 1) = zarticulo1
                                ZAuxiliar(ZRenglon, 2) = ZCantidad
                                ZAuxiliar(ZRenglon, 3) = ZVector(ZCicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
    
    ZCosto1 = 0
    ZCosto2 = 0
    ZCosto3 = 0
                    
    For ZDa = 1 To ZRenglon
        ZArticulo = ZAuxiliar(ZDa, 1)
        ZCantidad = ZAuxiliar(ZDa, 2)
        ZWVector = ZAuxiliar(ZDa, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCos1 = (ZCantidad * rstArticulo!Costo2 * Val(ZWVector))
            ZCosto1 = ZCosto1 + WCos1
            WCos2 = (ZCantidad * rstArticulo!Costo1 * Val(ZWVector))
            ZCosto2 = ZCosto2 + WCos2
            WCos3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
            WCos3 = (ZCantidad * WCos3 * Val(ZWVector))
            ZCosto3 = ZCosto3 + WCos3
            rstArticulo.Close
        End If
    Next ZDa
    
    Call Redondeo(XCosto1)
    Call Redondeo(XCosto2)
    Call Redondeo(XCosto3)
    
End Sub


Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Dim Auxiliar(100, 7) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        ZZCantidad = rstComposicion!Cantidad
                        
                        Rem If Left$(Articulo1, 2) = "DW" Then
                        Rem     Tipo = "T"
                        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        Rem End If
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(ZZCantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = ZZCantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
            Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
            Rem     Renglon = Renglon + 1
            Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
            Rem     Auxiliar(Renglon, 2) = 1
            Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For DA = 1 To Renglon
        Articulo = Auxiliar(DA, 1)
        ZZCantidad = Auxiliar(DA, 2)
        WVector = Auxiliar(DA, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            Rem Select Case TipoListado.ListIndex
            Rem     Case 0
            Rem         WCosto = (Cantidad * rstArticulo!Costo2 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(WVector))
            Rem     Case 1
            Rem         WCosto = (Cantidad * rstArticulo!Costo1 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * rstArticulo!Costo1 * Val(WVector))
            Rem     Case 2
            Rem         Costo3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
            Rem         WCosto = (Cantidad * Costo3 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * Costo3 * Val(WVector))
            Rem     Case 3
            Rem         Costo4 = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
            Rem         If Costo4 = 0 Then
            Rem             Costo4 = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
            Rem         End If
            Rem         WCosto = (Cantidad * Costo4 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * Costo4 * Val(WVector))
            Rem     Case 4
            Rem         Costo4 = IIf(IsNull(rstArticulo!UltimoFob), "0", rstArticulo!UltimoFob)
            Rem         If Costo4 = 0 Then
            Rem             Costo4 = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
            Rem         End If
            Rem         WCosto = (Cantidad * Costo4 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * Costo4 * Val(WVector))
            Rem     Case Else
            Rem         WCosto = 0
            Rem         Costo = 0
            Rem End Select
            
            Select Case TipoCosto.ListIndex
                Case 0
                    ZZCosto = rstArticulo!Costo2
                    ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                    If ZTipoCosto = 1 Then
                        WDescriTipo = "Estimado"
                            Else
                        WDescriTipo = ""
                    End If
                    rstArticulo.Close
                    
                    
                Case Else
                    ZZCosto = rstArticulo!Costo1

                    Costo1 = rstArticulo!Costo1
                    WWCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                    ZCosto1 = IIf(IsNull(rstArticulo!ZCosto1), "0", rstArticulo!ZCosto1)
                    ZZOrdenI = IIf(IsNull(rstArticulo!OrdenI), "", rstArticulo!OrdenI)
                    ZZOrdenII = IIf(IsNull(rstArticulo!OrdenII), "", rstArticulo!OrdenII)
                    ZZOrdenIII = IIf(IsNull(rstArticulo!OrdenIII), "", rstArticulo!OrdenIII)
                    ZZPtaOrdenI = IIf(IsNull(rstArticulo!PtaOrdenI), "0", rstArticulo!PtaOrdenI)
                    ZZPtaOrdenII = IIf(IsNull(rstArticulo!PtaOrdenII), "0", rstArticulo!PtaOrdenII)
                    ZZPtaOrdenIII = IIf(IsNull(rstArticulo!PtaOrdenIII), "0", rstArticulo!PtaOrdenIII)
                        
                    ZZFechaOrdenI = ""
                    ZZFechaOrdenII = ""
                    ZZFechaOrdenIII = ""
                    
                    ZZMoneda = ""
                    
                    rstArticulo.Close
                    
                    XEmpresa = WEmpresa

                    If ZZPtaOrdenI <> 0 And ZZOrdenI <> 0 Then
                    
                        Select Case ZZPtaOrdenI
                            Case 1
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                WEmpresa = "0002"
                                txtOdbc = "Empresa02"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                WEmpresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 4
                                WEmpresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 5
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 6
                                WEmpresa = "0006"
                                txtOdbc = "Empresa06"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 7
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 8
                                WEmpresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 9
                                WEmpresa = "0009"
                                txtOdbc = "Empresa09"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 10
                                WEmpresa = "0010"
                                txtOdbc = "Empresa10"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 11
                                WEmpresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenI) + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZFechaOrdenI = rstOrden!Fecha
                            Select Case rstOrden!Moneda
                                Case 0
                                    ZZMoneda = "U$S"
                                Case 1
                                    ZZMoneda = "$"
                                Case Else
                                    ZZMoneda = "Eur"
                            End Select
                            rstOrden.Close
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                    
                    
                    If ZZPtaOrdenII <> 0 And ZZOrdenII <> 0 Then
                        
                        Select Case ZZPtaOrdenII
                            Case 1
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                WEmpresa = "0002"
                                txtOdbc = "Empresa02"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                WEmpresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 4
                                WEmpresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 5
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 6
                                WEmpresa = "0006"
                                txtOdbc = "Empresa06"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 7
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 8
                                WEmpresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 9
                                WEmpresa = "0009"
                                txtOdbc = "Empresa09"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 10
                                WEmpresa = "0010"
                                txtOdbc = "Empresa10"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 11
                                WEmpresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenII) + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZFechaOrdenII = rstOrden!Fecha
                            rstOrden.Close
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                    
                    If ZZPtaOrdenIII <> 0 And ZZOrdenIII <> 0 Then
                        
                        Select Case ZZPtaOrdenIII
                            Case 1
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                WEmpresa = "0002"
                                txtOdbc = "Empresa02"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                WEmpresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 4
                                WEmpresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 5
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 6
                                WEmpresa = "0006"
                                txtOdbc = "Empresa06"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 7
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 8
                                WEmpresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 9
                                WEmpresa = "0009"
                                txtOdbc = "Empresa09"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 10
                                WEmpresa = "0010"
                                txtOdbc = "Empresa10"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 11
                                WEmpresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenIII) + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZFechaOrdenIII = rstOrden!Fecha
                            rstOrden.Close
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                    
                    Rem DADA
                    Rem spCambio = "ConsultaCambio " + "'" + ZZFecha + "'"
                    Rem Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstCambio.RecordCount > 0 Then
                    Rem     ZZParidad = rstCambio!Cambio
                    Rem     rstCambio.Close
                    Rem End If
                    
                    Rem If ZZParidad <> 0 Then
                    Rem     ZZCostoPesos = WPrecio * ZZParidad
                    Rem End If
                            
                    If ZZFechaOrdenI <> "" Then
                        WFechaOrdI = Right$(ZZFechaOrdenI, 4) + Mid$(ZZFechaOrdenI, 4, 2) + Left$(ZZFechaOrdenI, 2)
                            Else
                        WFechaOrdI = ""
                    End If
                    If ZZFechaOrdenII <> "" Then
                        WFechaOrdII = Right$(ZZFechaOrdenII, 4) + Mid$(ZZFechaOrdenII, 4, 2) + Left$(ZZFechaOrdenII, 2)
                            Else
                        WFechaOrdII = ""
                    End If
                    If ZZFechaOrdenIII <> "" Then
                        WFechaOrdIII = Right$(ZZFechaOrdenIII, 4) + Mid$(ZZFechaOrdenIII, 4, 2) + Left$(ZZFechaOrdenIII, 2)
                            Else
                        WFechaOrdIII = ""
                    End If
                    
                    If WFechaOrdI <> "" And WFechaOrdI > WFechaOrdII And WFechaOrdI > WFechaOrdIII Then
                        ZZCosto = Costo1
                    End If
                    
                    If WFechaOrdII <> "" And WFechaOrdII > WFechaOrdI And WFechaOrdII > WFechaOrdIII Then
                    
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                         spCambios = "ConsultaCambio  " + "'" + ZZFechaOrdenII + "'"
                         Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                         If rstCambios.RecordCount > 0 Then
                             ZZZZParidad = rstCambios!Cambio
                             rstCambios.Close
                             If ZZZZParidad <> 0 Then
                                 ZZCosto1Dol = WWCosto1 / ZZZZParidad
                             End If
                         End If
                        Call Conecta_Empresa
                    
                        ZZCosto = ZZCosto1Dol
                    End If
                    
                    If WFechaOrdIII <> "" And WFechaOrdIII > WFechaOrdI And WFechaOrdIII > WFechaOrdII Then
                        ZZCosto = ZCosto1
                    End If
                    WDescriTipo = ""
                                            
                    Rem WEmpresa = "0001"
                    Rem txtOdbc = "Empresa01"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    Rem spCambios = "ConsultaCambio  " + "'" + FechaOrdenII.Text + "'"
                    Rem Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstCambios.RecordCount > 0 Then
                    Rem     ZZParidad = rstCambios!Cambio
                    Rem     rstCambios.Close
                    Rem     If ZZParidad <> 0 Then
                    Rem         ZZCosto1Dol = Val(WCosto1.Text) / ZZParidad
                    Rem         WCosto1Dol.Text = Str$(ZZCosto1Dol)
                    Rem         WCosto1Dol.Text = Pusing("###,###.##", WCosto1Dol.Text)
                    Rem     End If
                    Rem End If
            End Select
            
            WCosto = (ZZCantidad * ZZCosto * Val(WVector))
            Costo = Costo + WCosto
                    
            If ZZCosto = 0 Then
                m$ = "El costo de la materia prima " + Articulo + " que compone el " + Producto + " esta en cero"
                a% = MsgBox(m$, 0, "Calculo de Costos")
            End If
            
        End If
    Next DA
    
End Sub

















Private Sub Alta_Producto()


    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
                    
    For Cicla = 1 To 9
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    If WSalidaError = "N" Then Exit Sub

    On Error GoTo WError
    
    WCodigo = Orden.Text + "-100"
    WDescripcion = Descripcion.Text
    WLinea = "0"
    WUnidad = ""
    WInicial = "0"
    WEntradas = "0"
    WSalidas = "0"
    WMinimo = "0"
    WMinimo1 = "0"
    WDeposito = ""
    WPedido = ""
    WEnvase1 = ""
    WEnvase2 = ""
    WEnvase3 = ""
    WEnvase4 = ""
    WEnvase5 = ""
    WEnvase6 = ""
    WProceso = "0"
    WCosto = ""
    WFactor = ""
    WDate = Date$
    WImpreadi = ""
    WIntervencion = ""
    WClase = ""
    WNaciones = ""
    WEmbalaje = ""
    WControla = "0"
    WEscrito = "0"
    WObservaciones = ""
    WTipoeti = ""
    WConservacion = ""
    WConservacionII = ""
    WVida = "0"
    WSeguridad = ""
            
    WVersion = ""
    WVersionI = ""
    WVersionII = ""
            
    WFechaVersion = "  /  /    "
    WFechaVersionI = "  /  /    "
    WFechaVersionII = "  /  /    "
            
    WEstado = ""
    WEstadoI = ""
    WEstadoII = ""
            
    WObserva = ""
    WObservaI = ""
    WObservaII = ""
            
    WMetodo = ""
    WEfluentes = ""
    
    XEmpresa = WEmpresa
    Erase CargaEmpresa
        
    Select Case Val(WEmpresa)
        Case 3
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
            CargaEmpresa(2, 1) = "0003"
            CargaEmpresa(2, 2) = "Empresa03"
            CargaEmpresa(3, 1) = "0005"
            CargaEmpresa(3, 2) = "Empresa05"
            CargaEmpresa(4, 1) = "0006"
            CargaEmpresa(4, 2) = "Empresa06"
            CargaEmpresa(5, 1) = "0007"
            CargaEmpresa(5, 2) = "Empresa07"
        Case 4
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Rem by nan 4-6-2013
        Case 5
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        
        
       Rem end by nan
        Case 10
            CargaEmpresa(1, 1) = "0010"
            CargaEmpresa(1, 2) = "Empresa10"
        Case Else
    End Select
                
    For Cicla = 1 To 5
            
        If CargaEmpresa(Cicla, 1) <> "" Then
        
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            ZSql = ""
            ZSql = ZSql & "INSERT INTO Terminado ("
            ZSql = ZSql & "Codigo ,"
            ZSql = ZSql & "Descripcion ,"
            ZSql = ZSql & "Linea ,"
            ZSql = ZSql & "Unidad ,"
            ZSql = ZSql & "Inicial ,"
            ZSql = ZSql & "Entradas ,"
            ZSql = ZSql & "Salidas ,"
            ZSql = ZSql & "Minimo ,"
            ZSql = ZSql & "Deposito ,"
            ZSql = ZSql & "Pedido ,"
            ZSql = ZSql & "Envase1 ,"
            ZSql = ZSql & "Envase2 ,"
            ZSql = ZSql & "Envase3 ,"
            ZSql = ZSql & "Envase4 ,"
            ZSql = ZSql & "Envase5 ,"
            ZSql = ZSql & "Envase6 ,"
            ZSql = ZSql & "Proceso ,"
            ZSql = ZSql & "Costo ,"
            ZSql = ZSql & "Factor ,"
            ZSql = ZSql & "WDate ,"
            ZSql = ZSql & "ImpreAdi ,"
            ZSql = ZSql & "Clase ,"
            ZSql = ZSql & "Intervencion ,"
            ZSql = ZSql & "Naciones ,"
            ZSql = ZSql & "Embalaje ,"
            ZSql = ZSql & "Controla ,"
            ZSql = ZSql & "Observaciones ,"
            ZSql = ZSql & "TipoEti ,"
            ZSql = ZSql & "Escrito ,"
            ZSql = ZSql & "Minimo1 ,"
            ZSql = ZSql & "Conservacion ,"
            ZSql = ZSql & "ConservacionII ,"
            ZSql = ZSql & "Vida ,"
            ZSql = ZSql & "Seguridad ,"
            ZSql = ZSql & "Version ,"
            ZSql = ZSql & "VersionI ,"
            ZSql = ZSql & "VersionII ,"
            ZSql = ZSql & "FechaVersion ,"
            ZSql = ZSql & "FechaVersionI ,"
            ZSql = ZSql & "FechaVersionII ,"
            ZSql = ZSql & "Estado ,"
            ZSql = ZSql & "EstadoI ,"
            ZSql = ZSql & "EstadoII ,"
            ZSql = ZSql & "Observa ,"
            ZSql = ZSql & "ObservaI ,"
            ZSql = ZSql & "ObservaII ,"
            ZSql = ZSql & "Metodo ,"
            ZSql = ZSql & "Efluentes )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + WCodigo + "',"
            ZSql = ZSql & "'" + WDescripcion + "',"
            ZSql = ZSql & "'" + WLinea + "',"
            ZSql = ZSql & "'" + WUnidad + "',"
            ZSql = ZSql & "'" + WInicial + "',"
            ZSql = ZSql & "'" + WEntradas + "',"
            ZSql = ZSql & "'" + WSalidas + "',"
            ZSql = ZSql & "'" + WMinimo + "',"
            ZSql = ZSql & "'" + WDeposito + "',"
            ZSql = ZSql & "'" + WPedido + "',"
            ZSql = ZSql & "'" + WEnvase1 + "',"
            ZSql = ZSql & "'" + WEnvase2 + "',"
            ZSql = ZSql & "'" + WEnvase3 + "',"
            ZSql = ZSql & "'" + WEnvase4 + "',"
            ZSql = ZSql & "'" + WEnvase5 + "',"
            ZSql = ZSql & "'" + WEnvase6 + "',"
            ZSql = ZSql & "'" + WProceso + "',"
            ZSql = ZSql & "'" + WCosto + "',"
            ZSql = ZSql & "'" + WFactor + "',"
            ZSql = ZSql & "'" + WDate + "',"
            ZSql = ZSql & "'" + WImpreadi + "',"
            ZSql = ZSql & "'" + WClase + "',"
            ZSql = ZSql & "'" + WIntervencion + "',"
            ZSql = ZSql & "'" + WNaciones + "',"
            ZSql = ZSql & "'" + WEmbalaje + "',"
            ZSql = ZSql & "'" + WControla + "',"
            ZSql = ZSql & "'" + WObservaciones + "',"
            ZSql = ZSql & "'" + WTipoeti + "',"
            ZSql = ZSql & "'" + WEscrito + "',"
            ZSql = ZSql & "'" + WMinimo1 + "',"
            ZSql = ZSql & "'" + WConservacion + "',"
            ZSql = ZSql & "'" + WConservacionII + "',"
            ZSql = ZSql & "'" + WVida + "',"
            ZSql = ZSql & "'" + WSeguridad + "',"
            ZSql = ZSql & "'" + WVersion + "',"
            ZSql = ZSql & "'" + WVersionI + "',"
            ZSql = ZSql & "'" + WVersionII + "',"
            ZSql = ZSql & "'" + WFechaVersion + "',"
            ZSql = ZSql & "'" + WFechaVersionI + "',"
            ZSql = ZSql & "'" + WFechaVersionII + "',"
            ZSql = ZSql & "'" + WEstado + "',"
            ZSql = ZSql & "'" + WEstadoI + "',"
            ZSql = ZSql & "'" + WEstadoII + "',"
            ZSql = ZSql & "'" + WObserva + "',"
            ZSql = ZSql & "'" + WObservaI + "',"
            ZSql = ZSql & "'" + WObvservaII + "',"
            ZSql = ZSql & "'" + WMetodo + "',"
            ZSql = ZSql & "'" + WEfluentes + "')"
      
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Cicla
    
    Call Conecta_Empresa
        
    Exit Sub

WError:
    Resume Next
    
Control_error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoErrorII.Visible = True
    Resume Next
    
End Sub


Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub


Sub Impresion()

    Sql1 = "DELETE ImpreHoja"
    spImpreHoja = Sql1
    Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
    Linea = 0
    
    For iRow = 1 To 99
    
        Tipo = WVector1.TextMatrix(iRow, 1)
        Articulo = WVector1.TextMatrix(iRow, 2)
        Terminado = WVector1.TextMatrix(iRow, 3)
        Detalle = WVector1.TextMatrix(iRow, 4)
        Canti = WVector1.TextMatrix(iRow, 5)
        Lote = WVector1.TextMatrix(iRow, 6)
        
        If Val(Canti) <> 0 Then
    
        Linea = Linea + 1
        
        WHoja = Hoja.Text
        WLinea = Str$(Linea)
        WFecha = Fecha.Text
        WCodigo1 = Left$(Orden.Text, 2)
        WCodigo2 = Right$(Orden.Text, 5) + "/100"
        WMaquina = ""
        If Tipo = "M" Then
            WArticulo1 = Left$(Articulo, 2)
            WArticulo2 = Mid$(Articulo, 4, 3) + "-" + Right$(Articulo, 3)
                Else
            WArticulo1 = Left$(Terminado, 2)
            WArticulo2 = Mid$(Terminado, 4, 5) + "-" + Right$(Terminado, 3)
        End If
        WCantidad = Canti
        WDetalle = Detalle
        ZZCanti = ""
        ZZLote = ""
        WTeorico = Cantidad.Text
        ZMetodo = ""
        ZEfluentes = ""
        ZVersionI = ""
        ZVersionII = ""
        ZVersionIII = ""
        ZDesEfluentesI = ""
        ZDesEfluentesII = ""
        WEquipo = ""
        ZZMetodo = ""
        ZZEspecificacion = ""
                        
        ZSql = ""
        ZSql = ZSql & "INSERT INTO ImpreHoja ("
        ZSql = ZSql & "Hoja ,"
        ZSql = ZSql & "Renglon ,"
        ZSql = ZSql & "Fecha ,"
        ZSql = ZSql & "Codigo1 ,"
        ZSql = ZSql & "Codigo2 ,"
        ZSql = ZSql & "Maquina ,"
        ZSql = ZSql & "Articulo1 ,"
        ZSql = ZSql & "Articulo2 ,"
        ZSql = ZSql & "Cantidad ,"
        ZSql = ZSql & "Teorico ,"
        ZSql = ZSql & "Metodo ,"
        ZSql = ZSql & "Efluentes ,"
        ZSql = ZSql & "DesEfluentesI ,"
        ZSql = ZSql & "DesEfluentesII ,"
        ZSql = ZSql & "VersionI ,"
        ZSql = ZSql & "VersionII ,"
        ZSql = ZSql & "VersionIII ,"
        ZSql = ZSql & "Detalle ,"
        ZSql = ZSql & "Equipo )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + WHoja + "',"
        ZSql = ZSql & "'" + WLinea + "',"
        ZSql = ZSql & "'" + WFecha + "',"
        ZSql = ZSql & "'" + WCodigo1 + "',"
        ZSql = ZSql & "'" + WCodigo2 + "',"
        ZSql = ZSql & "'" + WMaquina + "',"
        ZSql = ZSql & "'" + WArticulo1 + "',"
        ZSql = ZSql & "'" + WArticulo2 + "',"
        ZSql = ZSql & "'" + WCantidad + "',"
        ZSql = ZSql & "'" + WTeorico + "',"
        ZSql = ZSql & "'" + ZMetodo + "',"
        ZSql = ZSql & "'" + ZEfluentes + "',"
        ZSql = ZSql & "'" + ZDesEfluentesI + "',"
        ZSql = ZSql & "'" + ZDesEfluentesII + "',"
        ZSql = ZSql & "'" + ZVersionI + "',"
        ZSql = ZSql & "'" + ZVersionII + "',"
        ZSql = ZSql & "'" + ZVersionIII + "',"
        ZSql = ZSql & "'" + WDetalle + "',"
        ZSql = ZSql & "'" + WEquipo + "')"
        
        spImpreHoja = ZSql
        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next iRow
                        
            
    XLinea = Linea
    For Ciclo = XLinea To 14
            
        Linea = Linea + 1
        WLinea = Str$(Linea)
                        
        WArticulo1 = ""
        WArticulo2 = ""
        WDetalle = ""
        WCantidad = ""
        WCanti1 = ""
        WLote1 = ""
        WCanti2 = ""
        WLote2 = ""
        WCanti3 = ""
        WLote3 = ""
        
        ZSql = ""
        ZSql = ZSql & "INSERT INTO ImpreHoja ("
        ZSql = ZSql & "Hoja ,"
        ZSql = ZSql & "Renglon ,"
        ZSql = ZSql & "Fecha ,"
        ZSql = ZSql & "Codigo1 ,"
        ZSql = ZSql & "Codigo2 ,"
        ZSql = ZSql & "Maquina ,"
        ZSql = ZSql & "Articulo1 ,"
        ZSql = ZSql & "Articulo2 ,"
        ZSql = ZSql & "Cantidad ,"
        ZSql = ZSql & "Teorico ,"
        ZSql = ZSql & "Metodo ,"
        ZSql = ZSql & "Efluentes ,"
        ZSql = ZSql & "DesEfluentesI ,"
        ZSql = ZSql & "DesEfluentesII ,"
        ZSql = ZSql & "VersionI ,"
        ZSql = ZSql & "VersionII ,"
        ZSql = ZSql & "VersionIII ,"
        ZSql = ZSql & "Detalle ,"
        ZSql = ZSql & "Equipo )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + WHoja + "',"
        ZSql = ZSql & "'" + WLinea + "',"
        ZSql = ZSql & "'" + WFecha + "',"
        ZSql = ZSql & "'" + WCodigo1 + "',"
        ZSql = ZSql & "'" + WCodigo2 + "',"
        ZSql = ZSql & "'" + WMaquina + "',"
        ZSql = ZSql & "'" + WArticulo1 + "',"
        ZSql = ZSql & "'" + WArticulo2 + "',"
        ZSql = ZSql & "'" + WCantidad + "',"
        ZSql = ZSql & "'" + WTeorico + "',"
        ZSql = ZSql & "'" + ZMetodo + "',"
        ZSql = ZSql & "'" + ZEfluentes + "',"
        ZSql = ZSql & "'" + ZDesEfluentesI + "',"
        ZSql = ZSql & "'" + ZDesEfluentesII + "',"
        ZSql = ZSql & "'" + ZVersionI + "',"
        ZSql = ZSql & "'" + ZVersionII + "',"
        ZSql = ZSql & "'" + ZVersionIII + "',"
        ZSql = ZSql & "'" + WDetalle + "',"
        ZSql = ZSql & "'" + WEquipo + "')"
        
        spImpreHoja = ZSql
        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)

    Next Ciclo
            
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
            
    Listado.ReportFileName = "ImpreHojaDesarrollo.rpt"
    Listado.GroupSelectionFormula = "{ImpreHoja.Hoja} in 0 to 999999"
                
    Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.VersionI, ImpreHoja.VersionII, ImpreHoja.VersionIII, ImpreHoja.Equipo, ImpreHoja.Detalle " _
                + "From " _
                + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                + "Where " _
                + "ImpreHoja.Hoja >= 0 AND " _
                + "ImpreHoja.Hoja <= 999999"
    
    Listado.Connect = Connect()
    Listado.Action = 1

End Sub

Private Sub BusquedaEnsayo_Click()
    ZClienteII = ""
    Call Busca_Ensayo
End Sub

Private Sub BusquedaEnsayoII_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    ZProceso = 2
    
    Pantalla.Height = 5340
    Pantalla.Left = 480
    Pantalla.Top = 1680
    Pantalla.Width = 10575
    
    Erase ZCarga
    ZLugar = 0
    
    Pasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM OrdenTrabajo"
    ZSql = ZSql + " Order by OrdenTRabajo.Cliente"
    spOrdenTrabajo = ZSql
    Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrdenTrabajo.RecordCount > 0 Then
        With rstOrdenTrabajo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    aa = rstOrdenTrabajo!Cliente
                    aa1 = WEmpresa
                
                    If Pasa = 0 Then
                        Pasa = 1
                        ZCliente = rstOrdenTrabajo!Cliente
                    End If
                    
                    If ZCliente <> rstOrdenTrabajo!Cliente Then
                        ZLugar = ZLugar + 1
                        ZCarga(ZLugar, 1) = ZCliente
                        ZCliente = rstOrdenTrabajo!Cliente
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrdenTrabajo.Close
    End If

    If Pasa <> 0 Then
        ZLugar = ZLugar + 1
        ZCarga(ZLugar, 1) = ZCliente
    End If
    
    For Ciclo = 1 To ZLugar
        
        ZCliente = ZCarga(Ciclo, 1)
        
        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZDesCliente = Trim(rstCliente!razon)
            rstCliente.Close
                Else
            ZDesCliente = ""
        End If
        
        IngresaItem = ZCliente + "  " + ZDesCliente
        Pantalla.AddItem IngresaItem
        IngresaItem = ZCliente
        WIndice.AddItem IngresaItem
        
    Next Ciclo
    
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case ZProceso
        Case 1
            Indice = Pantalla.ListIndex
            Orden.Text = Left$(WIndice.List(Indice), 8)
            Version.Text = Mid$(Str$(Val(Mid$(WIndice.List(Indice), 10, 10))), 2, 10)
            Call Version_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            ZClienteII = WIndice.List(Indice)
            Call Busca_Ensayo
            
        Case Else
    End Select
End Sub

Private Sub Busca_Ensayo()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear
    ZProceso = 1
    
    Pantalla.Height = 5340
    Pantalla.Left = 480
    Pantalla.Top = 1680
    Pantalla.Width = 10575
    
    Erase ZCarga
    ZLugar = 0
    
    Pasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayo"
    ZSql = ZSql + " Order by CargaEnsayo.Clave"
    spCargaEnsayo = ZSql
    Set rstCargaEnsayo = db.OpenRecordset(spCargaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayo.RecordCount > 0 Then
        With rstCargaEnsayo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    aa = rstCargaEnsayo!Orden
                    aa1 = WEmpresa
                
                    If Pasa = 0 Then
                        Pasa = 1
                        ZOrden = rstCargaEnsayo!Orden
                    End If
                    
                    If ZOrden <> rstCargaEnsayo!Orden Then
                        ZLugar = ZLugar + 1
                        ZCarga(ZLugar, 1) = ZOrden
                        ZCarga(ZLugar, 2) = ZVersion
                        ZCarga(ZLugar, 3) = ZClave
                        ZOrden = rstCargaEnsayo!Orden
                    End If
                    
                    ZVersion = rstCargaEnsayo!Version
                    ZClave = rstCargaEnsayo!Clave
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayo.Close
    End If

    If Pasa <> 0 Then
        ZLugar = ZLugar + 1
        ZCarga(ZLugar, 1) = ZOrden
        ZCarga(ZLugar, 2) = ZVersion
        ZCarga(ZLugar, 3) = ZClave
    End If
    
    For Ciclo = 1 To ZLugar
        
        ZOrden = ZCarga(Ciclo, 1)
        ZVersion = ZCarga(Ciclo, 2)
        ZClave = ZCarga(Ciclo, 3)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenTrabajo"
        ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + ZOrden + "'"
        spOrdenTrabajo = ZSql
        Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenTrabajo.RecordCount > 0 Then
            ZObservaciones = Trim(rstOrdenTrabajo!Observaciones)
            ZCliente = Trim(rstOrdenTrabajo!Cliente)
            rstOrdenTrabajo.Close
                Else
            ZObservaciones = ""
            ZCliente = ""
        End If
        
        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZDesCliente = Trim(rstCliente!razon)
            rstCliente.Close
                Else
            ZDesCliente = ""
        End If
        
        If ZClienteII = "" Or ZClienteII = ZCliente Then
        
        IngresaItem = ZOrden + "/" + Str$(ZVersion) + "  " + ZObservaciones + "  (" + ZDesCliente + ")"
        Pantalla.AddItem IngresaItem
        IngresaItem = ZClave
        WIndice.AddItem IngresaItem
        
        End If
        
    Next Ciclo
    
    Pantalla.Visible = True

End Sub


