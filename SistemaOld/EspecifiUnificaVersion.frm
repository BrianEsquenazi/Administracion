VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspecifiUnificaVersion 
   Caption         =   "Consulta de Versiones de Especificaciones de Materia Prima"
   ClientHeight    =   8250
   ClientLeft      =   375
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11160
   Begin VB.TextBox ControlCambio 
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
      MaxLength       =   100
      TabIndex        =   85
      Text            =   " "
      Top             =   5520
      Width           =   5280
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   17
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox FechaFinal 
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
      Left            =   8520
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   16
      Text            =   " "
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox FechaInicio 
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
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   15
      Text            =   " "
      Top             =   0
      Width           =   1335
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
      Left            =   3840
      MaxLength       =   50
      TabIndex        =   14
      Text            =   " "
      Top             =   0
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport lista 
      Left            =   7920
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wespec1Unifica.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
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
      Height          =   1260
      Left            =   960
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   9000
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "EspecifiUnificaVersion.frx":0000
      Left            =   120
      List            =   "EspecifiUnificaVersion.frx":0007
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   3
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9960
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8070
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Especificacion 1 - 10"
      TabPicture(0)   =   "EspecifiUnificaVersion.frx":0015
      Tab(0).ControlCount=   33
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblresultado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblensayo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDescri"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Descri1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "descri2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Descri3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Descri4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Descri5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Descri6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Descri7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Descri8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Descri9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Descri10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Valor1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "valor2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Valor3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "valor4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "valor5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "valor6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "valor7"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "valor8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "valor9"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "valor10"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Ensayo1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Ensayo2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Ensayo3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Ensayo4"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Ensayo5"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Ensayo6"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Ensayo7"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Ensayo8"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Ensayo9"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Ensayo10"
      Tab(0).Control(32).Enabled=   0   'False
      TabCaption(1)   =   "Especificacion 11  - 20"
      TabPicture(1)   =   "EspecifiUnificaVersion.frx":0031
      Tab(1).ControlCount=   33
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Valor11"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Valor12"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "Valor13"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "Valor14"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "Valor15"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "Valor16"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "Valor17"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "Valor18"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "Valor19"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Valor20"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Ensayo11"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Ensayo12"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Ensayo13"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Ensayo14"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Ensayo15"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Ensayo16"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Ensayo17"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "Ensayo18"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Ensayo19"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "Ensayo20"
      Tab(1).Control(19).Enabled=   -1  'True
      Tab(1).Control(20)=   "Label6"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label7"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label8"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Descri11"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Descri12"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Descri13"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Descri14"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Descri15"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Descri16"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Descri17"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Descri18"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Descri19"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Descri20"
      Tab(1).Control(32).Enabled=   0   'False
      TabCaption(2)   =   "Especificacion 21  - 30"
      TabPicture(2)   =   "EspecifiUnificaVersion.frx":004D
      Tab(2).ControlCount=   33
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TituloIII"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label28"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label27"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Descri21"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Descri22"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Descri23"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Descri24"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Descri25"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Descri26"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Descri27"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Descri28"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Descri29"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Descri30"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Ensayo21"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "Ensayo22"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "Ensayo23"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "Ensayo24"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "Ensayo25"
      Tab(2).Control(17).Enabled=   -1  'True
      Tab(2).Control(18)=   "Ensayo26"
      Tab(2).Control(18).Enabled=   -1  'True
      Tab(2).Control(19)=   "Ensayo27"
      Tab(2).Control(19).Enabled=   -1  'True
      Tab(2).Control(20)=   "Ensayo28"
      Tab(2).Control(20).Enabled=   -1  'True
      Tab(2).Control(21)=   "Ensayo29"
      Tab(2).Control(21).Enabled=   -1  'True
      Tab(2).Control(22)=   "Ensayo30"
      Tab(2).Control(22).Enabled=   -1  'True
      Tab(2).Control(23)=   "Valor22"
      Tab(2).Control(23).Enabled=   -1  'True
      Tab(2).Control(24)=   "Valor23"
      Tab(2).Control(24).Enabled=   -1  'True
      Tab(2).Control(25)=   "Valor24"
      Tab(2).Control(25).Enabled=   -1  'True
      Tab(2).Control(26)=   "Valor25"
      Tab(2).Control(26).Enabled=   -1  'True
      Tab(2).Control(27)=   "Valor26"
      Tab(2).Control(27).Enabled=   -1  'True
      Tab(2).Control(28)=   "Valor27"
      Tab(2).Control(28).Enabled=   -1  'True
      Tab(2).Control(29)=   "Valor28"
      Tab(2).Control(29).Enabled=   -1  'True
      Tab(2).Control(30)=   "Valor29"
      Tab(2).Control(30).Enabled=   -1  'True
      Tab(2).Control(31)=   "Valor30"
      Tab(2).Control(31).Enabled=   -1  'True
      Tab(2).Control(32)=   "Valor21"
      Tab(2).Control(32).Enabled=   -1  'True
      Begin VB.TextBox Valor21 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   119
         Text            =   " "
         Top             =   840
         Width           =   5040
      End
      Begin VB.TextBox Valor30 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   105
         Text            =   " "
         Top             =   4080
         Width           =   5040
      End
      Begin VB.TextBox Valor29 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   104
         Text            =   " "
         Top             =   3720
         Width           =   5040
      End
      Begin VB.TextBox Valor28 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   103
         Text            =   " "
         Top             =   3360
         Width           =   5040
      End
      Begin VB.TextBox Valor27 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   102
         Text            =   " "
         Top             =   3000
         Width           =   5040
      End
      Begin VB.TextBox Valor26 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   101
         Text            =   " "
         Top             =   2640
         Width           =   5040
      End
      Begin VB.TextBox Valor25 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   100
         Text            =   " "
         Top             =   2280
         Width           =   5040
      End
      Begin VB.TextBox Valor24 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   99
         Text            =   " "
         Top             =   1920
         Width           =   5040
      End
      Begin VB.TextBox Valor23 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   98
         Text            =   " "
         Top             =   1560
         Width           =   5040
      End
      Begin VB.TextBox Valor22 
         Height          =   285
         Left            =   -69240
         MaxLength       =   70
         TabIndex        =   97
         Text            =   " "
         Top             =   1200
         Width           =   5040
      End
      Begin VB.TextBox Ensayo30 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   96
         Text            =   " "
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Ensayo29 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   95
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Ensayo28 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   94
         Text            =   " "
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Ensayo27 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   93
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Ensayo26 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   92
         Text            =   " "
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Ensayo25 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   91
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Ensayo24 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   90
         Text            =   " "
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Ensayo23 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   89
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Ensayo22 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   88
         Text            =   " "
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Ensayo21 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   87
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Ensayo10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   58
         Text            =   " "
         Top             =   4020
         Width           =   735
      End
      Begin VB.TextBox Ensayo9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   57
         Text            =   " "
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox Ensayo8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   56
         Text            =   " "
         Top             =   3300
         Width           =   735
      End
      Begin VB.TextBox Ensayo7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   55
         Text            =   " "
         Top             =   2940
         Width           =   735
      End
      Begin VB.TextBox Ensayo6 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   4
         TabIndex        =   54
         Text            =   " "
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox Ensayo5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   53
         Text            =   " "
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox Ensayo4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   52
         Text            =   " "
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox Ensayo3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   51
         Text            =   " "
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Ensayo2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   50
         Text            =   " "
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox Ensayo1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   49
         Text            =   " "
         Top             =   780
         Width           =   735
      End
      Begin VB.TextBox valor10 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   48
         Text            =   " "
         Top             =   4020
         Width           =   5055
      End
      Begin VB.TextBox valor9 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   3660
         Width           =   5055
      End
      Begin VB.TextBox valor8 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   3300
         Width           =   5055
      End
      Begin VB.TextBox valor7 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   45
         Text            =   " "
         Top             =   2940
         Width           =   5055
      End
      Begin VB.TextBox valor6 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   44
         Text            =   " "
         Top             =   2580
         Width           =   5055
      End
      Begin VB.TextBox valor5 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   43
         Text            =   " "
         Top             =   2220
         Width           =   5055
      End
      Begin VB.TextBox valor4 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   42
         Text            =   " "
         Top             =   1860
         Width           =   5055
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   41
         Text            =   " "
         Top             =   1500
         Width           =   5055
      End
      Begin VB.TextBox valor2 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   40
         Text            =   " "
         Top             =   1140
         Width           =   5055
      End
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   39
         Text            =   " "
         Top             =   780
         Width           =   5055
      End
      Begin VB.TextBox Valor11 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   38
         Text            =   " "
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox Valor12 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   37
         Text            =   " "
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox Valor13 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   36
         Text            =   " "
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox Valor14 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   35
         Text            =   " "
         Top             =   1920
         Width           =   5055
      End
      Begin VB.TextBox Valor15 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   34
         Text            =   " "
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox Valor16 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   33
         Text            =   " "
         Top             =   2640
         Width           =   5055
      End
      Begin VB.TextBox Valor17 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   32
         Text            =   " "
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox Valor18 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   31
         Text            =   " "
         Top             =   3360
         Width           =   5055
      End
      Begin VB.TextBox Valor19 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   30
         Text            =   " "
         Top             =   3720
         Width           =   5055
      End
      Begin VB.TextBox Valor20 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   29
         Text            =   " "
         Top             =   4080
         Width           =   5055
      End
      Begin VB.TextBox Ensayo11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   28
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Ensayo12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   27
         Text            =   " "
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Ensayo13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   26
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Ensayo14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   25
         Text            =   " "
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Ensayo15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   24
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Ensayo16 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   23
         Text            =   " "
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Ensayo17 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   22
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Ensayo18 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   21
         Text            =   " "
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Ensayo19 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   20
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Ensayo20 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   19
         Text            =   " "
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Descri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   118
         Top             =   4080
         Width           =   4740
      End
      Begin VB.Label Descri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   117
         Top             =   3720
         Width           =   4740
      End
      Begin VB.Label Descri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   116
         Top             =   3360
         Width           =   4740
      End
      Begin VB.Label Descri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   115
         Top             =   3000
         Width           =   4740
      End
      Begin VB.Label Descri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   114
         Top             =   2640
         Width           =   4740
      End
      Begin VB.Label Descri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   113
         Top             =   2280
         Width           =   4740
      End
      Begin VB.Label Descri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   112
         Top             =   1920
         Width           =   4740
      End
      Begin VB.Label Descri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   111
         Top             =   1560
         Width           =   4740
      End
      Begin VB.Label Descri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   110
         Top             =   1200
         Width           =   4740
      End
      Begin VB.Label Descri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   109
         Top             =   840
         Width           =   4740
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         TabIndex        =   108
         Top             =   480
         Width           =   4740
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   107
         Top             =   480
         Width           =   735
      End
      Begin VB.Label TituloIII 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   106
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   84
         Top             =   4020
         Width           =   4740
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   83
         Top             =   3660
         Width           =   4740
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   82
         Top             =   3300
         Width           =   4740
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   81
         Top             =   2940
         Width           =   4740
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   80
         Top             =   2580
         Width           =   4740
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   79
         Top             =   2220
         Width           =   4740
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   78
         Top             =   1860
         Width           =   4740
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   77
         Top             =   1500
         Width           =   4740
      End
      Begin VB.Label descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   76
         Top             =   1140
         Width           =   4740
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   75
         Top             =   780
         Width           =   4740
      End
      Begin VB.Label lblDescri 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   74
         Top             =   420
         Width           =   4695
      End
      Begin VB.Label lblensayo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   72
         Top             =   420
         Width           =   5055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   71
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   70
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   69
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Descri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   68
         Top             =   840
         Width           =   4740
      End
      Begin VB.Label Descri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   67
         Top             =   1200
         Width           =   4740
      End
      Begin VB.Label Descri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   66
         Top             =   1560
         Width           =   4740
      End
      Begin VB.Label Descri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   65
         Top             =   1920
         Width           =   4740
      End
      Begin VB.Label Descri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   64
         Top             =   2280
         Width           =   4740
      End
      Begin VB.Label Descri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   63
         Top             =   2640
         Width           =   4740
      End
      Begin VB.Label Descri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   62
         Top             =   3000
         Width           =   4740
      End
      Begin VB.Label Descri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   61
         Top             =   3360
         Width           =   4740
      End
      Begin VB.Label Descri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   60
         Top             =   3720
         Width           =   4740
      End
      Begin VB.Label Descri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   59
         Top             =   4080
         Width           =   4740
      End
   End
   Begin VB.Label lblLabels 
      Caption         =   "Control de Cambios"
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
      Index           =   3
      Left            =   120
      TabIndex        =   86
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblLabels 
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
      Index           =   2
      Left            =   6000
      TabIndex        =   13
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
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
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   12
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
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
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEspecifiUnificaVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecificacionesUnificaVersion As Recordset
Dim spEspecificacionesUnificaVersion As String
Dim rstEspecificacionesUnificaVersionII As Recordset
Dim spEspecificacionesUnificaVersionII As String
Dim XParam As String
Dim EmpresaActual As String
Dim ZFecha As String
Dim ZVersion As String

Dim ZVector(10000) As String
Dim ZEnsayo(30) As String
Dim ZValor(30) As String
Dim ZDescri(40) As String

Private Sub Imprime_Datos()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaVersion"
    Sql3 = " Where EspecificacionesUnificaVersion.Producto = " + "'" + Codigo.Text + "'"
    Sql4 = " and EspecificacionesUnificaVersion.Version = " + "'" + Version.Text + "'"
    spEspecificacionesUnificaVersion = Sql1 + Sql2 + Sql3 + Sql4
    Set rstEspecificacionesUnificaVersion = db.OpenRecordset(spEspecificacionesUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaVersion.RecordCount > 0 Then
    
        Ensayo1.Text = rstEspecificacionesUnificaVersion!Ensayo1
        Ensayo2.Text = rstEspecificacionesUnificaVersion!Ensayo2
        Ensayo3.Text = rstEspecificacionesUnificaVersion!Ensayo3
        Ensayo4.Text = rstEspecificacionesUnificaVersion!Ensayo4
        Ensayo5.Text = rstEspecificacionesUnificaVersion!Ensayo5
        Ensayo6.Text = rstEspecificacionesUnificaVersion!Ensayo6
        Ensayo7.Text = rstEspecificacionesUnificaVersion!Ensayo7
        Ensayo8.Text = rstEspecificacionesUnificaVersion!Ensayo8
        Ensayo9.Text = rstEspecificacionesUnificaVersion!Ensayo9
        Ensayo10.Text = rstEspecificacionesUnificaVersion!Ensayo10
        Ensayo11.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "", rstEspecificacionesUnificaVersion!Ensayo11)
        Ensayo12.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "", rstEspecificacionesUnificaVersion!Ensayo12)
        Ensayo13.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "", rstEspecificacionesUnificaVersion!Ensayo13)
        Ensayo14.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "", rstEspecificacionesUnificaVersion!Ensayo14)
        Ensayo15.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "", rstEspecificacionesUnificaVersion!Ensayo15)
        Ensayo16.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "", rstEspecificacionesUnificaVersion!Ensayo16)
        Ensayo17.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "", rstEspecificacionesUnificaVersion!Ensayo17)
        Ensayo18.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "", rstEspecificacionesUnificaVersion!Ensayo18)
        Ensayo19.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "", rstEspecificacionesUnificaVersion!Ensayo19)
        Ensayo20.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "", rstEspecificacionesUnificaVersion!Ensayo20)
        
        ZEnsayo(1) = rstEspecificacionesUnificaVersion!Ensayo1
        ZEnsayo(2) = rstEspecificacionesUnificaVersion!Ensayo2
        ZEnsayo(3) = rstEspecificacionesUnificaVersion!Ensayo3
        ZEnsayo(4) = rstEspecificacionesUnificaVersion!Ensayo4
        ZEnsayo(5) = rstEspecificacionesUnificaVersion!Ensayo5
        ZEnsayo(6) = rstEspecificacionesUnificaVersion!Ensayo6
        ZEnsayo(7) = rstEspecificacionesUnificaVersion!Ensayo7
        ZEnsayo(8) = rstEspecificacionesUnificaVersion!Ensayo8
        ZEnsayo(9) = rstEspecificacionesUnificaVersion!Ensayo9
        ZEnsayo(10) = rstEspecificacionesUnificaVersion!Ensayo10
        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "", rstEspecificacionesUnificaVersion!Ensayo11)
        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "", rstEspecificacionesUnificaVersion!Ensayo12)
        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "", rstEspecificacionesUnificaVersion!Ensayo13)
        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "", rstEspecificacionesUnificaVersion!Ensayo14)
        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "", rstEspecificacionesUnificaVersion!Ensayo15)
        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "", rstEspecificacionesUnificaVersion!Ensayo16)
        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "", rstEspecificacionesUnificaVersion!Ensayo17)
        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "", rstEspecificacionesUnificaVersion!Ensayo18)
        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "", rstEspecificacionesUnificaVersion!Ensayo19)
        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "", rstEspecificacionesUnificaVersion!Ensayo20)
        
        
        Valor1.Text = rstEspecificacionesUnificaVersion!Valor1
        valor2.Text = rstEspecificacionesUnificaVersion!valor2
        Valor3.Text = rstEspecificacionesUnificaVersion!Valor3
        valor4.Text = rstEspecificacionesUnificaVersion!valor4
        valor5.Text = rstEspecificacionesUnificaVersion!valor5
        valor6.Text = rstEspecificacionesUnificaVersion!valor6
        valor7.Text = rstEspecificacionesUnificaVersion!valor7
        valor8.Text = rstEspecificacionesUnificaVersion!valor8
        valor9.Text = rstEspecificacionesUnificaVersion!valor9
        valor10.Text = rstEspecificacionesUnificaVersion!valor10
        Valor11.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor1), "", rstEspecificacionesUnificaVersion!ZValor1)
        Valor12.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor2), "", rstEspecificacionesUnificaVersion!ZValor2)
        Valor13.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor3), "", rstEspecificacionesUnificaVersion!ZValor3)
        Valor14.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor4), "", rstEspecificacionesUnificaVersion!ZValor4)
        Valor15.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor5), "", rstEspecificacionesUnificaVersion!ZValor5)
        Valor16.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor6), "", rstEspecificacionesUnificaVersion!ZValor6)
        Valor17.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor7), "", rstEspecificacionesUnificaVersion!ZValor7)
        Valor18.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor8), "", rstEspecificacionesUnificaVersion!ZValor8)
        Valor19.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor9), "", rstEspecificacionesUnificaVersion!ZValor9)
        Valor20.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor10), "", rstEspecificacionesUnificaVersion!ZValor10)
        
        FechaInicio.Text = rstEspecificacionesUnificaVersion!FechaInicio
        FechaFinal.Text = rstEspecificacionesUnificaVersion!FechaFinal
        Rem Observaciones.Text = rstEspecificacionesUnificaVersion!Observaciones
        ControlCambio.Text = IIf(IsNull(rstEspecificacionesUnificaVersion!ControlCambio), "", rstEspecificacionesUnificaVersion!ControlCambio)
        
        rstEspecificacionesUnificaVersion.Close
                        
    End If
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaVersionII"
    Sql3 = " Where EspecificacionesUnificaVersionII.Producto = " + "'" + Codigo.Text + "'"
    Sql4 = " and EspecificacionesUnificaVersionII.Version = " + "'" + Version.Text + "'"
    spEspecificacionesUnificaVersionII = Sql1 + Sql2 + Sql3 + Sql4
    Set rstEspecificacionesUnificaVersionII = db.OpenRecordset(spEspecificacionesUnificaVersionII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaVersionII.RecordCount > 0 Then
    
        Ensayo21.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo21), "", rstEspecificacionesUnificaVersionII!Ensayo21)
        Ensayo22.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo22), "", rstEspecificacionesUnificaVersionII!Ensayo22)
        Ensayo23.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo23), "", rstEspecificacionesUnificaVersionII!Ensayo23)
        Ensayo24.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo24), "", rstEspecificacionesUnificaVersionII!Ensayo24)
        Ensayo25.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo25), "", rstEspecificacionesUnificaVersionII!Ensayo25)
        Ensayo26.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo26), "", rstEspecificacionesUnificaVersionII!Ensayo26)
        Ensayo27.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo27), "", rstEspecificacionesUnificaVersionII!Ensayo27)
        Ensayo28.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo28), "", rstEspecificacionesUnificaVersionII!Ensayo28)
        Ensayo29.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo29), "", rstEspecificacionesUnificaVersionII!Ensayo29)
        Ensayo30.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo30), "", rstEspecificacionesUnificaVersionII!Ensayo30)
        
        ZEnsayo(21) = Ensayo21.Text
        ZEnsayo(22) = Ensayo22.Text
        ZEnsayo(23) = Ensayo23.Text
        ZEnsayo(24) = Ensayo24.Text
        ZEnsayo(25) = Ensayo25.Text
        ZEnsayo(26) = Ensayo26.Text
        ZEnsayo(27) = Ensayo27.Text
        ZEnsayo(28) = Ensayo28.Text
        ZEnsayo(29) = Ensayo29.Text
        ZEnsayo(30) = Ensayo30.Text
        
        Valor21.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor21), "", rstEspecificacionesUnificaVersionII!Valor21)
        Valor22.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor22), "", rstEspecificacionesUnificaVersionII!Valor22)
        Valor23.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor23), "", rstEspecificacionesUnificaVersionII!Valor23)
        Valor24.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor24), "", rstEspecificacionesUnificaVersionII!Valor24)
        Valor25.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor25), "", rstEspecificacionesUnificaVersionII!Valor25)
        Valor26.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor26), "", rstEspecificacionesUnificaVersionII!Valor26)
        Valor27.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor27), "", rstEspecificacionesUnificaVersionII!Valor27)
        Valor28.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor28), "", rstEspecificacionesUnificaVersionII!Valor28)
        Valor29.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor29), "", rstEspecificacionesUnificaVersionII!Valor29)
        Valor30.Text = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor30), "", rstEspecificacionesUnificaVersionII!Valor30)
        
        rstEspecificacionesUnificaVersionII.Close
                        
    End If
    
    For Cicla = 1 To 30
        ZZDescri = ""
        If Val(ZEnsayo(Cicla)) <> 0 Then
            spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDescri = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        End If
        Select Case Cicla
            Case 1
                Descri1.Caption = ZZDescri
            Case 2
                descri2.Caption = ZZDescri
            Case 3
                Descri3.Caption = ZZDescri
            Case 4
                Descri4.Caption = ZZDescri
            Case 5
                Descri5.Caption = ZZDescri
            Case 6
                Descri6.Caption = ZZDescri
            Case 7
                Descri7.Caption = ZZDescri
            Case 8
                Descri8.Caption = ZZDescri
            Case 9
                Descri9.Caption = ZZDescri
            Case 10
                Descri10.Caption = ZZDescri
            Case 11
                Descri11.Caption = ZZDescri
            Case 12
                Descri12.Caption = ZZDescri
            Case 13
                Descri13.Caption = ZZDescri
            Case 14
                Descri14.Caption = ZZDescri
            Case 15
                Descri15.Caption = ZZDescri
            Case 16
                Descri16.Caption = ZZDescri
            Case 17
                Descri17.Caption = ZZDescri
            Case 18
                Descri18.Caption = ZZDescri
            Case 19
                Descri19.Caption = ZZDescri
            Case 20
                Descri20.Caption = ZZDescri
            Case 21
                Descri21.Caption = ZZDescri
            Case 22
                Descri22.Caption = ZZDescri
            Case 23
                Descri23.Caption = ZZDescri
            Case 24
                Descri24.Caption = ZZDescri
            Case 25
                Descri25.Caption = ZZDescri
            Case 26
                Descri26.Caption = ZZDescri
            Case 27
                Descri27.Caption = ZZDescri
            Case 28
                Descri28.Caption = ZZDescri
            Case 29
                Descri29.Caption = ZZDescri
            Case 30
                Descri30.Caption = ZZDescri
            Case Else
        End Select
                
    Next Cicla
    
    Call Conecta_Empresa
        
End Sub

Private Sub CmdLimpiar_Click()
    Codigo.Text = "  -   -   "
    Ensayo1.Text = ""
    Valor1.Text = ""
    Ensayo2.Text = ""
    valor2.Text = ""
    Ensayo3.Text = ""
    Valor3.Text = ""
    Ensayo4.Text = ""
    valor4.Text = ""
    Ensayo5.Text = ""
    valor5.Text = ""
    Ensayo6.Text = ""
    valor6.Text = ""
    Ensayo7.Text = ""
    valor7.Text = ""
    Ensayo8.Text = ""
    valor8.Text = ""
    Ensayo9.Text = ""
    valor9.Text = ""
    Ensayo10.Text = ""
    valor10.Text = ""
    Ensayo11.Text = ""
    Valor11.Text = ""
    Ensayo12.Text = ""
    Valor12.Text = ""
    Ensayo13.Text = ""
    Valor13.Text = ""
    Ensayo14.Text = ""
    Valor14.Text = ""
    Ensayo15.Text = ""
    Valor15.Text = ""
    Ensayo16.Text = ""
    Valor16.Text = ""
    Ensayo17.Text = ""
    Valor17.Text = ""
    Ensayo18.Text = ""
    Valor18.Text = ""
    Ensayo19.Text = ""
    Valor19.Text = ""
    Ensayo20.Text = ""
    Valor20.Text = ""
    Ensayo21.Text = ""
    Valor21.Text = ""
    Ensayo22.Text = ""
    Valor22.Text = ""
    Ensayo23.Text = ""
    Valor23.Text = ""
    Ensayo24.Text = ""
    Valor24.Text = ""
    Ensayo25.Text = ""
    Valor25.Text = ""
    Ensayo26.Text = ""
    Valor26.Text = ""
    Ensayo27.Text = ""
    Valor27.Text = ""
    Ensayo28.Text = ""
    Valor28.Text = ""
    Ensayo29.Text = ""
    Valor29.Text = ""
    Ensayo30.Text = ""
    Valor30.Text = ""
    
    Descriprod.Caption = ""
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    Descri7.Caption = ""
    Descri8.Caption = ""
    Descri9.Caption = ""
    Descri10.Caption = ""
    Descri11.Caption = ""
    Descri12.Caption = ""
    Descri13.Caption = ""
    Descri14.Caption = ""
    Descri15.Caption = ""
    Descri16.Caption = ""
    Descri17.Caption = ""
    Descri18.Caption = ""
    Descri19.Caption = ""
    Descri20.Caption = ""
    Descri21.Caption = ""
    Descri22.Caption = ""
    Descri23.Caption = ""
    Descri24.Caption = ""
    Descri25.Caption = ""
    Descri26.Caption = ""
    Descri27.Caption = ""
    Descri28.Caption = ""
    Descri29.Caption = ""
    Descri30.Caption = ""
    
    Version.Text = ""
    FechaInicio.Text = ""
    FechaFinal.Text = ""
    ControlCambio.Text = ""
    
    Codigo.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgEspecifiUnificaVersion.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
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
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEspecifiUnificaVersion.Caption = "Consulta de Versiones de Especificaciones de Materia Prima :  " + !Nombre
        End If
    End With
    EmpresaActual = WEmpresa
End Sub

Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            Codigo.Text = UCase(Codigo.Text)
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Descriprod.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                Version.SetFocus
                    Else
                Codigo.SetFocus
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Listado_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    ZSql = "DELETE ListaEspe"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
        
    ZSql = ""
    ZSql = ZSql + "INSERT INTO ListaEspe ("
    ZSql = ZSql + "Codigo ,"
    ZSql = ZSql + "Descripcion,"
    ZSql = ZSql + "Codigo1,"
    ZSql = ZSql + "Codigo2,"
    ZSql = ZSql + "Codigo3,"
    ZSql = ZSql + "Codigo4,"
    ZSql = ZSql + "Codigo5,"
    ZSql = ZSql + "Codigo6,"
    ZSql = ZSql + "Codigo7,"
    ZSql = ZSql + "Codigo8,"
    ZSql = ZSql + "Codigo9,"
    ZSql = ZSql + "Codigo10,"
    ZSql = ZSql + "Codigo11,"
    ZSql = ZSql + "Codigo12,"
    ZSql = ZSql + "Codigo13,"
    ZSql = ZSql + "Codigo14,"
    ZSql = ZSql + "Codigo15,"
    ZSql = ZSql + "Codigo16,"
    ZSql = ZSql + "Codigo17,"
    ZSql = ZSql + "Codigo18,"
    ZSql = ZSql + "Codigo19,"
    ZSql = ZSql + "Codigo20,"
    ZSql = ZSql + "Codigo21,"
    ZSql = ZSql + "Codigo22,"
    ZSql = ZSql + "Codigo23,"
    ZSql = ZSql + "Codigo24,"
    ZSql = ZSql + "Codigo25,"
    ZSql = ZSql + "Codigo26,"
    ZSql = ZSql + "Codigo27,"
    ZSql = ZSql + "Codigo28,"
    ZSql = ZSql + "Codigo29,"
    ZSql = ZSql + "Codigo30,"
    ZSql = ZSql + "Descri1,"
    ZSql = ZSql + "Descri2,"
    ZSql = ZSql + "Descri3,"
    ZSql = ZSql + "Descri4,"
    ZSql = ZSql + "Descri5,"
    ZSql = ZSql + "Descri6,"
    ZSql = ZSql + "Descri7,"
    ZSql = ZSql + "Descri8,"
    ZSql = ZSql + "Descri9,"
    ZSql = ZSql + "Descri10,"
    ZSql = ZSql + "Descri11,"
    ZSql = ZSql + "Descri12,"
    ZSql = ZSql + "Descri13,"
    ZSql = ZSql + "Descri14,"
    ZSql = ZSql + "Descri15,"
    ZSql = ZSql + "Descri16,"
    ZSql = ZSql + "Descri17,"
    ZSql = ZSql + "Descri18,"
    ZSql = ZSql + "Descri19,"
    ZSql = ZSql + "Descri20,"
    ZSql = ZSql + "Descri21,"
    ZSql = ZSql + "Descri22,"
    ZSql = ZSql + "Descri23,"
    ZSql = ZSql + "Descri24,"
    ZSql = ZSql + "Descri25,"
    ZSql = ZSql + "Descri26,"
    ZSql = ZSql + "Descri27,"
    ZSql = ZSql + "Descri28,"
    ZSql = ZSql + "Descri29,"
    ZSql = ZSql + "Descri30,"
    ZSql = ZSql + "Valor1,"
    ZSql = ZSql + "Valor2,"
    ZSql = ZSql + "Valor3,"
    ZSql = ZSql + "Valor4,"
    ZSql = ZSql + "Valor5,"
    ZSql = ZSql + "Valor6,"
    ZSql = ZSql + "Valor7,"
    ZSql = ZSql + "Valor8,"
    ZSql = ZSql + "Valor9,"
    ZSql = ZSql + "Valor10,"
    ZSql = ZSql + "ZValor1,"
    ZSql = ZSql + "ZValor2,"
    ZSql = ZSql + "ZValor3,"
    ZSql = ZSql + "ZValor4,"
    ZSql = ZSql + "ZValor5,"
    ZSql = ZSql + "ZValor6,"
    ZSql = ZSql + "ZValor7,"
    ZSql = ZSql + "ZValor8,"
    ZSql = ZSql + "ZValor9,"
    ZSql = ZSql + "ZValor10,"
    ZSql = ZSql + "ZValor121,"
    ZSql = ZSql + "ZValor122,"
    ZSql = ZSql + "ZValor123,"
    ZSql = ZSql + "ZValor124,"
    ZSql = ZSql + "ZValor125,"
    ZSql = ZSql + "ZValor126,"
    ZSql = ZSql + "ZValor127,"
    ZSql = ZSql + "ZValor128,"
    ZSql = ZSql + "ZValor129,"
    ZSql = ZSql + "ZValor130,"
    ZSql = ZSql + "Version ,"
    ZSql = ZSql + "Responsable,"
    ZSql = ZSql + "Fecha )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + Codigo.Text + "',"
    ZSql = ZSql + "'" + Descriprod.Caption + "',"
    ZSql = ZSql + "'" + Ensayo1.Text + "',"
    ZSql = ZSql + "'" + Ensayo2.Text + "',"
    ZSql = ZSql + "'" + Ensayo3.Text + "',"
    ZSql = ZSql + "'" + Ensayo4.Text + "',"
    ZSql = ZSql + "'" + Ensayo5.Text + "',"
    ZSql = ZSql + "'" + Ensayo6.Text + "',"
    ZSql = ZSql + "'" + Ensayo7.Text + "',"
    ZSql = ZSql + "'" + Ensayo8.Text + "',"
    ZSql = ZSql + "'" + Ensayo9.Text + "',"
    ZSql = ZSql + "'" + Ensayo10.Text + "',"
    ZSql = ZSql + "'" + Ensayo11.Text + "',"
    ZSql = ZSql + "'" + Ensayo12.Text + "',"
    ZSql = ZSql + "'" + Ensayo13.Text + "',"
    ZSql = ZSql + "'" + Ensayo14.Text + "',"
    ZSql = ZSql + "'" + Ensayo15.Text + "',"
    ZSql = ZSql + "'" + Ensayo16.Text + "',"
    ZSql = ZSql + "'" + Ensayo17.Text + "',"
    ZSql = ZSql + "'" + Ensayo18.Text + "',"
    ZSql = ZSql + "'" + Ensayo19.Text + "',"
    ZSql = ZSql + "'" + Ensayo20.Text + "',"
    ZSql = ZSql + "'" + Ensayo21.Text + "',"
    ZSql = ZSql + "'" + Ensayo22.Text + "',"
    ZSql = ZSql + "'" + Ensayo23.Text + "',"
    ZSql = ZSql + "'" + Ensayo24.Text + "',"
    ZSql = ZSql + "'" + Ensayo25.Text + "',"
    ZSql = ZSql + "'" + Ensayo26.Text + "',"
    ZSql = ZSql + "'" + Ensayo27.Text + "',"
    ZSql = ZSql + "'" + Ensayo28.Text + "',"
    ZSql = ZSql + "'" + Ensayo29.Text + "',"
    ZSql = ZSql + "'" + Ensayo30.Text + "',"
    ZSql = ZSql + "'" + Descri1.Caption + "',"
    ZSql = ZSql + "'" + descri2.Caption + "',"
    ZSql = ZSql + "'" + Descri3.Caption + "',"
    ZSql = ZSql + "'" + Descri4.Caption + "',"
    ZSql = ZSql + "'" + Descri5.Caption + "',"
    ZSql = ZSql + "'" + Descri6.Caption + "',"
    ZSql = ZSql + "'" + Descri7.Caption + "',"
    ZSql = ZSql + "'" + Descri8.Caption + "',"
    ZSql = ZSql + "'" + Descri9.Caption + "',"
    ZSql = ZSql + "'" + Descri10.Caption + "',"
    ZSql = ZSql + "'" + Descri11.Caption + "',"
    ZSql = ZSql + "'" + Descri12.Caption + "',"
    ZSql = ZSql + "'" + Descri13.Caption + "',"
    ZSql = ZSql + "'" + Descri14.Caption + "',"
    ZSql = ZSql + "'" + Descri15.Caption + "',"
    ZSql = ZSql + "'" + Descri16.Caption + "',"
    ZSql = ZSql + "'" + Descri17.Caption + "',"
    ZSql = ZSql + "'" + Descri18.Caption + "',"
    ZSql = ZSql + "'" + Descri19.Caption + "',"
    ZSql = ZSql + "'" + Descri20.Caption + "',"
    ZSql = ZSql + "'" + Descri21.Caption + "',"
    ZSql = ZSql + "'" + Descri22.Caption + "',"
    ZSql = ZSql + "'" + Descri23.Caption + "',"
    ZSql = ZSql + "'" + Descri24.Caption + "',"
    ZSql = ZSql + "'" + Descri25.Caption + "',"
    ZSql = ZSql + "'" + Descri26.Caption + "',"
    ZSql = ZSql + "'" + Descri27.Caption + "',"
    ZSql = ZSql + "'" + Descri28.Caption + "',"
    ZSql = ZSql + "'" + Descri29.Caption + "',"
    ZSql = ZSql + "'" + Descri30.Caption + "',"
    ZSql = ZSql + "'" + Valor1.Text + "',"
    ZSql = ZSql + "'" + valor2.Text + "',"
    ZSql = ZSql + "'" + Valor3.Text + "',"
    ZSql = ZSql + "'" + valor4.Text + "',"
    ZSql = ZSql + "'" + valor5.Text + "',"
    ZSql = ZSql + "'" + valor6.Text + "',"
    ZSql = ZSql + "'" + valor7.Text + "',"
    ZSql = ZSql + "'" + valor8.Text + "',"
    ZSql = ZSql + "'" + valor9.Text + "',"
    ZSql = ZSql + "'" + valor10.Text + "',"
    ZSql = ZSql + "'" + Valor11.Text + "',"
    ZSql = ZSql + "'" + Valor12.Text + "',"
    ZSql = ZSql + "'" + Valor13.Text + "',"
    ZSql = ZSql + "'" + Valor14.Text + "',"
    ZSql = ZSql + "'" + Valor15.Text + "',"
    ZSql = ZSql + "'" + Valor16.Text + "',"
    ZSql = ZSql + "'" + Valor17.Text + "',"
    ZSql = ZSql + "'" + Valor18.Text + "',"
    ZSql = ZSql + "'" + Valor19.Text + "',"
    ZSql = ZSql + "'" + Valor20.Text + "',"
    ZSql = ZSql + "'" + Valor21.Text + "',"
    ZSql = ZSql + "'" + Valor22.Text + "',"
    ZSql = ZSql + "'" + Valor23.Text + "',"
    ZSql = ZSql + "'" + Valor24.Text + "',"
    ZSql = ZSql + "'" + Valor25.Text + "',"
    ZSql = ZSql + "'" + Valor26.Text + "',"
    ZSql = ZSql + "'" + Valor27.Text + "',"
    ZSql = ZSql + "'" + Valor28.Text + "',"
    ZSql = ZSql + "'" + Valor29.Text + "',"
    ZSql = ZSql + "'" + Valor30.Text + "',"
    ZSql = ZSql + "'" + Version.Text + "',"
    ZSql = ZSql + "'" + FechaInicio.Text + "',"
    ZSql = ZSql + "'" + FechaFinal.Text + "')"
    
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
        
    
    Lista.WindowTitle = "Listado de Especificaciones de Materia Prima (Unificado)"
    Lista.WindowTop = 0
    Lista.WindowLeft = 0
    Lista.WindowWidth = Screen.Width
    Lista.WindowHeight = Screen.Height

    Rem lista.GroupSelectionFormula = "{EspecificacionesUnifica.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Lista.Destination = 1
    Rem lista.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Val(WEmpresa) = 3 Then
        Lista.ReportFileName = "ListaEspeVersion.rpt"
            Else
        Lista.ReportFileName = "ListaEspeVersionPelli.rpt"
    End If
    
    Rem Lista.SQLQuery = "SELECT ListaEspe.Codigo, ListaEspe.Descripcion, " _
    rem             + "ListaEspe.Codigo1, ListaEspe.Codigo2, ListaEspe.Codigo3, ListaEspe.Codigo4, ListaEspe.Codigo5, ListaEspe.Codigo6, ListaEspe.Codigo7, ListaEspe.Codigo8, ListaEspe.Codigo9, ListaEspe.Codigo10, ListaEspe.Codigo11, ListaEspe.Codigo12, ListaEspe.Codigo13, ListaEspe.Codigo14, ListaEspe.Codigo15, ListaEspe.Codigo16, ListaEspe.Codigo17, ListaEspe.Codigo18, ListaEspe.Codigo19, ListaEspe.Codigo20, ListaEspe.Codigo21, ListaEspe.Codigo22, ListaEspe.Codigo23, ListaEspe.Codigo24, ListaEspe.Codigo25, ListaEspe.Codigo26, ListaEspe.Codigo27, ListaEspe.Codigo28, ListaEspe.Codigo20, ListaEspe.Codigo30, " _
    rem             + "ListaEspe.Descri1, ListaEspe.Descri2, ListaEspe.Descri3, ListaEspe.Descri4, ListaEspe.Descri5, ListaEspe.Descri6, ListaEspe.Descri7, ListaEspe.Descri8, ListaEspe.Descri9, ListaEspe.Descri10, ListaEspe.Descri11, List1Espe.Descri12, ListaEspe.Descri13, ListaEspe.Descri14, ListaEspe.Descri15, ListaEspe.Descri16, ListaEspe.Descri17, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, ListaEspe.Descri10, " _
    rem             + "ListaEspe.Valor1, ListaEspe.Valor2, ListaEspe.Valor3, ListaEspe.Valor4, ListaEspe.Valor5, ListaEspe.Valor6, ListaEspe.Valor7, ListaEspe.Valor8, ListaEspe.Valor9, ListaEspe.Valor10, " _
    rem             + "ListaEspe.Version, ListaEspe.Responsable, ListaEspe.Fecha " _
    rem             + "From " _
    rem             + DSQ + ".dbo.ListaEspe ListaEspe " _
    rem             + "Where " _
    rem             + "ListaEspe.Codigo >= '" + Codigo.Text + "' AND " _
    rem             + "ListaEspe.Codigo <= '" + Codigo.Text + "'"
    
    Lista.SQLQuery = "SELECT * " _
                + "From " _
                + DSQ + ".dbo.ListaEspe ListaEspe " _
                + "Where " _
                + "ListaEspe.Codigo >= '" + Codigo.Text + "' AND " _
                + "ListaEspe.Codigo <= '" + Codigo.Text + "'"
    
    Lista.Connect = Connect()
    
    Lista.Action = 1
    
    Call Conecta_Empresa

End Sub

Sub Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            Codigo.Text = UCase(Codigo.Text)
            
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
                
            Sql1 = "Select EspecificacionesUnificaVersion.Producto"
            Sql2 = " FROM EspecificacionesUnificaVersion"
            Sql3 = " Where EspecificacionesUnificaVersion.Producto = " + "'" + Codigo.Text + "'"
            Sql4 = " and EspecificacionesUnificaVersion.Version = " + "'" + Version.Text + "'"
            spEspecificacionesUnificaVersion = Sql1 + Sql2 + Sql3 + Sql4
            Set rstEspecificacionesUnificaVersion = db.OpenRecordset(spEspecificacionesUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificacionesUnificaVersion.RecordCount > 0 Then
                rstEspecificacionesUnificaVersion.Close
                Call Conecta_Empresa
                Call Imprime_Datos
                    Else
                XCodigo = Codigo.Text
                XVersion = Version.Text
                Call CmdLimpiar_Click
                Codigo.Text = XCodigo
                Version.Text = XVersion
                Call Conecta_Empresa
                Version.SetFocus
            End If
            
        End If
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Codigos"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()
    Opcion.Visible = False
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    
End Sub



