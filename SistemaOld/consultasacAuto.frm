VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaSacauto 
   Caption         =   "Consulta de SAC"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11775
   Begin VB.CommandButton Impresion 
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
      Left            =   1680
      MouseIcon       =   "consultasacAuto.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "consultasacAuto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   209
      ToolTipText     =   "Impresion de Orden de Pago"
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox Ayuda 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFFC0&
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
      ItemData        =   "consultasacAuto.frx":0B4C
      Left            =   2040
      List            =   "consultasacAuto.frx":0B53
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Opcion 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3120
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "NUMERO"
      Height          =   735
      Left            =   9240
      TabIndex        =   164
      Top             =   7200
      Width           =   2175
      Begin VB.Image AnteriorNro 
         Height          =   480
         Left            =   840
         MouseIcon       =   "consultasacAuto.frx":0B61
         MousePointer    =   99  'Custom
         Picture         =   "consultasacAuto.frx":0E6B
         ToolTipText     =   "Registro Anterior"
         Top             =   120
         Width           =   480
      End
      Begin VB.Image SiguienteNro 
         Height          =   480
         Left            =   1440
         MouseIcon       =   "consultasacAuto.frx":12AD
         MousePointer    =   99  'Custom
         Picture         =   "consultasacAuto.frx":15B7
         ToolTipText     =   "Registro Posterior"
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TIPO"
      Height          =   735
      Left            =   7080
      TabIndex        =   163
      Top             =   7200
      Width           =   2055
      Begin VB.Image SiguienteTipo 
         Height          =   480
         Left            =   1320
         MouseIcon       =   "consultasacAuto.frx":19F9
         MousePointer    =   99  'Custom
         Picture         =   "consultasacAuto.frx":1D03
         ToolTipText     =   "Registro Posterior"
         Top             =   120
         Width           =   480
      End
      Begin VB.Image AnteriorTipo 
         Height          =   480
         Left            =   600
         MouseIcon       =   "consultasacAuto.frx":2145
         MousePointer    =   99  'Custom
         Picture         =   "consultasacAuto.frx":244F
         ToolTipText     =   "Registro Anterior"
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.TextBox Referencia 
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
      MaxLength       =   100
      TabIndex        =   161
      Text            =   " "
      Top             =   1200
      Width           =   10455
   End
   Begin VB.TextBox Numero 
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
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   150
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Centro 
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
      Left            =   8400
      MaxLength       =   6
      TabIndex        =   149
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Estado 
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
      Left            =   9000
      TabIndex        =   148
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox Origen 
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
      Left            =   5520
      TabIndex        =   147
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Ano 
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
      Left            =   4320
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   146
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Tipo 
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox ResponsableEmisor 
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   129
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox ResponsableDestino 
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
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   128
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Titulo 
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
      MaxLength       =   100
      TabIndex        =   127
      Text            =   " "
      Top             =   1560
      Width           =   10455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   9128
      _Version        =   327680
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion de la NC"
      TabPicture(0)   =   "consultasacAuto.frx":2891
      Tab(0).ControlCount=   4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "IngresoCausa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "IngresoNoCon"
      Tab(0).Control(3).Enabled=   0   'False
      TabCaption(1)   =   "Acciones"
      TabPicture(1)   =   "consultasacAuto.frx":28AD
      Tab(1).ControlCount=   39
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DesResponsable1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DesResponsable2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DesResponsable3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DesResponsable4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DesResponsable5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DesResponsable6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label26"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label27"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label28"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label29"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label30"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label31"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Plazo6"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Plazo5"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Plazo4"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Plazo3"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Plazo2"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Plazo1"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Accion11"
      Tab(1).Control(21).Enabled=   -1  'True
      Tab(1).Control(22)=   "Accion12"
      Tab(1).Control(22).Enabled=   -1  'True
      Tab(1).Control(23)=   "Accion21"
      Tab(1).Control(23).Enabled=   -1  'True
      Tab(1).Control(24)=   "Accion22"
      Tab(1).Control(24).Enabled=   -1  'True
      Tab(1).Control(25)=   "Accion31"
      Tab(1).Control(25).Enabled=   -1  'True
      Tab(1).Control(26)=   "Accion32"
      Tab(1).Control(26).Enabled=   -1  'True
      Tab(1).Control(27)=   "Accion41"
      Tab(1).Control(27).Enabled=   -1  'True
      Tab(1).Control(28)=   "Accion42"
      Tab(1).Control(28).Enabled=   -1  'True
      Tab(1).Control(29)=   "Accion51"
      Tab(1).Control(29).Enabled=   -1  'True
      Tab(1).Control(30)=   "Accion52"
      Tab(1).Control(30).Enabled=   -1  'True
      Tab(1).Control(31)=   "Accion61"
      Tab(1).Control(31).Enabled=   -1  'True
      Tab(1).Control(32)=   "Accion62"
      Tab(1).Control(32).Enabled=   -1  'True
      Tab(1).Control(33)=   "Responsable1"
      Tab(1).Control(33).Enabled=   -1  'True
      Tab(1).Control(34)=   "Responsable2"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Responsable3"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Responsable4"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Responsable5"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Responsable6"
      Tab(1).Control(38).Enabled=   -1  'True
      TabCaption(2)   =   "Implementacion"
      TabPicture(2)   =   "consultasacAuto.frx":28C9
      Tab(2).ControlCount=   59
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Estado11"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Estado12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Estado13"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Estado14"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Estado15"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Estado16"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Responsable15"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Responsable16"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Responsable14"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Responsable13"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Responsable12"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Comentario11"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Comentario12"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Comentario21"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Comentario22"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Comentario31"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Comentario32"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Comentario41"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Comentario42"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Comentario51"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Comentario52"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Comentario61"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Comentario62"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Responsable11"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Accion162"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Accion161"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Accion152"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Accion151"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Accion142"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Accion141"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Accion132"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Accion131"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Accion122"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Accion121"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Accion112"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Accion111"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Fecha1"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Fecha2"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Fecha3"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Fecha4"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Fecha5"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Fecha6"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Label37"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Label36"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Label35"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Label34"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Label33"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Label32"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Label24"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "DesResponsable16"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "DesResponsable15"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "DesResponsable14"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "DesResponsable13"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "DesResponsable12"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Label20"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "Label18"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "Label17"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "DesResponsable11"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "Label7"
      Tab(2).Control(58).Enabled=   0   'False
      TabCaption(3)   =   "Verificacion"
      TabPicture(3)   =   "consultasacAuto.frx":28E5
      Tab(3).ControlCount=   70
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Responsable31"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "Responsable32"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "Responsable33"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "Responsable34"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "Responsable35"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "Responsable36"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "Estado31"
      Tab(3).Control(6).Enabled=   -1  'True
      Tab(3).Control(7)=   "Estado32"
      Tab(3).Control(7).Enabled=   -1  'True
      Tab(3).Control(8)=   "Estado33"
      Tab(3).Control(8).Enabled=   -1  'True
      Tab(3).Control(9)=   "Estado34"
      Tab(3).Control(9).Enabled=   -1  'True
      Tab(3).Control(10)=   "Estado35"
      Tab(3).Control(10).Enabled=   -1  'True
      Tab(3).Control(11)=   "Estado36"
      Tab(3).Control(11).Enabled=   -1  'True
      Tab(3).Control(12)=   "Estado6"
      Tab(3).Control(12).Enabled=   -1  'True
      Tab(3).Control(13)=   "Estado5"
      Tab(3).Control(13).Enabled=   -1  'True
      Tab(3).Control(14)=   "Estado4"
      Tab(3).Control(14).Enabled=   -1  'True
      Tab(3).Control(15)=   "Estado3"
      Tab(3).Control(15).Enabled=   -1  'True
      Tab(3).Control(16)=   "Estado2"
      Tab(3).Control(16).Enabled=   -1  'True
      Tab(3).Control(17)=   "Estado1"
      Tab(3).Control(17).Enabled=   -1  'True
      Tab(3).Control(18)=   "Responsable26"
      Tab(3).Control(18).Enabled=   -1  'True
      Tab(3).Control(19)=   "Responsable25"
      Tab(3).Control(19).Enabled=   -1  'True
      Tab(3).Control(20)=   "Responsable24"
      Tab(3).Control(20).Enabled=   -1  'True
      Tab(3).Control(21)=   "Responsable23"
      Tab(3).Control(21).Enabled=   -1  'True
      Tab(3).Control(22)=   "Responsable22"
      Tab(3).Control(22).Enabled=   -1  'True
      Tab(3).Control(23)=   "Comentario211"
      Tab(3).Control(23).Enabled=   -1  'True
      Tab(3).Control(24)=   "Comentario212"
      Tab(3).Control(24).Enabled=   -1  'True
      Tab(3).Control(25)=   "Comentario221"
      Tab(3).Control(25).Enabled=   -1  'True
      Tab(3).Control(26)=   "Comentario222"
      Tab(3).Control(26).Enabled=   -1  'True
      Tab(3).Control(27)=   "Comentario231"
      Tab(3).Control(27).Enabled=   -1  'True
      Tab(3).Control(28)=   "Comentario232"
      Tab(3).Control(28).Enabled=   -1  'True
      Tab(3).Control(29)=   "Comentario241"
      Tab(3).Control(29).Enabled=   -1  'True
      Tab(3).Control(30)=   "Comentario242"
      Tab(3).Control(30).Enabled=   -1  'True
      Tab(3).Control(31)=   "Comentario251"
      Tab(3).Control(31).Enabled=   -1  'True
      Tab(3).Control(32)=   "Comentario252"
      Tab(3).Control(32).Enabled=   -1  'True
      Tab(3).Control(33)=   "Comentario261"
      Tab(3).Control(33).Enabled=   -1  'True
      Tab(3).Control(34)=   "Comentario262"
      Tab(3).Control(34).Enabled=   -1  'True
      Tab(3).Control(35)=   "Responsable21"
      Tab(3).Control(35).Enabled=   -1  'True
      Tab(3).Control(36)=   "Fecha21"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "Fecha22"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Fecha23"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Fecha24"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Fecha25"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Fecha26"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Fecha31"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Fecha32"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "Fecha33"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "Fecha34"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).Control(46)=   "Fecha35"
      Tab(3).Control(46).Enabled=   0   'False
      Tab(3).Control(47)=   "Fecha36"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).Control(48)=   "Label51"
      Tab(3).Control(48).Enabled=   0   'False
      Tab(3).Control(49)=   "DesResponsable31"
      Tab(3).Control(49).Enabled=   0   'False
      Tab(3).Control(50)=   "DesResponsable32"
      Tab(3).Control(50).Enabled=   0   'False
      Tab(3).Control(51)=   "DesResponsable33"
      Tab(3).Control(51).Enabled=   0   'False
      Tab(3).Control(52)=   "DesResponsable34"
      Tab(3).Control(52).Enabled=   0   'False
      Tab(3).Control(53)=   "DesResponsable35"
      Tab(3).Control(53).Enabled=   0   'False
      Tab(3).Control(54)=   "DesResponsable36"
      Tab(3).Control(54).Enabled=   0   'False
      Tab(3).Control(55)=   "Label43"
      Tab(3).Control(55).Enabled=   0   'False
      Tab(3).Control(56)=   "Label42"
      Tab(3).Control(56).Enabled=   0   'False
      Tab(3).Control(57)=   "Label41"
      Tab(3).Control(57).Enabled=   0   'False
      Tab(3).Control(58)=   "Label40"
      Tab(3).Control(58).Enabled=   0   'False
      Tab(3).Control(59)=   "Label39"
      Tab(3).Control(59).Enabled=   0   'False
      Tab(3).Control(60)=   "Label38"
      Tab(3).Control(60).Enabled=   0   'False
      Tab(3).Control(61)=   "DesResponsable26"
      Tab(3).Control(61).Enabled=   0   'False
      Tab(3).Control(62)=   "DesResponsable25"
      Tab(3).Control(62).Enabled=   0   'False
      Tab(3).Control(63)=   "DesResponsable24"
      Tab(3).Control(63).Enabled=   0   'False
      Tab(3).Control(64)=   "DesResponsable23"
      Tab(3).Control(64).Enabled=   0   'False
      Tab(3).Control(65)=   "DesResponsable22"
      Tab(3).Control(65).Enabled=   0   'False
      Tab(3).Control(66)=   "Label23"
      Tab(3).Control(66).Enabled=   0   'False
      Tab(3).Control(67)=   "DesResponsable21"
      Tab(3).Control(67).Enabled=   0   'False
      Tab(3).Control(68)=   "Label8"
      Tab(3).Control(68).Enabled=   0   'False
      Tab(3).Control(69)=   "Label6"
      Tab(3).Control(69).Enabled=   0   'False
      TabCaption(4)   =   "Comentarios"
      TabPicture(4)   =   "consultasacAuto.frx":2901
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Comentario"
      Tab(4).Control(0).Enabled=   -1  'True
      Begin VB.TextBox Responsable31 
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
         Left            =   -70800
         MaxLength       =   6
         TabIndex        =   195
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Responsable32 
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
         Left            =   -70800
         MaxLength       =   6
         TabIndex        =   194
         Text            =   " "
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Responsable33 
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
         Left            =   -70800
         MaxLength       =   6
         TabIndex        =   193
         Text            =   " "
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Responsable34 
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
         Left            =   -70800
         MaxLength       =   6
         TabIndex        =   192
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Responsable35 
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
         Left            =   -70800
         MaxLength       =   6
         TabIndex        =   191
         Text            =   " "
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Responsable36 
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
         Left            =   -70800
         MaxLength       =   6
         TabIndex        =   190
         Text            =   " "
         Top             =   4320
         Width           =   615
      End
      Begin VB.ComboBox Estado31 
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
         Left            =   -68880
         TabIndex        =   189
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Estado32 
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
         Left            =   -68880
         TabIndex        =   188
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox Estado33 
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
         Left            =   -68880
         TabIndex        =   187
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Estado34 
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
         Left            =   -68880
         TabIndex        =   186
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Estado35 
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
         Left            =   -68880
         TabIndex        =   185
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox Estado36 
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
         Left            =   -68880
         TabIndex        =   184
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Comentario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   171
         Top             =   480
         Width           =   11175
      End
      Begin VB.ComboBox Estado11 
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
         Left            =   -68160
         TabIndex        =   144
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Estado12 
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
         Left            =   -68160
         TabIndex        =   143
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox Estado13 
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
         Left            =   -68160
         TabIndex        =   142
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Estado14 
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
         Left            =   -68160
         TabIndex        =   141
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Estado15 
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
         Left            =   -68160
         TabIndex        =   140
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox Estado16 
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
         Left            =   -68160
         TabIndex        =   139
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox IngresoNoCon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   136
         Top             =   840
         Width           =   11175
      End
      Begin VB.TextBox IngresoCausa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   135
         Top             =   3120
         Width           =   11175
      End
      Begin VB.ComboBox Estado6 
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
         Left            =   -72240
         TabIndex        =   126
         Top             =   4320
         Width           =   1335
      End
      Begin VB.ComboBox Estado5 
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
         Left            =   -72240
         TabIndex        =   125
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox Estado4 
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
         Left            =   -72240
         TabIndex        =   124
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox Estado3 
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
         Left            =   -72240
         TabIndex        =   123
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox Estado2 
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
         Left            =   -72240
         TabIndex        =   122
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox Estado1 
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
         Left            =   -72240
         TabIndex        =   121
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Responsable26 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   118
         Text            =   " "
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Responsable25 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   115
         Text            =   " "
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Responsable24 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   112
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Responsable23 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   109
         Text            =   " "
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Responsable22 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   106
         Text            =   " "
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Comentario211 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   105
         Text            =   " "
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox Comentario212 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   104
         Text            =   " "
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Comentario221 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   103
         Text            =   " "
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Comentario222 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   102
         Text            =   " "
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox Comentario231 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   101
         Text            =   " "
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox Comentario232 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   100
         Text            =   " "
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox Comentario241 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   99
         Text            =   " "
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox Comentario242 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   98
         Text            =   " "
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox Comentario251 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   97
         Text            =   " "
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox Comentario252 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   96
         Text            =   " "
         Top             =   3840
         Width           =   3615
      End
      Begin VB.TextBox Comentario261 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   95
         Text            =   " "
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox Comentario262 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   94
         Text            =   " "
         Top             =   4560
         Width           =   3615
      End
      Begin VB.TextBox Responsable21 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   88
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Responsable15 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   85
         Text            =   " "
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Responsable16 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   82
         Text            =   " "
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Responsable14 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   79
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Responsable13 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   76
         Text            =   " "
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Responsable12 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   73
         Text            =   " "
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Comentario11 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   72
         Text            =   " "
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Comentario12 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   71
         Text            =   " "
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Comentario21 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   70
         Text            =   " "
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Comentario22 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   69
         Text            =   " "
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Comentario31 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   68
         Text            =   " "
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Comentario32 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   67
         Text            =   " "
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Comentario41 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   66
         Text            =   " "
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Comentario42 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   65
         Text            =   " "
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox Comentario51 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   64
         Text            =   " "
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox Comentario52 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   63
         Text            =   " "
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Comentario61 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   62
         Text            =   " "
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox Comentario62 
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
         Left            =   -66720
         MaxLength       =   50
         TabIndex        =   61
         Text            =   " "
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox Responsable11 
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
         Left            =   -71640
         MaxLength       =   6
         TabIndex        =   54
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Accion162 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   53
         Text            =   " "
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox Accion161 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         Text            =   " "
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox Accion152 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         Text            =   " "
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Accion151 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Text            =   " "
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Accion142 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Text            =   " "
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox Accion141 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Text            =   " "
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox Accion132 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox Accion131 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox Accion122 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   45
         Text            =   " "
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Accion121 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         Text            =   " "
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Accion112 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   43
         Text            =   " "
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Accion111 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         Text            =   " "
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Responsable6 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   39
         Text            =   " "
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Responsable5 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   36
         Text            =   " "
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Responsable4 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   33
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Responsable3 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   30
         Text            =   " "
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Responsable2 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   27
         Text            =   " "
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Responsable1 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   21
         Text            =   " "
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Accion62 
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
         MaxLength       =   60
         TabIndex        =   20
         Text            =   " "
         Top             =   4560
         Width           =   8055
      End
      Begin VB.TextBox Accion61 
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
         MaxLength       =   60
         TabIndex        =   19
         Text            =   " "
         Top             =   4320
         Width           =   8055
      End
      Begin VB.TextBox Accion52 
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
         MaxLength       =   60
         TabIndex        =   18
         Text            =   " "
         Top             =   3840
         Width           =   8055
      End
      Begin VB.TextBox Accion51 
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
         MaxLength       =   60
         TabIndex        =   17
         Text            =   " "
         Top             =   3600
         Width           =   8055
      End
      Begin VB.TextBox Accion42 
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
         MaxLength       =   60
         TabIndex        =   16
         Text            =   " "
         Top             =   3120
         Width           =   8055
      End
      Begin VB.TextBox Accion41 
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
         MaxLength       =   60
         TabIndex        =   15
         Text            =   " "
         Top             =   2880
         Width           =   8055
      End
      Begin VB.TextBox Accion32 
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
         MaxLength       =   60
         TabIndex        =   14
         Text            =   " "
         Top             =   2400
         Width           =   8055
      End
      Begin VB.TextBox Accion31 
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
         MaxLength       =   60
         TabIndex        =   13
         Text            =   " "
         Top             =   2160
         Width           =   8055
      End
      Begin VB.TextBox Accion22 
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
         MaxLength       =   60
         TabIndex        =   12
         Text            =   " "
         Top             =   1800
         Width           =   8055
      End
      Begin VB.TextBox Accion21 
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
         MaxLength       =   60
         TabIndex        =   11
         Text            =   " "
         Top             =   1560
         Width           =   8055
      End
      Begin VB.TextBox Accion12 
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
         MaxLength       =   60
         TabIndex        =   10
         Text            =   " "
         Top             =   1080
         Width           =   8055
      End
      Begin VB.TextBox Accion11 
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
         MaxLength       =   60
         TabIndex        =   9
         Text            =   " "
         Top             =   840
         Width           =   8055
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   -360
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3555 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
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
      Begin MSMask.MaskEdBox Plazo1 
         Height          =   285
         Left            =   -64560
         TabIndex        =   24
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo2 
         Height          =   285
         Left            =   -64560
         TabIndex        =   29
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo3 
         Height          =   285
         Left            =   -64560
         TabIndex        =   32
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo4 
         Height          =   285
         Left            =   -64560
         TabIndex        =   35
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo5 
         Height          =   285
         Left            =   -64560
         TabIndex        =   38
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Plazo6 
         Height          =   285
         Left            =   -64560
         TabIndex        =   41
         Top             =   4320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha1 
         Height          =   285
         Left            =   -69600
         TabIndex        =   57
         Top             =   840
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha2 
         Height          =   285
         Left            =   -69600
         TabIndex        =   74
         Top             =   1560
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha3 
         Height          =   285
         Left            =   -69600
         TabIndex        =   77
         Top             =   2160
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha4 
         Height          =   285
         Left            =   -69600
         TabIndex        =   80
         Top             =   2880
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha5 
         Height          =   285
         Left            =   -69600
         TabIndex        =   83
         Top             =   3600
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha6 
         Height          =   285
         Left            =   -69600
         TabIndex        =   86
         Top             =   4320
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha21 
         Height          =   285
         Left            =   -74160
         TabIndex        =   89
         Top             =   1080
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha22 
         Height          =   285
         Left            =   -74160
         TabIndex        =   107
         Top             =   1800
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha23 
         Height          =   285
         Left            =   -74160
         TabIndex        =   110
         Top             =   2400
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha24 
         Height          =   285
         Left            =   -74160
         TabIndex        =   113
         Top             =   3120
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha25 
         Height          =   285
         Left            =   -74160
         TabIndex        =   116
         Top             =   3840
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha26 
         Height          =   285
         Left            =   -74160
         TabIndex        =   119
         Top             =   4560
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha31 
         Height          =   285
         Left            =   -70800
         TabIndex        =   196
         Top             =   1080
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha32 
         Height          =   285
         Left            =   -70800
         TabIndex        =   197
         Top             =   1800
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha33 
         Height          =   285
         Left            =   -70800
         TabIndex        =   198
         Top             =   2400
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha34 
         Height          =   285
         Left            =   -70800
         TabIndex        =   199
         Top             =   3120
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha35 
         Height          =   285
         Left            =   -70800
         TabIndex        =   200
         Top             =   3840
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha36 
         Height          =   285
         Left            =   -70800
         TabIndex        =   201
         Top             =   4560
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Verificacion Efectividad"
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
         Left            =   -70800
         TabIndex        =   208
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label DesResponsable31 
         BackColor       =   &H00C0C000&
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
         Left            =   -70080
         TabIndex        =   207
         Top             =   840
         Width           =   975
      End
      Begin VB.Label DesResponsable32 
         BackColor       =   &H00C0C000&
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
         Left            =   -70080
         TabIndex        =   206
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label DesResponsable33 
         BackColor       =   &H00C0C000&
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
         Left            =   -70080
         TabIndex        =   205
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label DesResponsable34 
         BackColor       =   &H00C0C000&
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
         Left            =   -70080
         TabIndex        =   204
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label DesResponsable35 
         BackColor       =   &H00C0C000&
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
         Left            =   -70080
         TabIndex        =   203
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label DesResponsable36 
         BackColor       =   &H00C0C000&
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
         Left            =   -70080
         TabIndex        =   202
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label43 
         Caption         =   "1"
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
         Left            =   -74760
         TabIndex        =   183
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label42 
         Caption         =   "2"
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
         Left            =   -74760
         TabIndex        =   182
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label41 
         Caption         =   "3"
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
         Left            =   -74760
         TabIndex        =   181
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label40 
         Caption         =   "4"
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
         Left            =   -74760
         TabIndex        =   180
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label39 
         Caption         =   "5"
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
         Left            =   -74760
         TabIndex        =   179
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label38 
         Caption         =   "6"
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
         Left            =   -74760
         TabIndex        =   178
         Top             =   4320
         Width           =   135
      End
      Begin VB.Label Label37 
         Caption         =   "1"
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
         Left            =   -74920
         TabIndex        =   177
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label36 
         Caption         =   "2"
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
         Left            =   -74920
         TabIndex        =   176
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label35 
         Caption         =   "3"
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
         Left            =   -74920
         TabIndex        =   175
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label34 
         Caption         =   "4"
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
         Left            =   -74920
         TabIndex        =   174
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label33 
         Caption         =   "5"
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
         Left            =   -74920
         TabIndex        =   173
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label32 
         Caption         =   "6"
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
         Left            =   -74920
         TabIndex        =   172
         Top             =   4320
         Width           =   135
      End
      Begin VB.Label Label31 
         Caption         =   "6"
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
         Left            =   -74920
         TabIndex        =   170
         Top             =   4320
         Width           =   135
      End
      Begin VB.Label Label30 
         Caption         =   "5"
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
         Left            =   -74920
         TabIndex        =   169
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label29 
         Caption         =   "4"
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
         Left            =   -74920
         TabIndex        =   168
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label28 
         Caption         =   "3"
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
         Left            =   -74920
         TabIndex        =   167
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label27 
         Caption         =   "2"
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
         Left            =   -74920
         TabIndex        =   166
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label26 
         Caption         =   "1"
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
         Left            =   -74920
         TabIndex        =   165
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
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
         Left            =   -68160
         TabIndex        =   145
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Causas que lo originaron"
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
         Height          =   300
         Left            =   120
         TabIndex        =   138
         Top             =   2760
         Width           =   11175
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion de la No Conformidad"
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
         Height          =   300
         Left            =   120
         TabIndex        =   137
         Top             =   480
         Width           =   11175
      End
      Begin VB.Label DesResponsable26 
         BackColor       =   &H00C0C000&
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
         Left            =   -73440
         TabIndex        =   120
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label DesResponsable25 
         BackColor       =   &H00C0C000&
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
         Left            =   -73440
         TabIndex        =   117
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label DesResponsable24 
         BackColor       =   &H00C0C000&
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
         Left            =   -73440
         TabIndex        =   114
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label DesResponsable23 
         BackColor       =   &H00C0C000&
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
         Left            =   -73440
         TabIndex        =   111
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label DesResponsable22 
         BackColor       =   &H00C0C000&
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
         Left            =   -73440
         TabIndex        =   108
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Verificacion Implementacion"
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
         Left            =   -74160
         TabIndex        =   93
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label DesResponsable21 
         BackColor       =   &H00C0C000&
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
         Left            =   -73440
         TabIndex        =   92
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acc"
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
         Left            =   -74880
         TabIndex        =   91
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -67440
         TabIndex        =   90
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label DesResponsable16 
         BackColor       =   &H00C0C000&
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
         Left            =   -70920
         TabIndex        =   87
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label DesResponsable15 
         BackColor       =   &H00C0C000&
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
         Left            =   -70920
         TabIndex        =   84
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label DesResponsable14 
         BackColor       =   &H00C0C000&
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
         Left            =   -70920
         TabIndex        =   81
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label DesResponsable13 
         BackColor       =   &H00C0C000&
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
         Left            =   -70920
         TabIndex        =   78
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label DesResponsable12 
         BackColor       =   &H00C0C000&
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
         Left            =   -70920
         TabIndex        =   75
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -66720
         TabIndex        =   60
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
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
         Left            =   -74760
         TabIndex        =   59
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -69600
         TabIndex        =   58
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label DesResponsable11 
         BackColor       =   &H00C0C000&
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
         Left            =   -70920
         TabIndex        =   56
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
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
         Left            =   -71640
         TabIndex        =   55
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label DesResponsable6 
         BackColor       =   &H00C0C000&
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
         Left            =   -65880
         TabIndex        =   40
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label DesResponsable5 
         BackColor       =   &H00C0C000&
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
         Left            =   -65880
         TabIndex        =   37
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label DesResponsable4 
         BackColor       =   &H00C0C000&
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
         Left            =   -65880
         TabIndex        =   34
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label DesResponsable3 
         BackColor       =   &H00C0C000&
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
         Left            =   -65880
         TabIndex        =   31
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label DesResponsable2 
         BackColor       =   &H00C0C000&
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
         Left            =   -65880
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
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
         Left            =   -74760
         TabIndex        =   26
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plazo"
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
         Left            =   -64560
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label DesResponsable1 
         BackColor       =   &H00C0C000&
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
         Left            =   -65880
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
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
         Left            =   -66600
         TabIndex        =   22
         Top             =   480
         Width           =   1935
      End
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1320
      TabIndex        =   151
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.Label Label25 
      Caption         =   "Referencia"
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
      Left            =   240
      TabIndex        =   162
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "Año"
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
      Left            =   3600
      TabIndex        =   160
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label15 
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
      Left            =   240
      TabIndex        =   159
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Centro"
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
      Left            =   7320
      TabIndex        =   158
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesCentro 
      BackColor       =   &H00C0C000&
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
      Left            =   9360
      TabIndex        =   157
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
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
      Left            =   8040
      TabIndex        =   156
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Origen"
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
      Left            =   4080
      TabIndex        =   155
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Numero"
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
      Left            =   5280
      TabIndex        =   154
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
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
      Left            =   240
      TabIndex        =   153
      Top             =   120
      Width           =   735
   End
   Begin VB.Label DesTipo 
      BackColor       =   &H00C0C000&
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
      Left            =   2160
      TabIndex        =   152
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Emisor"
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
      Left            =   240
      TabIndex        =   134
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Resp. Inv."
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
      Left            =   5520
      TabIndex        =   133
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label DesResponsableEmisor 
      BackColor       =   &H00C0C000&
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
      Left            =   2280
      TabIndex        =   132
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label DesResponsableDestino 
      BackColor       =   &H00C0C000&
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
      Left            =   7560
      TabIndex        =   131
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo"
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
      Left            =   240
      TabIndex        =   130
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   4680
      MouseIcon       =   "consultasacAuto.frx":291D
      MousePointer    =   99  'Custom
      Picture         =   "consultasacAuto.frx":2C27
      ToolTipText     =   "Consulta de Datos"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6360
      MouseIcon       =   "consultasacAuto.frx":3469
      MousePointer    =   99  'Custom
      Picture         =   "consultasacAuto.frx":3773
      ToolTipText     =   "Salida"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "consultasacAuto.frx":3FB5
      MousePointer    =   99  'Custom
      Picture         =   "consultasacAuto.frx":42BF
      ToolTipText     =   "Elimina el Registro"
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "consultasacAuto.frx":4B01
      MousePointer    =   99  'Custom
      Picture         =   "consultasacAuto.frx":4E0B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "consultasacAuto.frx":564D
      MousePointer    =   99  'Custom
      Picture         =   "consultasacAuto.frx":5957
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7320
      Width           =   480
   End
End
Attribute VB_Name = "PrgConsultaSacauto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCargaSacII As Recordset
Dim spCargaSacII As String
Dim rstCargaSacIII As Recordset
Dim spCargaSacIII As String
Dim rstCargaSacIV As Recordset
Dim spCargaSacIV As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String

Dim XParam As String
Dim ZZLugar As Integer

Sub Imprime_Descripcion()
    
    Sql1 = "Select *"
    Sql2 = " FROM TipoSac"
    Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
    spTipoSac = Sql1 + Sql2 + Sql3
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        DesTipo.Caption = Trim(rstTipoSac!Descripcion)
        rstTipoSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    






    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable11.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable11.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable12.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable12.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable13.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable13.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable14.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable14.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable15.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable15.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable16.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable16.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    
    
    
    
    
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable21.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable21.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable22.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable22.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable23.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable23.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable24.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable24.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable25.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable25.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable26.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable26.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable31.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable31.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable32.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable32.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable33.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable33.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable34.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable34.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable35.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable35.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable36.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable36.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    
    
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    
    
    
    
    
    

End Sub

Sub Imprime_Datos()

    Rem On Error GoTo WError

    ZTipo = Tipo.Text
    ZAno = Ano.Text
    ZNumero = Numero.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Tipo.Text = ZTipo
    Ano.Text = ZAno
    Numero.Text = ZNumero
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
    
        Centro.Text = rstCargaSac!Centro
        Fecha.Text = rstCargaSac!Fecha
        Origen.ListIndex = rstCargaSac!Origen
        Estado.ListIndex = rstCargaSac!Estado
        ResponsableEmisor.Text = rstCargaSac!ResponsableEmisor
        ResponsableDestino.Text = rstCargaSac!ResponsableDestino
        Referencia.Text = Trim(rstCargaSac!Referencia)
        Titulo.Text = Trim(rstCargaSac!Titulo)
        IngresoNoCon.Text = IIf(IsNull(rstCargaSac!IngresoNoCon), "", rstCargaSac!IngresoNoCon)
        IngresoCausa.Text = IIf(IsNull(rstCargaSac!IngresoCausa), "", rstCargaSac!IngresoCausa)
        
        rstCargaSac.Close
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacII"
    ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacII = ZSql
    Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacII.RecordCount > 0 Then
    
        Accion11.Text = Trim(rstCargaSacII!Accion11)
        Accion12.Text = Trim(rstCargaSacII!Accion12)
        Accion21.Text = Trim(rstCargaSacII!Accion21)
        Accion22.Text = Trim(rstCargaSacII!Accion22)
        Accion31.Text = Trim(rstCargaSacII!Accion31)
        Accion32.Text = Trim(rstCargaSacII!Accion32)
        Accion41.Text = Trim(rstCargaSacII!Accion41)
        Accion42.Text = Trim(rstCargaSacII!Accion42)
        Accion51.Text = Trim(rstCargaSacII!Accion51)
        Accion52.Text = Trim(rstCargaSacII!Accion52)
        Accion61.Text = Trim(rstCargaSacII!Accion61)
        Accion62.Text = Trim(rstCargaSacII!Accion62)
        
        Responsable1.Text = rstCargaSacII!Responsable1
        Responsable2.Text = rstCargaSacII!Responsable2
        Responsable3.Text = rstCargaSacII!Responsable3
        Responsable4.Text = rstCargaSacII!Responsable4
        Responsable5.Text = rstCargaSacII!Responsable5
        Responsable6.Text = rstCargaSacII!Responsable6
        
        Plazo1.Text = rstCargaSacII!Plazo1
        Plazo2.Text = rstCargaSacII!Plazo2
        Plazo3.Text = rstCargaSacII!Plazo3
        Plazo4.Text = rstCargaSacII!Plazo4
        Plazo5.Text = rstCargaSacII!Plazo5
        Plazo6.Text = rstCargaSacII!Plazo6
        
        rstCargaSacII.Close
    End If
    
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIII"
    ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIII = ZSql
    Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIII.RecordCount > 0 Then
    
        Responsable11.Text = rstCargaSacIII!Responsable1
        Responsable12.Text = rstCargaSacIII!Responsable2
        Responsable13.Text = rstCargaSacIII!Responsable3
        Responsable14.Text = rstCargaSacIII!Responsable4
        Responsable15.Text = rstCargaSacIII!Responsable5
        Responsable16.Text = rstCargaSacIII!Responsable6
        
        Fecha1.Text = rstCargaSacIII!Fecha1
        Fecha2.Text = rstCargaSacIII!Fecha2
        Fecha3.Text = rstCargaSacIII!Fecha3
        Fecha4.Text = rstCargaSacIII!Fecha4
        Fecha5.Text = rstCargaSacIII!Fecha5
        Fecha6.Text = rstCargaSacIII!Fecha6
        
        Comentario11.Text = Trim(rstCargaSacIII!Comentario11)
        Comentario12.Text = Trim(rstCargaSacIII!Comentario12)
        Comentario21.Text = Trim(rstCargaSacIII!Comentario21)
        Comentario22.Text = Trim(rstCargaSacIII!Comentario22)
        Comentario31.Text = Trim(rstCargaSacIII!Comentario31)
        Comentario32.Text = Trim(rstCargaSacIII!Comentario32)
        Comentario41.Text = Trim(rstCargaSacIII!Comentario41)
        Comentario42.Text = Trim(rstCargaSacIII!Comentario42)
        Comentario51.Text = Trim(rstCargaSacIII!Comentario51)
        Comentario52.Text = Trim(rstCargaSacIII!Comentario52)
        Comentario61.Text = Trim(rstCargaSacIII!Comentario61)
        Comentario62.Text = Trim(rstCargaSacIII!Comentario62)
        
        Estado11.ListIndex = rstCargaSacIII!Estado1
        Estado12.ListIndex = rstCargaSacIII!Estado2
        Estado13.ListIndex = rstCargaSacIII!Estado3
        Estado14.ListIndex = rstCargaSacIII!Estado4
        Estado15.ListIndex = rstCargaSacIII!Estado5
        Estado16.ListIndex = rstCargaSacIII!Estado6
        
        rstCargaSacIII.Close
    End If
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIV"
    ZSql = ZSql + " Where CargaSacIV.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIV = ZSql
    Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIV.RecordCount > 0 Then
    
        Responsable21.Text = rstCargaSacIV!Responsable1
        Responsable22.Text = rstCargaSacIV!Responsable2
        Responsable23.Text = rstCargaSacIV!Responsable3
        Responsable24.Text = rstCargaSacIV!Responsable4
        Responsable25.Text = rstCargaSacIV!Responsable5
        Responsable26.Text = rstCargaSacIV!Responsable6
        
        Responsable31.Text = IIf(IsNull(rstCargaSacIV!Responsable11), "", rstCargaSacIV!Responsable11)
        Responsable32.Text = IIf(IsNull(rstCargaSacIV!Responsable12), "", rstCargaSacIV!Responsable12)
        Responsable33.Text = IIf(IsNull(rstCargaSacIV!Responsable13), "", rstCargaSacIV!Responsable13)
        Responsable34.Text = IIf(IsNull(rstCargaSacIV!Responsable14), "", rstCargaSacIV!Responsable14)
        Responsable35.Text = IIf(IsNull(rstCargaSacIV!Responsable15), "", rstCargaSacIV!Responsable15)
        Responsable36.Text = IIf(IsNull(rstCargaSacIV!Responsable16), "", rstCargaSacIV!Responsable16)
        
        Fecha21.Text = rstCargaSacIV!Fecha1
        Fecha22.Text = rstCargaSacIV!Fecha2
        Fecha23.Text = rstCargaSacIV!Fecha3
        Fecha24.Text = rstCargaSacIV!Fecha4
        Fecha25.Text = rstCargaSacIV!Fecha5
        Fecha26.Text = rstCargaSacIV!Fecha6
        
        Fecha31.Text = IIf(IsNull(rstCargaSacIV!Fecha11), "  /  /    ", rstCargaSacIV!Fecha11)
        Fecha32.Text = IIf(IsNull(rstCargaSacIV!Fecha12), "  /  /    ", rstCargaSacIV!Fecha12)
        Fecha33.Text = IIf(IsNull(rstCargaSacIV!Fecha13), "  /  /    ", rstCargaSacIV!Fecha13)
        Fecha34.Text = IIf(IsNull(rstCargaSacIV!Fecha14), "  /  /    ", rstCargaSacIV!Fecha14)
        Fecha35.Text = IIf(IsNull(rstCargaSacIV!Fecha15), "  /  /    ", rstCargaSacIV!Fecha15)
        Fecha36.Text = IIf(IsNull(rstCargaSacIV!Fecha16), "  /  /    ", rstCargaSacIV!Fecha16)
        
        Comentario211.Text = Trim(rstCargaSacIV!Comentario11)
        Comentario212.Text = Trim(rstCargaSacIV!Comentario12)
        Comentario221.Text = Trim(rstCargaSacIV!Comentario21)
        Comentario222.Text = Trim(rstCargaSacIV!Comentario22)
        Comentario231.Text = Trim(rstCargaSacIV!Comentario31)
        Comentario232.Text = Trim(rstCargaSacIV!Comentario32)
        Comentario241.Text = Trim(rstCargaSacIV!Comentario41)
        Comentario242.Text = Trim(rstCargaSacIV!Comentario42)
        Comentario251.Text = Trim(rstCargaSacIV!Comentario51)
        Comentario252.Text = Trim(rstCargaSacIV!Comentario52)
        Comentario261.Text = Trim(rstCargaSacIV!Comentario61)
        Comentario262.Text = Trim(rstCargaSacIV!Comentario62)
        
        Estado1.ListIndex = rstCargaSacIV!Estado1
        Estado2.ListIndex = rstCargaSacIV!Estado2
        Estado3.ListIndex = rstCargaSacIV!Estado3
        Estado4.ListIndex = rstCargaSacIV!Estado4
        Estado5.ListIndex = rstCargaSacIV!Estado5
        Estado6.ListIndex = rstCargaSacIV!Estado6
        
        Estado31.ListIndex = IIf(IsNull(rstCargaSacIV!Estado11), "0", rstCargaSacIV!Estado11)
        Estado32.ListIndex = IIf(IsNull(rstCargaSacIV!Estado12), "0", rstCargaSacIV!Estado12)
        Estado33.ListIndex = IIf(IsNull(rstCargaSacIV!Estado13), "0", rstCargaSacIV!Estado13)
        Estado34.ListIndex = IIf(IsNull(rstCargaSacIV!Estado14), "0", rstCargaSacIV!Estado14)
        Estado35.ListIndex = IIf(IsNull(rstCargaSacIV!Estado15), "0", rstCargaSacIV!Estado15)
        Estado36.ListIndex = IIf(IsNull(rstCargaSacIV!Estado16), "0", rstCargaSacIV!Estado16)
        
        rstCargaSacIV.Close
    End If
    
    ZResponsableEmisor = Val(ResponsableEmisor.Text)
    ZResponsableDestino = Val(ResponsableDestino.Text)
    
    ZResponsable1 = Val(Responsable1.Text)
    ZResponsable2 = Val(Responsable2.Text)
    ZResponsable3 = Val(Responsable3.Text)
    ZResponsable4 = Val(Responsable4.Text)
    ZResponsable5 = Val(Responsable5.Text)
    ZResponsable6 = Val(Responsable6.Text)
    
    ZResponsable11 = Val(Responsable11.Text)
    ZResponsable12 = Val(Responsable12.Text)
    ZResponsable13 = Val(Responsable13.Text)
    ZResponsable14 = Val(Responsable14.Text)
    ZResponsable15 = Val(Responsable15.Text)
    ZResponsable16 = Val(Responsable16.Text)
    
    ZResponsable21 = Val(Responsable21.Text)
    ZResponsable22 = Val(Responsable22.Text)
    ZResponsable23 = Val(Responsable23.Text)
    ZResponsable24 = Val(Responsable24.Text)
    ZResponsable25 = Val(Responsable25.Text)
    ZResponsable26 = Val(Responsable26.Text)
    
    ZResponsable31 = Val(Responsable31.Text)
    ZResponsable32 = Val(Responsable32.Text)
    ZResponsable33 = Val(Responsable33.Text)
    ZResponsable34 = Val(Responsable34.Text)
    ZResponsable35 = Val(Responsable35.Text)
    ZResponsable36 = Val(Responsable36.Text)
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacAdicional"
    ZSql = ZSql + " Where CargaSacAdicional.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Numero = " + "'" + Numero.Text + "'"
    spCargaSacAdicional = ZSql
    Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacAdicional.RecordCount > 0 Then
        Comentario.Text = IIf(IsNull(rstCargaSacAdicional!Dato1), "", rstCargaSacAdicional!Dato1)
        rstCargaSacAdicional.Close
    End If
    
    
    
    ZZPasa = "N"
    If ZZCodigoResponsable = 1 Then
        ZZPasa = "S"
    End If
    If WOperador = ZResponsableEmisor Or WOperador = ZResponsableDestino Then
        ZZPasa = "S"
    End If
    If WOperador = ZResponsable1 Or WOperador = ZResponsable2 Or WOperador = ZResponsable3 Or WOperador = ZResponsable4 Or WOperador = ZResponsable5 Or WOperador = ZResponsable6 Then
        ZZPasa = "S"
    End If
    If WOperador = ZResponsable21 Or WOperador = ZResponsable22 Or WOperador = ZResponsable23 Or WOperador = ZResponsable24 Or WOperador = ZResponsable25 Or WOperador = ZResponsable26 Then
        ZZPasa = "S"
    End If
    If WOperador = ZResponsable31 Or WOperador = ZResponsable32 Or WOperador = ZResponsable33 Or WOperador = ZResponsable34 Or WOperador = ZResponsable35 Or WOperador = ZResponsable36 Then
        ZZPasa = "S"
    End If
    
    
    If ZZPasa = "N" Then
        CmdAdd.Enabled = False
    End If
    
    
    
    
    
    
    
    
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub


Private Sub AnteriorNro_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero < " + "'" + Numero.Text + "'"
    ZSql = ZSql + " Order by Numero"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        With rstCargaSac
            .MoveLast
            Numero.Text = rstCargaSac!Numero
            rstCargaSac.Close
            Call Numero_KeyPress(13)
        End With
    End If

End Sub

Private Sub AnteriorTipo_Click()

    ZZTipo = Tipo.Text
    ZZNumero = Numero.Text
    Tipo.Text = Str$(Val(Tipo.Text) - 1)
    Numero.Text = "0"

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero > " + "'" + Numero.Text + "'"
    ZSql = ZSql + " Order by Numero"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        With rstCargaSac
            .MoveFirst
            Numero.Text = rstCargaSac!Numero
            rstCargaSac.Close
            Call Numero_KeyPress(13)
        End With
            Else
        Tipo.Text = ZZTipo
        Numero.Text = ZZNumero
        Call Numero_KeyPress(13)
    End If


End Sub

Private Sub cmdAdd_Click()

    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
    Auxi3 = Tipo.Text
    Auxi1 = Ano.Text
    Auxi2 = Numero.Text
    Call Ceros(Auxi3, 4)
    Call Ceros(Auxi1, 4)
    Call Ceros(Auxi2, 6)
    WClave = Auxi3 + Auxi1 + Auxi2
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
    
        rstCargaSac.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaSac SET "
        ZSql = ZSql + " Centro = " + "'" + Centro.Text + "',"
        ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
        ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
        ZSql = ZSql + " Origen = " + "'" + Str$(Origen.ListIndex) + "',"
        ZSql = ZSql + " Estado = " + "'" + Str$(Estado.ListIndex) + "',"
        ZSql = ZSql + " ResponsableEmisor = " + "'" + ResponsableEmisor.Text + "',"
        ZSql = ZSql + " ResponsableDestino = " + "'" + ResponsableDestino.Text + "',"
        ZSql = ZSql + " Referencia = " + "'" + Referencia.Text + "',"
        ZSql = ZSql + " Titulo = " + "'" + Titulo.Text + "',"
        ZSql = ZSql + " IngresoNoCon = " + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + " IngresoCausa = " + "'" + IngresoCausa.Text + "'"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Centro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "ResponsableEmisor ,"
        ZSql = ZSql + "ResponsableDestino ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + Centro.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + WOrdFecha + "',"
        ZSql = ZSql + "'" + Str$(Origen.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado.ListIndex) + "',"
        ZSql = ZSql + "'" + ResponsableEmisor.Text + "',"
        ZSql = ZSql + "'" + ResponsableDestino.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "')"
        
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    
    
    
    Auxi3 = Tipo.Text
    Auxi1 = Ano.Text
    Auxi2 = Numero.Text
    Call Ceros(Auxi3, 4)
    Call Ceros(Auxi1, 4)
    Call Ceros(Auxi2, 6)
    WClave = Auxi3 + Auxi1 + Auxi2

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacII"
    ZSql = ZSql + " Where CargaSacII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacII = ZSql
    Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacII.RecordCount > 0 Then
    
        rstCargaSacII.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaSacII SET "
        ZSql = ZSql + " Accion11 = " + "'" + Accion11.Text + "',"
        ZSql = ZSql + " Accion12 = " + "'" + Accion12.Text + "',"
        ZSql = ZSql + " Accion21 = " + "'" + Accion21.Text + "',"
        ZSql = ZSql + " Accion22 = " + "'" + Accion22.Text + "',"
        ZSql = ZSql + " Accion31 = " + "'" + Accion31.Text + "',"
        ZSql = ZSql + " Accion32 = " + "'" + Accion32.Text + "',"
        ZSql = ZSql + " Accion41 = " + "'" + Accion41.Text + "',"
        ZSql = ZSql + " Accion42 = " + "'" + Accion42.Text + "',"
        ZSql = ZSql + " Accion51 = " + "'" + Accion51.Text + "',"
        ZSql = ZSql + " Accion52 = " + "'" + Accion52.Text + "',"
        ZSql = ZSql + " Accion61 = " + "'" + Accion61.Text + "',"
        ZSql = ZSql + " Accion62 = " + "'" + Accion62.Text + "',"
        ZSql = ZSql + " Responsable1 = " + "'" + Responsable1.Text + "',"
        ZSql = ZSql + " Responsable2 = " + "'" + Responsable2.Text + "',"
        ZSql = ZSql + " Responsable3 = " + "'" + Responsable3.Text + "',"
        ZSql = ZSql + " Responsable4 = " + "'" + Responsable4.Text + "',"
        ZSql = ZSql + " Responsable5 = " + "'" + Responsable5.Text + "',"
        ZSql = ZSql + " Responsable6 = " + "'" + Responsable6.Text + "',"
        ZSql = ZSql + " Plazo1 = " + "'" + Plazo1.Text + "',"
        ZSql = ZSql + " Plazo2 = " + "'" + Plazo2.Text + "',"
        ZSql = ZSql + " Plazo3 = " + "'" + Plazo3.Text + "',"
        ZSql = ZSql + " Plazo4 = " + "'" + Plazo4.Text + "',"
        ZSql = ZSql + " Plazo5 = " + "'" + Plazo5.Text + "',"
        ZSql = ZSql + " Plazo6 = " + "'" + Plazo6.Text + "'"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacII = ZSql
        Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaSacII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Accion11 ,"
        ZSql = ZSql + "Accion12 ,"
        ZSql = ZSql + "Accion21 ,"
        ZSql = ZSql + "Accion22 ,"
        ZSql = ZSql + "Accion31 ,"
        ZSql = ZSql + "Accion32 ,"
        ZSql = ZSql + "Accion41 ,"
        ZSql = ZSql + "Accion42 ,"
        ZSql = ZSql + "Accion51 ,"
        ZSql = ZSql + "Accion52 ,"
        ZSql = ZSql + "Accion61 ,"
        ZSql = ZSql + "Accion62 ,"
        ZSql = ZSql + "Responsable1 ,"
        ZSql = ZSql + "Responsable2 ,"
        ZSql = ZSql + "Responsable3 ,"
        ZSql = ZSql + "Responsable4 ,"
        ZSql = ZSql + "Responsable5 ,"
        ZSql = ZSql + "Responsable6 ,"
        ZSql = ZSql + "Plazo1 ,"
        ZSql = ZSql + "Plazo2 ,"
        ZSql = ZSql + "Plazo3 ,"
        ZSql = ZSql + "Plazo4 ,"
        ZSql = ZSql + "Plazo5 ,"
        ZSql = ZSql + "Plazo6 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + Accion11.Text + "',"
        ZSql = ZSql + "'" + Accion12.Text + "',"
        ZSql = ZSql + "'" + Accion21.Text + "',"
        ZSql = ZSql + "'" + Accion22.Text + "',"
        ZSql = ZSql + "'" + Accion31.Text + "',"
        ZSql = ZSql + "'" + Accion32.Text + "',"
        ZSql = ZSql + "'" + Accion41.Text + "',"
        ZSql = ZSql + "'" + Accion42.Text + "',"
        ZSql = ZSql + "'" + Accion51.Text + "',"
        ZSql = ZSql + "'" + Accion52.Text + "',"
        ZSql = ZSql + "'" + Accion61.Text + "',"
        ZSql = ZSql + "'" + Accion62.Text + "',"
        ZSql = ZSql + "'" + Responsable1.Text + "',"
        ZSql = ZSql + "'" + Responsable2.Text + "',"
        ZSql = ZSql + "'" + Responsable3.Text + "',"
        ZSql = ZSql + "'" + Responsable4.Text + "',"
        ZSql = ZSql + "'" + Responsable5.Text + "',"
        ZSql = ZSql + "'" + Responsable6.Text + "',"
        ZSql = ZSql + "'" + Plazo1.Text + "',"
        ZSql = ZSql + "'" + Plazo2.Text + "',"
        ZSql = ZSql + "'" + Plazo3.Text + "',"
        ZSql = ZSql + "'" + Plazo4.Text + "',"
        ZSql = ZSql + "'" + Plazo5.Text + "',"
        ZSql = ZSql + "'" + Plazo6.Text + "')"
        
        spCargaSacII = ZSql
        Set rstCargaSacII = db.OpenRecordset(spCargaSacII, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
        
        
        












    Auxi3 = Tipo.Text
    Auxi1 = Ano.Text
    Auxi2 = Numero.Text
    Call Ceros(Auxi3, 4)
    Call Ceros(Auxi1, 4)
    Call Ceros(Auxi2, 6)
    WClave = Auxi3 + Auxi1 + Auxi2

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIII"
    ZSql = ZSql + " Where CargaSacIII.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIII.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIII = ZSql
    Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIII.RecordCount > 0 Then
    
        rstCargaSacIII.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaSacIII SET "
        ZSql = ZSql + " Responsable1 = " + "'" + Responsable11.Text + "',"
        ZSql = ZSql + " Responsable2 = " + "'" + Responsable12.Text + "',"
        ZSql = ZSql + " Responsable3 = " + "'" + Responsable13.Text + "',"
        ZSql = ZSql + " Responsable4 = " + "'" + Responsable14.Text + "',"
        ZSql = ZSql + " Responsable5 = " + "'" + Responsable15.Text + "',"
        ZSql = ZSql + " Responsable6 = " + "'" + Responsable16.Text + "',"
        ZSql = ZSql + " Fecha1 = " + "'" + Fecha1.Text + "',"
        ZSql = ZSql + " Fecha2 = " + "'" + Fecha2.Text + "',"
        ZSql = ZSql + " Fecha3 = " + "'" + Fecha3.Text + "',"
        ZSql = ZSql + " Fecha4 = " + "'" + Fecha4.Text + "',"
        ZSql = ZSql + " Fecha5 = " + "'" + Fecha5.Text + "',"
        ZSql = ZSql + " Fecha6 = " + "'" + Fecha6.Text + "',"
        ZSql = ZSql + " Comentario11 = " + "'" + Comentario11.Text + "',"
        ZSql = ZSql + " Comentario12 = " + "'" + Comentario12.Text + "',"
        ZSql = ZSql + " Comentario21 = " + "'" + Comentario21.Text + "',"
        ZSql = ZSql + " Comentario22 = " + "'" + Comentario22.Text + "',"
        ZSql = ZSql + " Comentario31 = " + "'" + Comentario31.Text + "',"
        ZSql = ZSql + " Comentario32 = " + "'" + Comentario32.Text + "',"
        ZSql = ZSql + " Comentario41 = " + "'" + Comentario41.Text + "',"
        ZSql = ZSql + " Comentario42 = " + "'" + Comentario42.Text + "',"
        ZSql = ZSql + " Comentario51 = " + "'" + Comentario51.Text + "',"
        ZSql = ZSql + " Comentario52 = " + "'" + Comentario52.Text + "',"
        ZSql = ZSql + " Comentario61 = " + "'" + Comentario61.Text + "',"
        ZSql = ZSql + " Comentario62 = " + "'" + Comentario62.Text + "',"
        ZSql = ZSql + " Estado1 = " + "'" + Str$(Estado11.ListIndex) + "',"
        ZSql = ZSql + " Estado2 = " + "'" + Str$(Estado12.ListIndex) + "',"
        ZSql = ZSql + " Estado3 = " + "'" + Str$(Estado13.ListIndex) + "',"
        ZSql = ZSql + " Estado4 = " + "'" + Str$(Estado14.ListIndex) + "',"
        ZSql = ZSql + " Estado5 = " + "'" + Str$(Estado15.ListIndex) + "',"
        ZSql = ZSql + " Estado6 = " + "'" + Str$(Estado16.ListIndex) + "'"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacIII = ZSql
        Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaSacIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Responsable1 ,"
        ZSql = ZSql + "Responsable2 ,"
        ZSql = ZSql + "Responsable3 ,"
        ZSql = ZSql + "Responsable4 ,"
        ZSql = ZSql + "Responsable5 ,"
        ZSql = ZSql + "Responsable6 ,"
        ZSql = ZSql + "Fecha1 ,"
        ZSql = ZSql + "Fecha2 ,"
        ZSql = ZSql + "Fecha3 ,"
        ZSql = ZSql + "Fecha4 ,"
        ZSql = ZSql + "Fecha5 ,"
        ZSql = ZSql + "Fecha6 ,"
        ZSql = ZSql + "Comentario11 ,"
        ZSql = ZSql + "Comentario12 ,"
        ZSql = ZSql + "Comentario21 ,"
        ZSql = ZSql + "Comentario22 ,"
        ZSql = ZSql + "Comentario31 ,"
        ZSql = ZSql + "Comentario32 ,"
        ZSql = ZSql + "Comentario41 ,"
        ZSql = ZSql + "Comentario42 ,"
        ZSql = ZSql + "Comentario51 ,"
        ZSql = ZSql + "Comentario52 ,"
        ZSql = ZSql + "Comentario61 ,"
        ZSql = ZSql + "Comentario62 ,"
        ZSql = ZSql + "Estado1 ,"
        ZSql = ZSql + "Estado2 ,"
        ZSql = ZSql + "Estado3 ,"
        ZSql = ZSql + "Estado4 ,"
        ZSql = ZSql + "Estado5 ,"
        ZSql = ZSql + "Estado6 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + Responsable11.Text + "',"
        ZSql = ZSql + "'" + Responsable12.Text + "',"
        ZSql = ZSql + "'" + Responsable13.Text + "',"
        ZSql = ZSql + "'" + Responsable14.Text + "',"
        ZSql = ZSql + "'" + Responsable15.Text + "',"
        ZSql = ZSql + "'" + Responsable16.Text + "',"
        ZSql = ZSql + "'" + Fecha1.Text + "',"
        ZSql = ZSql + "'" + Fecha2.Text + "',"
        ZSql = ZSql + "'" + Fecha3.Text + "',"
        ZSql = ZSql + "'" + Fecha4.Text + "',"
        ZSql = ZSql + "'" + Fecha5.Text + "',"
        ZSql = ZSql + "'" + Fecha6.Text + "',"
        ZSql = ZSql + "'" + Comentario11.Text + "',"
        ZSql = ZSql + "'" + Comentario12.Text + "',"
        ZSql = ZSql + "'" + Comentario21.Text + "',"
        ZSql = ZSql + "'" + Comentario22.Text + "',"
        ZSql = ZSql + "'" + Comentario31.Text + "',"
        ZSql = ZSql + "'" + Comentario32.Text + "',"
        ZSql = ZSql + "'" + Comentario41.Text + "',"
        ZSql = ZSql + "'" + Comentario42.Text + "',"
        ZSql = ZSql + "'" + Comentario51.Text + "',"
        ZSql = ZSql + "'" + Comentario52.Text + "',"
        ZSql = ZSql + "'" + Comentario61.Text + "',"
        ZSql = ZSql + "'" + Comentario62.Text + "',"
        ZSql = ZSql + "'" + Str$(Estado11.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado12.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado13.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado14.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado15.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado16.ListIndex) + "')"
        
        spCargaSacIII = ZSql
        Set rstCargaSacIII = db.OpenRecordset(spCargaSacIII, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    Auxi3 = Tipo.Text
    Auxi1 = Ano.Text
    Auxi2 = Numero.Text
    Call Ceros(Auxi3, 4)
    Call Ceros(Auxi1, 4)
    Call Ceros(Auxi2, 6)
    WClave = Auxi3 + Auxi1 + Auxi2

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacIV"
    ZSql = ZSql + " Where CargaSacIV.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacIV.Numero = " + "'" + Numero.Text + "'"
    spCargaSacIV = ZSql
    Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacIV.RecordCount > 0 Then
    
        rstCargaSacIV.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaSacIV SET "
        ZSql = ZSql + " Responsable1 = " + "'" + Responsable21.Text + "',"
        ZSql = ZSql + " Responsable2 = " + "'" + Responsable22.Text + "',"
        ZSql = ZSql + " Responsable3 = " + "'" + Responsable23.Text + "',"
        ZSql = ZSql + " Responsable4 = " + "'" + Responsable24.Text + "',"
        ZSql = ZSql + " Responsable5 = " + "'" + Responsable25.Text + "',"
        ZSql = ZSql + " Responsable6 = " + "'" + Responsable26.Text + "',"
        ZSql = ZSql + " Responsable11 = " + "'" + Responsable31.Text + "',"
        ZSql = ZSql + " Responsable12 = " + "'" + Responsable32.Text + "',"
        ZSql = ZSql + " Responsable13 = " + "'" + Responsable33.Text + "',"
        ZSql = ZSql + " Responsable14 = " + "'" + Responsable34.Text + "',"
        ZSql = ZSql + " Responsable15 = " + "'" + Responsable35.Text + "',"
        ZSql = ZSql + " Responsable16 = " + "'" + Responsable36.Text + "',"
        ZSql = ZSql + " Fecha1 = " + "'" + Fecha21.Text + "',"
        ZSql = ZSql + " Fecha2 = " + "'" + Fecha22.Text + "',"
        ZSql = ZSql + " Fecha3 = " + "'" + Fecha23.Text + "',"
        ZSql = ZSql + " Fecha4 = " + "'" + Fecha24.Text + "',"
        ZSql = ZSql + " Fecha5 = " + "'" + Fecha25.Text + "',"
        ZSql = ZSql + " Fecha6 = " + "'" + Fecha26.Text + "',"
        ZSql = ZSql + " Fecha11 = " + "'" + Fecha31.Text + "',"
        ZSql = ZSql + " Fecha12 = " + "'" + Fecha32.Text + "',"
        ZSql = ZSql + " Fecha13 = " + "'" + Fecha33.Text + "',"
        ZSql = ZSql + " Fecha14 = " + "'" + Fecha34.Text + "',"
        ZSql = ZSql + " Fecha15 = " + "'" + Fecha35.Text + "',"
        ZSql = ZSql + " Fecha16 = " + "'" + Fecha36.Text + "',"
        ZSql = ZSql + " Comentario11 = " + "'" + Comentario211.Text + "',"
        ZSql = ZSql + " Comentario12 = " + "'" + Comentario212.Text + "',"
        ZSql = ZSql + " Comentario21 = " + "'" + Comentario221.Text + "',"
        ZSql = ZSql + " Comentario22 = " + "'" + Comentario222.Text + "',"
        ZSql = ZSql + " Comentario31 = " + "'" + Comentario231.Text + "',"
        ZSql = ZSql + " Comentario32 = " + "'" + Comentario232.Text + "',"
        ZSql = ZSql + " Comentario41 = " + "'" + Comentario241.Text + "',"
        ZSql = ZSql + " Comentario42 = " + "'" + Comentario242.Text + "',"
        ZSql = ZSql + " Comentario51 = " + "'" + Comentario251.Text + "',"
        ZSql = ZSql + " Comentario52 = " + "'" + Comentario252.Text + "',"
        ZSql = ZSql + " Comentario61 = " + "'" + Comentario261.Text + "',"
        ZSql = ZSql + " Comentario62 = " + "'" + Comentario262.Text + "',"
        ZSql = ZSql + " Estado1 = " + "'" + Str$(Estado1.ListIndex) + "',"
        ZSql = ZSql + " Estado2 = " + "'" + Str$(Estado2.ListIndex) + "',"
        ZSql = ZSql + " Estado3 = " + "'" + Str$(Estado3.ListIndex) + "',"
        ZSql = ZSql + " Estado4 = " + "'" + Str$(Estado4.ListIndex) + "',"
        ZSql = ZSql + " Estado5 = " + "'" + Str$(Estado5.ListIndex) + "',"
        ZSql = ZSql + " Estado6 = " + "'" + Str$(Estado6.ListIndex) + "',"
        ZSql = ZSql + " Estado11 = " + "'" + Str$(Estado31.ListIndex) + "',"
        ZSql = ZSql + " Estado12 = " + "'" + Str$(Estado32.ListIndex) + "',"
        ZSql = ZSql + " Estado13 = " + "'" + Str$(Estado33.ListIndex) + "',"
        ZSql = ZSql + " Estado14 = " + "'" + Str$(Estado34.ListIndex) + "',"
        ZSql = ZSql + " Estado15 = " + "'" + Str$(Estado35.ListIndex) + "',"
        ZSql = ZSql + " Estado16 = " + "'" + Str$(Estado36.ListIndex) + "'"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacIV = ZSql
        Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaSacIV ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Responsable1 ,"
        ZSql = ZSql + "Responsable2 ,"
        ZSql = ZSql + "Responsable3 ,"
        ZSql = ZSql + "Responsable4 ,"
        ZSql = ZSql + "Responsable5 ,"
        ZSql = ZSql + "Responsable6 ,"
        ZSql = ZSql + "Responsable11 ,"
        ZSql = ZSql + "Responsable12 ,"
        ZSql = ZSql + "Responsable13 ,"
        ZSql = ZSql + "Responsable14 ,"
        ZSql = ZSql + "Responsable15 ,"
        ZSql = ZSql + "Responsable16 ,"
        ZSql = ZSql + "Fecha1 ,"
        ZSql = ZSql + "Fecha2 ,"
        ZSql = ZSql + "Fecha3 ,"
        ZSql = ZSql + "Fecha4 ,"
        ZSql = ZSql + "Fecha5 ,"
        ZSql = ZSql + "Fecha6 ,"
        ZSql = ZSql + "Fecha11 ,"
        ZSql = ZSql + "Fecha12 ,"
        ZSql = ZSql + "Fecha13 ,"
        ZSql = ZSql + "Fecha14 ,"
        ZSql = ZSql + "Fecha15 ,"
        ZSql = ZSql + "Fecha16 ,"
        ZSql = ZSql + "Comentario11 ,"
        ZSql = ZSql + "Comentario12 ,"
        ZSql = ZSql + "Comentario21 ,"
        ZSql = ZSql + "Comentario22 ,"
        ZSql = ZSql + "Comentario31 ,"
        ZSql = ZSql + "Comentario32 ,"
        ZSql = ZSql + "Comentario41 ,"
        ZSql = ZSql + "Comentario42 ,"
        ZSql = ZSql + "Comentario51 ,"
        ZSql = ZSql + "Comentario52 ,"
        ZSql = ZSql + "Comentario61 ,"
        ZSql = ZSql + "Comentario62 ,"
        ZSql = ZSql + "Estado1 ,"
        ZSql = ZSql + "Estado2 ,"
        ZSql = ZSql + "Estado3 ,"
        ZSql = ZSql + "Estado4 ,"
        ZSql = ZSql + "Estado5 ,"
        ZSql = ZSql + "Estado6 ,"
        ZSql = ZSql + "Estado11 ,"
        ZSql = ZSql + "Estado12 ,"
        ZSql = ZSql + "Estado13 ,"
        ZSql = ZSql + "Estado14 ,"
        ZSql = ZSql + "Estado15 ,"
        ZSql = ZSql + "Estado16 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + Responsable21.Text + "',"
        ZSql = ZSql + "'" + Responsable22.Text + "',"
        ZSql = ZSql + "'" + Responsable23.Text + "',"
        ZSql = ZSql + "'" + Responsable24.Text + "',"
        ZSql = ZSql + "'" + Responsable25.Text + "',"
        ZSql = ZSql + "'" + Responsable26.Text + "',"
        ZSql = ZSql + "'" + Responsable31.Text + "',"
        ZSql = ZSql + "'" + Responsable32.Text + "',"
        ZSql = ZSql + "'" + Responsable33.Text + "',"
        ZSql = ZSql + "'" + Responsable34.Text + "',"
        ZSql = ZSql + "'" + Responsable35.Text + "',"
        ZSql = ZSql + "'" + Responsable36.Text + "',"
        ZSql = ZSql + "'" + Fecha21.Text + "',"
        ZSql = ZSql + "'" + Fecha22.Text + "',"
        ZSql = ZSql + "'" + Fecha23.Text + "',"
        ZSql = ZSql + "'" + Fecha24.Text + "',"
        ZSql = ZSql + "'" + Fecha25.Text + "',"
        ZSql = ZSql + "'" + Fecha26.Text + "',"
        ZSql = ZSql + "'" + Fecha31.Text + "',"
        ZSql = ZSql + "'" + Fecha32.Text + "',"
        ZSql = ZSql + "'" + Fecha33.Text + "',"
        ZSql = ZSql + "'" + Fecha34.Text + "',"
        ZSql = ZSql + "'" + Fecha35.Text + "',"
        ZSql = ZSql + "'" + Fecha36.Text + "',"
        ZSql = ZSql + "'" + Comentario211.Text + "',"
        ZSql = ZSql + "'" + Comentario212.Text + "',"
        ZSql = ZSql + "'" + Comentario221.Text + "',"
        ZSql = ZSql + "'" + Comentario222.Text + "',"
        ZSql = ZSql + "'" + Comentario231.Text + "',"
        ZSql = ZSql + "'" + Comentario232.Text + "',"
        ZSql = ZSql + "'" + Comentario241.Text + "',"
        ZSql = ZSql + "'" + Comentario242.Text + "',"
        ZSql = ZSql + "'" + Comentario251.Text + "',"
        ZSql = ZSql + "'" + Comentario252.Text + "',"
        ZSql = ZSql + "'" + Comentario261.Text + "',"
        ZSql = ZSql + "'" + Comentario262.Text + "',"
        ZSql = ZSql + "'" + Str$(Estado1.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado2.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado3.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado4.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado5.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado6.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado31.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado32.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado33.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado34.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado35.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado36.ListIndex) + "')"
        
        spCargaSacIV = ZSql
        Set rstCargaSacIV = db.OpenRecordset(spCargaSacIV, dbOpenSnapshot, dbSQLPassThrough)
        
    End If



    ZLargo = Len(Comentario.Text)
    ZDato = Comentario.Text
    For CicloII = 1 To ZLargo
        If Mid$(ZDato, CicloII, 1) = Chr$(39) Then
            ZDato = Left$(ZDato, CicloII - 1) + " " + Mid$(ZDato, CicloII + 1, ZLargo)
        End If
    Next CicloII
    Comentario.Text = ZDato
        
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacAdicional"
    ZSql = ZSql + " Where CargaSacAdicional.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Numero = " + "'" + Numero.Text + "'"
    spCargaSacAdicional = ZSql
    Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacAdicional.RecordCount > 0 Then
    
        rstCargaSacAdicional.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaSacAdicional SET "
        ZSql = ZSql + " Dato1 = " + "'" + Comentario.Text + "'"
        ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSacAdicional = ZSql
        Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaSacAdicional ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Dato1 ,"
        ZSql = ZSql + "Dato2 ,"
        ZSql = ZSql + "Dato3 ,"
        ZSql = ZSql + "Dato4 ,"
        ZSql = ZSql + "Dato5 ,"
        ZSql = ZSql + "Foto1 ,"
        ZSql = ZSql + "Foto2 ,"
        ZSql = ZSql + "Foto3 ,"
        ZSql = ZSql + "Foto4 ,"
        ZSql = ZSql + "Foto5 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + Comentario.Text + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "')"
        
        spCargaSacAdicional = ZSql
        Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    m$ = "Actualizacion realizada con exito"
    A% = MsgBox(m$, 0, "Actualizacion de datos")


End Sub

Private Sub CmdLimpiar_Click()
    
    Tipo.Text = ""
    DesTipo.Caption = ""
    Ano.Text = ""
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    Referencia.Text = ""
    Titulo.Text = ""
    IngresoNoCon.Text = ""
    IngresoCausa.Text = ""
    Comentario.Text = ""
    
    Origen.ListIndex = 0
    Estado.ListIndex = 0
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""

    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Accion111.Text = ""
    Accion112.Text = ""
    Accion121.Text = ""
    Accion122.Text = ""
    Accion131.Text = ""
    Accion132.Text = ""
    Accion141.Text = ""
    Accion142.Text = ""
    Accion151.Text = ""
    Accion152.Text = ""
    Accion151.Text = ""
    Accion162.Text = ""
    
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""

    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    
    Fecha21.Text = "  /  /    "
    Fecha22.Text = "  /  /    "
    Fecha23.Text = "  /  /    "
    Fecha24.Text = "  /  /    "
    Fecha25.Text = "  /  /    "
    Fecha26.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    Estado11.ListIndex = 0
    Estado12.ListIndex = 0
    Estado13.ListIndex = 0
    Estado14.ListIndex = 0
    Estado15.ListIndex = 0
    Estado16.ListIndex = 0
    
    Responsable21.Text = ""
    Responsable22.Text = ""
    Responsable23.Text = ""
    Responsable24.Text = ""
    Responsable25.Text = ""
    Responsable26.Text = ""
    
    DesResponsable21.Caption = ""
    DesResponsable22.Caption = ""
    DesResponsable23.Caption = ""
    DesResponsable24.Caption = ""
    DesResponsable25.Caption = ""
    DesResponsable26.Caption = ""
    
    Responsable31.Text = ""
    Responsable32.Text = ""
    Responsable33.Text = ""
    Responsable34.Text = ""
    Responsable35.Text = ""
    Responsable36.Text = ""
    
    DesResponsable31.Caption = ""
    DesResponsable32.Caption = ""
    DesResponsable33.Caption = ""
    DesResponsable34.Caption = ""
    DesResponsable35.Caption = ""
    DesResponsable36.Caption = ""
    
    Fecha31.Text = "  /  /    "
    Fecha32.Text = "  /  /    "
    Fecha33.Text = "  /  /    "
    Fecha34.Text = "  /  /    "
    Fecha35.Text = "  /  /    "
    Fecha36.Text = "  /  /    "
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Estado31.ListIndex = 0
    Estado32.ListIndex = 0
    Estado33.ListIndex = 0
    Estado34.ListIndex = 0
    Estado35.ListIndex = 0
    Estado36.ListIndex = 0
    
    Comentario211.Text = ""
    Comentario212.Text = ""
    Comentario221.Text = ""
    Comentario222.Text = ""
    Comentario231.Text = ""
    Comentario232.Text = ""
    Comentario241.Text = ""
    Comentario242.Text = ""
    Comentario251.Text = ""
    Comentario252.Text = ""
    Comentario261.Text = ""
    Comentario262.Text = ""
    
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Tablas.Tab = 0
    
End Sub

Private Sub cmdClose_Click()

    PrgConsultaSacauto.Hide
    Unload Me
    PrgIndiceSac.Show
    
End Sub



Sub Form_Load()

    Tipo.Text = Mid$(WPasaNumero, 1, 4)
    Rem DesTipo.Caption = "SAC"
    Ano.Text = Mid$(WPasaNumero, 5, 4)
    Numero.Text = Mid$(WPasaNumero, 9, 10)
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    Referencia.Text = ""
    Titulo.Text = ""
    IngresoNoCon.Text = ""
    IngresoCausa.Text = ""
    Comentario.Text = ""
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""

    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Accion111.Text = ""
    Accion112.Text = ""
    Accion121.Text = ""
    Accion122.Text = ""
    Accion131.Text = ""
    Accion132.Text = ""
    Accion141.Text = ""
    Accion142.Text = ""
    Accion151.Text = ""
    Accion152.Text = ""
    Accion151.Text = ""
    Accion162.Text = ""
    
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""

    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "

    Fecha21.Text = "  /  /    "
    Fecha22.Text = "  /  /    "
    Fecha23.Text = "  /  /    "
    Fecha24.Text = "  /  /    "
    Fecha25.Text = "  /  /    "
    Fecha26.Text = "  /  /    "

    Fecha31.Text = "  /  /    "
    Fecha32.Text = "  /  /    "
    Fecha33.Text = "  /  /    "
    Fecha34.Text = "  /  /    "
    Fecha35.Text = "  /  /    "
    Fecha36.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    Responsable21.Text = ""
    Responsable22.Text = ""
    Responsable23.Text = ""
    Responsable24.Text = ""
    Responsable25.Text = ""
    Responsable26.Text = ""
    
    DesResponsable21.Caption = ""
    DesResponsable22.Caption = ""
    DesResponsable23.Caption = ""
    DesResponsable24.Caption = ""
    DesResponsable25.Caption = ""
    DesResponsable26.Caption = ""
    
    Responsable31.Text = ""
    Responsable32.Text = ""
    Responsable33.Text = ""
    Responsable34.Text = ""
    Responsable35.Text = ""
    Responsable36.Text = ""
    
    DesResponsable31.Caption = ""
    DesResponsable32.Caption = ""
    DesResponsable33.Caption = ""
    DesResponsable34.Caption = ""
    DesResponsable35.Caption = ""
    DesResponsable36.Caption = ""
    
    Comentario211.Text = ""
    Comentario212.Text = ""
    Comentario221.Text = ""
    Comentario222.Text = ""
    Comentario231.Text = ""
    Comentario232.Text = ""
    Comentario241.Text = ""
    Comentario242.Text = ""
    Comentario251.Text = ""
    Comentario252.Text = ""
    Comentario261.Text = ""
    Comentario262.Text = ""
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "INICIADA"
    Estado.AddItem "INVESTIGACION"
    Estado.AddItem "IMPLEMENTACION"
    Estado.AddItem "IMPLEMENTACION A VERIFICAR"
    Estado.AddItem "IMPLEMENTACION VERIFICADA"
    Estado.AddItem "CERRADA"
    Estado.AddItem "ANULADA"
    
    Estado.ListIndex = 0
    
    Origen.Clear
    
    Origen.AddItem ""
    Origen.AddItem "Auditoria"
    Origen.AddItem "Reclamo"
    Origen.AddItem "I. No Conformidad"
    Origen.AddItem "Proceso/Sist"
    Origen.AddItem "Otro"
    
    Origen.ListIndex = 0
    
    
    
    Estado11.Clear
    
    Estado11.AddItem ""
    Estado11.AddItem "Imple."
    Estado11.AddItem "Nula"
    
    Estado11.ListIndex = 0
    
    Estado12.Clear
    
    Estado12.AddItem ""
    Estado12.AddItem "Imple."
    Estado12.AddItem "Nula"
    
    Estado12.ListIndex = 0
    
    Estado13.Clear
    
    Estado13.AddItem ""
    Estado13.AddItem "Imple."
    Estado13.AddItem "Nula"
    
    Estado13.ListIndex = 0
    
    Estado14.Clear
    
    Estado14.AddItem ""
    Estado14.AddItem "Imple."
    Estado14.AddItem "Nula"
    
    Estado14.ListIndex = 0
    
    Estado15.Clear
    
    Estado15.AddItem ""
    Estado15.AddItem "Imple."
    Estado15.AddItem "Nula"
 
    Estado15.ListIndex = 0
    
    Estado16.Clear
    
    Estado16.AddItem ""
    Estado16.AddItem "Imple."
    Estado16.AddItem "Nula"
    
    Estado16.ListIndex = 0
    
    
    
    
    
    Estado1.Clear
    
    Estado1.AddItem "No Imple."
    Estado1.AddItem "Imple."
    Estado1.AddItem "Nula"
    Estado1.AddItem "Cerrada"
    Estado1.AddItem ""
    
    Estado1.ListIndex = 0
    
    Estado2.Clear
    
    Estado2.AddItem "No Imple."
    Estado2.AddItem "Imple."
    Estado2.AddItem "Nula"
    Estado2.AddItem "Cerrada"
    Estado2.AddItem ""
    
    Estado2.ListIndex = 0
    
    Estado3.Clear
    
    Estado3.AddItem "No Imple."
    Estado3.AddItem "Imple."
    Estado3.AddItem "Nula"
    Estado3.AddItem "Cerrada"
    Estado3.AddItem ""
    
    Estado3.ListIndex = 0
    
    Estado4.Clear
    
    Estado4.AddItem "No Imple."
    Estado4.AddItem "Imple."
    Estado4.AddItem "Nula"
    Estado4.AddItem "Cerrada"
    Estado4.AddItem ""
    
    Estado4.ListIndex = 0
    
    Estado5.Clear
    
    Estado5.AddItem "No Imple."
    Estado5.AddItem "Imple."
    Estado5.AddItem "Nula"
    Estado5.AddItem "Cerrada"
    Estado5.AddItem ""
 
    Estado5.ListIndex = 0
    
    Estado6.Clear
    
    Estado6.AddItem "No Imple."
    Estado6.AddItem "Imple."
    Estado6.AddItem "Nula"
    Estado6.AddItem "Cerrada"
    Estado6.AddItem ""
    
    Estado6.ListIndex = 0
    
    
    
    
    
    
    Estado31.Clear
    
    Estado31.AddItem "No Imple."
    Estado31.AddItem "Imple."
    Estado31.AddItem "Nula"
    Estado31.AddItem "Cerrada"
    Estado31.AddItem ""
    
    Estado31.ListIndex = 0
    
    Estado32.Clear
    
    Estado32.AddItem "No Imple."
    Estado32.AddItem "Imple."
    Estado32.AddItem "Nula"
    Estado32.AddItem "Cerrada"
    Estado32.AddItem ""
    
    Estado32.ListIndex = 0
    
    Estado33.Clear
    
    Estado33.AddItem "No Imple."
    Estado33.AddItem "Imple."
    Estado33.AddItem "Nula"
    Estado33.AddItem "Cerrada"
    Estado33.AddItem ""
    
    Estado33.ListIndex = 0
    
    Estado34.Clear
    
    Estado34.AddItem "No Imple."
    Estado34.AddItem "Imple."
    Estado34.AddItem "Nula"
    Estado34.AddItem "Cerrada"
    Estado34.AddItem ""
    
    Estado34.ListIndex = 0
    
    Estado35.Clear
    
    Estado35.AddItem "No Imple."
    Estado35.AddItem "Imple."
    Estado35.AddItem "Nula"
    Estado35.AddItem "Cerrada"
    Estado35.AddItem ""
 
    Estado35.ListIndex = 0
    
    Estado36.Clear
    
    Estado36.AddItem "No Imple."
    Estado36.AddItem "Imple."
    Estado36.AddItem "Nula"
    Estado36.AddItem "Cerrada"
    Estado36.AddItem ""
    
    Estado36.ListIndex = 0
    
    
    
    
    
    
    
    Call Numero_KeyPress(13)
    
End Sub


Private Sub SiguienteNro_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero > " + "'" + Numero.Text + "'"
    ZSql = ZSql + " Order by Numero"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        With rstCargaSac
            .MoveFirst
            Numero.Text = rstCargaSac!Numero
            rstCargaSac.Close
            Call Numero_KeyPress(13)
        End With
    End If

End Sub

Private Sub SiguienteTipo_Click()

    ZZTipo = Tipo.Text
    ZZNumero = Numero.Text
    Tipo.Text = Str$(Val(Tipo.Text) + 1)
    Numero.Text = "0"

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero > " + "'" + Numero.Text + "'"
    ZSql = ZSql + " Order by Numero"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        With rstCargaSac
            .MoveFirst
            Numero.Text = rstCargaSac!Numero
            rstCargaSac.Close
            Call Numero_KeyPress(13)
        End With
            Else
        Tipo.Text = ZZTipo
        Numero.Text = ZZNumero
        Call Numero_KeyPress(13)
    End If


End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 2
            Accion111.Text = Accion11.Text
            Accion112.Text = Accion12.Text
            Accion121.Text = Accion21.Text
            Accion122.Text = Accion22.Text
            Accion131.Text = Accion31.Text
            Accion132.Text = Accion32.Text
            Accion141.Text = Accion41.Text
            Accion142.Text = Accion42.Text
            Accion151.Text = Accion51.Text
            Accion152.Text = Accion52.Text
            Accion161.Text = Accion61.Text
            Accion162.Text = Accion62.Text
        Case Else
    End Select
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Numero.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                
                Centro.Text = rstCargaSac!Centro
                Fecha.Text = rstCargaSac!Fecha
                Origen.ListIndex = rstCargaSac!Origen
                Estado.ListIndex = rstCargaSac!Estado
                ResponsableEmisor.Text = rstCargaSac!ResponsableEmisor
                ResponsableDestino.Text = rstCargaSac!ResponsableDestino
                Referencia.Text = Trim(rstCargaSac!Referencia)
                Titulo.Text = rstCargaSac!Titulo
                IngresoNoCon.Text = IIf(IsNull(rstCargaSac!IngresoNoCon), "", rstCargaSac!IngresoNoCon)
                IngresoCausa.Text = IIf(IsNull(rstCargaSac!IngresoCausa), "", rstCargaSac!IngresoCausa)
                Rem dada
                rstCargaSac.Close
                
                Call Imprime_Datos
                Call Imprime_Descripcion
                
                    Else
                    
                WTipo = Tipo.Text
                WAno = Ano.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Ano.Text = WAno
                Numero.Text = WNumero
                Tipo.Text = WTipo
                Sql1 = "Select *"
                Sql2 = " FROM TipoSac"
                Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
                spTipoSac = Sql1 + Sql2 + Sql3
                Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstTipoSac.RecordCount > 0 Then
                    DesTipo.Caption = Trim(rstTipoSac!Descripcion)
                    rstTipoSac.Close
                    Ano.SetFocus
                End If
                Tipo.SetFocus
                
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub











Private Sub Centro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM CentroSac"
        Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
        spCentroSac = Sql1 + Sql2 + Sql3
        Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCentroSac.RecordCount > 0 Then
            DesCentro.Caption = Trim(rstCentroSac!Descripcion)
            rstCentroSac.Close
            ResponsableEmisor.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Centro.Text = ""
        DesCentro.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ResponsableEmisor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
            ResponsableDestino.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        ResponsableEmisor.Text = ""
        DesResponsableEmisor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ResponsableDestino_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM ResponsableSac"
        Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
        spResponsableSac = Sql1 + Sql2 + Sql3
        Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstResponsableSac.RecordCount > 0 Then
            DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
            rstResponsableSac.Close
            Referencia.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        ResponsableDestino.Text = ""
        DesResponsableDestino.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Referencia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Titulo.SetFocus
    End If
    If KeyAscii = 27 Then
        Referencia.Text = ""
    End If
End Sub

Private Sub Titulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Centro.SetFocus
    End If
    If KeyAscii = 27 Then
        Titulo.Text = ""
    End If
End Sub


Private Sub Accion11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion12.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion11.Text = ""
    End If
End Sub

Private Sub Accion12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable1.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion12.Text = ""
    End If
End Sub

Private Sub Responsable1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable1.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo1.SetFocus
            End If
                Else
            DesResponsable1.Caption = ""
            Plazo1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable1.Text = ""
        DesResponsable1.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo1.Text, Auxi)
        If Auxi = "S" Or Plazo1.Text = "  /  /    " Then
            Accion21.SetFocus
                Else
            Plazo1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo1.Text = "  /  /    "
    End If
End Sub








Private Sub Accion21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion22.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion21.Text = ""
    End If
End Sub

Private Sub Accion22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable2.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion22.Text = ""
    End If
End Sub

Private Sub Responsable2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable2.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo2.SetFocus
            End If
                Else
            DesResponsable2.Caption = ""
            Plazo2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable2.Text = ""
        DesResponsable2.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo2.Text, Auxi)
        If Auxi = "S" Or Plazo2.Text = "  /  /    " Then
            Accion31.SetFocus
                Else
            Plazo2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo2.Text = "  /  /    "
    End If
End Sub







Private Sub Accion31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion32.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion31.Text = ""
    End If
End Sub

Private Sub Accion32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable3.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion32.Text = ""
    End If
End Sub

Private Sub Responsable3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable3.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo3.SetFocus
            End If
                Else
            DesResponsable3.Caption = ""
            Plazo3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable3.Text = ""
        DesResponsable3.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo3.Text, Auxi)
        If Auxi = "S" Or Plazo3.Text = "  /  /    " Then
            Accion41.SetFocus
                Else
            Plazo3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo3.Text = "  /  /    "
    End If
End Sub







Private Sub Accion41_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion42.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion41.Text = ""
    End If
End Sub

Private Sub Accion42_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable4.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion42.Text = ""
    End If
End Sub

Private Sub Responsable4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable4.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo4.SetFocus
            End If
                Else
            DesResponsable4.Caption = ""
            Plazo4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable4.Text = ""
        DesResponsable4.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo4.Text, Auxi)
        If Auxi = "S" Or Plazo4.Text = "  /  /    " Then
            Accion51.SetFocus
                Else
            Plazo4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo4.Text = "  /  /    "
    End If
End Sub







Private Sub Accion51_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion52.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion51.Text = ""
    End If
End Sub

Private Sub Accion52_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable5.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion52.Text = ""
    End If
End Sub

Private Sub Responsable5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable5.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo5.SetFocus
            End If
                Else
            DesResponsable5.Caption = ""
            Plazo5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable5.Text = ""
        DesResponsable5.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo5.Text, Auxi)
        If Auxi = "S" Or Plazo5.Text = "  /  /    " Then
            Accion61.SetFocus
                Else
            Plazo5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo5.Text = "  /  /    "
    End If
End Sub







Private Sub Accion61_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion62.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion61.Text = ""
    End If
End Sub

Private Sub Accion62_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable6.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion62.Text = ""
    End If
End Sub

Private Sub Responsable6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable6.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo6.SetFocus
            End If
                Else
            DesResponsable6.Caption = ""
            Plazo6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable6.Text = ""
        DesResponsable6.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo6.Text, Auxi)
        If Auxi = "S" Or Plazo6.Text = "  /  /    " Then
            Accion11.SetFocus
                Else
            Plazo6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo6.Text = "  /  /    "
    End If
End Sub





Private Sub Responsable11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable11.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable11.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable11.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha1.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable11.Text = ""
        DesResponsable11.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha1.Text, Auxi)
        If Auxi = "S" Or Fecha1.Text = "  /  /    " Then
            Estado11.SetFocus
                Else
            Fecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha1.Text = "  /  /    "
    End If
End Sub

Private Sub Estado11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario11.SetFocus
    End If
End Sub

Private Sub Comentario11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario12.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario11.Text = ""
    End If
End Sub

Private Sub Comentario12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable12.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario12.Text = ""
    End If
End Sub




Private Sub Responsable12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable12.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable12.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable12.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha2.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable12.Text = ""
        DesResponsable12.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha2.Text, Auxi)
        If Auxi = "S" Or Fecha2.Text = "  /  /    " Then
            Estado12.SetFocus
                Else
            Fecha2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha2.Text = "  /  /    "
    End If
End Sub

Private Sub Estado12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario21.SetFocus
    End If
End Sub

Private Sub Comentario21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario22.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario21.Text = ""
    End If
End Sub

Private Sub Comentario22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable13.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario22.Text = ""
    End If
End Sub





Private Sub Responsable13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable13.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable13.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable13.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha3.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable13.Text = ""
        DesResponsable13.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha3.Text, Auxi)
        If Auxi = "S" Or Fecha3.Text = "  /  /    " Then
            Estado13.SetFocus
                Else
            Fecha3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha3.Text = "  /  /    "
    End If
End Sub

Private Sub Estado13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario31.SetFocus
    End If
End Sub

Private Sub Comentario31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario32.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario31.Text = ""
    End If
End Sub

Private Sub Comentario32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable14.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario32.Text = ""
    End If
End Sub






Private Sub Responsable14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable14.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable14.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable14.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha4.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable14.Text = ""
        DesResponsable14.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha4.Text, Auxi)
        If Auxi = "S" Or Fecha4.Text = "  /  /    " Then
            Estado14.SetFocus
                Else
            Fecha4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha4.Text = "  /  /    "
    End If
End Sub

Private Sub Estado14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario41.SetFocus
    End If
End Sub

Private Sub Comentario41_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario42.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario41.Text = ""
    End If
End Sub

Private Sub Comentario42_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable15.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario42.Text = ""
    End If
End Sub





Private Sub Responsable15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable15.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable15.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable15.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha5.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable15.Text = ""
        DesResponsable15.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha5.Text, Auxi)
        If Auxi = "S" Or Fecha5.Text = "  /  /    " Then
            Estado15.SetFocus
                Else
            Fecha5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha5.Text = "  /  /    "
    End If
End Sub

Private Sub Estado15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario51.SetFocus
    End If
End Sub

Private Sub Comentario51_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario52.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario51.Text = ""
    End If
End Sub

Private Sub Comentario52_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable16.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario52.Text = ""
    End If
End Sub






Private Sub Responsable16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable16.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable16.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable16.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha6.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable16.Text = ""
        DesResponsable16.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha6.Text, Auxi)
        If Auxi = "S" Or Fecha6.Text = "  /  /    " Then
            Estado16.SetFocus
                Else
            Fecha6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha6.Text = "  /  /    "
    End If
End Sub

Private Sub Estado16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario61.SetFocus
    End If
End Sub

Private Sub Comentario61_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario62.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario61.Text = ""
    End If
End Sub

Private Sub Comentario62_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable11.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario62.Text = ""
    End If
End Sub







































Private Sub Responsable21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable21.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable21.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable21.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha21.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable21.Text = ""
        DesResponsable21.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha21.Text, Auxi)
        If Auxi = "S" Or Fecha21.Text = "  /  /    " Then
            Estado1.SetFocus
                Else
            Fecha21.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha21.Text = "  /  /    "
    End If
End Sub

Private Sub Estado1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable31.SetFocus
    End If
End Sub

Private Sub Responsable31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable31.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable31.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable31.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha31.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable31.Text = ""
        DesResponsable31.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha31.Text, Auxi)
        If Auxi = "S" Or Fecha31.Text = "  /  /    " Then
            Estado31.SetFocus
                Else
            Fecha31.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha31.Text = "  /  /    "
    End If
End Sub

Private Sub Estado31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario211.SetFocus
    End If
End Sub

Private Sub Comentario211_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario212.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario211.Text = ""
    End If
End Sub

Private Sub Comentario212_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable22.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario212.Text = ""
    End If
End Sub






Private Sub Responsable22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable22.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable22.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable22.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha22.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable22.Text = ""
        DesResponsable22.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha22.Text, Auxi)
        If Auxi = "S" Or Fecha22.Text = "  /  /    " Then
            Estado2.SetFocus
                Else
            Fecha22.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha22.Text = "  /  /    "
    End If
End Sub

Private Sub Estado2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable32.SetFocus
    End If
End Sub

Private Sub Responsable32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable32.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable32.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable32.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha32.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable32.Text = ""
        DesResponsable32.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha32.Text, Auxi)
        If Auxi = "S" Or Fecha32.Text = "  /  /    " Then
            Estado32.SetFocus
                Else
            Fecha32.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha32.Text = "  /  /    "
    End If
End Sub

Private Sub Estado32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario221.SetFocus
    End If
End Sub

Private Sub Comentario221_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario222.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario221.Text = ""
    End If
End Sub

Private Sub Comentario222_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable23.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario222.Text = ""
    End If
End Sub










Private Sub Responsable23_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable23.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable23.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable23.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha23.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable23.Text = ""
        DesResponsable23.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha23_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha23.Text, Auxi)
        If Auxi = "S" Or Fecha23.Text = "  /  /    " Then
            Estado3.SetFocus
                Else
            Fecha23.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha23.Text = "  /  /    "
    End If
End Sub

Private Sub Estado3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable33.SetFocus
    End If
End Sub

Private Sub Responsable33_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable33.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable33.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable33.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha33.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable33.Text = ""
        DesResponsable33.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha33_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha3.Text, Auxi)
        If Auxi = "S" Or Fecha33.Text = "  /  /    " Then
            Estado33.SetFocus
                Else
            Fecha33.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha33.Text = "  /  /    "
    End If
End Sub

Private Sub Estado33_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario231.SetFocus
    End If
End Sub

Private Sub Comentario231_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario232.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario231.Text = ""
    End If
End Sub

Private Sub Comentario232_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable24.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario232.Text = ""
    End If
End Sub






Private Sub Responsable24_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable24.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable24.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable24.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha24.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable24.Text = ""
        DesResponsable24.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha24_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha24.Text, Auxi)
        If Auxi = "S" Or Fecha24.Text = "  /  /    " Then
            Estado4.SetFocus
                Else
            Fecha24.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha24.Text = "  /  /    "
    End If
End Sub

Private Sub Estado4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable34.SetFocus
    End If
End Sub

Private Sub Responsable34_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable34.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable34.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable34.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha34.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable34.Text = ""
        DesResponsable34.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha34_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha34.Text, Auxi)
        If Auxi = "S" Or Fecha34.Text = "  /  /    " Then
            Estado34.SetFocus
                Else
            Fecha34.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha34.Text = "  /  /    "
    End If
End Sub

Private Sub Estado34_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario241.SetFocus
    End If
End Sub

Private Sub Comentario241_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario242.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario241.Text = ""
    End If
End Sub

Private Sub Comentario242_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable25.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario242.Text = ""
    End If
End Sub








Private Sub Responsable25_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable25.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable25.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable25.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha25.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable25.Text = ""
        DesResponsable25.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha25_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha25.Text, Auxi)
        If Auxi = "S" Or Fecha25.Text = "  /  /    " Then
            Estado5.SetFocus
                Else
            Fecha25.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha25.Text = "  /  /    "
    End If
End Sub

Private Sub Estado5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable35.SetFocus
    End If
End Sub

Private Sub Responsable35_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable35.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable35.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable35.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha35.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable35.Text = ""
        DesResponsable35.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha35_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha35.Text, Auxi)
        If Auxi = "S" Or Fecha35.Text = "  /  /    " Then
            Estado35.SetFocus
                Else
            Fecha35.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha35.Text = "  /  /    "
    End If
End Sub

Private Sub Estado35_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario251.SetFocus
    End If
End Sub

Private Sub Comentario251_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario252.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario251.Text = ""
    End If
End Sub

Private Sub Comentario252_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable26.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario252.Text = ""
    End If
End Sub







Private Sub Responsable26_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable26.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable26.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable26.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha26.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable26.Text = ""
        DesResponsable26.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha26_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha26.Text, Auxi)
        If Auxi = "S" Or Fecha26.Text = "  /  /    " Then
            Estado6.SetFocus
                Else
            Fecha26.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha26.Text = "  /  /    "
    End If
End Sub

Private Sub Estado6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable36.SetFocus
    End If
End Sub

Private Sub Responsable36_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable36.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable36.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable36.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha36.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Responsable36.Text = ""
        DesResponsable36.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha36_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha36.Text, Auxi)
        If Auxi = "S" Or Fecha36.Text = "  /  /    " Then
            Estado36.SetFocus
                Else
            Fecha36.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha36.Text = "  /  /    "
    End If
End Sub

Private Sub Estado36_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario261.SetFocus
    End If
End Sub

Private Sub Comentario261_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario262.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario261.Text = ""
    End If
End Sub

Private Sub Comentario262_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable21.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario262.Text = ""
    End If
End Sub














Private Sub Centro_DblClick()

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub


Private Sub ResponsableEmisor_DblClick()

    ZZLugar = 1

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub ResponsableDestino_DblClick()

    ZZLugar = 2

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem ""
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable1_DblClick()

    ZZLugar = 3

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable2_DblClick()

    ZZLugar = 4

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable3_DblClick()

    ZZLugar = 5

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable4_DblClick()

    ZZLugar = 6

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable5_DblClick()

    ZZLugar = 7

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable6_DblClick()

    ZZLugar = 8

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub


Private Sub Responsable31_DblClick()

    ZZLugar = 9

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable32_DblClick()

    ZZLugar = 10

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable33_DblClick()

    ZZLugar = 11

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable34_DblClick()

    ZZLugar = 12

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable35_DblClick()

    ZZLugar = 13

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable36_DblClick()

    ZZLugar = 14

    Opcion.Clear
    Opcion.AddItem "Tipo"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub


























Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False
     Opcion.Clear
     Opcion.AddItem "Centro"
     Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Ayuda.Text = ""
    Ayuda.Visible = True
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstCentroSac!Codigo) + " " + rstCentroSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCentroSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCentroSac.Close
            End If
        
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstResponsableSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
        
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Centro.Text = WIndice.List(Indice)
            Call Centro_Keypress(13)
            
        Case 1
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Select Case ZZLugar
                Case 1
                    ResponsableEmisor.Text = WIndice.List(Indice)
                    Call ResponsableEmisor_Keypress(13)
                Case 2
                    ResponsableDestino.Text = WIndice.List(Indice)
                    Call ResponsableDestino_Keypress(13)
                Case 3
                    Responsable1.Text = WIndice.List(Indice)
                    Call Responsable1_Keypress(13)
                Case 4
                    Responsable2.Text = WIndice.List(Indice)
                    Call Responsable2_Keypress(13)
                Case 5
                    Responsable3.Text = WIndice.List(Indice)
                    Call Responsable3_Keypress(13)
                Case 6
                    Responsable4.Text = WIndice.List(Indice)
                    Call Responsable4_Keypress(13)
                Case 7
                    Responsable5.Text = WIndice.List(Indice)
                    Call Responsable5_Keypress(13)
                Case 8
                    Responsable6.Text = WIndice.List(Indice)
                    Call Responsable6_Keypress(13)
                Case 9
                    Responsable31.Text = WIndice.List(Indice)
                    Call Responsable31_Keypress(13)
                Case 10
                    Responsable32.Text = WIndice.List(Indice)
                    Call Responsable32_Keypress(13)
                Case 11
                    Responsable33.Text = WIndice.List(Indice)
                    Call Responsable33_Keypress(13)
                Case 12
                    Responsable34.Text = WIndice.List(Indice)
                    Call Responsable34_Keypress(13)
                Case 13
                    Responsable35.Text = WIndice.List(Indice)
                    Call Responsable35_Keypress(13)
                Case 14
                    Responsable36.Text = WIndice.List(Indice)
                    Call Responsable36_Keypress(13)
            
                Case Else
            End Select
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    LugarAyuda = 0
    WIndice.Clear
    Pantalla.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstCentroSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstCentroSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstCentroSac!Codigo) + " " + rstCentroSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstCentroSac!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCentroSac.Close
            End If

        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstResponsableSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstResponsableSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstResponsableSac!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub


Private Sub Impresion_Click()

    Auxi3 = Tipo.Text
    Auxi1 = Ano.Text
    Auxi2 = Numero.Text
    Call Ceros(Auxi3, 4)
    Call Ceros(Auxi1, 4)
    Call Ceros(Auxi2, 6)
    ZZClave = Auxi3 + Auxi1 + Auxi2


    ZSql = ""
    ZSql = ZSql + "DELETE ImpreSac"
    Rem ZSql = ZSql + " Where ImpreSac.Tipo = " + "'" + Tipo.Text + "'"
    Rem ZSql = ZSql + " and ImpreSac.Ano = " + "'" + Ano.Text + "'"
    Rem ZSql = ZSql + " and ImpreSac.Numero = " + "'" + Numero.Text + "'"
    rsImpreSac = ZSql
    Set rstImpreSac = db.OpenRecordset(rsImpreSac, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZGraba = "S"
    
    If Trim(Accion11.Text) <> "" Then
        
        ZZGraba = "N"
        
        ZZImpre11 = Accion11.Text
        ZZImpre12 = Accion12.Text
        ZZImpre13 = DesResponsable1.Caption
        ZZImpre14 = Plazo1.Text
        ZZImpre15 = ""
        
        ZZImpre21 = ""
        ZZImpre22 = ""
        ZZImpre23 = DesResponsable11.Caption
        ZZImpre24 = Fecha1.Text
        If Trim(DesResponsable11.Caption) <> "" Then
            ZZImpre25 = Estado11.Text
                Else
            ZZImpre25 = ""
        End If
        ZZImpre26 = Comentario11.Text
        ZZImpre27 = Comentario12.Text
        ZZImpre28 = ""
        
        ZZImpre31 = DesResponsable21.Caption
        If Trim(DesResponsable21.Caption) <> "" Then
            ZZImpre32 = Estado1.Text
                Else
            ZZImpre32 = ""
        End If
        ZZImpre33 = Fecha21.Text
        
        ZZImpre34 = DesResponsable31.Caption
        If Trim(DesResponsable31.Caption) <> "" Then
            ZZImpre35 = Estado31.Text
                Else
            ZZImpre35 = ""
        End If
        ZZImpre36 = Fecha31.Text
        ZZImpre37 = Comentario211.Text
        ZZImpre38 = Comentario212.Text
        ZZImpre39 = ""
        ZZImpre40 = "1"
        
        ZZCorte = "1"
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "DesTipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "DesCentro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa ,"
        ZSql = ZSql + "DesResponsableEmisor ,"
        ZSql = ZSql + "DesResponsableDestino ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Impre32 ,"
        ZSql = ZSql + "Impre33 ,"
        ZSql = ZSql + "Impre34 ,"
        ZSql = ZSql + "Impre35 ,"
        ZSql = ZSql + "Impre36 ,"
        ZSql = ZSql + "Impre37 ,"
        ZSql = ZSql + "Impre38 ,"
        ZSql = ZSql + "Impre39 ,"
        ZSql = ZSql + "Impre40 ,"
        ZSql = ZSql + "Comentario )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + DesTipo.Caption + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + DesCentro.Caption + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Estado.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "',"
        ZSql = ZSql + "'" + DesResponsableEmisor.Caption + "',"
        ZSql = ZSql + "'" + DesResponsableDestino.Caption + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + ZZCorte + "',"
        ZSql = ZSql + "'" + ZZImpre11 + "',"
        ZSql = ZSql + "'" + ZZImpre12 + "',"
        ZSql = ZSql + "'" + ZZImpre13 + "',"
        ZSql = ZSql + "'" + ZZImpre14 + "',"
        ZSql = ZSql + "'" + ZZImpre15 + "',"
        ZSql = ZSql + "'" + ZZImpre21 + "',"
        ZSql = ZSql + "'" + ZZImpre22 + "',"
        ZSql = ZSql + "'" + ZZImpre23 + "',"
        ZSql = ZSql + "'" + ZZImpre24 + "',"
        ZSql = ZSql + "'" + ZZImpre25 + "',"
        ZSql = ZSql + "'" + ZZImpre26 + "',"
        ZSql = ZSql + "'" + ZZImpre27 + "',"
        ZSql = ZSql + "'" + ZZImpre28 + "',"
        ZSql = ZSql + "'" + ZZImpre31 + "',"
        ZSql = ZSql + "'" + ZZImpre32 + "',"
        ZSql = ZSql + "'" + ZZImpre33 + "',"
        ZSql = ZSql + "'" + ZZImpre34 + "',"
        ZSql = ZSql + "'" + ZZImpre35 + "',"
        ZSql = ZSql + "'" + ZZImpre36 + "',"
        ZSql = ZSql + "'" + ZZImpre37 + "',"
        ZSql = ZSql + "'" + ZZImpre38 + "',"
        ZSql = ZSql + "'" + ZZImpre39 + "',"
        ZSql = ZSql + "'" + ZZImpre40 + "',"
        ZSql = ZSql + "'" + Comentario.Text + "')"
        
        spImpreSac = ZSql
        Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    If Trim(Accion21.Text) <> "" Then
        
        ZZGraba = "N"
        
        ZZImpre11 = Accion21.Text
        ZZImpre12 = Accion22.Text
        ZZImpre13 = DesResponsable2.Caption
        ZZImpre14 = Plazo2.Text
        ZZImpre15 = ""
        
        ZZImpre21 = ""
        ZZImpre22 = ""
        ZZImpre23 = DesResponsable12.Caption
        ZZImpre24 = Fecha2.Text
        If Trim(DesResponsable12.Caption) <> "" Then
            ZZImpre25 = Estado12.Text
                Else
            ZZImpre25 = ""
        End If
        ZZImpre26 = Comentario21.Text
        ZZImpre27 = Comentario22.Text
        ZZImpre28 = ""
        
        ZZImpre31 = DesResponsable22.Caption
        If Trim(DesResponsable22.Caption) <> "" Then
            ZZImpre32 = Estado2.Text
                Else
            ZZImpre32 = ""
        End If
        ZZImpre33 = Fecha22.Text
        
        ZZImpre34 = DesResponsable32.Caption
        If Trim(DesResponsable32.Caption) <> "" Then
            ZZImpre35 = Estado32.Text
                Else
            ZZImpre35 = ""
        End If
        ZZImpre36 = Fecha32.Text
        ZZImpre37 = Comentario221.Text
        ZZImpre38 = Comentario222.Text
        ZZImpre39 = ""
        ZZImpre40 = "2"
        
        ZZCorte = "1"
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "DesTipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "DesCentro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa ,"
        ZSql = ZSql + "DesResponsableEmisor ,"
        ZSql = ZSql + "DesResponsableDestino ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Impre32 ,"
        ZSql = ZSql + "Impre33 ,"
        ZSql = ZSql + "Impre34 ,"
        ZSql = ZSql + "Impre35 ,"
        ZSql = ZSql + "Impre36 ,"
        ZSql = ZSql + "Impre37 ,"
        ZSql = ZSql + "Impre38 ,"
        ZSql = ZSql + "Impre39 ,"
        ZSql = ZSql + "Impre40 ,"
        ZSql = ZSql + "Comentario )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + DesTipo.Caption + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + DesCentro.Caption + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Estado.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "',"
        ZSql = ZSql + "'" + DesResponsableEmisor.Caption + "',"
        ZSql = ZSql + "'" + DesResponsableDestino.Caption + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + ZZCorte + "',"
        ZSql = ZSql + "'" + ZZImpre11 + "',"
        ZSql = ZSql + "'" + ZZImpre12 + "',"
        ZSql = ZSql + "'" + ZZImpre13 + "',"
        ZSql = ZSql + "'" + ZZImpre14 + "',"
        ZSql = ZSql + "'" + ZZImpre15 + "',"
        ZSql = ZSql + "'" + ZZImpre21 + "',"
        ZSql = ZSql + "'" + ZZImpre22 + "',"
        ZSql = ZSql + "'" + ZZImpre23 + "',"
        ZSql = ZSql + "'" + ZZImpre24 + "',"
        ZSql = ZSql + "'" + ZZImpre25 + "',"
        ZSql = ZSql + "'" + ZZImpre26 + "',"
        ZSql = ZSql + "'" + ZZImpre27 + "',"
        ZSql = ZSql + "'" + ZZImpre28 + "',"
        ZSql = ZSql + "'" + ZZImpre31 + "',"
        ZSql = ZSql + "'" + ZZImpre32 + "',"
        ZSql = ZSql + "'" + ZZImpre33 + "',"
        ZSql = ZSql + "'" + ZZImpre34 + "',"
        ZSql = ZSql + "'" + ZZImpre35 + "',"
        ZSql = ZSql + "'" + ZZImpre36 + "',"
        ZSql = ZSql + "'" + ZZImpre37 + "',"
        ZSql = ZSql + "'" + ZZImpre38 + "',"
        ZSql = ZSql + "'" + ZZImpre39 + "',"
        ZSql = ZSql + "'" + ZZImpre40 + "',"
        ZSql = ZSql + "'" + Comentario.Text + "')"
        
        spImpreSac = ZSql
        Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    
    
    
    If Trim(Accion31.Text) <> "" Then
        
        ZZGraba = "N"
        
        ZZImpre11 = Accion31.Text
        ZZImpre12 = Accion32.Text
        ZZImpre13 = DesResponsable3.Caption
        ZZImpre14 = Plazo3.Text
        ZZImpre15 = ""
        
        ZZImpre21 = ""
        ZZImpre22 = ""
        ZZImpre23 = DesResponsable13.Caption
        ZZImpre24 = Fecha3.Text
        If Trim(DesResponsable13.Caption) <> "" Then
            ZZImpre25 = Estado13.Text
                Else
            ZZImpre25 = ""
        End If
        ZZImpre26 = Comentario31.Text
        ZZImpre27 = Comentario32.Text
        ZZImpre28 = ""
        
        ZZImpre31 = DesResponsable23.Caption
        If Trim(DesResponsable23.Caption) <> "" Then
            ZZImpre32 = Estado3.Text
                Else
            ZZImpre32 = ""
        End If
        ZZImpre33 = Fecha23.Text
        
        ZZImpre34 = DesResponsable33.Caption
        If Trim(DesResponsable33.Caption) <> "" Then
            ZZImpre35 = Estado33.Text
                Else
            ZZImpre35 = ""
        End If
        ZZImpre36 = Fecha33.Text
        ZZImpre37 = Comentario231.Text
        ZZImpre38 = Comentario232.Text
        ZZImpre39 = ""
        ZZImpre40 = "3"
        
        ZZCorte = "1"
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "DesTipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "DesCentro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa ,"
        ZSql = ZSql + "DesResponsableEmisor ,"
        ZSql = ZSql + "DesResponsableDestino ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Impre32 ,"
        ZSql = ZSql + "Impre33 ,"
        ZSql = ZSql + "Impre34 ,"
        ZSql = ZSql + "Impre35 ,"
        ZSql = ZSql + "Impre36 ,"
        ZSql = ZSql + "Impre37 ,"
        ZSql = ZSql + "Impre38 ,"
        ZSql = ZSql + "Impre39 ,"
        ZSql = ZSql + "Impre40 ,"
        ZSql = ZSql + "Comentario )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + DesTipo.Caption + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + DesCentro.Caption + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Estado.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "',"
        ZSql = ZSql + "'" + DesResponsableEmisor.Caption + "',"
        ZSql = ZSql + "'" + DesResponsableDestino.Caption + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + ZZCorte + "',"
        ZSql = ZSql + "'" + ZZImpre11 + "',"
        ZSql = ZSql + "'" + ZZImpre12 + "',"
        ZSql = ZSql + "'" + ZZImpre13 + "',"
        ZSql = ZSql + "'" + ZZImpre14 + "',"
        ZSql = ZSql + "'" + ZZImpre15 + "',"
        ZSql = ZSql + "'" + ZZImpre21 + "',"
        ZSql = ZSql + "'" + ZZImpre22 + "',"
        ZSql = ZSql + "'" + ZZImpre23 + "',"
        ZSql = ZSql + "'" + ZZImpre24 + "',"
        ZSql = ZSql + "'" + ZZImpre25 + "',"
        ZSql = ZSql + "'" + ZZImpre26 + "',"
        ZSql = ZSql + "'" + ZZImpre27 + "',"
        ZSql = ZSql + "'" + ZZImpre28 + "',"
        ZSql = ZSql + "'" + ZZImpre31 + "',"
        ZSql = ZSql + "'" + ZZImpre32 + "',"
        ZSql = ZSql + "'" + ZZImpre33 + "',"
        ZSql = ZSql + "'" + ZZImpre34 + "',"
        ZSql = ZSql + "'" + ZZImpre35 + "',"
        ZSql = ZSql + "'" + ZZImpre36 + "',"
        ZSql = ZSql + "'" + ZZImpre37 + "',"
        ZSql = ZSql + "'" + ZZImpre38 + "',"
        ZSql = ZSql + "'" + ZZImpre39 + "',"
        ZSql = ZSql + "'" + ZZImpre40 + "',"
        ZSql = ZSql + "'" + Comentario.Text + "')"
        
        spImpreSac = ZSql
        Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    
    
    
    
    If Trim(Accion41.Text) <> "" Then
        
        ZZGraba = "N"
        
        ZZImpre11 = Accion41.Text
        ZZImpre12 = Accion42.Text
        ZZImpre13 = DesResponsable4.Caption
        ZZImpre14 = Plazo4.Text
        ZZImpre15 = ""
        
        ZZImpre21 = ""
        ZZImpre22 = ""
        ZZImpre23 = DesResponsable14.Caption
        ZZImpre24 = Fecha4.Text
        If Trim(DesResponsable14.Caption) <> "" Then
            ZZImpre25 = Estado14.Text
                Else
            ZZImpre25 = ""
        End If
        ZZImpre26 = Comentario41.Text
        ZZImpre27 = Comentario42.Text
        ZZImpre28 = ""
        
        ZZImpre31 = DesResponsable24.Caption
        If Trim(DesResponsable24.Caption) <> "" Then
            ZZImpre32 = Estado4.Text
                Else
            ZZImpre32 = ""
        End If
        ZZImpre33 = Fecha24.Text
        
        ZZImpre34 = DesResponsable34.Caption
        If Trim(DesResponsable34.Caption) <> "" Then
            ZZImpre35 = Estado34.Text
                Else
            ZZImpre35 = ""
        End If
        ZZImpre36 = Fecha34.Text
        ZZImpre37 = Comentario241.Text
        ZZImpre38 = Comentario242.Text
        ZZImpre39 = ""
        ZZImpre40 = "4"
        
        ZZCorte = "1"
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "DesTipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "DesCentro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa ,"
        ZSql = ZSql + "DesResponsableEmisor ,"
        ZSql = ZSql + "DesResponsableDestino ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Impre32 ,"
        ZSql = ZSql + "Impre33 ,"
        ZSql = ZSql + "Impre34 ,"
        ZSql = ZSql + "Impre35 ,"
        ZSql = ZSql + "Impre36 ,"
        ZSql = ZSql + "Impre37 ,"
        ZSql = ZSql + "Impre38 ,"
        ZSql = ZSql + "Impre39 ,"
        ZSql = ZSql + "Impre40 ,"
        ZSql = ZSql + "Comentario )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + DesTipo.Caption + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + DesCentro.Caption + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Estado.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "',"
        ZSql = ZSql + "'" + DesResponsableEmisor.Caption + "',"
        ZSql = ZSql + "'" + DesResponsableDestino.Caption + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + ZZCorte + "',"
        ZSql = ZSql + "'" + ZZImpre11 + "',"
        ZSql = ZSql + "'" + ZZImpre12 + "',"
        ZSql = ZSql + "'" + ZZImpre13 + "',"
        ZSql = ZSql + "'" + ZZImpre14 + "',"
        ZSql = ZSql + "'" + ZZImpre15 + "',"
        ZSql = ZSql + "'" + ZZImpre21 + "',"
        ZSql = ZSql + "'" + ZZImpre22 + "',"
        ZSql = ZSql + "'" + ZZImpre23 + "',"
        ZSql = ZSql + "'" + ZZImpre24 + "',"
        ZSql = ZSql + "'" + ZZImpre25 + "',"
        ZSql = ZSql + "'" + ZZImpre26 + "',"
        ZSql = ZSql + "'" + ZZImpre27 + "',"
        ZSql = ZSql + "'" + ZZImpre28 + "',"
        ZSql = ZSql + "'" + ZZImpre31 + "',"
        ZSql = ZSql + "'" + ZZImpre32 + "',"
        ZSql = ZSql + "'" + ZZImpre33 + "',"
        ZSql = ZSql + "'" + ZZImpre34 + "',"
        ZSql = ZSql + "'" + ZZImpre35 + "',"
        ZSql = ZSql + "'" + ZZImpre36 + "',"
        ZSql = ZSql + "'" + ZZImpre37 + "',"
        ZSql = ZSql + "'" + ZZImpre38 + "',"
        ZSql = ZSql + "'" + ZZImpre39 + "',"
        ZSql = ZSql + "'" + ZZImpre40 + "',"
        ZSql = ZSql + "'" + Comentario.Text + "')"
        
        spImpreSac = ZSql
        Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    If Trim(Accion51.Text) <> "" Then
        
        ZZGraba = "N"
        
        ZZImpre11 = Accion51.Text
        ZZImpre12 = Accion52.Text
        ZZImpre13 = DesResponsable5.Caption
        ZZImpre14 = Plazo5.Text
        ZZImpre15 = ""
        
        ZZImpre21 = ""
        ZZImpre22 = ""
        ZZImpre23 = DesResponsable15.Caption
        ZZImpre24 = Fecha5.Text
        If Trim(DesResponsable15.Caption) <> "" Then
            ZZImpre25 = Estado15.Text
                Else
            ZZImpre25 = ""
        End If
        ZZImpre26 = Comentario51.Text
        ZZImpre27 = Comentario52.Text
        ZZImpre28 = ""
        
        ZZImpre31 = DesResponsable25.Caption
        If Trim(DesResponsable25.Caption) <> "" Then
            ZZImpre32 = Estado5.Text
                Else
            ZZImpre32 = ""
        End If
        ZZImpre33 = Fecha25.Text
        
        ZZImpre34 = DesResponsable35.Caption
        If Trim(DesResponsable35.Caption) <> "" Then
            ZZImpre35 = Estado35.Text
                Else
            ZZImpre35 = ""
        End If
        ZZImpre36 = Fecha35.Text
        ZZImpre37 = Comentario251.Text
        ZZImpre38 = Comentario252.Text
        ZZImpre39 = ""
        ZZImpre40 = "5"
        
        ZZCorte = "1"
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "DesTipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "DesCentro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa ,"
        ZSql = ZSql + "DesResponsableEmisor ,"
        ZSql = ZSql + "DesResponsableDestino ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Impre32 ,"
        ZSql = ZSql + "Impre33 ,"
        ZSql = ZSql + "Impre34 ,"
        ZSql = ZSql + "Impre35 ,"
        ZSql = ZSql + "Impre36 ,"
        ZSql = ZSql + "Impre37 ,"
        ZSql = ZSql + "Impre38 ,"
        ZSql = ZSql + "Impre39 ,"
        ZSql = ZSql + "Impre40 ,"
        ZSql = ZSql + "Comentario )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + DesTipo.Caption + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + DesCentro.Caption + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Estado.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "',"
        ZSql = ZSql + "'" + DesResponsableEmisor.Caption + "',"
        ZSql = ZSql + "'" + DesResponsableDestino.Caption + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + ZZCorte + "',"
        ZSql = ZSql + "'" + ZZImpre11 + "',"
        ZSql = ZSql + "'" + ZZImpre12 + "',"
        ZSql = ZSql + "'" + ZZImpre13 + "',"
        ZSql = ZSql + "'" + ZZImpre14 + "',"
        ZSql = ZSql + "'" + ZZImpre15 + "',"
        ZSql = ZSql + "'" + ZZImpre21 + "',"
        ZSql = ZSql + "'" + ZZImpre22 + "',"
        ZSql = ZSql + "'" + ZZImpre23 + "',"
        ZSql = ZSql + "'" + ZZImpre24 + "',"
        ZSql = ZSql + "'" + ZZImpre25 + "',"
        ZSql = ZSql + "'" + ZZImpre26 + "',"
        ZSql = ZSql + "'" + ZZImpre27 + "',"
        ZSql = ZSql + "'" + ZZImpre28 + "',"
        ZSql = ZSql + "'" + ZZImpre31 + "',"
        ZSql = ZSql + "'" + ZZImpre32 + "',"
        ZSql = ZSql + "'" + ZZImpre33 + "',"
        ZSql = ZSql + "'" + ZZImpre34 + "',"
        ZSql = ZSql + "'" + ZZImpre35 + "',"
        ZSql = ZSql + "'" + ZZImpre36 + "',"
        ZSql = ZSql + "'" + ZZImpre37 + "',"
        ZSql = ZSql + "'" + ZZImpre38 + "',"
        ZSql = ZSql + "'" + ZZImpre39 + "',"
        ZSql = ZSql + "'" + ZZImpre40 + "',"
        ZSql = ZSql + "'" + Comentario.Text + "')"
        
        spImpreSac = ZSql
        Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    If Trim(Accion61.Text) <> "" Then
        
        ZZGraba = "N"
        
        ZZImpre11 = Accion61.Text
        ZZImpre12 = Accion62.Text
        ZZImpre13 = DesResponsable6.Caption
        ZZImpre14 = Plazo6.Text
        ZZImpre15 = ""
        
        ZZImpre21 = ""
        ZZImpre22 = ""
        ZZImpre23 = DesResponsable16.Caption
        ZZImpre24 = Fecha6.Text
        If Trim(DesResponsable16.Caption) <> "" Then
            ZZImpre25 = Estado16.Text
                Else
            ZZImpre25 = ""
        End If
        ZZImpre26 = Comentario61.Text
        ZZImpre27 = Comentario62.Text
        ZZImpre28 = ""
        
        ZZImpre31 = DesResponsable26.Caption
        If Trim(DesResponsable26.Caption) <> "" Then
            ZZImpre32 = Estado6.Text
                Else
            ZZImpre32 = ""
        End If
        ZZImpre33 = Fecha26.Text
        
        ZZImpre34 = DesResponsable36.Caption
        If Trim(DesResponsable36.Caption) <> "" Then
            ZZImpre35 = Estado36.Text
                Else
            ZZImpre35 = ""
        End If
        ZZImpre36 = Fecha36.Text
        ZZImpre37 = Comentario261.Text
        ZZImpre38 = Comentario262.Text
        ZZImpre39 = ""
        ZZImpre40 = "6"
        
        ZZCorte = "1"
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "DesTipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "DesCentro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa ,"
        ZSql = ZSql + "DesResponsableEmisor ,"
        ZSql = ZSql + "DesResponsableDestino ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Impre32 ,"
        ZSql = ZSql + "Impre33 ,"
        ZSql = ZSql + "Impre34 ,"
        ZSql = ZSql + "Impre35 ,"
        ZSql = ZSql + "Impre36 ,"
        ZSql = ZSql + "Impre37 ,"
        ZSql = ZSql + "Impre38 ,"
        ZSql = ZSql + "Impre39 ,"
        ZSql = ZSql + "Impre40 ,"
        ZSql = ZSql + "Comentario )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + DesTipo.Caption + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + DesCentro.Caption + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Estado.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "',"
        ZSql = ZSql + "'" + DesResponsableEmisor.Caption + "',"
        ZSql = ZSql + "'" + DesResponsableDestino.Caption + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + ZZCorte + "',"
        ZSql = ZSql + "'" + ZZImpre11 + "',"
        ZSql = ZSql + "'" + ZZImpre12 + "',"
        ZSql = ZSql + "'" + ZZImpre13 + "',"
        ZSql = ZSql + "'" + ZZImpre14 + "',"
        ZSql = ZSql + "'" + ZZImpre15 + "',"
        ZSql = ZSql + "'" + ZZImpre21 + "',"
        ZSql = ZSql + "'" + ZZImpre22 + "',"
        ZSql = ZSql + "'" + ZZImpre23 + "',"
        ZSql = ZSql + "'" + ZZImpre24 + "',"
        ZSql = ZSql + "'" + ZZImpre25 + "',"
        ZSql = ZSql + "'" + ZZImpre26 + "',"
        ZSql = ZSql + "'" + ZZImpre27 + "',"
        ZSql = ZSql + "'" + ZZImpre28 + "',"
        ZSql = ZSql + "'" + ZZImpre31 + "',"
        ZSql = ZSql + "'" + ZZImpre32 + "',"
        ZSql = ZSql + "'" + ZZImpre33 + "',"
        ZSql = ZSql + "'" + ZZImpre34 + "',"
        ZSql = ZSql + "'" + ZZImpre35 + "',"
        ZSql = ZSql + "'" + ZZImpre36 + "',"
        ZSql = ZSql + "'" + ZZImpre37 + "',"
        ZSql = ZSql + "'" + ZZImpre38 + "',"
        ZSql = ZSql + "'" + ZZImpre39 + "',"
        ZSql = ZSql + "'" + ZZImpre40 + "',"
        ZSql = ZSql + "'" + Comentario.Text + "')"
        
        spImpreSac = ZSql
        Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    If ZZGraba = "S" Then
        
        ZZImpre11 = ""
        ZZImpre12 = ""
        ZZImpre13 = ""
        ZZImpre14 = ""
        ZZImpre15 = ""
        
        ZZImpre21 = ""
        ZZImpre22 = ""
        ZZImpre23 = ""
        ZZImpre24 = ""
        ZZImpre25 = ""
        ZZImpre26 = ""
        ZZImpre27 = ""
        ZZImpre28 = ""
        
        ZZImpre31 = ""
        ZZImpre32 = ""
        ZZImpre33 = ""
        ZZImpre34 = ""
        ZZImpre35 = ""
        ZZImpre36 = ""
        ZZImpre37 = ""
        ZZImpre38 = ""
        ZZImpre39 = ""
        ZZImpre40 = "2"
        
        ZZCorte = "1"
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreSac ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "DesTipo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "DesCentro ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "IngresoNoCon ,"
        ZSql = ZSql + "IngresoCausa ,"
        ZSql = ZSql + "DesResponsableEmisor ,"
        ZSql = ZSql + "DesResponsableDestino ,"
        ZSql = ZSql + "Titulo ,"
        ZSql = ZSql + "Referencia ,"
        ZSql = ZSql + "Corte ,"
        ZSql = ZSql + "Impre11 ,"
        ZSql = ZSql + "Impre12 ,"
        ZSql = ZSql + "Impre13 ,"
        ZSql = ZSql + "Impre14 ,"
        ZSql = ZSql + "Impre15 ,"
        ZSql = ZSql + "Impre21 ,"
        ZSql = ZSql + "Impre22 ,"
        ZSql = ZSql + "Impre23 ,"
        ZSql = ZSql + "Impre24 ,"
        ZSql = ZSql + "Impre25 ,"
        ZSql = ZSql + "Impre26 ,"
        ZSql = ZSql + "Impre27 ,"
        ZSql = ZSql + "Impre28 ,"
        ZSql = ZSql + "Impre31 ,"
        ZSql = ZSql + "Impre32 ,"
        ZSql = ZSql + "Impre33 ,"
        ZSql = ZSql + "Impre34 ,"
        ZSql = ZSql + "Impre35 ,"
        ZSql = ZSql + "Impre36 ,"
        ZSql = ZSql + "Impre37 ,"
        ZSql = ZSql + "Impre38 ,"
        ZSql = ZSql + "Impre39 ,"
        ZSql = ZSql + "Impre40 ,"
        ZSql = ZSql + "Comentario )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + DesTipo.Caption + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Numero.Text + "',"
        ZSql = ZSql + "'" + DesCentro.Caption + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Origen.Text + "',"
        ZSql = ZSql + "'" + Estado.Text + "',"
        ZSql = ZSql + "'" + IngresoNoCon.Text + "',"
        ZSql = ZSql + "'" + IngresoCausa.Text + "',"
        ZSql = ZSql + "'" + DesResponsableEmisor.Caption + "',"
        ZSql = ZSql + "'" + DesResponsableDestino.Caption + "',"
        ZSql = ZSql + "'" + Titulo.Text + "',"
        ZSql = ZSql + "'" + Referencia.Text + "',"
        ZSql = ZSql + "'" + ZZCorte + "',"
        ZSql = ZSql + "'" + ZZImpre11 + "',"
        ZSql = ZSql + "'" + ZZImpre12 + "',"
        ZSql = ZSql + "'" + ZZImpre13 + "',"
        ZSql = ZSql + "'" + ZZImpre14 + "',"
        ZSql = ZSql + "'" + ZZImpre15 + "',"
        ZSql = ZSql + "'" + ZZImpre21 + "',"
        ZSql = ZSql + "'" + ZZImpre22 + "',"
        ZSql = ZSql + "'" + ZZImpre23 + "',"
        ZSql = ZSql + "'" + ZZImpre24 + "',"
        ZSql = ZSql + "'" + ZZImpre25 + "',"
        ZSql = ZSql + "'" + ZZImpre26 + "',"
        ZSql = ZSql + "'" + ZZImpre27 + "',"
        ZSql = ZSql + "'" + ZZImpre28 + "',"
        ZSql = ZSql + "'" + ZZImpre31 + "',"
        ZSql = ZSql + "'" + ZZImpre32 + "',"
        ZSql = ZSql + "'" + ZZImpre33 + "',"
        ZSql = ZSql + "'" + ZZImpre34 + "',"
        ZSql = ZSql + "'" + ZZImpre35 + "',"
        ZSql = ZSql + "'" + ZZImpre36 + "',"
        ZSql = ZSql + "'" + ZZImpre37 + "',"
        ZSql = ZSql + "'" + ZZImpre38 + "',"
        ZSql = ZSql + "'" + ZZImpre39 + "',"
        ZSql = ZSql + "'" + ZZImpre40 + "',"
        ZSql = ZSql + "'" + Comentario.Text + "')"
        
        spImpreSac = ZSql
        Set rstImpreSac = db.OpenRecordset(spImpreSac, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    
    
    
    Listado.WindowTitle = "Impresion de Ficha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ImpreSac.Tipo, ImpreSac.DesTipo, ImpreSac.Ano, ImpreSac.Numero, ImpreSac.DesCentro, ImpreSac.Fecha, ImpreSac.Origen, ImpreSac.Estado, ImpreSac.IngresoNoCon, ImpreSac.IngresoCausa, ImpreSac.DesResponsableEmisor, ImpreSac.DesResponsableDestino, ImpreSac.Titulo, ImpreSac.Referencia, ImpreSac.Corte, ImpreSac.Impre11, ImpreSac.Impre12, ImpreSac.Impre13, ImpreSac.Impre14, ImpreSac.Impre23, ImpreSac.Impre24, ImpreSac.Impre25, ImpreSac.Impre26, ImpreSac.Impre27, ImpreSac.Impre31, ImpreSac.Impre32, ImpreSac.Impre33, ImpreSac.Impre34, ImpreSac.Impre35, ImpreSac.Impre36, ImpreSac.Impre40, ImpreSac.Comentario " _
            + "From " _
            + DSQ + ".dbo.ImpreSac ImpreSac"
            
    Rem Uno = "{Planifica.ResponsableII} in " + ZDesdeII + " to " + ZHastaII
    Rem Dos = " and {Planifica.OrdVencimiento} in " + Chr$(34) + DesdeFecha + Chr$(34) + " to " + Chr$(34) + HastaFecha + Chr$(34)
    Rem Tres = " and {Planifica.Estado} in " + ZDesdeIII + " to " + ZHastaIII
    Rem Cuatro = " and {Planifica.Responsable} in " + ZDesdeI + " to " + ZHastaI
    
    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
    Rem Listado.SelectionFormula = Uno + Dos + Tres + Cuatro
    
    Listado.Connect = Connect()
    
    Listado.Destination = 1
    Listado.Destination = 0
    
    Listado.ReportFileName = "ImpreSac.Rpt"
    Listado.Action = 1
    

End Sub



