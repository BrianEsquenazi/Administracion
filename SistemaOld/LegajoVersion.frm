VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgLegajoVersion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Version de Legajos"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11865
   LinkTopic       =   "Form2"
   ScaleHeight     =   8535
   ScaleWidth      =   11865
   Visible         =   0   'False
   Begin VB.TextBox cantiversion 
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
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   93
      Top             =   480
      Width           =   975
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
      Left            =   1440
      TabIndex        =   90
      Top             =   480
      Width           =   975
   End
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
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   59
      Top             =   120
      Width           =   4695
   End
   Begin VB.TextBox Codigo 
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9975
      _Version        =   327680
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Requerimientos"
      TabPicture(0)   =   "LegajoVersion.frx":0000
      Tab(0).ControlCount=   68
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label14"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label15"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DesSector"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label18"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "DesPerfil"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label19"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label20"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TareasI"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TareasII"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "DescriI"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "DescriII"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "DescriIII"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "DescriIV"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "DescriV"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Equivalencias"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "ObservaI"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ObservaII"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "ObservaIII"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ObservaIV"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "ObservaV"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "NecesariaI"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "NecesariaII"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "NecesariaIII"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "NecesariaIV"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "NecesariaV"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "DeseableI"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "DeseableII"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "DeseableIII"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "DeseableIV"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "DeseableV"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Fisica"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "OtrosI"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "OtrosII"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Sector"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TareasIII"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Perfil"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "EstadoI"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "EstadoII"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "EstadoIII"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "EstadoIV"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "EstadoV"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "EstadoVI"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "EstadoVII"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "EstadoVIII"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "EstadoIX"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "EstaI"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "EstaII"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "EstaIV"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "EstaIII"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "EstaVI"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "EstaV"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "EstaVIII"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "EstaVII"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "EstaIX"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "PerfilVersion"
      Tab(0).Control(67).Enabled=   0   'False
      TabCaption(1)   =   "Conocimientos para el Puesto"
      TabPicture(1)   =   "LegajoVersion.frx":001C
      Tab(1).ControlCount=   11
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "WTitulo(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "WVector1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WTexto3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "WTitulo(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "WTitulo(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "WTexto1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "WCombo1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "WTexto2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "WTitulo(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "WTitulo(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "WTitulo(5)"
      Tab(1).Control(10).Enabled=   0   'False
      Begin VB.TextBox PerfilVersion 
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
         Left            =   -72960
         MaxLength       =   4
         TabIndex        =   95
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox EstaIX 
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
         Left            =   -67320
         TabIndex        =   85
         Top             =   5160
         Width           =   1335
      End
      Begin VB.ComboBox EstaVII 
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
         Left            =   -67320
         TabIndex        =   84
         Top             =   4440
         Width           =   1335
      End
      Begin VB.ComboBox EstaVIII 
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
         Left            =   -67320
         TabIndex        =   83
         Top             =   4800
         Width           =   1335
      End
      Begin VB.ComboBox EstaV 
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
         Left            =   -67320
         TabIndex        =   82
         Top             =   3720
         Width           =   1335
      End
      Begin VB.ComboBox EstaVI 
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
         Left            =   -67320
         TabIndex        =   81
         Top             =   4080
         Width           =   1335
      End
      Begin VB.ComboBox EstaIII 
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
         Left            =   -67320
         TabIndex        =   80
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox EstaIV 
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
         Left            =   -67320
         TabIndex        =   79
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox EstaII 
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
         Left            =   -67320
         TabIndex        =   78
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox EstaI 
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
         Left            =   -67320
         TabIndex        =   77
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   5
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox EstadoIX 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   74
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox EstadoVIII 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   73
         Top             =   4800
         Width           =   2415
      End
      Begin VB.TextBox EstadoVII 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   72
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox EstadoVI 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   71
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox EstadoV 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   70
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox EstadoIV 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   69
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox EstadoIII 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   68
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox EstadoII 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   67
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox EstadoI 
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
         Left            =   -65880
         MaxLength       =   50
         TabIndex        =   66
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Perfil 
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
         Left            =   -73920
         MaxLength       =   4
         TabIndex        =   63
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TareasIII 
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
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   62
         Top             =   1560
         Width           =   9855
      End
      Begin VB.TextBox Sector 
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
         Left            =   -68160
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   57
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox OtrosII 
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
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   56
         Top             =   5160
         Width           =   6135
      End
      Begin VB.TextBox OtrosI 
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
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   55
         Top             =   4800
         Width           =   6135
      End
      Begin VB.TextBox Fisica 
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
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   53
         Top             =   4440
         Width           =   6135
      End
      Begin VB.CheckBox DeseableV 
         Enabled         =   0   'False
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
         Left            =   -70440
         TabIndex        =   51
         Top             =   3720
         Width           =   375
      End
      Begin VB.CheckBox DeseableIV 
         Enabled         =   0   'False
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
         Left            =   -70440
         TabIndex        =   50
         Top             =   3360
         Width           =   375
      End
      Begin VB.CheckBox DeseableIII 
         Enabled         =   0   'False
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
         Left            =   -70440
         TabIndex        =   49
         Top             =   3000
         Width           =   375
      End
      Begin VB.CheckBox DeseableII 
         Enabled         =   0   'False
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
         Left            =   -70440
         TabIndex        =   48
         Top             =   2640
         Width           =   375
      End
      Begin VB.CheckBox DeseableI 
         Enabled         =   0   'False
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
         Left            =   -70440
         TabIndex        =   47
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox NecesariaV 
         Enabled         =   0   'False
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
         TabIndex        =   46
         Top             =   3720
         Width           =   495
      End
      Begin VB.CheckBox NecesariaIV 
         Enabled         =   0   'False
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
         TabIndex        =   45
         Top             =   3360
         Width           =   495
      End
      Begin VB.CheckBox NecesariaIII 
         Enabled         =   0   'False
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
         TabIndex        =   44
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox NecesariaII 
         Enabled         =   0   'False
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
         TabIndex        =   43
         Top             =   2640
         Width           =   495
      End
      Begin VB.CheckBox NecesariaI 
         Enabled         =   0   'False
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
         TabIndex        =   42
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox ObservaV 
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
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   41
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox ObservaIV 
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
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox ObservaIII 
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
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   39
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox ObservaII 
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
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox ObservaI 
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
         Left            =   -70080
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   37
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Equivalencias 
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
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   36
         Top             =   4080
         Width           =   6135
      End
      Begin VB.TextBox DescriV 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox DescriIV 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox DescriIII 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox DescriII 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox DescriI 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox TareasII 
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
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1200
         Width           =   9855
      End
      Begin VB.TextBox TareasI 
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
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   18
         Top             =   840
         Width           =   9855
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   2
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1860
         Width           =   375
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
         Left            =   720
         TabIndex        =   11
         Top             =   1260
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         Top             =   1140
         Visible         =   0   'False
         Width           =   390
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
         Left            =   1320
         TabIndex        =   9
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   3
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   4
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1860
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   1260
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
         Height          =   4455
         Left            =   240
         TabIndex        =   15
         Top             =   660
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7858
         _Version        =   393216
         BackColor       =   16777152
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Index           =   6
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label20 
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
         Left            =   -67320
         TabIndex        =   86
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Observaciones"
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
         Left            =   -65760
         TabIndex        =   76
         Top             =   1620
         Width           =   1935
      End
      Begin VB.Label DesPerfil 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -72360
         TabIndex        =   65
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label18 
         Caption         =   "Perfil"
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
         TabIndex        =   64
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label DesSector 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -67200
         TabIndex        =   58
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label16 
         Caption         =   "Otros"
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
         Left            =   -74880
         TabIndex        =   54
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Cond.Fisica"
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
         Left            =   -74880
         TabIndex        =   52
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Equivalencias"
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
         Left            =   -74880
         TabIndex        =   30
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Experienecia"
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
         Left            =   -74880
         TabIndex        =   29
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Idioma"
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
         Left            =   -74880
         TabIndex        =   28
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Terciaria"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Secundaria"
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
         Left            =   -74880
         TabIndex        =   26
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Primaria"
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
         Left            =   -74880
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Des."
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
         TabIndex        =   24
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Nec."
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
         Left            =   -71040
         TabIndex        =   23
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Observaciones"
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
         Left            =   -70080
         TabIndex        =   22
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Especialidad o Equivalencia"
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
         Left            =   -73560
         TabIndex        =   21
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "EDUCACION"
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tareas"
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
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Sector"
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
         Left            =   -69000
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
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
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
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
      Height          =   1500
      ItemData        =   "LegajoVersion.frx":0038
      Left            =   120
      List            =   "LegajoVersion.frx":003F
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox FIngreso 
      Height          =   300
      Left            =   8760
      TabIndex        =   60
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
   Begin MSMask.MaskEdBox FechaVersionI 
      Height          =   300
      Left            =   8760
      TabIndex        =   88
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
   Begin MSMask.MaskEdBox FechaVersionII 
      Height          =   300
      Left            =   10200
      TabIndex        =   91
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
   Begin VB.Label Label23 
      Caption         =   "Cant. de versiones"
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
      Left            =   2640
      TabIndex        =   94
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label22 
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
      Left            =   240
      TabIndex        =   92
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label21 
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
      Left            =   7440
      TabIndex        =   89
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "F. Ingreso"
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
      Left            =   7440
      TabIndex        =   61
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9360
      MouseIcon       =   "LegajoVersion.frx":004D
      MousePointer    =   99  'Custom
      Picture         =   "LegajoVersion.frx":0357
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "LegajoVersion.frx":0B99
      MousePointer    =   99  'Custom
      Picture         =   "LegajoVersion.frx":0EA3
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   7920
      MouseIcon       =   "LegajoVersion.frx":16E5
      MousePointer    =   99  'Custom
      Picture         =   "LegajoVersion.frx":19EF
      ToolTipText     =   "Consulta de Datos"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   8640
      MouseIcon       =   "LegajoVersion.frx":2231
      MousePointer    =   99  'Custom
      Picture         =   "LegajoVersion.frx":253B
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Legajo"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgLegajoVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLegajoVersion As Recordset
Dim spLegajoVersion As String
Dim rstLegajo As Recordset
Dim spLegajo As String
Dim rstTarea As Recordset
Dim spTarea As String
Dim rstSector As Recordset
Dim spSector As String
Dim rstCurso As Recordset
Dim spCurso As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim Renglon As Integer
Dim WVector(100, 30) As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Consulta_Click()
     Opcion.Clear
     Opcion.AddItem "Legajo"
     Opcion.Visible = True
End Sub

Private Sub FechaVesionII_Change()

End Sub





Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Legajo"
            ZSql = ZSql + " Where Renglon = 1"
            ZSql = ZSql + " Order by Codigo"
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLegajo.RecordCount > 0 Then
                With rstLegajo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstLegajo!Codigo) + " " + rstLegajo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstLegajo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLegajo.Close
            End If
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()
    Call Limpia_Click
    PrgLegajoVersion.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Codigo.Text = ""
    Descripcion.Text = ""
    FechaVersionI.Text = "  /  /    "
    FechaVersionII.Text = "  /  /    "
    Version.Text = ""
    FIngreso.Text = "  /  /    "
    Perfil.Text = ""
    DesPerfil.Caption = ""
    Sector.Text = ""
    DesSector.Caption = ""
    TareasI.Text = ""
    TareasII.Text = ""
    TareasIII.Text = ""
    DescriI.Text = ""
    DescriII.Text = ""
    DescriIII.Text = ""
    DescriIV.Text = ""
    DescriV.Text = ""
    ObservaI.Text = ""
    ObservaII.Text = ""
    ObservaIII.Text = ""
    ObservaIV.Text = ""
    ObservaV.Text = ""
    NecesariaI.Value = 0
    NecesariaII.Value = 0
    NecesariaIII.Value = 0
    NecesariaIV.Value = 0
    NecesariaV.Value = 0
    DeseableI.Value = 0
    DeseableII.Value = 0
    DeseableIII.Value = 0
    DeseableIV.Value = 0
    DeseableV.Value = 0
    Equivalencias.Text = ""
    Fisica.Text = ""
    OtrosI.Text = ""
    OtrosII.Text = ""
    
    EstadoI.Text = ""
    EstadoII.Text = ""
    EstadoIII.Text = ""
    EstadoIV.Text = ""
    EstadoV.Text = ""
    EstadoVI.Text = ""
    EstadoVII.Text = ""
    EstadoVIII.Text = ""
    EstadoIX.Text = ""
    
    EstaI.ListIndex = 0
    EstaII.ListIndex = 0
    EstaIII.ListIndex = 0
    EstaIV.ListIndex = 0
    EstaV.ListIndex = 0
    EstaVI.ListIndex = 0
    EstaVII.ListIndex = 0
    EstaVIII.ListIndex = 0
    EstaIX.ListIndex = 0
    
    Renglon = 0
    
    Tablas.Tab = 0
    
    WVector1.Col = 5
    WVector1.Row = 1
    
    Codigo.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    EstaI.Clear
    
    EstaI.AddItem ""
    EstaI.AddItem "Exede"
    EstaI.AddItem "Cumple"
    EstaI.AddItem "Reforzar"
    EstaI.AddItem "En Entren."
    EstaI.AddItem "No Cumple"
    EstaI.AddItem "No Aplica"
    EstaI.AddItem "No Evalua"
    EstaI.AddItem "Cumple Act"
    
    EstaI.ListIndex = 0
    
    EstaII.Clear
    
    EstaII.AddItem ""
    EstaII.AddItem "Exede"
    EstaII.AddItem "Cumple"
    EstaII.AddItem "Reforzar"
    EstaII.AddItem "En Entren."
    EstaII.AddItem "No Cumple"
    EstaII.AddItem "No Aplica"
    EstaII.AddItem "No Evalua"
    EstaII.AddItem "Cumple Act"
    
    EstaII.ListIndex = 0
    
    EstaIII.Clear
    
    EstaIII.AddItem ""
    EstaIII.AddItem "Exede"
    EstaIII.AddItem "Cumple"
    EstaIII.AddItem "Reforzar"
    EstaIII.AddItem "En Entren."
    EstaIII.AddItem "No Cumple"
    EstaIII.AddItem "No Aplica"
    EstaIII.AddItem "No Evalua"
    EstaIII.AddItem "Cumple Act"
    
    EstaIII.ListIndex = 0
    
    EstaIV.Clear
    
    EstaIV.AddItem ""
    EstaIV.AddItem "Exede"
    EstaIV.AddItem "Cumple"
    EstaIV.AddItem "Reforzar"
    EstaIV.AddItem "En Entren."
    EstaIV.AddItem "No Cumple"
    EstaIV.AddItem "No Aplica"
    EstaIV.AddItem "No Evalua"
    EstaIV.AddItem "Cumple Act"
    
    EstaIV.ListIndex = 0
    
    EstaV.Clear
    
    EstaV.AddItem ""
    EstaV.AddItem "Exede"
    EstaV.AddItem "Cumple"
    EstaV.AddItem "Reforzar"
    EstaV.AddItem "En Entren."
    EstaV.AddItem "No Cumple"
    EstaV.AddItem "No Aplica"
    EstaV.AddItem "No Evalua"
    EstaV.AddItem "Cumple Act"
    
    EstaV.ListIndex = 0
    
    EstaVI.Clear
    
    EstaVI.AddItem ""
    EstaVI.AddItem "Exede"
    EstaVI.AddItem "Cumple"
    EstaVI.AddItem "Reforzar"
    EstaVI.AddItem "En Entren."
    EstaVI.AddItem "No Cumple"
    EstaVI.AddItem "No Aplica"
    EstaVI.AddItem "No Evalua"
    EstaVI.AddItem "Cumple Act"
    
    EstaVI.ListIndex = 0
    
    EstaVII.Clear
    
    EstaVII.AddItem ""
    EstaVII.AddItem "Exede"
    EstaVII.AddItem "Cumple"
    EstaVII.AddItem "Reforzar"
    EstaVII.AddItem "En Entren."
    EstaVII.AddItem "No Cumple"
    EstaVII.AddItem "No Aplica"
    EstaVII.AddItem "No Evalua"
    EstaVII.AddItem "Cumple Act"
    
    EstaVII.ListIndex = 0
    
    EstaVIII.Clear
    
    EstaVIII.AddItem ""
    EstaVIII.AddItem "Exede"
    EstaVIII.AddItem "Cumple"
    EstaVIII.AddItem "Reforzar"
    EstaVIII.AddItem "En Entren."
    EstaVIII.AddItem "No Cumple"
    EstaVIII.AddItem "No Aplica"
    EstaVIII.AddItem "No Evalua"
    EstaVIII.AddItem "Cumple Act"
    
    EstaVIII.ListIndex = 0
    
    EstaIX.Clear
    
    EstaIX.AddItem ""
    EstaIX.AddItem "Exede"
    EstaIX.AddItem "Cumple"
    EstaIX.AddItem "Reforzar"
    EstaIX.AddItem "En Entren."
    EstaIX.AddItem "No Cumple"
    EstaIX.AddItem "No Aplica"
    EstaIX.AddItem "No Evalua"
    EstaIX.AddItem "Cumple Act"
    
    EstaIX.ListIndex = 0
    
    Codigo.Text = ""
    Descripcion.Text = ""
    FechaVersionI.Text = "  /  /    "
    FechaVersionII.Text = "  /  /    "
    Version.Text = ""
    FIngreso.Text = "  /  /    "
    Perfil.Text = ""
    DesPerfil.Caption = ""
    Sector.Text = ""
    DesSector.Caption = ""
    TareasI.Text = ""
    TareasII.Text = ""
    TareasIII.Text = ""
    DescriI.Text = ""
    DescriII.Text = ""
    DescriIII.Text = ""
    DescriIV.Text = ""
    DescriV.Text = ""
    ObservaI.Text = ""
    ObservaII.Text = ""
    ObservaIII.Text = ""
    ObservaIV.Text = ""
    ObservaV.Text = ""
    NecesariaI.Value = 0
    NecesariaII.Value = 0
    NecesariaIII.Value = 0
    NecesariaIV.Value = 0
    NecesariaV.Value = 0
    DeseableI.Value = 0
    DeseableII.Value = 0
    DeseableIII.Value = 0
    DeseableIV.Value = 0
    DeseableV.Value = 0
    Equivalencias.Text = ""
    Fisica.Text = ""
    OtrosI.Text = ""
    OtrosII.Text = ""
    
    EstadoI.Text = ""
    EstadoII.Text = ""
    EstadoIII.Text = ""
    EstadoIV.Text = ""
    EstadoV.Text = ""
    EstadoVI.Text = ""
    EstadoVII.Text = ""
    EstadoVIII.Text = ""
    EstadoIX.Text = ""
    
    Renglon = 0
    
    WVector1.Col = 5
    WVector1.Row = 1
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    WRenglon2 = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM LegajoVersion"
    ZSql = ZSql + " Where LegajoVersion.Codigo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " and LegajoVersion.version = " + "'" + Version.Text + "'"
    ZSql = ZSql + " Order by Clave"
    spLegajoVersion = ZSql
    Set rstLegajoVersion = db.OpenRecordset(spLegajoVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstLegajoVersion.RecordCount > 0 Then
        Perfil.Text = rstLegajoVersion!Perfil
        PerfilVersion.Text = rstLegajoVersion!PerfilVersion
        EstadoI.Text = Trim(rstLegajoVersion!EstadoI)
        EstadoII.Text = Trim(rstLegajoVersion!EstadoII)
        EstadoIII.Text = Trim(rstLegajoVersion!EstadoIII)
        EstadoIV.Text = Trim(rstLegajoVersion!EstadoIV)
        EstadoV.Text = Trim(rstLegajoVersion!EstadoV)
        EstadoVI.Text = Trim(rstLegajoVersion!EstadoVI)
        EstadoVII.Text = Trim(rstLegajoVersion!EstadoVII)
        EstadoVIII.Text = Trim(rstLegajoVersion!EstadoVIII)
        EstadoIX.Text = Trim(rstLegajoVersion!EstadoIX)
        EstaI.ListIndex = rstLegajoVersion!EstaI
        EstaII.ListIndex = rstLegajoVersion!EstaII
        EstaIII.ListIndex = rstLegajoVersion!EstaIII
        EstaIV.ListIndex = rstLegajoVersion!EstaIV
        EstaV.ListIndex = rstLegajoVersion!EstaV
        EstaVI.ListIndex = rstLegajoVersion!EstaVI
        EstaVII.ListIndex = rstLegajoVersion!EstaVII
        EstaVIII.ListIndex = rstLegajoVersion!EstaVIII
        EstaIX.ListIndex = rstLegajoVersion!EstaIX
        FechaVersionI.Text = IIf(IsNull(rstLegajoVersion!FechaVersionI), "  /  /    ", rstLegajoVersion!FechaVersionI)
        FechaVersionII.Text = IIf(IsNull(rstLegajoVersion!FechaVersionII), "  /  /    ", rstLegajoVersion!FechaVersionII)
        rstLegajoVersion.Close
    End If
    
    Rem HERNAN
    
    ZVersion = 0
    Sql1 = "Select *"
    Sql2 = " FROM Tarea"
    Sql3 = " Where Tarea.Codigo = " + "'" + Perfil.Text + "'"
    spTarea = Sql1 + Sql2 + Sql3
    Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
    If rstTarea.RecordCount > 0 Then
        ZVersion = rstTarea!Version
        rstTarea.Close
    End If
    
    If ZVersion = 0 Or ZVersion = Val(PerfilVersion.Text) Then
    
        ZSql = ""
        ZSql = ZSql + "Select *, Curso.Descripcion as [WDesCurso], Sector.Descripcion as [WDesSector]"
        ZSql = ZSql + " FROM Tarea, Curso, Sector"
        ZSql = ZSql + " Where Tarea.Codigo = " + "'" + Perfil.Text + "'"
        ZSql = ZSql + " and Tarea.Curso = Curso.Codigo"
        ZSql = ZSql + " And Tarea.Sector = Sector.Codigo"
        ZSql = ZSql + " Order by Tarea.curso"
        Rem    ZSql = ZSql + " Order by Tarea.Clave"
    
        spTarea = ZSql
        Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
        If rstTarea.RecordCount > 0 Then
            With rstTarea
                .MoveFirst
                Do
                    If .EOF = False Then
                        WRenglon = WRenglon + 1
                        WVector1.Row = WRenglon
                        Renglon = WRenglon
                    
                        If Renglon = 1 Then
                            DesPerfil.Caption = Trim(rstTarea!Descripcion)
                            Sector.Text = rstTarea!Sector
                            DesSector.Caption = rstTarea!WDesSector
                            TareasI.Text = Trim(rstTarea!TareasI)
                            TareasII.Text = Trim(rstTarea!TareasII)
                            TareasIII.Text = Trim(rstTarea!TareasIII)
                            DescriI.Text = Trim(rstTarea!DescriI)
                            DescriII.Text = Trim(rstTarea!DescriII)
                            DescriIII.Text = Trim(rstTarea!DescriIII)
                            DescriIV.Text = Trim(rstTarea!DescriIV)
                            DescriV.Text = Trim(rstTarea!DescriV)
                            ObservaI.Text = Trim(rstTarea!ObservaI)
                            ObservaII.Text = Trim(rstTarea!ObservaII)
                            ObservaIII.Text = Trim(rstTarea!ObservaIII)
                            ObservaIV.Text = Trim(rstTarea!ObservaIV)
                            ObservaV.Text = Trim(rstTarea!ObservaV)
                            NecesariaI.Value = rstTarea!NecesariaI
                            NecesariaII.Value = rstTarea!NecesariaII
                            NecesariaIII.Value = rstTarea!NecesariaIII
                            NecesariaIV.Value = rstTarea!NecesariaIV
                            NecesariaV.Value = rstTarea!NecesariaV
                            DeseableI.Value = rstTarea!DeseableI
                            DeseableII.Value = rstTarea!DeseableII
                            DeseableIII.Value = rstTarea!DeseableIII
                            DeseableIV.Value = rstTarea!DeseableIV
                            DeseableV.Value = rstTarea!DeseableV
                            Equivalencias.Text = Trim(rstTarea!Equivalencias)
                            Fisica.Text = Trim(rstTarea!Fisica)
                            OtrosI.Text = Trim(rstTarea!OtrosI)
                            OtrosII.Text = Trim(rstTarea!OtrosII)
                        End If
                    
                        WVector1.Col = 1
                        WVector1.Text = rstTarea!Curso
            
                        WVector1.Col = 2
                        WVector1.Text = rstTarea!WDesCurso
                    
                        WVector1.Col = 3
                        WVector1.Text = Trim(rstTarea!NecesariaCurso)
                    
                        WVector1.Col = 4
                        WVector1.Text = Trim(rstTarea!DeseableCurso)
                    
                        WVector1.Col = 5
                        WVector1.Text = ""
                    
                        WVector1.Col = 6
                        WVector1.Text = ""
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTarea.Close
        End If
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "Select *, Curso.Descripcion as [WDesCurso], Sector.Descripcion as [WDesSector]"
        ZSql = ZSql + " FROM TareaVersion, Curso, Sector"
        ZSql = ZSql + " Where TareaVersion.Codigo = " + "'" + Perfil.Text + "'"
        ZSql = ZSql + " and TareaVersion.Version = " + "'" + PerfilVersion.Text + "'"
        ZSql = ZSql + " and TareaVersion.Curso = Curso.Codigo"
        ZSql = ZSql + " And TareaVersion.Sector = Sector.Codigo"
        ZSql = ZSql + " Order by TareaVersion.Curso"
        Rem    ZSql = ZSql + " Order by Tarea.Clave"
    
        spTareaVersion = ZSql
        Set rstTareaVersion = db.OpenRecordset(spTareaVersion, dbOpenSnapshot, dbSQLPassThrough)
        If rstTareaVersion.RecordCount > 0 Then
            With rstTareaVersion
                .MoveFirst
                Do
                    If .EOF = False Then
                        WRenglon = WRenglon + 1
                        WVector1.Row = WRenglon
                        Renglon = WRenglon
                        If Renglon = 1 Then
                            DesPerfil.Caption = Trim(rstTareaVersion!Descripcion)
                            Sector.Text = rstTareaVersion!Sector
                            DesSector.Caption = rstTareaVersion!WDesSector
                            TareasI.Text = Trim(rstTareaVersion!TareasI)
                            TareasII.Text = Trim(rstTareaVersion!TareasII)
                            TareasIII.Text = Trim(rstTareaVersion!TareasIII)
                            DescriI.Text = Trim(rstTareaVersion!DescriI)
                            DescriII.Text = Trim(rstTareaVersion!DescriII)
                            DescriIII.Text = Trim(rstTareaVersion!DescriIII)
                            DescriIV.Text = Trim(rstTareaVersion!DescriIV)
                            DescriV.Text = Trim(rstTareaVersion!DescriV)
                            ObservaI.Text = Trim(rstTareaVersion!ObservaI)
                            ObservaII.Text = Trim(rstTareaVersion!ObservaII)
                            ObservaIII.Text = Trim(rstTareaVersion!ObservaIII)
                            ObservaIV.Text = Trim(rstTareaVersion!ObservaIV)
                            ObservaV.Text = Trim(rstTareaVersion!ObservaV)
                            NecesariaI.Value = rstTareaVersion!NecesariaI
                            NecesariaII.Value = rstTareaVersion!NecesariaII
                            NecesariaIII.Value = rstTareaVersion!NecesariaIII
                            NecesariaIV.Value = rstTareaVersion!NecesariaIV
                            NecesariaV.Value = rstTareaVersion!NecesariaV
                            DeseableI.Value = rstTareaVersion!DeseableI
                            DeseableII.Value = rstTareaVersion!DeseableII
                            DeseableIII.Value = rstTareaVersion!DeseableIII
                            DeseableIV.Value = rstTareaVersion!DeseableIV
                            DeseableV.Value = rstTareaVersion!DeseableV
                            Equivalencias.Text = Trim(rstTareaVersion!Equivalencias)
                            Fisica.Text = Trim(rstTareaVersion!Fisica)
                            OtrosI.Text = Trim(rstTareaVersion!OtrosI)
                            OtrosII.Text = Trim(rstTareaVersion!OtrosII)
                        End If
                    
                        WVector1.Col = 1
                        WVector1.Text = rstTareaVersion!Curso
            
                        WVector1.Col = 2
                        WVector1.Text = rstTareaVersion!WDesCurso
                    
                        WVector1.Col = 3
                        WVector1.Text = Trim(rstTareaVersion!NecesariaCurso)
                    
                        WVector1.Col = 4
                        WVector1.Text = Trim(rstTareaVersion!DeseableCurso)
                    
                        WVector1.Col = 5
                        WVector1.Text = ""
                    
                        WVector1.Col = 6
                        WVector1.Text = ""
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTareaVersion.Close
        End If
            
    End If
    
    For Ciclo = 1 To WRenglon
    
        ZCurso = WVector1.TextMatrix(Ciclo, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM LegajoVersion"
        ZSql = ZSql + " Where LegajoVersion.Codigo = " + "'" + Codigo.Text + "'"
        ZSql = ZSql + " and LegajoVersion.Version = " + "'" + Version.Text + "'"
        ZSql = ZSql + " and LegajoVersion.Curso = " + "'" + ZCurso + "'"
        spLegajoVersion = ZSql
        Set rstLegajoVersion = db.OpenRecordset(spLegajoVersion, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajoVersion.RecordCount > 0 Then
        
            ZEstaCurso = Str$(rstLegajoVersion!EstaCurso)
            
            Select Case Val(ZEstaCurso)
                Case 1
                    WVector1.TextMatrix(Ciclo, 5) = "Exede"
                Case 2
                    WVector1.TextMatrix(Ciclo, 5) = "Cumple"
                Case 3
                    WVector1.TextMatrix(Ciclo, 5) = "Reforzar"
                Case 4
                    WVector1.TextMatrix(Ciclo, 5) = "En entren."
                Case 5
                    WVector1.TextMatrix(Ciclo, 5) = "No Cumple"
                Case Else
                    WVector1.TextMatrix(Ciclo, 5) = ""
            End Select
            WVector1.TextMatrix(Ciclo, 6) = Trim(rstLegajoVersion!EstadoCurso)
            rstLegajoVersion.Close
        End If
        
    Next Ciclo
    
    Tablas.Tab = 0
    Descripcion.SetFocus

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        cantiversion = ""
        Sql1 = "Select *"
        Sql2 = " FROM Legajo"
        Sql3 = " Where Legajo.Codigo = " + "'" + Codigo.Text + "'"
        spLegajo = Sql1 + Sql2 + Sql3
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
            Descripcion.Text = rstLegajo!Descripcion
            FIngreso.Text = rstLegajo!FIngreso
           cantiversion.Text = rstLegajo!Version
            rstLegajo.Close
            Version.SetFocus
        
                
                Else
            WCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = WCodigo
            Codigo.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Codigo.Text = ""
        Descripcion.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Version_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM LegajoVersion"
        ZSql = ZSql + " Where LegajoVersion.Codigo = " + "'" + Codigo.Text + "'"
        ZSql = ZSql + " and LegajoVersion.version = " + "'" + Version.Text + "'"
        spLegajoVersion = ZSql
        Set rstLegajoVersion = db.OpenRecordset(spLegajoVersion, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajoVersion.RecordCount > 0 Then
            rstLegajoVersion.Close
            Call Proceso_Click
        End If
    End If
    
    If KeyAscii = 27 Then
        Version.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Legajo"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLegajo.RecordCount > 0 Then
                With rstLegajo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstLegajo!Renglon = 1 Then
                                da = Len(rstLegajo!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstLegajo!Descripcion), aa, WEspacios) Then
                                        IngresaItem = Str$(rstLegajo!Codigo) + " " + rstLegajo!Descripcion
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstLegajo!Codigo
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next aa
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLegajo.Close
            End If
            
            
        Case Else
    End Select
            
    End If

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
            WCombo1.AddItem ""
            WCombo1.AddItem "Exede"
            WCombo1.AddItem "Cumple"
            WCombo1.AddItem "Reforzar"
            WCombo1.AddItem "En entren."
            WCombo1.AddItem "No Cumple"
            WCombo1.AddItem "No Aplica"
            WCombo1.AddItem "No Evalua"
            WCombo1.AddItem "Cumple Act"
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
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 5
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
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If WAuxi1 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
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
    WVector1.Cols = 7
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
                WVector1.Text = "Tema"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Necesaria"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Deseable"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Estado"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 1
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
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
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
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
    
    WVector1.Col = 5
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    Select Case Tablas.Tab
        Case 0
            Sector.SetFocus
        Case 1
            WVector1.Col = 5
            WVector1.Row = 1
            Call StartEdit
        Case Else
    End Select
End Sub

