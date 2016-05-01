VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgLegajo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Legajos"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11865
   LinkTopic       =   "Form2"
   ScaleHeight     =   8535
   ScaleWidth      =   11865
   Visible         =   0   'False
   Begin VB.CommandButton Observaciones 
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
      Height          =   495
      Left            =   7200
      TabIndex        =   116
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox Version 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   9720
      TabIndex        =   100
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Descripcion 
      BackColor       =   &H00FFFFFF&
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
      Width           =   4455
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
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9975
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Requerimientos"
      TabPicture(0)   =   "legajo.frx":0000
      Tab(0).ControlCount=   69
      Tab(0).ControlEnabled=   -1  'True
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
      Tab(0).Control(67)=   "IngresaObservacionesII"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "PerfilVersion"
      Tab(0).Control(68).Enabled=   0   'False
      TabCaption(1)   =   "Conocimientos para el Puesto"
      TabPicture(1)   =   "legajo.frx":001C
      Tab(1).ControlCount=   12
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WVector1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "WTexto3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WTitulo(4)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "WTitulo(3)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "WTexto1"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "WCombo1"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "WTexto2"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "WTitulo(1)"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "WTitulo(2)"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "WTitulo(5)"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "WTitulo(6)"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "IngresaObservacionesI"
      Tab(1).Control(11).Enabled=   0   'False
      TabCaption(2)   =   "Cursos Realizados"
      TabPicture(2)   =   "legajo.frx":0038
      Tab(2).ControlCount=   11
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "WVector2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "WTexto32"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "WCombo12"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "WTexto22"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "WTexto12"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "WTitulo2(5)"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "WTitulo2(2)"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "WTitulo2(1)"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "WTitulo2(3)"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "WTitulo2(4)"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "WTitulo2(6)"
      Tab(2).Control(10).Enabled=   -1  'True
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   118
         Top             =   480
         Width           =   495
      End
      Begin VB.Frame IngresaObservacionesII 
         Height          =   2295
         Left            =   960
         TabIndex        =   110
         Top             =   1320
         Visible         =   0   'False
         Width           =   10095
         Begin VB.TextBox ObservaII2 
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
            MaxLength       =   50
            TabIndex        =   115
            Top             =   600
            Width           =   9495
         End
         Begin VB.TextBox ObservaII1 
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
            MaxLength       =   50
            TabIndex        =   114
            Top             =   240
            Width           =   9495
         End
         Begin VB.TextBox ObservaII3 
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
            MaxLength       =   50
            TabIndex        =   113
            Top             =   960
            Width           =   9495
         End
         Begin VB.TextBox ObservaII4 
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
            MaxLength       =   50
            TabIndex        =   112
            Top             =   1320
            Width           =   9495
         End
         Begin VB.TextBox ObservaII5 
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
            MaxLength       =   50
            TabIndex        =   111
            Top             =   1680
            Width           =   9495
         End
      End
      Begin VB.Frame IngresaObservacionesI 
         Height          =   2295
         Left            =   -74520
         TabIndex        =   104
         Top             =   1800
         Visible         =   0   'False
         Width           =   10095
         Begin VB.TextBox ObservaI5 
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
            MaxLength       =   50
            TabIndex        =   109
            Top             =   1680
            Width           =   9495
         End
         Begin VB.TextBox ObservaI4 
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
            MaxLength       =   50
            TabIndex        =   108
            Top             =   1320
            Width           =   9495
         End
         Begin VB.TextBox ObservaI3 
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
            MaxLength       =   50
            TabIndex        =   107
            Top             =   960
            Width           =   9495
         End
         Begin VB.TextBox ObservaI1 
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
            MaxLength       =   50
            TabIndex        =   106
            Top             =   240
            Width           =   9495
         End
         Begin VB.TextBox ObservaI2 
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
            MaxLength       =   50
            TabIndex        =   105
            Top             =   600
            Width           =   9495
         End
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
         Left            =   -71280
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox WTitulo2 
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
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   2280
         Width           =   375
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
         Left            =   7680
         TabIndex        =   95
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
         Left            =   7680
         TabIndex        =   94
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
         Left            =   7680
         TabIndex        =   93
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
         Left            =   7680
         TabIndex        =   92
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
         Left            =   7680
         TabIndex        =   91
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
         Left            =   7680
         TabIndex        =   90
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
         Left            =   7680
         TabIndex        =   89
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
         Left            =   7680
         TabIndex        =   88
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
         Left            =   7680
         TabIndex        =   87
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox WTitulo2 
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
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox WTitulo2 
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
         Left            =   -72120
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox WTitulo2 
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox WTitulo2 
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox WTitulo2 
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
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   2280
         Width           =   375
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
         TabIndex        =   80
         Top             =   1680
         Width           =   375
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
         Left            =   -72720
         TabIndex        =   79
         Top             =   1680
         Width           =   375
      End
      Begin VB.ComboBox WCombo12 
         Height          =   315
         Left            =   -71160
         TabIndex        =   77
         Top             =   1680
         Visible         =   0   'False
         Width           =   390
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
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1800
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   9120
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
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   63
         Top             =   480
         Width           =   735
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
         Left            =   1200
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
         Left            =   6840
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
         Left            =   1440
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
         Left            =   1440
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
         Left            =   1440
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
         Left            =   4560
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
         Left            =   4560
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
         Left            =   4560
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
         Left            =   4560
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
         Left            =   4560
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
         Left            =   4080
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
         Left            =   4080
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
         Left            =   4080
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
         Left            =   4080
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
         Left            =   4080
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
         Left            =   4920
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
         Left            =   4920
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
         Left            =   4920
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
         Left            =   4920
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
         Left            =   4920
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
         Left            =   1440
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
         Left            =   1440
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
         Left            =   1440
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
         Left            =   1440
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
         Left            =   1440
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
         Left            =   1440
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
         Left            =   1200
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
         Left            =   1200
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1800
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
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
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
         Left            =   -74280
         TabIndex        =   11
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
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
         Left            =   -73680
         TabIndex        =   9
         Top             =   1200
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
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1800
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
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   -73080
         TabIndex        =   14
         Top             =   1200
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
         Height          =   4935
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8705
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   -72000
         TabIndex        =   78
         Top             =   1680
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
         Height          =   3855
         Left            =   -74760
         TabIndex        =   81
         Top             =   720
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6800
         _Version        =   327680
         BackColor       =   16777152
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
         Left            =   7680
         TabIndex        =   96
         Top             =   1920
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
         Left            =   9240
         TabIndex        =   76
         Top             =   1920
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
         Left            =   2640
         TabIndex        =   65
         Top             =   480
         Width           =   3255
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
         Left            =   240
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
         Left            =   7800
         TabIndex        =   58
         Top             =   480
         Width           =   3255
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   4440
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
         Left            =   3960
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
         Left            =   4920
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
         Left            =   1440
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
         Left            =   120
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
         Left            =   240
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
         Left            =   6000
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
      ItemData        =   "legajo.frx":0054
      Left            =   120
      List            =   "legajo.frx":005B
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox FIngreso 
      Height          =   300
      Left            =   8160
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
   Begin MSMask.MaskEdBox FechaVersion 
      Height          =   300
      Left            =   8280
      TabIndex        =   98
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
   Begin MSMask.MaskEdBox Fegreso 
      Height          =   300
      Left            =   10440
      TabIndex        =   102
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   327680
      BackColor       =   16777215
      HideSelection   =   0   'False
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
   Begin VB.Label AvisoActualiza 
      BackColor       =   &H000000FF&
      Caption         =   "SE ACTUALIZO LA VERSION DEL PERFIL Y AUN NO SE ACTUALIZO LA CALIFICACION"
      Height          =   255
      Left            =   120
      TabIndex        =   117
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label23 
      Caption         =   "F. Egreso"
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
      Left            =   9600
      TabIndex        =   101
      Top             =   120
      Width           =   1095
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
      Left            =   7200
      TabIndex        =   99
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "legajo.frx":0069
      MousePointer    =   99  'Custom
      Picture         =   "legajo.frx":0373
      ToolTipText     =   "Impresion "
      Top             =   6600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   9840
      MouseIcon       =   "legajo.frx":0BB5
      MousePointer    =   99  'Custom
      Picture         =   "legajo.frx":0EBF
      ToolTipText     =   "Elimina el Registro"
      Top             =   7680
      Width           =   480
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
      Left            =   7200
      TabIndex        =   61
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   10800
      MouseIcon       =   "legajo.frx":1701
      MousePointer    =   99  'Custom
      Picture         =   "legajo.frx":1A0B
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "legajo.frx":224D
      MousePointer    =   99  'Custom
      Picture         =   "legajo.frx":2557
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8040
      MouseIcon       =   "legajo.frx":2D99
      MousePointer    =   99  'Custom
      Picture         =   "legajo.frx":30A3
      ToolTipText     =   "Consulta de Datos"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "legajo.frx":38E5
      MousePointer    =   99  'Custom
      Picture         =   "legajo.frx":3BEF
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgLegajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim ZControl As String

Dim WParametros2(10, 10) As Double
Dim WFormato2(10) As String

Private Sub cmdDelete_Click()

    If Val(Codigo.Text) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM Legajo"
        Sql3 = " Where Legajo.Codigo = " + "'" + Codigo.Text + "'"
        spLegajo = Sql1 + Sql2 + Sql3
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
            rstLegajo.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            respuesta% = MsgBox(m$, 32 + 4, T$)
            If respuesta% = 6 Then
                Sql1 = "DELETE Legajo"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spLegajo = Sql1 + Sql2
                Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
                Call Limpia_Click
            End If
        End If
    
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub Command1_Click()

    WFechaVersion = "01/01/2006"
    WVersion = "1"
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Legajo SET "
    ZSql = ZSql + " FechaVersion = " + "'" + WFechaVersion + "',"
    ZSql = ZSql + " Version = " + "'" + WVersion + "'"
    spLegajo = ZSql
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
    
End Sub

Private Sub Codigo_LostFocus()
    If Val(Codigo.Text) <> Val(ZControl) Then
        Call Codigo_KeyPress(13)
    End If
End Sub

Private Sub Codigo_GotFocus()
    ZControl = Codigo.Text
End Sub


Private Sub Consulta_Click()
     Opcion.Clear
     Opcion.AddItem "Legajo"
     Opcion.AddItem "Perfiles"
     Opcion.Visible = True
End Sub

Private Sub Fegreso_KeyPress(KeyAscii As Integer)
Dim respuesta, Mess As String

Mess = " La Fecha de Egreso no puede ser menor a la de Ingreso"

If KeyAscii = 13 Then
        Call Valida_fecha1(Fegreso.Text, Auxi)
        If Fegreso < FIngreso Then
       respuesta = MsgBox(Mess, vbokeyonly, "FECHA DE EGRESO")
       End If
        If Auxi = "S" Then
            Perfil.SetFocus
                Else
            Fegreso.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fegreso.Text = "  /  /    "
    End If
End Sub

Private Sub Observaciones_Click()
    If Tablas.Tab = 0 Then
        IngresaObservacionesII.Visible = True
        ObservaII1.SetFocus
            Else
        If Tablas.Tab = 1 Then
            IngresaObservacionesI.Visible = True
            ObservaI1.SetFocus
        End If
    End If
End Sub

Private Sub ObservaI1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI2.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaI1.Text = ""
    End If
End Sub

Private Sub ObservaI2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI3.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaI2.Text = ""
    End If
End Sub

Private Sub ObservaI3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI4.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaI3.Text = ""
    End If
End Sub

Private Sub ObservaI4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI5.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaI4.Text = ""
    End If
End Sub

Private Sub ObservaI5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IngresaObservacionesI.Visible = False
    End If
    If KeyAscii = 27 Then
        ObservaI5.Text = ""
    End If
End Sub

Private Sub ObservaII1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII2.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII1.Text = ""
    End If
End Sub

Private Sub ObservaII2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII3.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII2.Text = ""
    End If
End Sub

Private Sub ObservaII3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII4.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII3.Text = ""
    End If
End Sub

Private Sub ObservaII4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII5.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII4.Text = ""
    End If
End Sub

Private Sub ObservaII5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IngresaObservacionesII.Visible = False
    End If
    If KeyAscii = 27 Then
        ObservaII5.Text = ""
    End If
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
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Tarea"
            ZSql = ZSql + " Where Renglon = 1"
            ZSql = ZSql + " Order by Codigo"
            spTarea = ZSql
            Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
            If rstTarea.RecordCount > 0 Then
                With rstTarea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstTarea!Codigo) + " " + rstTarea!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTarea!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTarea.Close
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
    PrgLegajo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Graba_Click()

    Dim Respuesta1, Mess As String
      
   Rem If Fegreso < FIngreso And Fegreso <> "00/00/0000" Then
   Rem     Mess = " La Fecha de Egreso no puede ser menor a la de Ingreso"
   Rem     Respuesta1 = MsgBox(Mess, vbokeyonly, "FECHA DE EGRESO")
   Rem     Exit Sub
   Rem  End If
    
    
    
    For IRow = 1 To 100
        ZCurso = WVector1.TextMatrix(IRow, 1)
        ZEstaCurso = WVector1.TextMatrix(IRow, 5)
        If Val(ZCurso) <> 0 And Trim(ZEstaCurso) = "" Then
            Mess = " Se deben calificar todas las tareas"
            Respuesta1 = MsgBox(Mess, vbokeyonly, "Calificacion de Tareas")
            Exit Sub
        End If
    Next IRow
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Legajo"
    Sql3 = " Where Legajo.Codigo = " + "'" + Codigo.Text + "'"
    spLegajo = Sql1 + Sql2 + Sql3
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLegajo.RecordCount > 0 Then
        rstLegajo.Close
      
        WActualiza = "N"
        T$ = "Actualizacion de Datos de Legajos"
        m$ = "Desea Actualizar la Version"
        respuesta% = MsgBox(m$, 32 + 4, T$)
        If respuesta% = 6 Then
            WActualiza = "S"
        End If
        If Fegreso = "  /  /    " Then
            Fegreso = "00/00/0000"
        End If
           
        If Fegreso <> "00/00/0000" Then
            If CDate(Fegreso) < CDate(FIngreso) Then
            Mess = " La Fecha de Egreso no puede ser menor a la de Ingreso"
               Respuesta1 = MsgBox(Mess, vbokeyonly, "FECHA DE EGRESO")
               Exit Sub
           End If
        End If
            
            Else
            
        WActualiza = "S"
    
    End If

    If WActualiza = "S" Then
    
        FechaVersion.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        Version.Text = Str$(Val(Version.Text) + 1)
        
        WRenglon = 0
        Erase WVector
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Legajo"
        ZSql = ZSql + " Where Legajo.Codigo = " + "'" + Codigo.Text + "'"
        ZSql = ZSql + " Order by Clave"
    
        spLegajo = ZSql
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
            With rstLegajo
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        WRenglon = WRenglon + 1
                    
                        WVector(WRenglon, 1) = rstLegajo!FIngreso
                        WVector(WRenglon, 2) = Str$(rstLegajo!Perfil)
                        WVector(WRenglon, 3) = Trim(rstLegajo!EstadoI)
                        WVector(WRenglon, 4) = Trim(rstLegajo!EstadoII)
                        WVector(WRenglon, 5) = Trim(rstLegajo!EstadoIII)
                        WVector(WRenglon, 6) = Trim(rstLegajo!EstadoIV)
                        WVector(WRenglon, 7) = Trim(rstLegajo!EstadoV)
                        WVector(WRenglon, 8) = Trim(rstLegajo!EstadoVI)
                        WVector(WRenglon, 9) = Trim(rstLegajo!EstadoVII)
                        WVector(WRenglon, 10) = Trim(rstLegajo!EstadoVIII)
                        WVector(WRenglon, 11) = Trim(rstLegajo!EstadoIX)
                        WVector(WRenglon, 12) = rstLegajo!EstaI
                        WVector(WRenglon, 13) = rstLegajo!EstaII
                        WVector(WRenglon, 14) = rstLegajo!EstaIII
                        WVector(WRenglon, 15) = rstLegajo!EstaIV
                        WVector(WRenglon, 16) = rstLegajo!EstaV
                        WVector(WRenglon, 17) = rstLegajo!EstaVI
                        WVector(WRenglon, 18) = rstLegajo!EstaVII
                        WVector(WRenglon, 19) = rstLegajo!EstaVIII
                        WVector(WRenglon, 20) = rstLegajo!EstaIX
                        WVector(WRenglon, 21) = rstLegajo!Curso
                        WVector(WRenglon, 22) = rstLegajo!EstaCurso
                        WVector(WRenglon, 23) = rstLegajo!ClavePerfil
                        WVector(WRenglon, 24) = rstLegajo!NecesariaCurso
                        WVector(WRenglon, 25) = rstLegajo!DeseableCurso
                        WVector(WRenglon, 26) = rstLegajo!Version
                        WVector(WRenglon, 27) = rstLegajo!FechaVersion
                        WVector(WRenglon, 28) = rstLegajo!EstadoCurso
                        WVector(WRenglon, 29) = IIf(IsNull(rstLegajo!Fegreso), "  /  /    ", rstLegajo!Fegreso)
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLegajo.Close
        End If
        
        
        For Ciclo = 1 To WRenglon
        
            ZZFIngreso = WVector(Ciclo, 1)
            ZZPerfil = WVector(Ciclo, 2)
            ZZPerfilVersion = PerfilVersion.Text
            
            ZZEstadoI = WVector(Ciclo, 3)
            ZZEstadoII = WVector(Ciclo, 4)
            ZZEstadoIII = WVector(Ciclo, 5)
            ZZEstadoIV = WVector(Ciclo, 6)
            ZZEstadoV = WVector(Ciclo, 7)
            ZZEstadoVI = WVector(Ciclo, 8)
            ZZEstadoVII = WVector(Ciclo, 9)
            ZZEstadoVIII = WVector(Ciclo, 10)
            ZZEstadoIX = WVector(Ciclo, 11)
            
            ZZEstaI = WVector(Ciclo, 12)
            ZZEstaII = WVector(Ciclo, 13)
            ZZEstaIII = WVector(Ciclo, 14)
            ZZEstaIV = WVector(Ciclo, 15)
            ZZEstaV = WVector(Ciclo, 16)
            ZZEstaVI = WVector(Ciclo, 17)
            ZZEstaVII = WVector(Ciclo, 18)
            ZZEstaVIII = WVector(Ciclo, 19)
            ZZEstaIX = WVector(Ciclo, 20)
            
            ZZCurso = WVector(Ciclo, 21)
            ZZEstaCurso = WVector(Ciclo, 22)
            ZZClavePerfil = WVector(Ciclo, 23)
            ZZNecesariaCurso = WVector(Ciclo, 24)
            ZZDeseableCurso = WVector(Ciclo, 25)
            
            ZZVersion = WVector(Ciclo, 26)
            ZZFechaVersionI = WVector(Ciclo, 27)
            ZZFechaVersionII = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZZEstadoCurso = WVector(Ciclo, 28)
            ZZFegreso = WVector(Ciclo, 29)
            Auxi1 = Codigo.Text
            Call Ceros(Auxi1, 6)
            
            Auxi2 = ZZVersion
            Call Ceros(Auxi2, 4)
                    
            Auxi = Str$(Ciclo)
            Call Ceros(Auxi, 2)
                        
            ZZClave = Auxi1 + Auxi2 + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO LegajoVersion ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "FIngreso ,"
            ZSql = ZSql + "Perfil ,"
            ZSql = ZSql + "PerfilVersion ,"
            ZSql = ZSql + "EstadoI ,"
            ZSql = ZSql + "EstadoII ,"
            ZSql = ZSql + "EstadoIII ,"
            ZSql = ZSql + "EstadoIV ,"
            ZSql = ZSql + "EstadoV ,"
            ZSql = ZSql + "EstadoVI ,"
            ZSql = ZSql + "EstadoVII ,"
            ZSql = ZSql + "EstadoVIII ,"
            ZSql = ZSql + "EstadoIX ,"
            ZSql = ZSql + "Curso ,"
            ZSql = ZSql + "EstaI ,"
            ZSql = ZSql + "EstaII ,"
            ZSql = ZSql + "EstaIII ,"
            ZSql = ZSql + "EstaIV ,"
            ZSql = ZSql + "EstaV ,"
            ZSql = ZSql + "EstaVI ,"
            ZSql = ZSql + "EstaVII ,"
            ZSql = ZSql + "EstaVIII ,"
            ZSql = ZSql + "EstaIX ,"
            ZSql = ZSql + "NecesariaCurso ,"
            ZSql = ZSql + "DeseableCurso ,"
            ZSql = ZSql + "ClavePerfil ,"
            ZSql = ZSql + "EstaCurso ,"
            ZSql = ZSql + "EstadoCurso ,"
            ZSql = ZSql + "FechaVersionI ,"
            ZSql = ZSql + "FechaVersionII,"
            ZSql = ZSql + "Fegreso )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + ZZVersion + "',"
            ZSql = ZSql + "'" + Str$(Ciclo) + "',"
            ZSql = ZSql + "'" + ZZFIngreso + "',"
            ZSql = ZSql + "'" + ZZPerfil + "',"
            ZSql = ZSql + "'" + ZZPerfilVersion + "',"
            ZSql = ZSql + "'" + ZZEstadoI + "',"
            ZSql = ZSql + "'" + ZZEstadoII + "',"
            ZSql = ZSql + "'" + ZZEstadoIII + "',"
            ZSql = ZSql + "'" + ZZEstadoIV + "',"
            ZSql = ZSql + "'" + ZZEstadoV + "',"
            ZSql = ZSql + "'" + ZZEstadoVI + "',"
            ZSql = ZSql + "'" + ZZEstadoVII + "',"
            ZSql = ZSql + "'" + ZZEstadoVIII + "',"
            ZSql = ZSql + "'" + ZZEstadoIX + "',"
            ZSql = ZSql + "'" + ZZCurso + "',"
            ZSql = ZSql + "'" + ZZEstaI + "',"
            ZSql = ZSql + "'" + ZZEstaII + "',"
            ZSql = ZSql + "'" + ZZEstaIII + "',"
            ZSql = ZSql + "'" + ZZEstaIV + "',"
            ZSql = ZSql + "'" + ZZEstaV + "',"
            ZSql = ZSql + "'" + ZZEstaVI + "',"
            ZSql = ZSql + "'" + ZZEstaVII + "',"
            ZSql = ZSql + "'" + ZZEstaVIII + "',"
            ZSql = ZSql + "'" + ZZEstaIX + "',"
            ZSql = ZSql + "'" + ZZNecesariaCurso + "',"
            ZSql = ZSql + "'" + ZZDeseableCurso + "',"
            ZSql = ZSql + "'" + ZZClavePerfil + "',"
            ZSql = ZSql + "'" + ZZEstaCurso + "',"
            ZSql = ZSql + "'" + ZZEstadoCurso + "',"
            ZSql = ZSql + "'" + ZZFechaVersionI + "',"
            ZSql = ZSql + "'" + ZZFechaVersionII + "',"
            ZSql = ZSql + "'" + ZZFegreso + "')"
             
            Rem  ZSql = ZSql + "'" + ZZFechaVersionII + "')"
            spLegajoVersion = ZSql
            Set rstLegajoVersion = db.OpenRecordset(spLegajoVersion, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Ciclo
        
    End If

    ZEstaI = Str$(EstaI.ListIndex)
    ZEstaII = Str$(EstaII.ListIndex)
    ZEstaIII = Str$(EstaIII.ListIndex)
    ZEstaIV = Str$(EstaIV.ListIndex)
    ZEstaV = Str$(EstaV.ListIndex)
    ZEstaVI = Str$(EstaVI.ListIndex)
    ZEstaVII = Str$(EstaVII.ListIndex)
    ZEstaVIII = Str$(EstaVIII.ListIndex)
    ZEstaIX = Str$(EstaIX.ListIndex)

    ZSql = ""
    ZSql = ZSql + "DELETE Legajo"
    ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
    spLegajo = ZSql
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)

    WRenglon = 0
    For IRow = 1 To 100
        
        WVector1.Row = IRow
            
        WVector1.Col = 1
        ZCurso = WVector1.Text
        
        WVector1.Col = 3
        ZNecesariaCurso = WVector1.Text
            
        WVector1.Col = 4
        ZDeseableCurso = WVector1.Text
            
        WVector1.Col = 5
        ZEstaCurso = WVector1.Text
        
        Select Case ZEstaCurso
            Case "Exede"
                ZZEstaCurso = "1"
            Case "Cumple"
                ZZEstaCurso = "2"
            Case "Reforzar"
                ZZEstaCurso = "3"
            Case "En entren."
                ZZEstaCurso = "4"
            Case "No Cumple"
                ZZEstaCurso = "5"
            Case "No Aplica"
                ZZEstaCurso = "6"
            Case "No Evalua"
                ZZEstaCurso = "7"
            Case "Cumple Act"
                ZZEstaCurso = "8"
            Case Else
                ZZEstaCurso = "0"
        End Select
        
        WVector1.Col = 6
        ZEstadoCurso = WVector1.Text
        
        ZActualizado = ""
            
        If Val(ZCurso) <> 0 Or IRow = 1 Then
                    
            Auxi1 = Codigo.Text
            Call Ceros(Auxi1, 6)
                    
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
                        
            WClave = Auxi1 + Auxi
            
            Auxi2 = Perfil.Text
            Call Ceros(Auxi2, 6)
            ZClavePerfil = Auxi2 + "01"
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Legajo ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "FIngreso ,"
            ZSql = ZSql + "Perfil ,"
            ZSql = ZSql + "EstadoI ,"
            ZSql = ZSql + "EstadoII ,"
            ZSql = ZSql + "EstadoIII ,"
            ZSql = ZSql + "EstadoIV ,"
            ZSql = ZSql + "EstadoV ,"
            ZSql = ZSql + "EstadoVI ,"
            ZSql = ZSql + "EstadoVII ,"
            ZSql = ZSql + "EstadoVIII ,"
            ZSql = ZSql + "EstadoIX ,"
            ZSql = ZSql + "Curso ,"
            ZSql = ZSql + "EstadoCurso ,"
            ZSql = ZSql + "EstaI ,"
            ZSql = ZSql + "EstaII ,"
            ZSql = ZSql + "EstaIII ,"
            ZSql = ZSql + "EstaIV ,"
            ZSql = ZSql + "EstaV ,"
            ZSql = ZSql + "EstaVI ,"
            ZSql = ZSql + "EstaVII ,"
            ZSql = ZSql + "EstaVIII ,"
            ZSql = ZSql + "EstaIX ,"
            ZSql = ZSql + "NecesariaCurso ,"
            ZSql = ZSql + "DeseableCurso ,"
            ZSql = ZSql + "ClavePerfil ,"
            ZSql = ZSql + "EstaCurso ,"
            ZSql = ZSql + "FechaVersion ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Actualizado ,"
            ZSql = ZSql + "ObservaI1 ,"
            ZSql = ZSql + "ObservaI2 ,"
            ZSql = ZSql + "ObservaI3 ,"
            ZSql = ZSql + "ObservaI4 ,"
            ZSql = ZSql + "ObservaI5 ,"
            ZSql = ZSql + "ObservaII1 ,"
            ZSql = ZSql + "ObservaII2 ,"
            ZSql = ZSql + "ObservaII3 ,"
            ZSql = ZSql + "ObservaII4 ,"
            ZSql = ZSql + "ObservaII5 ,"
            ZSql = ZSql + "Fegreso )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + FIngreso.Text + "',"
            ZSql = ZSql + "'" + Perfil.Text + "',"
            ZSql = ZSql + "'" + EstadoI.Text + "',"
            ZSql = ZSql + "'" + EstadoII.Text + "',"
            ZSql = ZSql + "'" + EstadoIII.Text + "',"
            ZSql = ZSql + "'" + EstadoIV.Text + "',"
            ZSql = ZSql + "'" + EstadoV.Text + "',"
            ZSql = ZSql + "'" + EstadoVI.Text + "',"
            ZSql = ZSql + "'" + EstadoVII.Text + "',"
            ZSql = ZSql + "'" + EstadoVIII.Text + "',"
            ZSql = ZSql + "'" + EstadoIX.Text + "',"
            ZSql = ZSql + "'" + ZCurso + "',"
            ZSql = ZSql + "'" + ZEstadoCurso + "',"
            ZSql = ZSql + "'" + ZEstaI + "',"
            ZSql = ZSql + "'" + ZEstaII + "',"
            ZSql = ZSql + "'" + ZEstaIII + "',"
            ZSql = ZSql + "'" + ZEstaIV + "',"
            ZSql = ZSql + "'" + ZEstaV + "',"
            ZSql = ZSql + "'" + ZEstaVI + "',"
            ZSql = ZSql + "'" + ZEstaVII + "',"
            ZSql = ZSql + "'" + ZEstaVIII + "',"
            ZSql = ZSql + "'" + ZEstaIX + "',"
            ZSql = ZSql + "'" + ZNecesariaCurso + "',"
            ZSql = ZSql + "'" + ZDeseableCurso + "',"
            ZSql = ZSql + "'" + ZClavePerfil + "',"
            ZSql = ZSql + "'" + ZZEstaCurso + "',"
            ZSql = ZSql + "'" + FechaVersion.Text + "',"
            ZSql = ZSql + "'" + Version.Text + "',"
            ZSql = ZSql + "'" + ZActualizado + "',"
            ZSql = ZSql + "'" + ObservaI1.Text + "',"
            ZSql = ZSql + "'" + ObservaI2.Text + "',"
            ZSql = ZSql + "'" + ObservaI3.Text + "',"
            ZSql = ZSql + "'" + ObservaI4.Text + "',"
            ZSql = ZSql + "'" + ObservaI5.Text + "',"
            ZSql = ZSql + "'" + ObservaII1.Text + "',"
            ZSql = ZSql + "'" + ObservaII2.Text + "',"
            ZSql = ZSql + "'" + ObservaII3.Text + "',"
            ZSql = ZSql + "'" + ObservaII4.Text + "',"
            ZSql = ZSql + "'" + ObservaII5.Text + "',"
            ZSql = ZSql + "'" + Fegreso.Text + "')"
            
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
            
    Next IRow
    
    Call Limpia_Click

End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector
    Call Limpia_Vector2

    Codigo.Text = ""
    Descripcion.Text = ""
    FechaVersion.Text = "  /  /    "
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
    PerfilVersion.Text = ""
    Fegreso.Text = "  /  /    "
    ObservaII1.Text = ""
    ObservaII2.Text = ""
    ObservaII3.Text = ""
    ObservaII4.Text = ""
    ObservaII5.Text = ""
    
    
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
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Legajo"
    spLegajo = Sql1 + Sql2
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLegajo.RecordCount > 0 Then
        rstLegajo.MoveLast
        WCodigoMayor = IIf(IsNull(rstLegajo!CodigoMayor), "0", rstLegajo!CodigoMayor)
        Codigo.Text = Mid$(Str$(WCodigoMayor + 1), 2, 8)
        rstLegajo.Close
            Else
        Codigo.Text = "1"
    End If
    
    Renglon = 0
    
    Tablas.Tab = 0
    
    WVector1.Col = 5
    WVector1.Row = 1
    
    AvisoActualiza.Visible = False
    
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
            
        Case 1
            Indice = Pantalla.ListIndex
            Perfil.Text = WIndice.List(Indice)
            Call Perfil_KeyPress(13)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Call Limpia_Vector2
    
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
    FechaVersion.Text = "  /  /    "
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
    PerfilVersion.Text = ""
    Fegreso.Text = "  /  /    "
    ObservaII1.Text = ""
    ObservaII2.Text = ""
    ObservaII3.Text = ""
    ObservaII4.Text = ""
    ObservaII5.Text = ""
    
    
    EstadoI.Text = ""
    EstadoII.Text = ""
    EstadoIII.Text = ""
    EstadoIV.Text = ""
    EstadoV.Text = ""
    EstadoVI.Text = ""
    EstadoVII.Text = ""
    EstadoVIII.Text = ""
    EstadoIX.Text = ""
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Legajo"
    spLegajo = Sql1 + Sql2
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLegajo.RecordCount > 0 Then
        rstLegajo.MoveLast
        WCodigoMayor = IIf(IsNull(rstLegajo!CodigoMayor), "0", rstLegajo!CodigoMayor)
        Codigo.Text = Mid$(Str$(WCodigoMayor + 1), 2, 8)
        rstLegajo.Close
            Else
        Codigo.Text = "1"
    End If
    
    AvisoActualiza.Visible = False
    
    Renglon = 0
    
    WVector1.Col = 5
    WVector1.Row = 1
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    Call Limpia_Vector2
    WRenglon = 0
    WRenglon2 = 0
    Fegreso.BackColor = vbWhite
    
    Sql1 = "Select *"
    Sql2 = " FROM Legajo"
    Sql3 = " Where Legajo.Codigo = " + "'" + Codigo.Text + "'"
    spLegajo = Sql1 + Sql2 + Sql3
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLegajo.RecordCount > 0 Then
        WFechaVersion = IIf(IsNull(rstLegajo!FechaVersion), "  /  /    ", rstLegajo!FechaVersion)
        WVersion = IIf(IsNull(rstLegajo!Version), "0", rstLegajo!Version)
        Descripcion.Text = Trim(rstLegajo!Descripcion)
        FechaVersion.Text = WFechaVersion
        Version.Text = WVersion
        FIngreso.Text = rstLegajo!FIngreso
        Perfil.Text = rstLegajo!Perfil
        EstadoI.Text = Trim(rstLegajo!EstadoI)
        EstadoII.Text = Trim(rstLegajo!EstadoII)
        EstadoIII.Text = Trim(rstLegajo!EstadoIII)
        EstadoIV.Text = Trim(rstLegajo!EstadoIV)
        EstadoV.Text = Trim(rstLegajo!EstadoV)
        EstadoVI.Text = Trim(rstLegajo!EstadoVI)
        EstadoVII.Text = Trim(rstLegajo!EstadoVII)
        EstadoVIII.Text = Trim(rstLegajo!EstadoVIII)
        EstadoIX.Text = Trim(rstLegajo!EstadoIX)
        EstaI.ListIndex = rstLegajo!EstaI
        EstaII.ListIndex = rstLegajo!EstaII
        EstaIII.ListIndex = rstLegajo!EstaIII
        EstaIV.ListIndex = rstLegajo!EstaIV
        EstaV.ListIndex = rstLegajo!EstaV
        EstaVI.ListIndex = rstLegajo!EstaVI
        EstaVII.ListIndex = rstLegajo!EstaVII
        EstaVIII.ListIndex = rstLegajo!EstaVIII
        EstaIX.ListIndex = rstLegajo!EstaIX
        If IsNull(rstLegajo!Fegreso) Then
            Fegreso.Text = "00/00/0000"
                Else
            Fegreso.Text = rstLegajo!Fegreso
            If Fegreso.Text <> "00/00/0000" Then
                Fegreso.BackColor = vbRed
            End If
        End If
        Rem Fegreso.Text = IIf(rstLegajo!Fegreso = Null, " ", rstLegajo!Fegreso)
        
        ObservaI1.Text = IIf(IsNull(rstLegajo!ObservaI1), "", rstLegajo!ObservaI1)
        ObservaI2.Text = IIf(IsNull(rstLegajo!ObservaI2), "", rstLegajo!ObservaI2)
        ObservaI3.Text = IIf(IsNull(rstLegajo!ObservaI3), "", rstLegajo!ObservaI3)
        ObservaI4.Text = IIf(IsNull(rstLegajo!ObservaI4), "", rstLegajo!ObservaI4)
        ObservaI5.Text = IIf(IsNull(rstLegajo!ObservaI5), "", rstLegajo!ObservaI5)
        
        ObservaII1.Text = IIf(IsNull(rstLegajo!ObservaII1), "", rstLegajo!ObservaII1)
        ObservaII2.Text = IIf(IsNull(rstLegajo!ObservaII2), "", rstLegajo!ObservaII2)
        ObservaII3.Text = IIf(IsNull(rstLegajo!ObservaII3), "", rstLegajo!ObservaII3)
        ObservaII4.Text = IIf(IsNull(rstLegajo!ObservaII4), "", rstLegajo!ObservaII4)
        ObservaII5.Text = IIf(IsNull(rstLegajo!ObservaII5), "", rstLegajo!ObservaII5)
        
        ZActualizado = IIf(IsNull(rstLegajo!Actualizado), "", rstLegajo!Actualizado)
        If ZActualizado = "N" Then
            AvisoActualiza.Visible = True
                Else
            AvisoActualiza.Visible = False
        End If
        
        rstLegajo.Close
    End If
    
    ObservaI1.Text = Trim(ObservaI1.Text)
    ObservaI2.Text = Trim(ObservaI2.Text)
    ObservaI3.Text = Trim(ObservaI3.Text)
    ObservaI4.Text = Trim(ObservaI4.Text)
    ObservaI5.Text = Trim(ObservaI5.Text)
    
    ObservaII1.Text = Trim(ObservaII1.Text)
    ObservaII2.Text = Trim(ObservaII2.Text)
    ObservaII3.Text = Trim(ObservaII3.Text)
    ObservaII4.Text = Trim(ObservaII4.Text)
    ObservaII5.Text = Trim(ObservaII5.Text)
    
    Rem hERNAN
    
    ZSql = ""
    ZSql = ZSql + "Select *, Curso.Descripcion as [WDesCurso], Sector.Descripcion as [WDesSector]"
    ZSql = ZSql + " FROM Tarea, Curso, Sector"
    ZSql = ZSql + " Where Tarea.Codigo = " + "'" + Perfil.Text + "'"
    ZSql = ZSql + " and Tarea.Curso = Curso.Codigo"
    ZSql = ZSql + " And Tarea.Sector = Sector.Codigo"
    ZSql = ZSql + " Order by Curso.codigo"
    Rem ZSql = ZSql + " Order by Tarea.Clave"
    
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
                        PerfilVersion.Text = rstTarea!Version
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
    
    For Ciclo = 1 To WRenglon
    
        ZCurso = WVector1.TextMatrix(Ciclo, 1)
        
        Sql1 = "Select *"
        Sql2 = " FROM Legajo"
        Sql3 = " Where Legajo.Codigo = " + "'" + Codigo.Text + "'"
        Sql5 = " and Legajo.Curso = " + "'" + ZCurso + "'"
        spLegajo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
        
            ZEstaCurso = Str$(rstLegajo!EstaCurso)
            
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
                Case 6
                    WVector1.TextMatrix(Ciclo, 5) = "No Aplica"
                Case 7
                    WVector1.TextMatrix(Ciclo, 5) = "No Evalua"
                Case 8
                    WVector1.TextMatrix(Ciclo, 5) = "Cumple Act"
                Case Else
                    WVector1.TextMatrix(Ciclo, 5) = ""
            End Select
            
            WVector1.TextMatrix(Ciclo, 6) = Trim(rstLegajo!EstadoCurso)
            rstLegajo.Close
        End If
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    Rem hernan
    
    
    ZSql = ""
    ZSql = ZSql + "Select *, Curso.Descripcion as [WDesCurso]"
    ZSql = ZSql + " FROM Cronograma, Curso"
    ZSql = ZSql + " Where Cronograma.Legajo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " and Cronograma.Curso = Curso.Codigo"
    ZSql = ZSql + " and Cronograma.Realizado <> 0"
    ZSql = ZSql + " Order by Cronograma.Curso ASC"
    
    spCronograma = ZSql
    Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)
    If rstCronograma.RecordCount > 0 Then
        With rstCronograma
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon2 = WRenglon2 + 1
                    
                    WVector2.Row = WRenglon2
                    Renglon = WRenglon2
                    
                    WVector2.Col = 1
                    WVector2.Text = rstCronograma!Curso
            
                    WVector2.Col = 2
                    WVector2.Text = rstCronograma!WDesCurso
            
                    WVector2.Col = 3
                    WVector2.Text = rstCronograma!Horas
                    WVector2.Text = Pusing("###,###.##", WVector2.Text)
            
                    WVector2.Col = 4
                    WVector2.Text = rstCronograma!Realizado
                    
                    WVector2.Col = 5
                    WVector2.Text = rstCronograma!Ano
            
                    WVector2.Col = 6
                    WVector2.Text = rstCronograma!observacionesii
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCronograma.Close
    End If
    
    Tablas.Tab = 0
    Descripcion.SetFocus

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ZControl = Codigo.Text
        Sql1 = "Select *"
        Sql2 = " FROM Legajo"
        Sql3 = " Where Legajo.Codigo = " + "'" + Codigo.Text + "'"
        spLegajo = Sql1 + Sql2 + Sql3
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
            rstLegajo.Close
            ZZCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = ZZCodigo
            Call Proceso_Click
            WVector1.Col = 5
            WVector1.Row = 1
            Call StartEdit
                Else
            WCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = WCodigo
        End If
        Descripcion.SetFocus
    End If
    
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FIngreso.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub FIngreso_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(FIngreso.Text, Auxi)
        If Auxi = "S" Then
            Perfil.SetFocus
                Else
            FIngreso.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FIngreso.Text = "  /  /    "
    End If
End Sub

Private Sub Perfil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Tarea"
        Sql3 = " Where Tarea.Codigo = " + "'" + Perfil.Text + "'"
        spTarea = Sql1 + Sql2 + Sql3
        Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
        If rstTarea.RecordCount > 0 Then
            rstTarea.Close
            
            Call Limpia_Vector
            Call Limpia_Vector2
            WRenglon = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *, Curso.Descripcion as [WDesCurso], Sector.Descripcion as [WDesSector]"
            ZSql = ZSql + " FROM Tarea, Curso, Sector"
            ZSql = ZSql + " Where Tarea.Codigo = " + "'" + Perfil.Text + "'"
            ZSql = ZSql + " and Tarea.Curso = Curso.Codigo"
            ZSql = ZSql + " And Tarea.Sector = Sector.Codigo"
            ZSql = ZSql + " Order by Tarea.Clave"
    
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
                                PerfilVersion.Text = rstTarea!Version
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
                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTarea.Close
            End If
            
            EstadoI.SetFocus
            
                Else
                
            Perfil.SetFocus
            
        End If
    End If
    
    If KeyAscii = 27 Then
        Perfil.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub EstadoI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoII.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoI.Text = ""
    End If
End Sub

Private Sub EstadoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoIII.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoII.Text = ""
    End If
End Sub

Private Sub EstadoIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoIV.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoIII.Text = ""
    End If
End Sub

Private Sub EstadoIV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoV.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoIV.Text = ""
    End If
End Sub

Private Sub EstadoV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoVI.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoV.Text = ""
    End If
End Sub

Private Sub EstadoVI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoVII.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoVI.Text = ""
    End If
End Sub

Private Sub EstadoVII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoVIII.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoVII.Text = ""
    End If
End Sub

Private Sub EstadoVIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoIX.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoVIII.Text = ""
    End If
End Sub

Private Sub EstadoIX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EstadoI.SetFocus
    End If
    If KeyAscii = 27 Then
        EstadoIX.Text = ""
    End If
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
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Tarea"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spTarea = ZSql
            Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
            If rstTarea.RecordCount > 0 Then
                With rstTarea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstTarea!Renglon = 1 Then
                                da = Len(rstTarea!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstTarea!Descripcion), aa, WEspacios) Then
                                        IngresaItem = Str$(rstTarea!Codigo) + " " + rstTarea!Descripcion
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstTarea!Codigo
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
                rstTarea.Close
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
    Rem WCombo1.Height = 3000
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
            WVector1.Col = 7
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tema"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Neces."
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Desea."
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Estado"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 1
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 2000
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


Rem
Rem Controles de la WVector2
Rem

Private Sub GridEditText2(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametros2(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametros2(1, XColumna)
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
            WTexto22.MaxLength = WParametros2(1, XColumna)
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

Private Sub EndEdit2()
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
        If WFormato2(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormato2(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo2()
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
            Call Control_Campo2
            If WControl = "S" Then
                Call Control_WVector2
            End If
            Call StartEdit2

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_Campo2
                If WControl = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_Campo2
                If WControl = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEdit2
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEdit2
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEdit2

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
            Call Control_Campo2
            If WControl = "S" Then
                Call Control_WVector2
            End If
            Call StartEdit2
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEdit2
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEdit2
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEdit2

    End Select
End Sub

Private Sub WTexto32_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call Control_Campo2
            If WControl = "S" Then
                Call Control_WVector2
            End If
            Call StartEdit2

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_Campo2
                If WControl = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_Campo2
                If WControl = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEdit2
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEdit2
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_Campo2
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEdit2

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
Private Sub WTexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEdit2
End Sub

Private Sub WVector2_LeaveCell()
    EndEdit2
End Sub

Private Sub WVector2_GotFocus()
    EndEdit2
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametros2(4, WVector2.Col)
        Case 1
        Case Else
            If WParametros2(2, XColumna) = 0 Then
                GridEditText2 KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit2()
    Select Case WParametros2(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditCombo2
        Case Else
            If WParametros2(2, WVector2.Col) = 0 Then
                GridEditText2 Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 6
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 5
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditText2 KeyAscii
End Sub

Private Sub Control_Campo2()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControl = "S"
    Select Case XColumna
        Case Else
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector2.Rows - 1
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        If WAuxi1 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector2
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For da = 1 To WVector2.Cols - 1
            WVector2.Col = da
            WVector2.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If

End Sub

Private Sub Limpia_Vector2()

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
    WVector2.Cols = 7
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
    Rem WParametros2(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros2(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros2(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros2(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Tema"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 4
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 1
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 2
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 3000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros2(1, Ciclo) = 50
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 3
                WVector2.Text = "Horas"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 1
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 4
                WVector2.Text = "Realizado"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 1
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 5
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 1
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 6
                WVector2.Text = "Observaciones"
                WVector2.ColWidth(Ciclo) = 3000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros2(1, Ciclo) = 30
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WTitulo2(Ciclo).Text = WVector2.Text
        WTitulo2(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        WTitulo2(Ciclo).Top = WVector2.CellTop + WVector2.Top
        WTitulo2(Ciclo).Width = WVector2.CellWidth
        WTitulo2(Ciclo).Height = WVector2.CellHeight
        WTitulo2(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 5
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    Select Case Tablas.Tab
        Case 0
            Sector.SetFocus
        Case 1
            WVector1.Col = 4
            WVector1.Row = 1
            Call StartEdit
            WTexto1.Visible = False
            WTexto2.Visible = False
            Codigo.SetFocus
        Case Else
    End Select
End Sub

