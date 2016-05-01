VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgTarea 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Perfiles"
   ClientHeight    =   8250
   ClientLeft      =   180
   ClientTop       =   375
   ClientWidth     =   11670
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11670
   Visible         =   0   'False
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
      Left            =   10560
      MaxLength       =   4
      TabIndex        =   63
      Top             =   120
      Width           =   855
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
      Width           =   4575
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
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9763
      _Version        =   327680
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Requerimientos"
      TabPicture(0)   =   "tarea.frx":0000
      Tab(0).ControlCount=   44
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
      Tab(0).Control(16)=   "TareasI"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TareasII"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "DescriI"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DescriII"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DescriIII"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "DescriIV"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "DescriV"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Equivalencias"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ObservaI"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "ObservaII"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "ObservaIII"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ObservaIV"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "ObservaV"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "NecesariaI"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "NecesariaII"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "NecesariaIII"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "NecesariaIV"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "NecesariaV"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "DeseableI"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "DeseableII"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "DeseableIII"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "DeseableIV"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "DeseableV"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Fisica"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "OtrosI"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "OtrosII"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Sector"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TareasIII"
      Tab(0).Control(43).Enabled=   0   'False
      TabCaption(1)   =   "Conocimientos para el Puesto"
      TabPicture(1)   =   "tarea.frx":001C
      Tab(1).ControlCount=   9
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   62
         Top             =   1560
         Width           =   9495
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
         Left            =   1560
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   56
         Top             =   5160
         Width           =   9495
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   55
         Top             =   4800
         Width           =   9495
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   53
         Top             =   4440
         Width           =   9495
      End
      Begin VB.CheckBox DeseableV 
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
         Left            =   6000
         TabIndex        =   51
         Top             =   3720
         Width           =   495
      End
      Begin VB.CheckBox DeseableIV 
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
         Left            =   6000
         TabIndex        =   50
         Top             =   3360
         Width           =   495
      End
      Begin VB.CheckBox DeseableIII 
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
         Left            =   6000
         TabIndex        =   49
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox DeseableII 
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
         Left            =   6000
         TabIndex        =   48
         Top             =   2640
         Width           =   495
      End
      Begin VB.CheckBox DeseableI 
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
         Left            =   6000
         TabIndex        =   47
         Top             =   2280
         Width           =   495
      End
      Begin VB.CheckBox NecesariaV 
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
         Left            =   5160
         TabIndex        =   46
         Top             =   3720
         Width           =   495
      End
      Begin VB.CheckBox NecesariaIV 
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
         Left            =   5160
         TabIndex        =   45
         Top             =   3360
         Width           =   495
      End
      Begin VB.CheckBox NecesariaIII 
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
         Left            =   5160
         TabIndex        =   44
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox NecesariaII 
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
         Left            =   5160
         TabIndex        =   43
         Top             =   2640
         Width           =   495
      End
      Begin VB.CheckBox NecesariaI 
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
         Left            =   5160
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
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   41
         Top             =   3720
         Width           =   4335
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
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   40
         Top             =   3360
         Width           =   4335
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
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   39
         Top             =   3000
         Width           =   4335
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
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   38
         Top             =   2640
         Width           =   4335
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
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   37
         Top             =   2280
         Width           =   4335
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   36
         Top             =   4080
         Width           =   9495
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   35
         Top             =   3720
         Width           =   3255
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   34
         Top             =   3360
         Width           =   3255
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   33
         Top             =   3000
         Width           =   3255
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2640
         Width           =   3255
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2280
         Width           =   3255
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1200
         Width           =   9495
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
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   18
         Top             =   840
         Width           =   9495
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
         Left            =   -73560
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
         Left            =   -74160
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
         Left            =   -74760
         TabIndex        =   11
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   -74760
         TabIndex        =   10
         Top             =   1800
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
         Left            =   -74160
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
         Left            =   -73080
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
         Left            =   -74640
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   -73560
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
         Height          =   4455
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7858
         _Version        =   327680
         BackColor       =   16777152
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
         Left            =   2520
         TabIndex        =   58
         Top             =   480
         Width           =   5055
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
         Left            =   240
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
         Left            =   240
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
         Left            =   240
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
         Left            =   240
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
         Left            =   240
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
         Left            =   240
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
         Left            =   240
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
         Left            =   240
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Deseable"
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
         Left            =   5760
         TabIndex        =   24
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Necesaria"
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
         Left            =   4800
         TabIndex        =   23
         Top             =   1920
         Width           =   975
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
         Left            =   6720
         TabIndex        =   22
         Top             =   1920
         Width           =   4215
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
         Left            =   1560
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
         Left            =   240
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
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1335
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
      Top             =   6120
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
      Top             =   6960
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   7320
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
      Height          =   1740
      ItemData        =   "tarea.frx":0038
      Left            =   120
      List            =   "tarea.frx":003F
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox Vigencia 
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
   Begin VB.Label Label18 
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
      Left            =   9600
      TabIndex        =   64
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   10080
      MouseIcon       =   "tarea.frx":004D
      MousePointer    =   99  'Custom
      Picture         =   "tarea.frx":0357
      ToolTipText     =   "Impresion "
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   7920
      MouseIcon       =   "tarea.frx":0B99
      MousePointer    =   99  'Custom
      Picture         =   "tarea.frx":0EA3
      ToolTipText     =   "Elimina el Registro"
      Top             =   6120
      Width           =   480
   End
   Begin VB.Label Label17 
      Caption         =   "Vigencia"
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
      Width           =   1095
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   10800
      MouseIcon       =   "tarea.frx":16E5
      MousePointer    =   99  'Custom
      Picture         =   "tarea.frx":19EF
      ToolTipText     =   "Salida"
      Top             =   6120
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7200
      MouseIcon       =   "tarea.frx":2231
      MousePointer    =   99  'Custom
      Picture         =   "tarea.frx":253B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6120
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8640
      MouseIcon       =   "tarea.frx":2D7D
      MousePointer    =   99  'Custom
      Picture         =   "tarea.frx":3087
      ToolTipText     =   "Consulta de Datos"
      Top             =   6120
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9360
      MouseIcon       =   "tarea.frx":38C9
      MousePointer    =   99  'Custom
      Picture         =   "tarea.frx":3BD3
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6120
      Width           =   480
   End
   Begin VB.Label Label1 
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
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim ZZDescripcion As String
Dim ZZSector As String
Dim ZZDesSector As String
Dim ZZTareasI As String
Dim ZZTareasII As String
Dim ZZTareasIII As String
Dim ZZDescriI As String
Dim ZZDescriII As String
Dim ZZDescriIII As String
Dim ZZDescriIV As String
Dim ZZDescriV As String
Dim ZZObservaI As String
Dim ZZObservaII As String
Dim ZZObservaIII As String
Dim ZZObservaIV As String
Dim ZZObservaV As String
Dim ZZNecesariaI As String
Dim ZZNecesariaII As String
Dim ZZNecesariaIII As String
Dim ZZNecesariaIV As String
Dim ZZNecesariaV As String
Dim ZZDeseableI As String
Dim ZZDeseableII As String
Dim ZZDeseableIII As String
Dim ZZDeseableIV As String
Dim ZZDeseableV As String
Dim ZZEquivalencias As String
Dim ZZFisica As String
Dim ZZOtrosI As String
Dim ZZOtrosII As String
Dim ZZCurso As String
Dim ZZNecesariaCurso As String
Dim ZZDeseableCurso As String

Dim ZZRenglon As Integer
Dim ZZVector(100, 2) As String

Dim WWLegajo As String
Dim WWVersion As String
Dim WWFechaVersion As String
Dim WWFIngreso As String
Dim WWPerfil As String
Dim WWPerfilVersion As String
Dim WWEstadoI As String
Dim WWEstadoII As String
Dim WWEstadoIII As String
Dim WWEstadoIV As String
Dim WWEstadoV As String
Dim WWEstadoVI As String
Dim WWEstadoVII As String
Dim WWEstadoVIII As String
Dim WWEstadoIX As String
Dim WWEstaI As String
Dim WWEstaII As String
Dim WWEstaIII As String
Dim WWEstaIV As String
Dim WWEstaV As String
Dim WWEstaVI As String
Dim WWEstaVII As String
Dim WWEstaVIII As String
Dim WWEstaIX As String
Dim WWCurso As String
Dim WWEstaCurso As String
Dim WWClavePerfil As String
Dim WWNecesariaCurso As String
Dim WWDeseableCurso As String
Dim WWFechaVersionI As String
Dim WWFechaVersionII As String
Dim WWEstadoCurso As String
Dim WWFegreso As String

Dim WWRenglon As Integer
Dim WWVector(100, 30) As String


Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub cmdDelete_Click()

    If Val(Codigo.Text) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM Tarea"
        Sql3 = " Where Tarea.Codigo = " + "'" + Codigo.Text + "'"
        spTarea = Sql1 + Sql2 + Sql3
        Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
        If rstTarea.RecordCount > 0 Then
            rstTarea.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            respuesta% = MsgBox(m$, 32 + 4, T$)
            If respuesta% = 6 Then
                Sql1 = "DELETE Tarea"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spTarea = Sql1 + Sql2
                Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
                Call Limpia_Click
            End If
        End If
    
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Perfiles"
     Opcion.AddItem "Sectores"
     Opcion.AddItem "Temas"

     Opcion.Visible = True
     
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
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Sector"
            Sql3 = " Order by Codigo"
            spSector = Sql1 + Sql2 + Sql3
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
            If rstSector.RecordCount > 0 Then
                With rstSector
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Auxi = Str$(rstSector!Codigo)
                            Call Ceros(Auxi, 4)
                            IngresaItem = Auxi + " " + rstSector!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstSector!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSector.Close
            End If
            
        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM Curso"
            Sql3 = " Order by Codigo"
            spCurso = Sql1 + Sql2 + Sql3
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                With rstCurso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstCurso!Codigo) + " " + rstCurso!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCurso!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCurso.Close
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
    PrgTarea.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Graba_Click()

    T$ = "Actualizacion de Datos de Legajos"
    m$ = "Desea Actualizar la Version"
    respuesta% = MsgBox(m$, 32 + 4, T$)
    If respuesta% = 6 Then
    
        ZDesdeVigencia = Vigencia.Text
        ZHastaVigencia = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
        WRenglon = 0
        For IRow = 1 To 100
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Tarea"
            ZSql = ZSql + " Where Tarea.Codigo = " + "'" + Codigo.Text + "'"
            ZSql = ZSql + " and Tarea.Renglon = " + "'" + Str$(IRow) + "'"
    
            spTarea = ZSql
            Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
            If rstTarea.RecordCount > 0 Then
        
                ZZDescripcion = Trim(rstTarea!Descripcion)
                ZZSector = rstTarea!Sector
                ZZTareasI = Trim(rstTarea!TareasI)
                ZZTareasII = Trim(rstTarea!TareasII)
                ZZTareasIII = Trim(rstTarea!TareasIII)
                ZZDescriI = Trim(rstTarea!DescriI)
                ZZDescriII = Trim(rstTarea!DescriII)
                ZZDescriIII = Trim(rstTarea!DescriIII)
                ZZDescriIV = Trim(rstTarea!DescriIV)
                ZZDescriV = Trim(rstTarea!DescriV)
                ZZObservaI = Trim(rstTarea!ObservaI)
                ZZObservaII = Trim(rstTarea!ObservaII)
                ZZObservaIII = Trim(rstTarea!ObservaIII)
                ZZObservaIV = Trim(rstTarea!ObservaIV)
                ZZObservaV = Trim(rstTarea!ObservaV)
                ZZNecesariaI = rstTarea!NecesariaI
                ZZNecesariaII = rstTarea!NecesariaII
                ZZNecesariaIII = rstTarea!NecesariaIII
                ZZNecesariaIV = rstTarea!NecesariaIV
                ZZNecesariaV = rstTarea!NecesariaV
                ZZDeseableI = rstTarea!DeseableI
                ZZDeseableII = rstTarea!DeseableII
                ZZDeseableIII = rstTarea!DeseableIII
                ZZDeseableIV = rstTarea!DeseableIV
                ZZDeseableV = rstTarea!DeseableV
                ZZEquivalencias = Trim(rstTarea!Equivalencias)
                ZZFisica = Trim(rstTarea!Fisica)
                ZZOtrosI = Trim(rstTarea!OtrosI)
                ZZOtrosII = Trim(rstTarea!OtrosII)
                ZZCurso = rstTarea!Curso
                ZZNecesariaCurso = Trim(rstTarea!NecesariaCurso)
                ZZDeseableCurso = Trim(rstTarea!DeseableCurso)
                        
                rstTarea.Close
                
                Auxi1 = Codigo.Text
                Call Ceros(Auxi1, 6)
                
                Auxi2 = Version.Text
                Call Ceros(Auxi2, 4)
                    
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
                        
                WClave = Auxi1 + Auxi2 + Auxi
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO TareaVersion ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Version ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "DesdeVigencia ,"
                ZSql = ZSql + "HastaVigencia ,"
                ZSql = ZSql + "Sector ,"
                ZSql = ZSql + "TareasI ,"
                ZSql = ZSql + "TareasII ,"
                ZSql = ZSql + "TareasIII ,"
                ZSql = ZSql + "DescriI ,"
                ZSql = ZSql + "DescriII ,"
                ZSql = ZSql + "DescriIII ,"
                ZSql = ZSql + "DescriIV ,"
                ZSql = ZSql + "DescriV ,"
                ZSql = ZSql + "ObservaI ,"
                ZSql = ZSql + "ObservaII ,"
                ZSql = ZSql + "ObservaIII ,"
                ZSql = ZSql + "ObservaIV ,"
                ZSql = ZSql + "ObservaV ,"
                ZSql = ZSql + "NecesariaI ,"
                ZSql = ZSql + "NecesariaII ,"
                ZSql = ZSql + "NecesariaIII ,"
                ZSql = ZSql + "NecesariaIV ,"
                ZSql = ZSql + "NecesariaV ,"
                ZSql = ZSql + "DeseableI ,"
                ZSql = ZSql + "DeseableII ,"
                ZSql = ZSql + "DeseableIII ,"
                ZSql = ZSql + "DeseableIV ,"
                ZSql = ZSql + "DeseableV ,"
                ZSql = ZSql + "Equivalencias ,"
                ZSql = ZSql + "Fisica ,"
                ZSql = ZSql + "OtrosI ,"
                ZSql = ZSql + "OtrosII ,"
                ZSql = ZSql + "Curso ,"
                ZSql = ZSql + "NecesariaCurso ,"
                ZSql = ZSql + "DeseableCurso )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Codigo.Text + "',"
                ZSql = ZSql + "'" + Version.Text + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + ZZDescripcion + "',"
                ZSql = ZSql + "'" + ZDesdeVigencia + "',"
                ZSql = ZSql + "'" + ZHastaVigencia + "',"
                ZSql = ZSql + "'" + ZZSector + "',"
                ZSql = ZSql + "'" + ZZTareasI + "',"
                ZSql = ZSql + "'" + ZZTareasII + "',"
                ZSql = ZSql + "'" + ZZTareasIII + "',"
                ZSql = ZSql + "'" + ZZDescriI + "',"
                ZSql = ZSql + "'" + ZZDescriII + "',"
                ZSql = ZSql + "'" + ZZDescriIII + "',"
                ZSql = ZSql + "'" + ZZDescriIV + "',"
                ZSql = ZSql + "'" + ZZDescriV + "',"
                ZSql = ZSql + "'" + ZZObservaI + "',"
                ZSql = ZSql + "'" + ZZObservaII + "',"
                ZSql = ZSql + "'" + ZZObservaIII + "',"
                ZSql = ZSql + "'" + ZZObservaIV + "',"
                ZSql = ZSql + "'" + ZZObservaV + "',"
                ZSql = ZSql + "'" + ZZNecesariaI + "',"
                ZSql = ZSql + "'" + ZZNecesariaII + "',"
                ZSql = ZSql + "'" + ZZNecesariaIII + "',"
                ZSql = ZSql + "'" + ZZNecesariaIV + "',"
                ZSql = ZSql + "'" + ZZNecesariaV + "',"
                ZSql = ZSql + "'" + ZZDeseableI + "',"
                ZSql = ZSql + "'" + ZZDeseableII + "',"
                ZSql = ZSql + "'" + ZZDeseableIII + "',"
                ZSql = ZSql + "'" + ZZDeseableIV + "',"
                ZSql = ZSql + "'" + ZZDeseableV + "',"
                ZSql = ZSql + "'" + ZZEquivalencias + "',"
                ZSql = ZSql + "'" + ZZFisica + "',"
                ZSql = ZSql + "'" + ZZOtrosI + "',"
                ZSql = ZSql + "'" + ZZOtrosII + "',"
                ZSql = ZSql + "'" + ZZCurso + "',"
                ZSql = ZSql + "'" + ZZNecesariaCurso + "',"
                ZSql = ZSql + "'" + ZZDeseableCurso + "')"
            
                spTareaVersion = ZSql
                Set rstTareaVersion = db.OpenRecordset(spTareaVersion, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
        Next IRow
    
        WWPerfilVersion = Version.Text
        Vigencia.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        Version.Text = Str$(Val(Version.Text) + 1)
        
        
        
        
        
        ZZRenglon = 0
        Erase ZZVector
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Legajo"
        ZSql = ZSql + " Where Legajo.Perfil = " + "'" + Codigo.Text + "'"
        ZSql = ZSql + " and Legajo.Renglon = 1"
        ZSql = ZSql + " Order by Clave"
    
        spLegajo = ZSql
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
            With rstLegajo
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        ZZRenglon = ZZRenglon + 1
                    
                        ZZVector(ZZRenglon, 1) = rstLegajo!Codigo
                        ZZVector(ZZRenglon, 2) = rstLegajo!Version
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLegajo.Close
        End If
        
        For CicloII = 1 To ZZRenglon
        
            WWLegajo = ZZVector(CicloII, 1)
            WWVersion = ZZVector(CicloII, 2)
    
            WWFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            WWVersion = Str$(Val(WWVersion) + 1)
        
            WWRenglon = 0
            Erase WWVector
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Legajo"
            ZSql = ZSql + " Where Legajo.Codigo = " + "'" + WWLegajo + "'"
            ZSql = ZSql + " Order by Clave"
    
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLegajo.RecordCount > 0 Then
                With rstLegajo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                
                            WWRenglon = WWRenglon + 1
                    
                            WWVector(WWRenglon, 1) = rstLegajo!FIngreso
                            WWVector(WWRenglon, 2) = Str$(rstLegajo!Perfil)
                            WWVector(WWRenglon, 3) = Trim(rstLegajo!EstadoI)
                            WWVector(WWRenglon, 4) = Trim(rstLegajo!EstadoII)
                            WWVector(WWRenglon, 5) = Trim(rstLegajo!EstadoIII)
                            WWVector(WWRenglon, 6) = Trim(rstLegajo!EstadoIV)
                            WWVector(WWRenglon, 7) = Trim(rstLegajo!EstadoV)
                            WWVector(WWRenglon, 8) = Trim(rstLegajo!EstadoVI)
                            WWVector(WWRenglon, 9) = Trim(rstLegajo!EstadoVII)
                            WWVector(WWRenglon, 10) = Trim(rstLegajo!EstadoVIII)
                            WWVector(WWRenglon, 11) = Trim(rstLegajo!EstadoIX)
                            WWVector(WWRenglon, 12) = rstLegajo!EstaI
                            WWVector(WWRenglon, 13) = rstLegajo!EstaII
                            WWVector(WWRenglon, 14) = rstLegajo!EstaIII
                            WWVector(WWRenglon, 15) = rstLegajo!EstaIV
                            WWVector(WWRenglon, 16) = rstLegajo!EstaV
                            WWVector(WWRenglon, 17) = rstLegajo!EstaVI
                            WWVector(WWRenglon, 18) = rstLegajo!EstaVII
                            WWVector(WWRenglon, 19) = rstLegajo!EstaVIII
                            WWVector(WWRenglon, 20) = rstLegajo!EstaIX
                            WWVector(WWRenglon, 21) = rstLegajo!Curso
                            WWVector(WWRenglon, 22) = rstLegajo!EstaCurso
                            WWVector(WWRenglon, 23) = rstLegajo!ClavePerfil
                            WWVector(WWRenglon, 24) = rstLegajo!NecesariaCurso
                            WWVector(WWRenglon, 25) = rstLegajo!DeseableCurso
                            WWVector(WWRenglon, 26) = rstLegajo!Version
                            WWVector(WWRenglon, 27) = rstLegajo!FechaVersion
                            WWVector(WWRenglon, 28) = rstLegajo!EstadoCurso
                            WWVector(WWRenglon, 29) = IIf(IsNull(rstLegajo!Fegreso), "  /  /    ", rstLegajo!Fegreso)
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLegajo.Close
            End If
        
            For CicloIII = 1 To WWRenglon
        
                WWFIngreso = WWVector(CicloIII, 1)
                WWPerfil = WWVector(CicloIII, 2)
                Rem WWPerfilVersion = PerfilVersion.Text
            
                WWEstadoI = WWVector(CicloIII, 3)
                WWEstadoII = WWVector(CicloIII, 4)
                WWEstadoIII = WWVector(CicloIII, 5)
                WWEstadoIV = WWVector(CicloIII, 6)
                WWEstadoV = WWVector(CicloIII, 7)
                WWEstadoVI = WWVector(CicloIII, 8)
                WWEstadoVII = WWVector(CicloIII, 9)
                WWEstadoVIII = WWVector(CicloIII, 10)
                WWEstadoIX = WWVector(CicloIII, 11)
                
                WWEstaI = WWVector(CicloIII, 12)
                WWEstaII = WWVector(CicloIII, 13)
                WWEstaIII = WWVector(CicloIII, 14)
                WWEstaIV = WWVector(CicloIII, 15)
                WWEstaV = WWVector(CicloIII, 16)
                WWEstaVI = WWVector(CicloIII, 17)
                WWEstaVII = WWVector(CicloIII, 18)
                WWEstaVIII = WWVector(CicloIII, 19)
                WWEstaIX = WWVector(CicloIII, 20)
                
                WWCurso = WWVector(CicloIII, 21)
                WWEstaCurso = WWVector(CicloIII, 22)
                WWClavePerfil = WWVector(CicloIII, 23)
                WWNecesariaCurso = WWVector(CicloIII, 24)
                WWDeseableCurso = WWVector(CicloIII, 25)
                
                WWVersion = WWVector(CicloIII, 26)
                WWFechaVersionI = WWVector(CicloIII, 27)
                WWFechaVersionII = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                WWEstadoCurso = WWVector(CicloIII, 28)
                WWFegreso = WWVector(CicloIII, 29)
                
                Auxi1 = WWLegajo
                Call Ceros(Auxi1, 6)
                
                Auxi2 = WWVersion
                Call Ceros(Auxi2, 4)
                    
                Auxi = Str$(CicloIII)
                Call Ceros(Auxi, 2)
                        
                WWClave = Auxi1 + Auxi2 + Auxi
            
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
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWLegajo + "',"
                ZSql = ZSql + "'" + WWVersion + "',"
                ZSql = ZSql + "'" + Str$(CicloIII) + "',"
                ZSql = ZSql + "'" + WWFIngreso + "',"
                ZSql = ZSql + "'" + WWPerfil + "',"
                ZSql = ZSql + "'" + WWPerfilVersion + "',"
                ZSql = ZSql + "'" + WWEstadoI + "',"
                ZSql = ZSql + "'" + WWEstadoII + "',"
                ZSql = ZSql + "'" + WWEstadoIII + "',"
                ZSql = ZSql + "'" + WWEstadoIV + "',"
                ZSql = ZSql + "'" + WWEstadoV + "',"
                ZSql = ZSql + "'" + WWEstadoVI + "',"
                ZSql = ZSql + "'" + WWEstadoVII + "',"
                ZSql = ZSql + "'" + WWEstadoVIII + "',"
                ZSql = ZSql + "'" + WWEstadoIX + "',"
                ZSql = ZSql + "'" + WWCurso + "',"
                ZSql = ZSql + "'" + WWEstaI + "',"
                ZSql = ZSql + "'" + WWEstaII + "',"
                ZSql = ZSql + "'" + WWEstaIII + "',"
                ZSql = ZSql + "'" + WWEstaIV + "',"
                ZSql = ZSql + "'" + WWEstaV + "',"
                ZSql = ZSql + "'" + WWEstaVI + "',"
                ZSql = ZSql + "'" + WWEstaVII + "',"
                ZSql = ZSql + "'" + WWEstaVIII + "',"
                ZSql = ZSql + "'" + WWEstaIX + "',"
                ZSql = ZSql + "'" + WWNecesariaCurso + "',"
                ZSql = ZSql + "'" + WWDeseableCurso + "',"
                ZSql = ZSql + "'" + WWClavePerfil + "',"
                ZSql = ZSql + "'" + WWEstaCurso + "',"
                ZSql = ZSql + "'" + WWEstadoCurso + "',"
                ZSql = ZSql + "'" + WWFechaVersionI + "',"
                ZSql = ZSql + "'" + WWFechaVersionII + "',"
                ZSql = ZSql + "'" + WWFegreso + "')"
             
                Rem  ZSql = ZSql + "'" + ZZFechaVersionII + "')"
                spLegajoVersion = ZSql
                Set rstLegajoVersion = db.OpenRecordset(spLegajoVersion, dbOpenSnapshot, dbSQLPassThrough)
            
            Next CicloIII
            
            WWLegajo = ZZVector(CicloII, 1)
            WWVersion = ZZVector(CicloII, 2)
            WWVersion = Str$(Val(WWVersion) + 1)
            WWFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Legajo SET "
            ZSql = ZSql + " FechaVersion = " + "'" + WWFechaVersion + "',"
            ZSql = ZSql + " Version = " + "'" + WWVersion + "',"
            ZSql = ZSql + " Actualizado = " + "'" + "N" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WWLegajo + "'"
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            
        Next CicloII
        
    End If



    ZSql = ""
    ZSql = ZSql + "DELETE Tarea"
    ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
    spTarea = ZSql
    Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)

    WRenglon = 0
    For IRow = 1 To 100
        
        WVector1.Row = IRow
            
        WVector1.Col = 1
        ZCurso = WVector1.Text
        
        WVector1.Col = 3
        ZNecesaria = WVector1.Text
            
        WVector1.Col = 4
        ZDeseable = WVector1.Text
            
        If Val(ZCurso) <> 0 Or IRow = 1 Then
                    
            Auxi1 = Codigo.Text
            Call Ceros(Auxi1, 6)
                    
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
                        
            WClave = Auxi1 + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Tarea ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Vigencia ,"
            ZSql = ZSql + "Sector ,"
            ZSql = ZSql + "TareasI ,"
            ZSql = ZSql + "TareasII ,"
            ZSql = ZSql + "TareasIII ,"
            ZSql = ZSql + "DescriI ,"
            ZSql = ZSql + "DescriII ,"
            ZSql = ZSql + "DescriIII ,"
            ZSql = ZSql + "DescriIV ,"
            ZSql = ZSql + "DescriV ,"
            ZSql = ZSql + "ObservaI ,"
            ZSql = ZSql + "ObservaII ,"
            ZSql = ZSql + "ObservaIII ,"
            ZSql = ZSql + "ObservaIV ,"
            ZSql = ZSql + "ObservaV ,"
            ZSql = ZSql + "NecesariaI ,"
            ZSql = ZSql + "NecesariaII ,"
            ZSql = ZSql + "NecesariaIII ,"
            ZSql = ZSql + "NecesariaIV ,"
            ZSql = ZSql + "NecesariaV ,"
            ZSql = ZSql + "DeseableI ,"
            ZSql = ZSql + "DeseableII ,"
            ZSql = ZSql + "DeseableIII ,"
            ZSql = ZSql + "DeseableIV ,"
            ZSql = ZSql + "DeseableV ,"
            ZSql = ZSql + "Equivalencias ,"
            ZSql = ZSql + "Fisica ,"
            ZSql = ZSql + "OtrosI ,"
            ZSql = ZSql + "OtrosII ,"
            ZSql = ZSql + "Curso ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "NecesariaCurso ,"
            ZSql = ZSql + "DeseableCurso )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + Vigencia.Text + "',"
            ZSql = ZSql + "'" + Sector.Text + "',"
            ZSql = ZSql + "'" + TareasI.Text + "',"
            ZSql = ZSql + "'" + TareasII.Text + "',"
            ZSql = ZSql + "'" + TareasIII.Text + "',"
            ZSql = ZSql + "'" + DescriI.Text + "',"
            ZSql = ZSql + "'" + DescriII.Text + "',"
            ZSql = ZSql + "'" + DescriIII.Text + "',"
            ZSql = ZSql + "'" + DescriIV.Text + "',"
            ZSql = ZSql + "'" + DescriV.Text + "',"
            ZSql = ZSql + "'" + ObservaI.Text + "',"
            ZSql = ZSql + "'" + ObservaII.Text + "',"
            ZSql = ZSql + "'" + ObservaIII.Text + "',"
            ZSql = ZSql + "'" + ObservaIV.Text + "',"
            ZSql = ZSql + "'" + ObservaV.Text + "',"
            ZSql = ZSql + "'" + Str$(NecesariaI.Value) + "',"
            ZSql = ZSql + "'" + Str$(NecesariaII.Value) + "',"
            ZSql = ZSql + "'" + Str$(NecesariaIII.Value) + "',"
            ZSql = ZSql + "'" + Str$(NecesariaIV.Value) + "',"
            ZSql = ZSql + "'" + Str$(NecesariaV.Value) + "',"
            ZSql = ZSql + "'" + Str$(DeseableI.Value) + "',"
            ZSql = ZSql + "'" + Str$(DeseableII.Value) + "',"
            ZSql = ZSql + "'" + Str$(DeseableIII.Value) + "',"
            ZSql = ZSql + "'" + Str$(DeseableIV.Value) + "',"
            ZSql = ZSql + "'" + Str$(DeseableV.Value) + "',"
            ZSql = ZSql + "'" + Equivalencias.Text + "',"
            ZSql = ZSql + "'" + Fisica.Text + "',"
            ZSql = ZSql + "'" + OtrosI.Text + "',"
            ZSql = ZSql + "'" + OtrosII.Text + "',"
            ZSql = ZSql + "'" + ZCurso + "',"
            ZSql = ZSql + "'" + Version.Text + "',"
            ZSql = ZSql + "'" + ZNecesaria + "',"
            ZSql = ZSql + "'" + ZDeseable + "')"
            
            spTarea = ZSql
            Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
            
    Next IRow
    
    Call Limpia_Click

End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Codigo.Text = ""
    Descripcion.Text = ""
    Vigencia.Text = "  /  /    "
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
    Version.Text = "1"
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Tarea"
    spTarea = Sql1 + Sql2
    Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
    If rstTarea.RecordCount > 0 Then
        rstTarea.MoveLast
        WCodigoMayor = IIf(IsNull(rstTarea!CodigoMayor), "0", rstTarea!CodigoMayor)
        Codigo.Text = Mid$(Str$(WCodigoMayor + 1), 2, 8)
        rstTarea.Close
            Else
        Codigo.Text = "1"
    End If
    
    Renglon = 0
    
    Tablas.Tab = 0
    
    WVector1.Col = 1
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
            
        Case 1
            Indice = Pantalla.ListIndex
            Sector.Text = WIndice.List(Indice)
            Call Sector_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            WCurso = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Curso"
            Sql3 = " Where Curso.Codigo = " + "'" + WCurso + "'"
            spCurso = Sql1 + Sql2 + Sql3
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstCurso!Codigo
                WVector1.Col = 2
                WVector1.Text = rstCurso!Descripcion
                WVector1.Col = 1
                rstCurso.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Codigo.Text = ""
    Descripcion.Text = ""
    Vigencia.Text = "  /  /    "
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
    Version.Text = "1"
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Tarea"
    spTarea = Sql1 + Sql2
    Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
    If rstTarea.RecordCount > 0 Then
        rstTarea.MoveLast
        WCodigoMayor = IIf(IsNull(rstTarea!CodigoMayor), "0", rstTarea!CodigoMayor)
        Codigo.Text = Mid$(Str$(WCodigoMayor + 1), 2, 8)
        rstTarea.Close
            Else
        Codigo.Text = "1"
    End If
    
    Renglon = 0
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *, Curso.Descripcion as [WDesCurso], Sector.Descripcion as [WDesSector]"
    ZSql = ZSql + " FROM Tarea, Curso, Sector"
    ZSql = ZSql + " Where Tarea.Codigo = " + "'" + Codigo.Text + "'"
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
                        Descripcion.Text = Trim(rstTarea!Descripcion)
                        Version.Text = rstTarea!Version
                        Vigencia.Text = rstTarea!Vigencia
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
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTarea.Close
    End If
    
    Tablas.Tab = 0
    Codigo.SetFocus

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM Tarea"
        Sql3 = " Where Tarea.Codigo = " + "'" + Codigo.Text + "'"
        spTarea = Sql1 + Sql2 + Sql3
        Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
        If rstTarea.RecordCount > 0 Then
            rstTarea.Close
            Call Proceso_Click
            WVector1.Col = 1
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
        Vigencia.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Vigencia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Vigencia.Text, Auxi)
        If Auxi = "S" Then
            Sector.SetFocus
                Else
            Vigencia.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vigencia.Text = "  /  /    "
    End If
End Sub

Private Sub Sector_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Sector"
        Sql3 = " Where Sector.Codigo = " + "'" + Sector.Text + "'"
        spSector = Sql1 + Sql2 + Sql3
        Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        If rstSector.RecordCount > 0 Then
            DesSector.Caption = rstSector!Descripcion
            rstSector.Close
            TareasI.SetFocus
                Else
            Sector.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Sector.Text = ""
        DesSector.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub TareasI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TareasII.SetFocus
    End If
    If KeyAscii = 27 Then
        TareasI.Text = ""
    End If
End Sub

Private Sub TareasII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TareasIII.SetFocus
    End If
    If KeyAscii = 27 Then
        TareasII.Text = ""
    End If
End Sub

Private Sub TareasIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriI.SetFocus
    End If
    If KeyAscii = 27 Then
        TareasIII.Text = ""
    End If
End Sub

Private Sub DescriI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaI.SetFocus
    End If
    If KeyAscii = 27 Then
        DescriI.Text = ""
    End If
End Sub

Private Sub ObservaI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaI.Text = ""
    End If
End Sub

Private Sub DescriII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaII.SetFocus
    End If
    If KeyAscii = 27 Then
        DescriII.Text = ""
    End If
End Sub

Private Sub ObservaII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaII.Text = ""
    End If
End Sub

Private Sub DescriIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIII.SetFocus
    End If
    If KeyAscii = 27 Then
        DescriIII.Text = ""
    End If
End Sub

Private Sub ObservaIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriIV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaIII.Text = ""
    End If
End Sub

Private Sub DescriIV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaIV.SetFocus
    End If
    If KeyAscii = 27 Then
        DescriIV.Text = ""
    End If
End Sub

Private Sub ObservaIV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaIV.Text = ""
    End If
End Sub

Private Sub DescriV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservaV.SetFocus
    End If
    If KeyAscii = 27 Then
        DescriV.Text = ""
    End If
End Sub

Private Sub ObservaV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Equivalencias.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservaV.Text = ""
    End If
End Sub

Private Sub Equivalencias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fisica.SetFocus
    End If
    If KeyAscii = 27 Then
        Equivalencias.Text = ""
    End If
End Sub

Private Sub Fisica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OtrosI.SetFocus
    End If
    If KeyAscii = 27 Then
        Fisica.Text = ""
    End If
End Sub

Private Sub OtrosI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OtrosII.SetFocus
    End If
    If KeyAscii = 27 Then
        OtrosI.Text = ""
    End If
End Sub

Private Sub OtrosII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        OtrosII.Text = ""
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
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstTarea!Descripcion, aa, WEspacios) Then
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
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Sector"
            Sql3 = " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            Sql4 = " Order by Codigo"
            spSector = Sql1 + Sql2 + Sql3 + Sql4
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
            If rstSector.RecordCount > 0 Then
                With rstSector
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSector.Close
            End If
            
        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM Curso"
            Sql3 = " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            Sql4 = " Order by Codigo"
            spCurso = Sql1 + Sql2 + Sql3 + Sql4
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                With rstCurso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCurso.Close
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
        Case 4
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
            Sql1 = "Select *"
            Sql2 = " FROM Curso"
            Sql3 = " Where Curso.Codigo = " + "'" + WVector1.Text + "'"
            spCurso = Sql1 + Sql2 + Sql3
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstCurso!Descripcion
                rstCurso.Close
                    Else
                WControl = "N"
            End If
            
        Case 3, 4
            If Trim(WVector1.Text) <> "" Then
                If WVector1.Text <> "X" And WVector1.Text <> "" Then
                    WControl = "N"
                End If
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

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

        Opcion.Clear
    
        Opcion.AddItem "Perfiles"
        Opcion.AddItem "Sector"
        Opcion.AddItem "Temas"

        Rem Opcion.Visible = True
    
        Opcion.ListIndex = 2
    
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
    WVector1.Cols = 5
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
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Necesaria"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Deseable"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
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
    
    WVector1.Col = 1
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
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        Case Else
    End Select
End Sub

