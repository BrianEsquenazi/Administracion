VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgTrabajoIII 
   Caption         =   "Ingreso de Especificaciones de Materia Prima (Unificado)"
   ClientHeight    =   8010
   ClientLeft      =   450
   ClientTop       =   615
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   8010
   ScaleWidth      =   11160
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3720
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   39
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   40
         Top             =   240
         Width           =   2895
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   41
      Top             =   720
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8070
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Especificacion 1 - 10"
      TabPicture(0)   =   "TrabajoIII.frx":0000
      Tab(0).ControlCount=   33
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Descri10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Descri9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Descri8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Descri7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Descri6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Descri5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Descri4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Descri3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "descri2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Descri1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDescri"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblensayo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblresultado"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Ensayo10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Ensayo9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Ensayo8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Ensayo7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Ensayo6"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Ensayo5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Ensayo4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Ensayo3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Ensayo2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Ensayo1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "valor10"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "valor9"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "valor8"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "valor7"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "valor6"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "valor5"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "valor4"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Valor3"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "valor2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Valor1"
      Tab(0).Control(32).Enabled=   0   'False
      TabCaption(1)   =   "Especificacion 11  - 20"
      TabPicture(1)   =   "TrabajoIII.frx":001C
      Tab(1).ControlCount=   33
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Ensayo20"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Ensayo19"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "Ensayo18"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "Ensayo17"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "Ensayo16"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "Ensayo15"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "Ensayo14"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "Ensayo13"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "Ensayo12"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Ensayo11"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Valor20"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Valor19"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Valor18"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Valor17"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Valor16"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Valor15"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Valor14"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "Valor13"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Valor12"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "Valor11"
      Tab(1).Control(19).Enabled=   -1  'True
      Tab(1).Control(20)=   "Descri20"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Descri19"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Descri18"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Descri17"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Descri16"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Descri15"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Descri14"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Descri13"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Descri12"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Descri11"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label8"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label7"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label6"
      Tab(1).Control(32).Enabled=   0   'False
      Begin VB.TextBox Ensayo20 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   94
         Text            =   " "
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Ensayo19 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   93
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Ensayo18 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   92
         Text            =   " "
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Ensayo17 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   91
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Ensayo16 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   90
         Text            =   " "
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Ensayo15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   89
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Ensayo14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   88
         Text            =   " "
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Ensayo13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   87
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Ensayo12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   86
         Text            =   " "
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Ensayo11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   85
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Valor20 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   84
         Text            =   " "
         Top             =   4080
         Width           =   5055
      End
      Begin VB.TextBox Valor19 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   83
         Text            =   " "
         Top             =   3720
         Width           =   5055
      End
      Begin VB.TextBox Valor18 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   82
         Text            =   " "
         Top             =   3360
         Width           =   5055
      End
      Begin VB.TextBox Valor17 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   81
         Text            =   " "
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox Valor16 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   80
         Text            =   " "
         Top             =   2640
         Width           =   5055
      End
      Begin VB.TextBox Valor15 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   79
         Text            =   " "
         Top             =   2280
         Width           =   5055
      End
      Begin VB.TextBox Valor14 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   78
         Text            =   " "
         Top             =   1920
         Width           =   5055
      End
      Begin VB.TextBox Valor13 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   77
         Text            =   " "
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox Valor12 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   76
         Text            =   " "
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox Valor11 
         Height          =   285
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   75
         Text            =   " "
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   73
         Text            =   " "
         Top             =   780
         Width           =   5055
      End
      Begin VB.TextBox valor2 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   72
         Text            =   " "
         Top             =   1140
         Width           =   5055
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   71
         Text            =   " "
         Top             =   1500
         Width           =   5055
      End
      Begin VB.TextBox valor4 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   70
         Text            =   " "
         Top             =   1860
         Width           =   5055
      End
      Begin VB.TextBox valor5 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   69
         Text            =   " "
         Top             =   2220
         Width           =   5055
      End
      Begin VB.TextBox valor6 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   68
         Text            =   " "
         Top             =   2580
         Width           =   5055
      End
      Begin VB.TextBox valor7 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   67
         Text            =   " "
         Top             =   2940
         Width           =   5055
      End
      Begin VB.TextBox valor8 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   66
         Text            =   " "
         Top             =   3300
         Width           =   5055
      End
      Begin VB.TextBox valor9 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   65
         Text            =   " "
         Top             =   3660
         Width           =   5055
      End
      Begin VB.TextBox valor10 
         Height          =   285
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   64
         Text            =   " "
         Top             =   4020
         Width           =   5055
      End
      Begin VB.TextBox Ensayo1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   51
         Text            =   " "
         Top             =   780
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
      Begin VB.TextBox Ensayo3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   49
         Text            =   " "
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Ensayo4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   48
         Text            =   " "
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox Ensayo5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   47
         Text            =   " "
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox Ensayo6 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   4
         TabIndex        =   46
         Text            =   " "
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox Ensayo7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   45
         Text            =   " "
         Top             =   2940
         Width           =   735
      End
      Begin VB.TextBox Ensayo8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   44
         Text            =   " "
         Top             =   3300
         Width           =   735
      End
      Begin VB.TextBox Ensayo9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   43
         Text            =   " "
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox Ensayo10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   42
         Text            =   " "
         Top             =   4020
         Width           =   735
      End
      Begin VB.Label Descri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   107
         Top             =   4080
         Width           =   4740
      End
      Begin VB.Label Descri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   106
         Top             =   3720
         Width           =   4740
      End
      Begin VB.Label Descri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   105
         Top             =   3360
         Width           =   4740
      End
      Begin VB.Label Descri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   104
         Top             =   3000
         Width           =   4740
      End
      Begin VB.Label Descri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   103
         Top             =   2640
         Width           =   4740
      End
      Begin VB.Label Descri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   102
         Top             =   2280
         Width           =   4740
      End
      Begin VB.Label Descri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   101
         Top             =   1920
         Width           =   4740
      End
      Begin VB.Label Descri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   100
         Top             =   1560
         Width           =   4740
      End
      Begin VB.Label Descri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   99
         Top             =   1200
         Width           =   4740
      End
      Begin VB.Label Descri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   98
         Top             =   840
         Width           =   4740
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
         TabIndex        =   97
         Top             =   480
         Width           =   4695
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
         TabIndex        =   96
         Top             =   480
         Width           =   735
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
         TabIndex        =   95
         Top             =   480
         Width           =   5055
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
         TabIndex        =   74
         Top             =   420
         Width           =   5055
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
         TabIndex        =   63
         Top             =   420
         Width           =   735
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
         TabIndex        =   62
         Top             =   420
         Width           =   4695
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   61
         Top             =   780
         Width           =   4740
      End
      Begin VB.Label descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   60
         Top             =   1140
         Width           =   4740
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   59
         Top             =   1500
         Width           =   4740
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   58
         Top             =   1860
         Width           =   4740
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   57
         Top             =   2220
         Width           =   4740
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   56
         Top             =   2580
         Width           =   4740
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Top             =   2940
         Width           =   4740
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   54
         Top             =   3300
         Width           =   4740
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   53
         Top             =   3660
         Width           =   4740
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   52
         Top             =   4020
         Width           =   4740
      End
   End
   Begin VB.TextBox Fecha 
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
      TabIndex        =   34
      Text            =   " "
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Version 
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
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   360
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
      Left            =   4800
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wespec1Unifica.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   4200
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   3135
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton ImpreListado 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton ImprePantalla 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Impresora"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Pantalla"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
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
      Left            =   1080
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5400
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
      Height          =   2460
      ItemData        =   "TrabajoIII.frx":0038
      Left            =   240
      List            =   "TrabajoIII.frx":003F
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Listado"
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      Height          =   1335
      Left            =   9120
      TabIndex        =   6
      Top             =   5520
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   7920
      TabIndex        =   5
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7920
      TabIndex        =   4
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   300
      Left            =   7920
      TabIndex        =   3
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   7920
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
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
      Height          =   255
      Left            =   6720
      TabIndex        =   36
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
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
      Left            =   8280
      TabIndex        =   35
      Top             =   0
      Width           =   2175
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
      Left            =   8160
      TabIndex        =   32
      Top             =   360
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
      Left            =   6000
      TabIndex        =   31
      Top             =   360
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
      TabIndex        =   28
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   15
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
Attribute VB_Name = "PrgTrabajoIII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String

Dim XParam As String

Dim WSaldo As Double
Dim WEntra As String

Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer
Dim WTipoPedido As String

Dim XCantidad1 As String
Dim xCantidad2 As String

Dim XMes As String
Dim XAno As String

Dim ControlLote(12, 2) As String
Dim CargaEmpresa(12, 2) As String
Dim ZHasta As Integer

Dim WCanti As Double
Dim WLote As String
Dim WLugar As Integer
Dim ZFechaVto As String
Dim Empe(12, 10) As String
Dim ZMes As String
Dim ZAno As String
Dim ZZRestriccion As Integer
Dim ZZRestriccionI As Integer
Dim ZZRestriccionII As Integer
Dim ZZVerifica(100, 2) As String


Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String
Dim WControlII As String

Dim ZSaldo As Double

Private Sub Cancela_click()
    PrgModpedNuevoII.Hide
    Unload Me
    PrgModpedNuevo.Show
End Sub

Private Sub Confirma_Click()
    Call Verifica_Lote
    If WEstado = "S" Then

        ZZTrabajaLote(1) = WVector1.TextMatrix(1, 1)
        ZZTrabajaLote(3) = WVector1.TextMatrix(2, 1)
        ZZTrabajaLote(5) = WVector1.TextMatrix(3, 1)
        ZZTrabajaLote(7) = WVector1.TextMatrix(4, 1)
        ZZTrabajaLote(9) = WVector1.TextMatrix(5, 1)
        ZZTrabajaLote(11) = WVector1.TextMatrix(6, 1)
        ZZTrabajaLote(13) = WVector1.TextMatrix(7, 1)
        ZZTrabajaLote(15) = WVector1.TextMatrix(8, 1)
        ZZTrabajaLote(17) = WVector1.TextMatrix(9, 1)
        ZZTrabajaLote(19) = WVector1.TextMatrix(10, 1)
        ZZTrabajaLote(21) = WVector1.TextMatrix(11, 1)
        ZZTrabajaLote(23) = WVector1.TextMatrix(12, 1)
        
        ZZTrabajaLote(2) = WVector1.TextMatrix(1, 2)
        ZZTrabajaLote(4) = WVector1.TextMatrix(2, 2)
        ZZTrabajaLote(6) = WVector1.TextMatrix(3, 2)
        ZZTrabajaLote(8) = WVector1.TextMatrix(4, 2)
        ZZTrabajaLote(10) = WVector1.TextMatrix(5, 2)
        ZZTrabajaLote(12) = WVector1.TextMatrix(6, 2)
        ZZTrabajaLote(14) = WVector1.TextMatrix(7, 2)
        ZZTrabajaLote(16) = WVector1.TextMatrix(8, 2)
        ZZTrabajaLote(18) = WVector1.TextMatrix(9, 2)
        ZZTrabajaLote(20) = WVector1.TextMatrix(10, 2)
        ZZTrabajaLote(22) = WVector1.TextMatrix(11, 2)
        ZZTrabajaLote(24) = WVector1.TextMatrix(12, 2)
        
        ZZTrabajaLote(31) = WVector1.TextMatrix(1, 3)
        ZZTrabajaLote(32) = WVector1.TextMatrix(1, 5)
        ZZTrabajaLote(33) = WVector1.TextMatrix(2, 3)
        ZZTrabajaLote(34) = WVector1.TextMatrix(2, 5)
        ZZTrabajaLote(35) = WVector1.TextMatrix(3, 3)
        ZZTrabajaLote(36) = WVector1.TextMatrix(3, 5)
        ZZTrabajaLote(37) = WVector1.TextMatrix(4, 3)
        ZZTrabajaLote(38) = WVector1.TextMatrix(4, 5)
        ZZTrabajaLote(39) = WVector1.TextMatrix(5, 3)
        ZZTrabajaLote(40) = WVector1.TextMatrix(5, 5)
        ZZTrabajaLote(41) = WVector1.TextMatrix(6, 3)
        ZZTrabajaLote(42) = WVector1.TextMatrix(6, 5)
        ZZTrabajaLote(43) = WVector1.TextMatrix(7, 3)
        ZZTrabajaLote(44) = WVector1.TextMatrix(7, 5)
        ZZTrabajaLote(45) = WVector1.TextMatrix(8, 3)
        ZZTrabajaLote(46) = WVector1.TextMatrix(8, 5)
        ZZTrabajaLote(47) = WVector1.TextMatrix(9, 3)
        ZZTrabajaLote(48) = WVector1.TextMatrix(9, 5)
        ZZTrabajaLote(49) = WVector1.TextMatrix(10, 3)
        ZZTrabajaLote(50) = WVector1.TextMatrix(10, 5)
        ZZTrabajaLote(51) = WVector1.TextMatrix(11, 3)
        ZZTrabajaLote(52) = WVector1.TextMatrix(11, 5)
        ZZTrabajaLote(53) = WVector1.TextMatrix(12, 3)
        ZZTrabajaLote(54) = WVector1.TextMatrix(12, 5)
        
        ZZTrabajaLote(61) = WVector1.TextMatrix(1, 6)
        ZZTrabajaLote(62) = WVector1.TextMatrix(2, 6)
        ZZTrabajaLote(63) = WVector1.TextMatrix(3, 6)
        ZZTrabajaLote(64) = WVector1.TextMatrix(4, 6)
        ZZTrabajaLote(65) = WVector1.TextMatrix(5, 6)
        ZZTrabajaLote(66) = WVector1.TextMatrix(6, 6)
        ZZTrabajaLote(67) = WVector1.TextMatrix(7, 6)
        ZZTrabajaLote(68) = WVector1.TextMatrix(8, 6)
        ZZTrabajaLote(69) = WVector1.TextMatrix(9, 6)
        ZZTrabajaLote(70) = WVector1.TextMatrix(10, 6)
        ZZTrabajaLote(71) = WVector1.TextMatrix(11, 6)
        ZZTrabajaLote(72) = WVector1.TextMatrix(12, 6)
        
        ZZTrabajaLote(81) = WVector1.TextMatrix(1, 7)
        ZZTrabajaLote(82) = WVector1.TextMatrix(2, 7)
        ZZTrabajaLote(83) = WVector1.TextMatrix(3, 7)
        ZZTrabajaLote(84) = WVector1.TextMatrix(4, 7)
        ZZTrabajaLote(85) = WVector1.TextMatrix(5, 7)
        ZZTrabajaLote(86) = WVector1.TextMatrix(6, 7)
        ZZTrabajaLote(87) = WVector1.TextMatrix(7, 7)
        ZZTrabajaLote(88) = WVector1.TextMatrix(8, 7)
        ZZTrabajaLote(89) = WVector1.TextMatrix(9, 7)
        ZZTrabajaLote(90) = WVector1.TextMatrix(10, 7)
        ZZTrabajaLote(91) = WVector1.TextMatrix(11, 7)
        ZZTrabajaLote(92) = WVector1.TextMatrix(12, 7)
        
        PrgModpedNuevoII.Hide
        Unload Me
        PrgModpedNuevo.Show
    End If
End Sub

Private Sub Form_Load()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
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
            CargaEmpresa(6, 1) = "0010"
            CargaEmpresa(6, 2) = "Empresa10"
            CargaEmpresa(7, 1) = "0011"
            CargaEmpresa(7, 2) = "Empresa11"
            ZHasta = 7
        Case Else
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
            ZHasta = 4
    End Select
    
    
    Call Limpia_Vector
    
    Producto.Text = WPasaTerminado
    Cantidad.Text = Pusing("###,###.##", Str$(WPasaCantidad))
    WTipoPedido = WPasaTipoPedido
    WControlII = ""
    
    WVector1.TextMatrix(1, 1) = ZZTrabajaLote(1)
    WVector1.TextMatrix(2, 1) = ZZTrabajaLote(3)
    WVector1.TextMatrix(3, 1) = ZZTrabajaLote(5)
    WVector1.TextMatrix(4, 1) = ZZTrabajaLote(7)
    WVector1.TextMatrix(5, 1) = ZZTrabajaLote(9)
    WVector1.TextMatrix(6, 1) = ZZTrabajaLote(11)
    WVector1.TextMatrix(7, 1) = ZZTrabajaLote(13)
    WVector1.TextMatrix(8, 1) = ZZTrabajaLote(15)
    WVector1.TextMatrix(9, 1) = ZZTrabajaLote(17)
    WVector1.TextMatrix(10, 1) = ZZTrabajaLote(19)
    WVector1.TextMatrix(11, 1) = ZZTrabajaLote(21)
    WVector1.TextMatrix(12, 1) = ZZTrabajaLote(23)
    
    If Val(ZZTrabajaLote(2)) <> 0 Then
        WVector1.TextMatrix(1, 2) = ZZTrabajaLote(2)
    End If
    If Val(ZZTrabajaLote(4)) <> 0 Then
        WVector1.TextMatrix(2, 2) = ZZTrabajaLote(4)
    End If
    If Val(ZZTrabajaLote(6)) <> 0 Then
        WVector1.TextMatrix(3, 2) = ZZTrabajaLote(6)
    End If
    If Val(ZZTrabajaLote(8)) <> 0 Then
        WVector1.TextMatrix(4, 2) = ZZTrabajaLote(8)
    End If
    If Val(ZZTrabajaLote(10)) <> 0 Then
        WVector1.TextMatrix(5, 2) = ZZTrabajaLote(10)
    End If
    If Val(ZZTrabajaLote(12)) <> 0 Then
        WVector1.TextMatrix(6, 2) = ZZTrabajaLote(12)
    End If
    If Val(ZZTrabajaLote(14)) <> 0 Then
        WVector1.TextMatrix(7, 2) = ZZTrabajaLote(14)
    End If
    If Val(ZZTrabajaLote(16)) <> 0 Then
        WVector1.TextMatrix(8, 2) = ZZTrabajaLote(16)
    End If
    If Val(ZZTrabajaLote(18)) <> 0 Then
        WVector1.TextMatrix(9, 2) = ZZTrabajaLote(18)
    End If
    If Val(ZZTrabajaLote(20)) <> 0 Then
        WVector1.TextMatrix(10, 2) = ZZTrabajaLote(20)
    End If
    If Val(ZZTrabajaLote(22)) <> 0 Then
        WVector1.TextMatrix(11, 2) = ZZTrabajaLote(22)
    End If
    If Val(ZZTrabajaLote(24)) <> 0 Then
        WVector1.TextMatrix(12, 2) = ZZTrabajaLote(24)
    End If
    
    If Val(ZZTrabajaLote(31)) <> 0 Then
        WVector1.TextMatrix(1, 3) = ZZTrabajaLote(31)
    End If
    If Val(ZZTrabajaLote(32)) <> 0 Then
        WVector1.TextMatrix(1, 5) = ZZTrabajaLote(32)
    End If
    If Val(ZZTrabajaLote(33)) <> 0 Then
        WVector1.TextMatrix(2, 3) = ZZTrabajaLote(33)
    End If
    If Val(ZZTrabajaLote(34)) <> 0 Then
        WVector1.TextMatrix(2, 5) = ZZTrabajaLote(34)
    End If
    If Val(ZZTrabajaLote(35)) <> 0 Then
        WVector1.TextMatrix(3, 3) = ZZTrabajaLote(35)
    End If
    If Val(ZZTrabajaLote(36)) <> 0 Then
        WVector1.TextMatrix(3, 5) = ZZTrabajaLote(36)
    End If
    If Val(ZZTrabajaLote(37)) <> 0 Then
        WVector1.TextMatrix(4, 3) = ZZTrabajaLote(37)
    End If
    If Val(ZZTrabajaLote(38)) <> 0 Then
        WVector1.TextMatrix(4, 5) = ZZTrabajaLote(38)
    End If
    If Val(ZZTrabajaLote(39)) <> 0 Then
        WVector1.TextMatrix(5, 3) = ZZTrabajaLote(39)
    End If
    If Val(ZZTrabajaLote(40)) <> 0 Then
        WVector1.TextMatrix(5, 5) = ZZTrabajaLote(40)
    End If
    If Val(ZZTrabajaLote(41)) <> 0 Then
        WVector1.TextMatrix(6, 3) = ZZTrabajaLote(41)
    End If
    If Val(ZZTrabajaLote(42)) <> 0 Then
        WVector1.TextMatrix(6, 5) = ZZTrabajaLote(42)
    End If
    If Val(ZZTrabajaLote(43)) <> 0 Then
        WVector1.TextMatrix(7, 3) = ZZTrabajaLote(43)
    End If
    If Val(ZZTrabajaLote(44)) <> 0 Then
        WVector1.TextMatrix(7, 5) = ZZTrabajaLote(44)
    End If
    If Val(ZZTrabajaLote(45)) <> 0 Then
        WVector1.TextMatrix(8, 3) = ZZTrabajaLote(45)
    End If
    If Val(ZZTrabajaLote(46)) <> 0 Then
        WVector1.TextMatrix(8, 5) = ZZTrabajaLote(46)
    End If
    If Val(ZZTrabajaLote(47)) <> 0 Then
        WVector1.TextMatrix(9, 3) = ZZTrabajaLote(47)
    End If
    If Val(ZZTrabajaLote(48)) <> 0 Then
        WVector1.TextMatrix(9, 5) = ZZTrabajaLote(48)
    End If
    If Val(ZZTrabajaLote(49)) <> 0 Then
        WVector1.TextMatrix(10, 3) = ZZTrabajaLote(49)
    End If
    If Val(ZZTrabajaLote(50)) <> 0 Then
        WVector1.TextMatrix(10, 5) = ZZTrabajaLote(50)
    End If
    If Val(ZZTrabajaLote(51)) <> 0 Then
        WVector1.TextMatrix(11, 3) = ZZTrabajaLote(51)
    End If
    If Val(ZZTrabajaLote(52)) <> 0 Then
        WVector1.TextMatrix(11, 5) = ZZTrabajaLote(52)
    End If
    If Val(ZZTrabajaLote(53)) <> 0 Then
        WVector1.TextMatrix(12, 3) = ZZTrabajaLote(53)
    End If
    If Val(ZZTrabajaLote(54)) <> 0 Then
        WVector1.TextMatrix(12, 5) = ZZTrabajaLote(54)
    End If
    
    If Val(ZZTrabajaLote(61)) <> 0 Then
        WVector1.TextMatrix(1, 6) = ZZTrabajaLote(61)
    End If
    If Val(ZZTrabajaLote(62)) <> 0 Then
        WVector1.TextMatrix(2, 6) = ZZTrabajaLote(62)
    End If
    If Val(ZZTrabajaLote(63)) <> 0 Then
        WVector1.TextMatrix(3, 6) = ZZTrabajaLote(63)
    End If
    If Val(ZZTrabajaLote(64)) <> 0 Then
        WVector1.TextMatrix(4, 6) = ZZTrabajaLote(64)
    End If
    If Val(ZZTrabajaLote(65)) <> 0 Then
        WVector1.TextMatrix(5, 6) = ZZTrabajaLote(65)
    End If
    If Val(ZZTrabajaLote(66)) <> 0 Then
        WVector1.TextMatrix(6, 6) = ZZTrabajaLote(66)
    End If
    If Val(ZZTrabajaLote(67)) <> 0 Then
        WVector1.TextMatrix(7, 6) = ZZTrabajaLote(67)
    End If
    If Val(ZZTrabajaLote(68)) <> 0 Then
        WVector1.TextMatrix(8, 6) = ZZTrabajaLote(68)
    End If
    If Val(ZZTrabajaLote(69)) <> 0 Then
        WVector1.TextMatrix(9, 6) = ZZTrabajaLote(69)
    End If
    If Val(ZZTrabajaLote(70)) <> 0 Then
        WVector1.TextMatrix(10, 6) = ZZTrabajaLote(70)
    End If
    If Val(ZZTrabajaLote(71)) <> 0 Then
        WVector1.TextMatrix(11, 6) = ZZTrabajaLote(71)
    End If
    If Val(ZZTrabajaLote(72)) <> 0 Then
        WVector1.TextMatrix(12, 6) = ZZTrabajaLote(72)
    End If
    
    
    
    If Val(ZZTrabajaLote(81)) <> 0 Then
        WVector1.TextMatrix(1, 7) = ZZTrabajaLote(81)
    End If
    If Val(ZZTrabajaLote(82)) <> 0 Then
        WVector1.TextMatrix(2, 7) = ZZTrabajaLote(82)
    End If
    If Val(ZZTrabajaLote(83)) <> 0 Then
        WVector1.TextMatrix(3, 7) = ZZTrabajaLote(83)
    End If
    If Val(ZZTrabajaLote(84)) <> 0 Then
        WVector1.TextMatrix(4, 7) = ZZTrabajaLote(84)
    End If
    If Val(ZZTrabajaLote(85)) <> 0 Then
        WVector1.TextMatrix(5, 7) = ZZTrabajaLote(85)
    End If
    If Val(ZZTrabajaLote(86)) <> 0 Then
        WVector1.TextMatrix(6, 7) = ZZTrabajaLote(86)
    End If
    If Val(ZZTrabajaLote(87)) <> 0 Then
        WVector1.TextMatrix(7, 7) = ZZTrabajaLote(87)
    End If
    If Val(ZZTrabajaLote(88)) <> 0 Then
        WVector1.TextMatrix(8, 7) = ZZTrabajaLote(88)
    End If
    If Val(ZZTrabajaLote(89)) <> 0 Then
        WVector1.TextMatrix(9, 7) = ZZTrabajaLote(89)
    End If
    If Val(ZZTrabajaLote(90)) <> 0 Then
        WVector1.TextMatrix(10, 7) = ZZTrabajaLote(90)
    End If
    If Val(ZZTrabajaLote(91)) <> 0 Then
        WVector1.TextMatrix(11, 7) = ZZTrabajaLote(91)
    End If
    If Val(ZZTrabajaLote(92)) <> 0 Then
        WVector1.TextMatrix(12, 7) = ZZTrabajaLote(92)
    End If
    
    
    ZZLugar = 0
    For Cicla = 31 To 53 Step 2
        ZZLugar = ZZLugar + 1
        spEnvase = "ConsultaEnvases " + "'" + ZZTrabajaLote(Cicla) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WVector1.TextMatrix(ZZLugar, 4) = Left$(Trim(rstEnvase!Abreviatura), 7)
            rstEnvase.Close
                Else
            WVector1.TextMatrix(ZZLugar, 4) = ""
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    Call Suma_Lote
    WVector1.Row = 1
    WVector1.Col = 1
    Rem Call Cantidad_KeyPress(13)
    
    Rem Call StartEdit
     
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub



Private Sub Suma_Lote()
    Suma = 0
    For Ciclo = 1 To 12
        If Trim(WVector1.TextMatrix(Ciclo, 1)) <> "" Then
            Suma = Suma + Val(WVector1.TextMatrix(Ciclo, 2))
        End If
    Next Ciclo
    Asignada.Text = Str$(Suma)
    Diferencia.Text = Str$(Val(Cantidad.Text) - Val(Asignada.Text))
    Asignada.Text = Pusing("###,###.##", Asignada.Text)
    Diferencia.Text = Pusing("###,###.##", Diferencia.Text)
End Sub

Private Sub Verifica_Lote()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    ZZRestriccionI = 0
    ZZRestriccionII = 0

    spCliente = "ConsultaCliente " + "'" + WPasaCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZZRestriccionI = IIf(IsNull(rstCliente!Restriccion), "0", rstCliente!Restriccion)
        rstCliente.Close
    End If
    
    Call Conecta_Empresa

    WEstado = "N"
    Suma = 0
    XTerminado = Producto.Text
    
    For Ciclo = 1 To 12
        If Trim(WVector1.TextMatrix(Ciclo, 1)) <> "" Then
            Suma = Suma + Val(WVector1.TextMatrix(Ciclo, 2))
        End If
    Next Ciclo
        
    Asignada.Text = Str$(Suma)
    If Val(Asignada.Text) = Val(Cantidad.Text) Then
        WEstado = "S"
            Else
        Rem m$ = "Las cantidades asignadas no concuerdan con las cantidades a facturar"
        Rem A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
    End If
    
    If WEstado = "S" Then
    
        Erase ControlLote
        For Ciclo = 1 To 12
            ControlLote(Ciclo, 1) = WVector1.TextMatrix(Ciclo, 1)
            ControlLote(Ciclo, 2) = WVector1.TextMatrix(Ciclo, 2)
        Next Ciclo
    
        For Ciclo1 = 1 To 12
            If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                For Ciclo2 = 1 To 12
                    If Ciclo1 <> Ciclo2 Then
                        If Val(ControlLote(Ciclo1, 1)) = Val(ControlLote(Ciclo2, 1)) <> 0 Then
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem dada
                            Rem m$ = "A asignado una misma partida 2 veces"
                            Rem a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            Rem WEstado = "N"
                            Rem Exit For
                        End If
                    End If
                Next Ciclo2
            End If
            If WEstado = "N" Then
                Exit For
            End If
        Next Ciclo1
        
    End If

    If WEstado = "S" Then
    
        Erase ControlLote
        For Ciclo = 1 To 12
            ControlLote(Ciclo, 1) = WVector1.TextMatrix(Ciclo, 1)
            ControlLote(Ciclo, 2) = WVector1.TextMatrix(Ciclo, 2)
        Next Ciclo
    
        For Ciclo1 = 1 To 12
    
            WLote = ControlLote(Ciclo1, 1)
            WCanti = Val(ControlLote(Ciclo1, 2))
            
            If WLote <> "" Or Val(WCanti) <> 0 Then
            
                If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                    WTipopro = "M"
                        Else
                    WTipopro = "T"
                End If
                
                Select Case WTipopro
                    Case "M"
                        WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                        WEntra = "N"
                        
                        If ZZRestriccionI = 1 Then
                            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                                ZZRestriccion = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                                rstArticulo.Close
                            End If
                            If ZZRestriccion = 1 Then
                                ZZRestriccionII = 1
                            End If
                        End If
                        
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0011"
                                    txtOdbc = "Empresa11"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "TA"
                                    WEmpresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                        
                        ZSql = ""
                        If Val(WLote) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Laudo"
                            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Laudo"
                            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                        End If
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            With rstLaudo
                                .MoveFirst
                                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                                ZEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                rstLaudo.Close
                            End With
                        End If
                            
                        If WEntra = "N" Then
                            ZSql = ""
                            If Val(WLote) = 0 Then
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Guia"
                                ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                                ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote + "'"
                                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                    Else
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Guia"
                                ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                                ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                            End If
                            spMovguia = ZSql
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                With rstMovguia
                                    .MoveFirst
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    Call Redondeo(WSaldo)
                                    WEntra = "S"
                                    If WSaldo < WCanti Then
                                        m$ = "La cantidad informada supera al saldo disponible"
                                        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                        WEstado = "N"
                                    End If
                                    ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                    ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                    If ZEstado = "N" Then
                                        If ZEstadoII = "V" Then
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                                Else
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        End If
                                        WEstado = "N"
                                    End If
                                    rstMovguia.Close
                                End With
                            End If
                        End If
                        
                        Call Conecta_Empresa
                        
                        If WEntra = "N" Then
                            m$ = "Partida Inexistente"
                            a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            WEstado = "N"
                        End If
                    
                    Case Else
                        WEntra = "N"
                        WControla = 0
                        ZZMeses = 0
                        
                        XEmpresa = WEmpresa
                        Select Case Val(WEmpresa)
                            Case 1, 3, 5, 6, 7, 10, 11
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZZMeses = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                            WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                            rstTerminado.Close
                        End If
                        
                        Call Conecta_Empresa
                
                        If WControla = 0 Then
                        
                            XEmpresa = WEmpresa
                            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                Select Case WTipoPedido
                                    Case "PG", "CO"
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case "FA"
                                        WEmpresa = "0011"
                                        txtOdbc = "Empresa11"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case "TA"
                                        WEmpresa = "0003"
                                        txtOdbc = "Empresa03"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                        WEmpresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                End Select
                            End If
                        
                            XParam = "'" + WLote + "','" _
                                    + XTerminado + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                            
                                WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                                ZEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                WFechaHoja = rstHoja!Fecha
                                rstHoja.Close
                                
                            End If
                            
                    
                            If WEntra = "N" Then
                                XParam = "'" + XTerminado + "','" _
                                            + WLote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    Call Redondeo(WSaldo)
                                    WEntra = "S"
                                    If WSaldo < WCanti Then
                                        m$ = "La cantidad informada supera al saldo disponible"
                                        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                        WEstado = "N"
                                    End If
                                    ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                    ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                    If ZEstado = "N" Then
                                        If ZEstadoII = "V" Then
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                                Else
                                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        End If
                                        WEstado = "N"
                                    End If
                                    rstMovguia.Close
                                End If
                            End If
                                    
                            Call Conecta_Empresa
                       
                       
                       
                    
                            Rem **********************
                            Rem verifica que si el cliente tiene
                            Rem restricciones que el rpoducto no las tenga
                            Rem
                            Rem*************************************
                       
                            If ZZRestriccionI = 1 Then
                       
                                XEmpresa = WEmpresa
                                Erase ZZVerifica
                                ZZLugarVeri = 0
        
                                Select Case Val(WEmpresa)
                                    Case 1, 3, 5, 6, 7, 10, 11
                                        Empe(1, 1) = "0001"
                                        Empe(1, 2) = "Empresa01"
                                        Empe(2, 1) = "0003"
                                        Empe(2, 2) = "Empresa03"
                                        Empe(3, 1) = "0005"
                                        Empe(3, 2) = "Empresa05"
                                        Empe(4, 1) = "0006"
                                        Empe(4, 2) = "Empresa06"
                                        Empe(5, 1) = "0007"
                                        Empe(5, 2) = "Empresa07"
                                        Empe(6, 1) = "0010"
                                        Empe(6, 2) = "Empresa10"
                                        Empe(7, 1) = "0011"
                                        Empe(7, 2) = "Empresa11"
                                        ZZHasta = 7
                                    Case Else
                                        Empe(1, 1) = "0002"
                                        Empe(1, 2) = "Empresa02"
                                        Empe(2, 1) = "0004"
                                        Empe(2, 2) = "Empresa04"
                                        Empe(3, 1) = "0008"
                                        Empe(3, 2) = "Empresa08"
                                        Empe(4, 1) = "0009"
                                        Empe(4, 2) = "Empresa09"
                                        ZZHasta = 4
                                End Select
                                
                                For CiclaEmpresa = 1 To ZZHasta
                        
                                    WEmpresa = Empe(CiclaEmpresa, 1)
                                    txtOdbc = Empe(CiclaEmpresa, 2)
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    
                                    spHoja = "ListaHoja " + "'" + WLote + "'"
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        With rstHoja
                                            .MoveFirst
                                            Do
                                                If .EOF = False Then
                                                    ZZLugarVeri = ZZLugarVeri + 1
                                                    ZZVerifica(ZZLugarVeri, 1) = rstHoja!Tipo
                                                    If UCase(rstHoja!Tipo) = "M" Then
                                                        ZZVerifica(ZZLugarVeri, 2) = rstHoja!Articulo
                                                            Else
                                                        ZZVerifica(ZZLugarVeri, 2) = rstHoja!Terminado
                                                    End If
                                                    .MoveNext
                                                        Else
                                                    Exit Do
                                                End If
                                            Loop
                                        End With
                                        rstHoja.Close
                                    End If
                            
                                Next CiclaEmpresa
                                
                                For CicloVeri = 1 To ZZLugarVeri
                                
                                    ZZTipoVeri = ZZVerifica(CicloVeri, 1)
                                    
                                    If UCase(ZZTipoVeri) = "M" Then
                                    
                                        ZZArtiVeri = ZZVerifica(CicloVeri, 2)
                                        
                                        spArticulo = "ConsultaArticulo " + "'" + ZZArtiVeri + "'"
                                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstArticulo.RecordCount > 0 Then
                                            ZZRestriccion = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                                            rstArticulo.Close
                                        End If
                                        If ZZRestriccion = 1 Then
                                            ZZRestriccionII = 1
                                        End If
                                        
                                            Else
                                            
                                        ZZTermiVeri = ZZVerifica(CicloVeri, 2)
                                                                
                                        spTerminado = "ConsultaTerminado " + "'" + ZZTermiVeri + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZZRestriccion = IIf(IsNull(rstTerminado!Restriccion), "0", rstTerminado!Restriccion)
                                            rstTerminado.Close
                                        End If
                                        If ZZRestriccion = 1 Then
                                            ZZRestriccionII = 1
                                        End If
                                        
                                    End If
                                    
                                Next CicloVeri
                                
                                If ZZRestriccionII = 1 Then
                                    m$ = "El cliente posee restriccion para los productos" + Chr$(13) + _
                                         "y algun componente de esta partida lo posee"
                                    G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                                    WEstado = "N"
                                End If
                            
                            End If
                       
                       
                    
                            Rem **********************
                            Rem Se habilita el bloqueo de actualizacion por vencimiento 01-10-2012
                            Rem
                            Rem*************************************
                            
                            If ZZMeses <> 0 Then
                            
                                XEmpresa = WEmpresa
                                ZZFechaActual = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
        
                                Select Case Val(WEmpresa)
                                    Case 1, 3, 5, 6, 7, 10, 11
                                        Empe(1, 1) = "0001"
                                        Empe(1, 2) = "Empresa01"
                                        Empe(2, 1) = "0003"
                                        Empe(2, 2) = "Empresa03"
                                        Empe(3, 1) = "0005"
                                        Empe(3, 2) = "Empresa05"
                                        Empe(4, 1) = "0006"
                                        Empe(4, 2) = "Empresa06"
                                        Empe(5, 1) = "0007"
                                        Empe(5, 2) = "Empresa07"
                                        Empe(6, 1) = "0010"
                                        Empe(6, 2) = "Empresa10"
                                        Empe(7, 1) = "0011"
                                        Empe(7, 2) = "Empresa11"
                                        ZZHasta = 7
                                    Case Else
                                        Empe(1, 1) = "0002"
                                        Empe(1, 2) = "Empresa02"
                                        Empe(2, 1) = "0004"
                                        Empe(2, 2) = "Empresa04"
                                        Empe(3, 1) = "0008"
                                        Empe(3, 2) = "Empresa08"
                                        Empe(4, 1) = "0009"
                                        Empe(4, 2) = "Empresa09"
                                        ZZHasta = 4
                                End Select
        
                                ZZZZRenglon = 0
                                ZZZZCantidadLote = 0
                                
                                For CiclaEmpresa = 1 To ZZHasta
                        
                                    WEmpresa = Empe(CiclaEmpresa, 1)
                                    txtOdbc = Empe(CiclaEmpresa, 2)
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    
                                    spHoja = "ListaHoja " + "'" + WLote + "'"
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        With rstHoja
                                            .MoveFirst
                                            Do
                                                If .EOF = False Then
                                                    ZZZZRenglon = ZZZZRenglon + 1
                                                    ZZZZCantidadLote = rstHoja!Canti1
                                                    ZZZZCantidad = rstHoja!Cantidad
                                                    ZZZZTipo = rstHoja!Tipo
                                                    .MoveNext
                                                        Else
                                                    Exit Do
                                                End If
                                            Loop
                                        End With
                                        rstHoja.Close
                                    End If
                            
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Hoja"
                                    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + WLote + "'"
                                    ZSql = ZSql + " and Hoja.Producto = " + "'" + XTerminado + "'"
                                    spHoja = ZSql
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                                        ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                                        ZZFecha = rstHoja!Fecha
                                        rstHoja.Close
                                        Exit For
                                    End If
                                    
                                Next CiclaEmpresa
                                
                                Call Conecta_Empresa
                                
                                Rem VERIFICA EL 75%
                                
                                If Val(ZZRevalida) <> 0 Then
                                
                                    WVida = Int(Val(ZZMesesRevalida) * 0.75)
                                    WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                                    WAno = Val(Right$(ZZFechaRevalida, 4))
                                    
                                        Else
                                        
                                    WVida = Int(Val(ZZMeses) * 0.75)
                                    WMes = Val(Mid$(ZZFecha, 4, 2))
                                    WAno = Val(Right$(ZZFecha, 4))
                                        
                                End If
                                
                                For Ciclo = 1 To WVida
                                    WMes = WMes + 1
                                    If WMes > 12 Then
                                        WAno = WAno + 1
                                        WMes = 1
                                    End If
                                Next Ciclo
                                ZMes = Str$(WMes)
                                ZAno = Str$(WAno)
                                Call Ceros(ZMes, 2)
                                Call Ceros(ZAno, 4)
                                ZZOrdVto = ZAno + ZMes + "01"
                                
                                If ZZOrdVto < ZZFechaActual Then
                                    WMarcaVencida = "S"
                                End If
                                
                                Rem VERIFICA EL 100%
                                
                                If Val(ZZRevalida) <> 0 Then
                                
                                    WVida = Int(Val(ZZMesesRevalida))
                                    WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                                    WAno = Val(Right$(ZZFechaRevalida, 4))
                                    
                                        Else
                                        
                                    WVida = Int(Val(ZZMeses))
                                    WMes = Val(Mid$(ZZFecha, 4, 2))
                                    WAno = Val(Right$(ZZFecha, 4))
                                        
                                End If
                                
                                For Ciclo = 1 To WVida
                                    WMes = WMes + 1
                                    If WMes > 12 Then
                                        WAno = WAno + 1
                                        WMes = 1
                                    End If
                                Next Ciclo
                                ZMes = Str$(WMes)
                                ZAno = Str$(WAno)
                                Call Ceros(ZMes, 2)
                                Call Ceros(ZAno, 4)
                                ZZOrdVto = ZAno + ZMes + "01"
                                
                                If ZZOrdVto < ZZFechaActual Then
                                    WMarcaVencida = "V"
                                End If
    
                                Rem If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                Rem     If ZZZZRenglon = 1 And ZZZZCantidad = ZZZZCantidadLote And ZZZZTipo = "M" Then
                                Rem             Else
                                Rem         m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                rem              "Por favor comuniquese con el laboratorio para su revalida"
                                Rem         G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                                Rem         WEstado = "N"
                                Rem     End If
                                Rem End If
    
                                If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                    m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                         "Por favor comuniquese con el laboratorio para su revalida"
                                    G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                                    WEstado = "N"
                                End If
    
                           
                           
                               Rem WMes = Val(Mid$(WFechaHoja, 4, 2))
                               Rem WAno = Val(Right$(WFechaHoja, 4))
                               Rem For Ciclo = 1 To WVida
                               Rem    WMes = WMes + 1
                               Rem     If WMes > 12 Then
                               Rem         WAno = WAno + 1
                               Rem         WMes = 1
                               Rem    End If
                               Rem Next Ciclo
                               Rem XMes = Str$(WMes)
                               Rem XAno = Str$(WAno)
                               Rem Call Ceros(XMes, 2)
                               Rem Call Ceros(XAno, 4)
                               Rem Wvencimiento = "01/" + XMes + "/" + XAno
                           
                               Rem WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                               Rem WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
                           
                               Rem WFechaVencimiento = "01" + Mid$(Wvencimiento, 3, 10)
                               Rem WFechaVencimientoOrd = Right$(WFechaVencimiento, 4) + Mid$(WFechaVencimiento, 4, 2) + Left$(WFechaVencimiento, 2)
                           
                               Rem Pasa = "S"
                               Rem If WFechaActualOrd >= WFechaVencimientoOrd Then
                               Rem      Pasa = "N"
                               Rem         Else
                               Rem      Meses = 0
                               Rem      WMes = Val(Mid$(WFechaActual, 4, 2))
                               Rem      WAno = Val(Right$(WFechaActual, 4))
                               Rem      Do
                               Rem          Meses = Meses + 1
                               Rem          WMes = WMes + 1
                               Rem          If WMes > 12 Then
                               Rem              WAno = WAno + 1
                               Rem              WMes = 1
                               Rem          End If
                               Rem          XMes = Str$(WMes)
                               Rem          XAno = Str$(WAno)
                               Rem          Call Ceros(XMes, 2)
                               Rem          Call Ceros(XAno, 4)
                               Rem          WCompara = "01/" + XMes + "/" + XAno
                               Rem          If WCompara = WFechaVencimiento Then
                               Rem             Exit Do
                               Rem          End If
                               Rem      Loop
                               Rem      If Meses <= 12 Then
                               Rem          Pasa = "N"
                               Rem      End If
                               Rem  End If
                           
                               Rem  If Pasa = "N" Then
                               Rem      m$ = "EL Producto tiene menos de un año de vida util"
                               Rem      G% = MsgBox(m$, 0, "Actualizacion de Pedido")
                               Rem      WEstado = "N"
                               Rem  End If
                           
                            End If
                    
                                Else
                                
                            WEntra = "S"
                            
                        End If
                        
                        If WEntra = "N" Then
                            m$ = "Partida Inexistente"
                            a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            WEstado = "N"
                        End If
                    
                End Select
            
            End If
            
        Next Ciclo1

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
        Call Suma_Lote
        If WVector1.Col = 1 Then
            If WVector1.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WControlII = "N"
                    Call Confirma_Click
                End If
            End If
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
            If WControlII = "" Then
                Call Control_Campo
                If WControl = "S" Then
                    Call Control_wvector1
                End If
                Call StartEdit
            End If

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
            If WControlII = "" Then
                If WControl = "S" Then
                    Call Control_wvector1
                End If
                Call StartEdit
            End If
    
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
            If WControlII = "" Then
                If WControl = "S" Then
                    Call Control_wvector1
                End If
                Call StartEdit
            End If

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
            XTerminado = Producto.Text
            WLote = WVector1.Text
            WSaldo = 0
            WEntra = ""
            Call Verifica_Articulo
                
            If WEntra = "S" Then
                WVector1.TextMatrix(WVector1.Row, 8) = Str$(WSaldo)
                    Else
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case Val(WBuscaEmpresa)
                        Case 1
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta I"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 11
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta VII (FARMA)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 7
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta V"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case Else
                    End Select
                End If
                WControl = "N"
            End If
        
        Case 2
            XTerminado = Producto.Text
            WLote = WVector1.TextMatrix(WVector1.Row, 1)
            WSaldo = 0
            WEntra = ""
            Call Verifica_Articulo
            If WSaldo >= Val(WVector1.Text) Then
                    Else
                m$ = XTerminado + " Cantidad Insuficiente Stock : " + Str$(WSaldo)
                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                WControl = "N"
            End If
            
        Case 3
            If Val(WVector1.Text) <> 0 Then
        
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
        
                spEnvases = "ConsultaEnvases " + "'" + WVector1.Text + "'"
                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvases.RecordCount > 0 Then
                    WVector1.Col = 4
                    WVector1.Text = rstEnvases!Abreviatura
                    rstEnvases.Close
                        Else
                    WControl = "N"
                End If
            
                Call Conecta_Empresa
                
                    Else
                    
                WVector1.TextMatrix(WVector1.Row, 4) = ""
                WVector1.TextMatrix(WVector1.Row, 5) = ""
                WVector1.TextMatrix(WVector1.Row, 6) = ""
                WVector1.TextMatrix(WVector1.Row, 7) = ""
                
                If WVector1.Row < 12 Then
                    WVector1.Row = WVector1.Row + 1
                End If
                WVector1.Col = 1
                
                WControl = "N"
                    
            End If
            
        Case 5
            If Val(WEmpresa) = 1 Then
                Select Case Val(WVector1.TextMatrix(WVector1.Row, 3))
                    Case 20, 21, 22, 23, 24, 25, 26, 28, 30
                        WVector1.TextMatrix(WVector1.Row, 6) = WVector1.Text
                    Case Else
                End Select
            End If
        
            
        Case Else
            WVector1.Col = XColumna
    End Select
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
    WVector1.Cols = 9
    WVector1.FixedRows = 1
    WVector1.Rows = 13
    
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
                WVector1.Text = "Lote"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Envase"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Cant.Envases"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Bultos"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "P.Plastico"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
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
    
    WAncho = 340
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = WAncho

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


Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Row = 1
        WVector1.Col = 1
        Call StartEdit
    End If
End Sub

Private Sub WTexto1_DblClick()
    If WVector1.Col = 3 Then
        Call Consulta_Click
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub WTexto2_DblClick()
    If WVector1.Col = 3 Then
        Call Consulta_Click
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub ficha_Pt()

    XEmpresa = WEmpresa
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        Select Case WTipoPedido
            Case "PG", "CO"
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case "FA"
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case "TA"
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    End If


    Call Limpia_Vector2
    WTerminado = Producto.Text
    XRenglon = 0
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 Then
                Rem And rstHoja!Real <> 0 Then
                 
                    ZProducto = rstHoja!Producto
                    ZCantidad = rstHoja!Real
                    ZFecha = rstHoja!Fecha
                    ZHoja = rstHoja!Hoja
                    ZSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(ZSaldo)
                    
                    If ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                
                        WVector2.Col = 1
                        WVector2.Text = "Hoja"
                        
                        WVector2.Col = 2
                        WVector2.Text = ZHoja
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = ZHoja
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                    
                    End If
                    
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
    End If
    
    
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    ZTerminado = rstMovguia!Terminado
                    ZCantidad = rstMovguia!Cantidad
                    ZFecha = rstMovguia!Fecha
                    ZCodigo = rstMovguia!Codigo
                    ZMovi = rstMovguia!Movi
                    ZDestino = rstMovguia!Destino
                    ZTipomov = rstMovguia!Tipomov
                    WWLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    ZPartida = IIf(IsNull(rstMovguia!Partida), "", rstMovguia!Partida)
                    ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZSaldo)
                    If Val(ZCodigo) > 900000 Then
                        WWTipo = "Prestamo"
                        ZCodigo = WCodigo - 900000
                            Else
                        WWTipo = "Guia In"
                    End If
                    
                    If ZMovi = "E" And ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                
                        WVector2.Col = 1
                        WVector2.Text = WWTipo
                        
                        WVector2.Col = 2
                        WVector2.Text = ZCodigo
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = WWLote
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                        
                    End If
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
    End If
    
    
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEntdev!Marca = "X" Then
                
                        Else
                
                ZTerminado = rstEntdev!Terminado
                ZCantidad = rstEntdev!Cantidad
                ZFecha = rstEntdev!Fecha
                ZCodigo = rstEntdev!Codigo
                WWLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                ZSaldo = rstEntdev!Saldo
                Call Redondeo(ZSaldo)
                
                If ZSaldo <> 0 Then
                    
                    XRenglon = XRenglon + 1
                    WVector2.Row = XRenglon
                
                    WVector2.Col = 1
                    WVector2.Text = "Dev"
                        
                    WVector2.Col = 2
                    WVector2.Text = ZCodigo
                                               
                    WVector2.Col = 3
                    WVector2.Text = ZFecha
                        
                    WVector2.Col = 4
                    WVector2.Text = ""
                        
                    WVector2.Col = 5
                    WVector2.Text = ZCantidad
                
                    WVector2.Col = 6
                    WVector2.Text = ZSaldo
                
                    WVector2.Col = 7
                    WVector2.Text = WWLote
                    
                    WVector2.Col = 8
                    WVector2.Text = ""

                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstEntdev.Close
        
    End If
                    
    WBuscaEmpresa = WEmpresa
    Call Conecta_Empresa
    
    WVector2.Col = 1
    WVector2.Row = 1
    
    WVector2.TopRow = 1
    
End Sub


Private Sub Limpia_Vector2()

    WVector2.Height = 6095
    WVector2.Left = 120
    WVector2.Top = 1050
    WVector2.Width = 12000

    WVector2.Clear
    WVector2.Font.Bold = True
    
    WVector2.FixedCols = 1
    WVector2.Cols = 12
    WVector2.FixedRows = 1
    WVector2.Rows = 5001
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Tipo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector2.Text = "Numero"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1500
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector2.Text = "Orden"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector2.Text = "Cantidad"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector2.Text = "Saldo"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 1400
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Envase"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WVector2.Text = "Cant.Ped."
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 11
                WVector2.Text = "Disponible"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
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
    
    WVector2.Visible = True
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Click()
    If Trim(WVector2.TextMatrix(WVector2.Row, 7)) <> "" Then
        WVector1.TextMatrix(WVector1.Row, 1) = WVector2.TextMatrix(WVector2.Row, 7)
    End If
    WVector2.Visible = False
    Call StartEdit
End Sub

Private Sub Verifica_Articulo()
    
    Rem DADA
    Rem DADA
    Rem DADA
    Rem DADA
    Rem DADA
    Rem DADA
    
    
    If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
        WTipopro = "M"
            Else
        WTipopro = "T"
    End If
    
    WEstado = ""
    WBuscaEmpresa = ""
    
    Select Case WTipopro
        Case "M"
            WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
            WEntra = "N"
            
            XEmpresa = WEmpresa
            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                Select Case WTipoPedido
                    Case "PG", "CO"
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case "FA"
                        WEmpresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case "TA"
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
            End If
            
            If Val(WLote) = 0 Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
                ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    Else
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote + "'"
                ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
            End If
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                With rstLaudo
                    .MoveFirst
                    
                    WEntra = "S"
                    
                    WEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                    If WEstado <> "N" Then
                        WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            Else
                        WEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                        If WEstadoII = "V" Then
                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                Else
                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        End If
                        WSaldo = 0
                    End If
                    
                    rstLaudo.Close
                End With
            End If
                
            If WEntra = "N" Then
                If Val(WLote) = 0 Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote + "'"
                    ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        Else
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                    ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                End If
                spMovguia = ZSql
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    With rstMovguia
                        .MoveFirst
                        
                        WEntra = "S"
                        
                        WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        If WEstado <> "N" Then
                            WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Else
                            WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                            If WEstadoII = "V" Then
                                m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    Else
                                m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            End If
                            WSaldo = 0
                        End If
                        
                        rstMovguia.Close
                    End With
                End If
                
            End If
            
            WBuscaEmpresa = WEmpresa
            Call Conecta_Empresa
        
            If Val(WEmpresa) = 1 Then
        
                ZZRestriccion = 0
                ZZRestriccionI = 0
                ZZRestriccionII = 0
                
                spCliente = "ConsultaCliente " + "'" + WPasaCliente + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    ZZRestriccion = IIf(IsNull(rstCliente!Restriccion), "0", rstCliente!Restriccion)
                    rstCliente.Close
                End If
                
                If ZZRestriccion = 1 Then
            
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZZRestriccionII = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                        rstArticulo.Close
                    End If
                    If ZZRestriccionII = 1 Then
                        WSaldo = 0
                        m$ = "El cliente posee restriccion para los productos" + Chr$(13) + _
                             "y algun componente de esta partida lo posee"
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                    End If
                    
                End If
                
            End If
            
        Case Else
            WEntra = "N"
            WControla = 0
            
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
            
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            Call Conecta_Empresa
            
            If WControla = 0 Then
   
                XEmpresa = WEmpresa
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case WTipoPedido
                        Case "PG", "CO"
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case "FA"
                            WEmpresa = "0011"
                            txtOdbc = "Empresa11"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case "TA"
                            WEmpresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0007"
                            txtOdbc = "Empresa07"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                End If
            
                XParam = "'" + WLote + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                
                    WEntra = "S"
                    
                    WEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                    If WEstado <> "N" Then
                        WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            Else
                        WEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                        If WEstadoII = "V" Then
                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                Else
                            m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        End If
                        WSaldo = 0
                    End If
                    
                    WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                    If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                        m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                             "Por favor comuniquese con el laboratorio para su revalida"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        WSaldo = 0
                    End If
                    
                    rstHoja.Close
                    
                End If
        
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                    
                        
                        WEntra = "S"
                        
                        WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                        If WEstado <> "N" Then
                            WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Else
                            WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                            If WEstadoII = "V" Then
                                m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    Else
                                m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            End If
                            WSaldo = 0
                        End If
                        
                        WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                        If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                            m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                "Por favor comuniquese con el laboratorio para su revalida"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            WSaldo = 0
                        End If
                        
                     
                        rstMovguia.Close
                    End If
                End If
            
            
            
            
            
            
            
                Rem dada
                Rem dada
                Rem dada
                Rem dada
                Rem dada
                Rem dada
                Rem dada
            
                Tipopro = "PT"
                XCodigo = Val(Mid$(XTerminado, 4, 5))
                If Left$(XTerminado, 2) <> "PT" Then
                    Select Case Left$(XTerminado, 2)
                        Case "DY", "DS"
                            Tipopro = "CO"
                        Case "QC"
                            Tipopro = "FA"
                        Case Else
                            Tipopro = "PT"
                    End Select
                        Else
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        Tipopro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 12999 Then
                            Tipopro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                Tipopro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    Tipopro = "BI"
                                        Else
                                    Tipopro = "PT"
                                End If
                            End If
                        End If
                    End If
                End If
                
                If Left$(XTerminado, 2) = "YQ" Then
                    Tipopro = "PT"
                End If
                If Left$(XTerminado, 2) = "YH" Then
                    Tipopro = "PT"
                End If
                If Left$(XTerminado, 2) = "YP" Then
                    Tipopro = "PT"
                End If
                If Left$(XTerminado, 2) = "YF" Then
                    Tipopro = "FA"
                End If
                                
                                
                                
                                
                                
                                
                Rem If Tipopro <> "FA" Then
                Rem
                Rem     For ZCiclo = 1 To ZHasta
                Rem
                Rem         WEmpresa = CargaEmpresa(ZCiclo, 1)
                Rem         txtOdbc = CargaEmpresa(ZCiclo, 2)
                Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Rem
                Rem         ZZRenglon = 0
                Rem         ZZTipo = ""
                Rem         ZZTerminado = ""
                Rem         ZZArticulo = ""
                Rem         ZZCantidad = 0
                Rem         ZZCantidadLote = 0
                Rem         ZZLote = ""
                Rem
                Rem         ZZSalida = "N"
                Rem
                Rem         spHoja = "ListaHoja " + "'" + WLote + "'"
                Rem         Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                Rem         If rstHoja.RecordCount > 0 Then
                Rem             With rstHoja
                Rem                 .MoveFirst
                Rem                 Do
                Rem                     If .EOF = False Then
                Rem                         ZZRenglon = ZZRenglon + 1
                Rem                         ZZTipo = rstHoja!Tipo
                Rem                         ZZTerminado = rstHoja!Terminado
                Rem                         ZZArticulo = rstHoja!Articulo
                Rem                         ZZCantidad = rstHoja!Cantidad
                Rem                         ZZCantidadLote = rstHoja!Canti1
                Rem                         ZZLote = rstHoja!lote1
                Rem                         ZZSalida = "S"
                Rem                         .MoveNext
                Rem                             Else
                Rem                         Exit Do
                Rem                     End If
                Rem                 Loop
                Rem             End With
                Rem             rstHoja.Close
                Rem         End If
                Rem
                Rem         Call Conecta_Empresa
                Rem
                Rem         If ZZSalida = "S" Then
                Rem             Exit For
                Rem         End If
                Rem
                Rem     Next ZCiclo
                Rem
                Rem     If ZZRenglon = 1 And ZZCantidad = ZZCantidadLote And ZZTipo = "M" Then
                Rem
                Rem         ZVto = ""
                Rem         ZLaudo = ZZLote
                Rem         ZArticulo = ZZArticulo
                Rem         ZFecha = ""
                Rem         ZFechaVto = ""
                Rem
                Rem         For ZCiclo = 1 To ZHasta
                Rem
                Rem             WEmpresa = CargaEmpresa(ZCiclo, 1)
                Rem             txtOdbc = CargaEmpresa(ZCiclo, 2)
                Rem             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Rem             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Rem
                Rem             ZSql = ""
                Rem             ZSql = ZSql + "Select *"
                Rem             ZSql = ZSql + " FROM Laudo"
                Rem             ZSql = ZSql + " Where Laudo = " + "'" + Str$(ZLaudo) + "'"
                Rem             ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
                Rem             spLaudo = ZSql
                Rem             Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                Rem             If rstLaudo.RecordCount > 0 Then
                Rem                 ZFecha = rstLaudo!Fecha
                Rem                 ZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
                Rem                 rstLaudo.Close
                Rem                 Exit For
                Rem             End If
                Rem
                Rem         Next ZCiclo
                Rem
                Rem         Call Conecta_Empresa
                Rem
                Rem         ZVto = ""
                Rem         ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
                Rem         If ZFechaVto <> "" And ZFechaVto <> "  /  /    " And ZFechaVto <> "00/00/0000" Then
                Rem             Call Valida_fecha(ZFechaVto, Auxi)
                Rem             If Auxi = "S" Then
                Rem                 ZVto = ZFechaVto
                Rem             End If
                Rem         End If
                Rem
                Rem         If ZVto = "" Then
                Rem
                Rem             ZMeses = 0
                Rem             ZSql = ""
                Rem             ZSql = ZSql + "Select *"
                Rem             ZSql = ZSql + " FROM Articulo"
                Rem             ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
                Rem             spArticulo = ZSql
                Rem             Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Rem             If rstArticulo.RecordCount > 0 Then
                Rem                 ZMeses = rstArticulo!Meses
                Rem                 rstArticulo.Close
                Rem             End If
                Rem
                Rem             WMes = Val(Mid$(ZFecha, 4, 2))
                Rem             WAno = Val(Right$(ZFecha, 4))
                Rem             For ZCiclo = 1 To ZMeses
                Rem                 WMes = WMes + 1
                Rem                 If WMes > 12 Then
                Rem                     WAno = WAno + 1
                Rem                     WMes = 1
                Rem                 End If
                Rem             Next ZCiclo
                Rem
                Rem             XMes = Str$(WMes)
                Rem             XAno = Str$(WAno)
                Rem             Call Ceros(XMes, 2)
                Rem             Call Ceros(XAno, 4)
                Rem             If Val(Left$(ZFecha, 2)) <= 30 Then
                Rem                 If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
                Rem                     ZVto = "28/" + XMes + "/" + XAno
                Rem                         Else
                Rem                     ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
                Rem                 End If
                Rem                     Else
                Rem                 If Val(XMes) = 2 Then
                Rem                     ZVto = "28/" + XMes + "/" + XAno
                Rem                         Else
                Rem                     ZVto = "30/" + XMes + "/" + XAno
                Rem                 End If
                Rem             End If
                Rem
                Rem         End If
                Rem
                Rem         Rem
                Rem         Rem
                Rem         Rem verifica venciminiento
                Rem         Rem
                Rem         Rem
                Rem         Rem
                Rem
                Rem         ZZVidaUtil = 0
                Rem         spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                Rem         Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                Rem         If rstTerminado.RecordCount > 0 Then
                Rem             ZZVidaUtil = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                Rem             ZZVidaUtil = Int(ZZVidaUtil * 0.25)
                Rem             rstTerminado.Close
                Rem         End If
                Rem
                Rem         WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                Rem         WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
                Rem
                Rem         WFechaVencimiento = "01" + Mid$(ZVto, 3, 10)
                Rem         WFechaVencimientoOrd = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
                Rem
                Rem         Pasa = "S"
                Rem         If Left$(WFechaActualOrd, 6) >= Left$(WFechaVencimientoOrd, 6) Then
                Rem
                Rem             Pasa = "N"
                Rem
                Rem                 Else
                Rem
                Rem             Meses = 0
                Rem             WMes = Val(Mid$(WFechaActual, 4, 2))
                Rem             WAno = Val(Right$(WFechaActual, 4))
                Rem             Do
                Rem                 Meses = Meses + 1
                Rem                 WMes = WMes + 1
                Rem                 If WMes > 12 Then
                Rem                     WAno = WAno + 1
                Rem                     WMes = 1
                Rem                 End If
                Rem                 XMes = Str$(WMes)
                Rem                 XAno = Str$(WAno)
                Rem                 Call Ceros(XMes, 2)
                Rem                 Call Ceros(XAno, 4)
                Rem                 WCompara = "01/" + XMes + "/" + XAno
                Rem                 If WCompara = WFechaVencimiento Then
                Rem                     Exit Do
                Rem                 End If
                Rem             Loop
                Rem
                Rem             If ZZVidaUtil >= Meses Then
                Rem                 Pasa = "N"
                Rem             End If
                Rem
                Rem         End If
                Rem
                Rem         If Pasa = "N" Then
                Rem             m$ = "EL Producto tiene menos de 25% de la vida util del PT"
                Rem             G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Rem             WSaldo = 0
                Rem         End If
                Rem
                Rem     End If
                Rem
                Rem End If
                
                WBuscaEmpresa = WEmpresa
                Call Conecta_Empresa
        
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            
                    ZZRestriccion = 0
                    ZZRestriccionI = 0
                    ZZRestriccionII = 0
        
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                    spCliente = "ConsultaCliente " + "'" + WPasaCliente + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        ZZRestriccion = IIf(IsNull(rstCliente!Restriccion), "0", rstCliente!Restriccion)
                        rstCliente.Close
                    End If
                    
                    Call Conecta_Empresa
                    
                    If ZZRestriccion = 1 Then
                
                        XEmpresa = WEmpresa
                        
                        ZZLugarVeri = 0
                        Erase ZZVerifica
                        
                        For CiclaEmpresa = 1 To ZHasta
                        
                            WEmpresa = CargaEmpresa(CiclaEmpresa, 1)
                            txtOdbc = CargaEmpresa(CiclaEmpresa, 2)
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    
                            spHoja = "ListaHoja " + "'" + WLote + "'"
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                With rstHoja
                                    .MoveFirst
                                    Do
                                        If .EOF = False Then
                                            ZZLugarVeri = ZZLugarVeri + 1
                                            ZZVerifica(ZZLugarVeri, 1) = rstHoja!Tipo
                                            If UCase(rstHoja!Tipo) = "M" Then
                                                ZZVerifica(ZZLugarVeri, 2) = rstHoja!Articulo
                                                    Else
                                                ZZVerifica(ZZLugarVeri, 2) = rstHoja!Terminado
                                            End If
                                            .MoveNext
                                                Else
                                            Exit Do
                                        End If
                                    Loop
                                End With
                                rstHoja.Close
                            End If
                        
                        Next CiclaEmpresa
                        
                        For CicloVeri = 1 To ZZLugarVeri
                        
                            ZZTipoVeri = ZZVerifica(CicloVeri, 1)
                            
                            If UCase(ZZTipoVeri) = "M" Then
                            
                                ZZArtiVeri = ZZVerifica(CicloVeri, 2)
                                
                                spArticulo = "ConsultaArticulo " + "'" + ZZArtiVeri + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    ZZRestriccionI = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                                    rstArticulo.Close
                                End If
                                If ZZRestriccionI = 1 Then
                                    ZZRestriccionII = 1
                                End If
                                
                                    Else
                                    
                                ZZTermiVeri = ZZVerifica(CicloVeri, 2)
                                                        
                                spTerminado = "ConsultaTerminado " + "'" + ZZTermiVeri + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    ZZRestriccionI = IIf(IsNull(rstTerminado!Restriccion), "0", rstTerminado!Restriccion)
                                    rstTerminado.Close
                                End If
                                If ZZRestriccionI = 1 Then
                                    ZZRestriccionII = 1
                                End If
                                
                            End If
                            
                        Next CicloVeri
                                
                        If ZZRestriccionII = 1 Then
                            WSaldo = 0
                            m$ = "El cliente posee restriccion para los productos" + Chr$(13) + _
                                 "y algun componente de esta partida lo posee"
                            G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                        End If
                            
                    End If
      
                End If
        
        
                    Else
            
                WEntra = "S"
        
            End If
    End Select

End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            spEnvases = "ListaEnvases"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            With rstEnvases
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = rstEnvases!Envases
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEnvases.Close
            
            Call Conecta_Empresa
            
            
        Case Else
    End Select
            
    pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()
    
    pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            Indice = pantalla.ListIndex
            WEnvases = WIndice.List(Indice)
            spEnvases = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WTexto2.Visible = False
                WVector1.TextMatrix(WVector1.Row, 3) = rstEnvases!Envases
                WVector1.TextMatrix(WVector1.Row, 4) = rstEnvases!Descripcion
                WVector1.Col = 5
                Call StartEdit
                rstEnvases.Close
            End If
            
            Call Conecta_Empresa
            
        Case Else
    End Select
    
End Sub


Private Sub Ayuda_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

        pantalla.Clear
        WIndice.Clear
        
        WEspacios = Len(Ayuda.Text)
        
        spEnvases = "ListaEnvases"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        With rstEnvases
            .MoveFirst
            Do
                If .EOF = False Then
                    DA = Len(rstEnvases!Descripcion) - WEspacios
                    
                    For aa = 1 To DA + 1
                        If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstEnvases!Descripcion), aa, WEspacios) Then
                            IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                            pantalla.AddItem IngresaItem
                            IngresaItem = rstEnvases!Envases
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
        rstEnvases.Close
        
    End If

End Sub



