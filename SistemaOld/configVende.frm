VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgConfigVende 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreos de Atributos de Vende"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   390
   ClientWidth     =   11535
   LinkTopic       =   "Form2"
   ScaleHeight     =   8400
   ScaleWidth      =   11535
   Begin VB.TextBox Operador 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Salida 
      Caption         =   "Salida"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   7440
      Width           =   1215
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   11668
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "configVende.frx":0000
      Tab(0).ControlCount=   34
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Titulo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Titulo2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Titulo3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Titulo4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Titulo5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Titulo6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Titulo7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Titulo8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Titulo9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Titulo10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Titulo14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Titulo13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Titulo12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Titulo11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Titulo15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Titulo16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Titulo17"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Opcion1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Opcion2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Opcion3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Opcion4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Opcion5"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Opcion6"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Opcion7"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Opcion8"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Opcion9"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Opcion10"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Opcion14"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Opcion13"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Opcion12"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Opcion11"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Opcion15"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Opcion16"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Opcion17"
      Tab(0).Control(33).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "configVende.frx":001C
      Tab(1).ControlCount=   64
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Titulo101"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Titulo102"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Titulo103"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Titulo106"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Titulo105"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Titulo104"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Titulo109"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Titulo108"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Titulo107"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Titulo110"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Titulo113"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Titulo112"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Titulo111"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Titulo116"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Titulo115"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Titulo114"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Titulo119"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Titulo118"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Titulo117"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Titulo122"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Titulo121"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Titulo120"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Titulo125"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Titulo124"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Titulo123"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Titulo126"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Titulo129"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Titulo128"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Titulo127"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Titulo132"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Titulo131"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Titulo130"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Opcion101"
      Tab(1).Control(32).Enabled=   -1  'True
      Tab(1).Control(33)=   "Opcion102"
      Tab(1).Control(33).Enabled=   -1  'True
      Tab(1).Control(34)=   "Opcion103"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Opcion106"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Opcion105"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Opcion104"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Opcion109"
      Tab(1).Control(38).Enabled=   -1  'True
      Tab(1).Control(39)=   "Opcion108"
      Tab(1).Control(39).Enabled=   -1  'True
      Tab(1).Control(40)=   "Opcion107"
      Tab(1).Control(40).Enabled=   -1  'True
      Tab(1).Control(41)=   "Opcion110"
      Tab(1).Control(41).Enabled=   -1  'True
      Tab(1).Control(42)=   "Opcion113"
      Tab(1).Control(42).Enabled=   -1  'True
      Tab(1).Control(43)=   "Opcion112"
      Tab(1).Control(43).Enabled=   -1  'True
      Tab(1).Control(44)=   "Opcion111"
      Tab(1).Control(44).Enabled=   -1  'True
      Tab(1).Control(45)=   "Opcion116"
      Tab(1).Control(45).Enabled=   -1  'True
      Tab(1).Control(46)=   "Opcion115"
      Tab(1).Control(46).Enabled=   -1  'True
      Tab(1).Control(47)=   "Opcion114"
      Tab(1).Control(47).Enabled=   -1  'True
      Tab(1).Control(48)=   "Opcion119"
      Tab(1).Control(48).Enabled=   -1  'True
      Tab(1).Control(49)=   "Opcion118"
      Tab(1).Control(49).Enabled=   -1  'True
      Tab(1).Control(50)=   "Opcion117"
      Tab(1).Control(50).Enabled=   -1  'True
      Tab(1).Control(51)=   "Opcion122"
      Tab(1).Control(51).Enabled=   -1  'True
      Tab(1).Control(52)=   "Opcion121"
      Tab(1).Control(52).Enabled=   -1  'True
      Tab(1).Control(53)=   "Opcion120"
      Tab(1).Control(53).Enabled=   -1  'True
      Tab(1).Control(54)=   "Opcion125"
      Tab(1).Control(54).Enabled=   -1  'True
      Tab(1).Control(55)=   "Opcion124"
      Tab(1).Control(55).Enabled=   -1  'True
      Tab(1).Control(56)=   "Opcion123"
      Tab(1).Control(56).Enabled=   -1  'True
      Tab(1).Control(57)=   "Opcion126"
      Tab(1).Control(57).Enabled=   -1  'True
      Tab(1).Control(58)=   "Opcion129"
      Tab(1).Control(58).Enabled=   -1  'True
      Tab(1).Control(59)=   "Opcion128"
      Tab(1).Control(59).Enabled=   -1  'True
      Tab(1).Control(60)=   "Opcion127"
      Tab(1).Control(60).Enabled=   -1  'True
      Tab(1).Control(61)=   "Opcion132"
      Tab(1).Control(61).Enabled=   -1  'True
      Tab(1).Control(62)=   "Opcion131"
      Tab(1).Control(62).Enabled=   -1  'True
      Tab(1).Control(63)=   "Opcion130"
      Tab(1).Control(63).Enabled=   -1  'True
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "configVende.frx":0038
      Tab(2).ControlCount=   6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Titulo203"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Titulo202"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Titulo201"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Opcion203"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "Opcion202"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "opcion201"
      Tab(2).Control(5).Enabled=   -1  'True
      Begin VB.CheckBox Opcion130 
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
         Left            =   -65400
         TabIndex        =   106
         Top             =   4380
         Width           =   615
      End
      Begin VB.CheckBox Opcion131 
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
         Left            =   -65400
         TabIndex        =   105
         Top             =   4665
         Width           =   615
      End
      Begin VB.CheckBox Opcion132 
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
         Left            =   -65400
         TabIndex        =   104
         Top             =   4965
         Width           =   615
      End
      Begin VB.CheckBox Opcion127 
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
         Left            =   -65400
         TabIndex        =   100
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox Opcion128 
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
         Left            =   -65400
         TabIndex        =   99
         Top             =   3765
         Width           =   615
      End
      Begin VB.CheckBox Opcion129 
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
         Left            =   -65400
         TabIndex        =   98
         Top             =   4065
         Width           =   615
      End
      Begin VB.CheckBox Opcion126 
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
         Left            =   -65400
         TabIndex        =   96
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox Opcion123 
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
         Left            =   -65400
         TabIndex        =   92
         Top             =   2280
         Width           =   615
      End
      Begin VB.CheckBox Opcion124 
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
         Left            =   -65400
         TabIndex        =   91
         Top             =   2565
         Width           =   615
      End
      Begin VB.CheckBox Opcion125 
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
         Left            =   -65400
         TabIndex        =   90
         Top             =   2865
         Width           =   615
      End
      Begin VB.CheckBox Opcion120 
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
         Left            =   -65400
         TabIndex        =   86
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion121 
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
         Left            =   -65400
         TabIndex        =   85
         Top             =   1725
         Width           =   615
      End
      Begin VB.CheckBox Opcion122 
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
         Left            =   -65400
         TabIndex        =   84
         Top             =   2025
         Width           =   615
      End
      Begin VB.CheckBox Opcion117 
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
         Left            =   -65400
         TabIndex        =   80
         Top             =   540
         Width           =   615
      End
      Begin VB.CheckBox Opcion118 
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
         Left            =   -65400
         TabIndex        =   79
         Top             =   825
         Width           =   615
      End
      Begin VB.CheckBox Opcion119 
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
         Left            =   -65400
         TabIndex        =   78
         Top             =   1125
         Width           =   615
      End
      Begin VB.CheckBox Opcion114 
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
         Left            =   -70320
         TabIndex        =   74
         Top             =   4380
         Width           =   615
      End
      Begin VB.CheckBox Opcion115 
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
         Left            =   -70320
         TabIndex        =   73
         Top             =   4665
         Width           =   615
      End
      Begin VB.CheckBox Opcion116 
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
         Left            =   -70320
         TabIndex        =   72
         Top             =   4965
         Width           =   615
      End
      Begin VB.CheckBox Opcion111 
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
         Left            =   -70320
         TabIndex        =   68
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox Opcion112 
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
         Left            =   -70320
         TabIndex        =   67
         Top             =   3765
         Width           =   615
      End
      Begin VB.CheckBox Opcion113 
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
         Left            =   -70320
         TabIndex        =   66
         Top             =   4065
         Width           =   615
      End
      Begin VB.CheckBox Opcion110 
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
         Left            =   -70320
         TabIndex        =   64
         Top             =   3180
         Width           =   615
      End
      Begin VB.CheckBox Opcion107 
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
         Left            =   -70320
         TabIndex        =   60
         Top             =   2280
         Width           =   615
      End
      Begin VB.CheckBox Opcion108 
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
         Left            =   -70320
         TabIndex        =   59
         Top             =   2565
         Width           =   615
      End
      Begin VB.CheckBox Opcion109 
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
         Left            =   -70320
         TabIndex        =   58
         Top             =   2865
         Width           =   615
      End
      Begin VB.CheckBox Opcion104 
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
         Left            =   -70320
         TabIndex        =   54
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion105 
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
         Left            =   -70320
         TabIndex        =   53
         Top             =   1725
         Width           =   615
      End
      Begin VB.CheckBox Opcion106 
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
         Left            =   -70320
         TabIndex        =   52
         Top             =   2025
         Width           =   615
      End
      Begin VB.CheckBox Opcion17 
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
         Left            =   4680
         TabIndex        =   50
         Top             =   6060
         Width           =   615
      End
      Begin VB.CheckBox Opcion16 
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
         Left            =   4680
         TabIndex        =   48
         Top             =   5700
         Width           =   615
      End
      Begin VB.CheckBox Opcion15 
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
         Left            =   4680
         TabIndex        =   46
         Top             =   5340
         Width           =   615
      End
      Begin VB.CheckBox Opcion11 
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
         Left            =   4680
         TabIndex        =   41
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox Opcion12 
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
         Left            =   4680
         TabIndex        =   40
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Opcion13 
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
         Left            =   4680
         TabIndex        =   39
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox Opcion14 
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
         Left            =   4680
         TabIndex        =   38
         Top             =   5040
         Width           =   615
      End
      Begin VB.CheckBox Opcion10 
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
         Left            =   4680
         TabIndex        =   36
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Opcion9 
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
         Left            =   4680
         TabIndex        =   34
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Opcion8 
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
         Left            =   4680
         TabIndex        =   32
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Opcion7 
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
         Left            =   4680
         TabIndex        =   30
         Top             =   2580
         Width           =   615
      End
      Begin VB.CheckBox Opcion6 
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
         Left            =   4680
         TabIndex        =   28
         Top             =   2220
         Width           =   615
      End
      Begin VB.CheckBox Opcion5 
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
         Left            =   4680
         TabIndex        =   26
         Top             =   1860
         Width           =   615
      End
      Begin VB.CheckBox opcion201 
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
         Left            =   -70320
         TabIndex        =   22
         Top             =   420
         Width           =   615
      End
      Begin VB.CheckBox Opcion202 
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
         Left            =   -70320
         TabIndex        =   21
         Top             =   780
         Width           =   615
      End
      Begin VB.CheckBox Opcion203 
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
         Left            =   -70320
         TabIndex        =   20
         Top             =   1140
         Width           =   615
      End
      Begin VB.CheckBox Opcion4 
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
         Left            =   4680
         TabIndex        =   19
         Top             =   1500
         Width           =   615
      End
      Begin VB.CheckBox Opcion3 
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
         Left            =   4680
         TabIndex        =   17
         Top             =   1140
         Width           =   615
      End
      Begin VB.CheckBox Opcion2 
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
         Left            =   4680
         TabIndex        =   15
         Top             =   780
         Width           =   615
      End
      Begin VB.CheckBox Opcion1 
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
         Left            =   4680
         TabIndex        =   13
         Top             =   420
         Width           =   615
      End
      Begin VB.CheckBox Opcion103 
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
         Left            =   -70320
         TabIndex        =   11
         Top             =   1125
         Width           =   615
      End
      Begin VB.CheckBox Opcion102 
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
         Left            =   -70320
         TabIndex        =   9
         Top             =   825
         Width           =   615
      End
      Begin VB.CheckBox Opcion101 
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
         Left            =   -70320
         TabIndex        =   7
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Titulo130 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   109
         Top             =   4380
         Width           =   3975
      End
      Begin VB.Label Titulo131 
         Caption         =   "Ingreso de "
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
         Left            =   -69600
         TabIndex        =   108
         Top             =   4665
         Width           =   3855
      End
      Begin VB.Label Titulo132 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   107
         Top             =   4965
         Width           =   3975
      End
      Begin VB.Label Titulo127 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   103
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label Titulo128 
         Caption         =   "Ingreso de "
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
         Left            =   -69600
         TabIndex        =   102
         Top             =   3765
         Width           =   3855
      End
      Begin VB.Label Titulo129 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   101
         Top             =   4065
         Width           =   3975
      End
      Begin VB.Label Titulo126 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   97
         Top             =   3180
         Width           =   3975
      End
      Begin VB.Label Titulo123 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   95
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label Titulo124 
         Caption         =   "Ingreso de "
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
         Left            =   -69600
         TabIndex        =   94
         Top             =   2565
         Width           =   3855
      End
      Begin VB.Label Titulo125 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   93
         Top             =   2865
         Width           =   3975
      End
      Begin VB.Label Titulo120 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   89
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Titulo121 
         Caption         =   "Ingreso de "
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
         Left            =   -69600
         TabIndex        =   88
         Top             =   1725
         Width           =   3855
      End
      Begin VB.Label Titulo122 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   87
         Top             =   2025
         Width           =   3975
      End
      Begin VB.Label Titulo117 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   83
         Top             =   540
         Width           =   3975
      End
      Begin VB.Label Titulo118 
         Caption         =   "Ingreso de "
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
         Left            =   -69600
         TabIndex        =   82
         Top             =   825
         Width           =   3855
      End
      Begin VB.Label Titulo119 
         Caption         =   "Ingrso de "
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
         Left            =   -69600
         TabIndex        =   81
         Top             =   1125
         Width           =   3975
      End
      Begin VB.Label Titulo114 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   77
         Top             =   4380
         Width           =   3975
      End
      Begin VB.Label Titulo115 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   76
         Top             =   4665
         Width           =   3855
      End
      Begin VB.Label Titulo116 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   75
         Top             =   4965
         Width           =   3975
      End
      Begin VB.Label Titulo111 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   71
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label Titulo112 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   70
         Top             =   3765
         Width           =   3855
      End
      Begin VB.Label Titulo113 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   69
         Top             =   4065
         Width           =   3975
      End
      Begin VB.Label Titulo110 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   65
         Top             =   3180
         Width           =   3975
      End
      Begin VB.Label Titulo107 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   63
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label Titulo108 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   62
         Top             =   2565
         Width           =   3855
      End
      Begin VB.Label Titulo109 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   61
         Top             =   2865
         Width           =   3975
      End
      Begin VB.Label Titulo104 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   57
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Titulo105 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   56
         Top             =   1725
         Width           =   3855
      End
      Begin VB.Label Titulo106 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   55
         Top             =   2025
         Width           =   3975
      End
      Begin VB.Label Titulo17 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   51
         Top             =   6060
         Width           =   3855
      End
      Begin VB.Label Titulo16 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   49
         Top             =   5700
         Width           =   3855
      End
      Begin VB.Label Titulo15 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   47
         Top             =   5340
         Width           =   3855
      End
      Begin VB.Label Titulo11 
         Caption         =   "Ingreso de "
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
         Left            =   480
         TabIndex        =   45
         Top             =   3960
         Width           =   3975
      End
      Begin VB.Label Titulo12 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   44
         Top             =   4320
         Width           =   3855
      End
      Begin VB.Label Titulo13 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   43
         Top             =   4680
         Width           =   3855
      End
      Begin VB.Label Titulo14 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   42
         Top             =   5040
         Width           =   3855
      End
      Begin VB.Label Titulo10 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   37
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label Titulo9 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   35
         Top             =   3240
         Width           =   3855
      End
      Begin VB.Label Titulo8 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   33
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Titulo7 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   31
         Top             =   2580
         Width           =   3855
      End
      Begin VB.Label Titulo6 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   29
         Top             =   2220
         Width           =   3855
      End
      Begin VB.Label Titulo5 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   27
         Top             =   1860
         Width           =   3855
      End
      Begin VB.Label Titulo201 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   25
         Top             =   420
         Width           =   4095
      End
      Begin VB.Label Titulo202 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   24
         Top             =   780
         Width           =   4095
      End
      Begin VB.Label Titulo203 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   23
         Top             =   1140
         Width           =   4095
      End
      Begin VB.Label Titulo4 
         Caption         =   "Ingreso de "
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
         Left            =   480
         TabIndex        =   18
         Top             =   1500
         Width           =   3975
      End
      Begin VB.Label Titulo3 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   16
         Top             =   1140
         Width           =   3855
      End
      Begin VB.Label Titulo2 
         Caption         =   "Ingreso de "
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
         Left            =   480
         TabIndex        =   14
         Top             =   780
         Width           =   3975
      End
      Begin VB.Label Titulo1 
         Caption         =   "Ingrso de "
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
         Left            =   480
         TabIndex        =   12
         Top             =   420
         Width           =   3855
      End
      Begin VB.Label Titulo103 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   10
         Top             =   1125
         Width           =   3975
      End
      Begin VB.Label Titulo102 
         Caption         =   "Ingreso de "
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
         Left            =   -74520
         TabIndex        =   8
         Top             =   825
         Width           =   3855
      End
      Begin VB.Label Titulo101 
         Caption         =   "Ingrso de "
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
         Left            =   -74520
         TabIndex        =   6
         Top             =   540
         Width           =   3975
      End
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
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Operador"
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
      Left            =   360
      TabIndex        =   4
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgConfigVende"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOperador As Recordset
Dim spOperador As String
Dim rstAtributos As Recordset
Dim spAtributos As String
Dim XParam As String


Sub Form_Load()

    Operador.Text = ""
    DesOperador.Caption = ""
    
    Tablas.TabCaption(0) = "Novedades"
    Tablas.TabCaption(1) = "Listados"
    Tablas.TabCaption(2) = "Procesos"
    
    Rem titulo1
    
    Titulo1.Caption = "Ingreso de Vendedores por cliente"
    Titulo2.Caption = "Ingreso de Comisiones por Linea"
    Titulo3.Caption = "Ingreso de Vendedores"
    Titulo4.Caption = "Consulta de Clientes"
    Titulo5.Caption = "Ingreso de Lineas"
    Titulo6.Caption = "Consulta de Cambios"
    Titulo7.Caption = "Consulta de Precios Por Cliente"
    Titulo8.Caption = "Ingreso de Pedidos"
    Titulo9.Caption = "Emision de Remitos"
    Titulo10.Caption = "Ingreso de Devolucion de Remitos"
    Titulo11.Caption = "Ingreso de Comprobantes Varios"
    Titulo12.Caption = "Ingreso de Conceptos"
    Titulo13.Caption = "Ingreso de Pagos"
    Titulo14.Caption = "Ingreso de Cobranzas"
    Titulo15.Caption = "Consulta de Cheques"
    Titulo16.Caption = "Ingreso de Compensacion de Valores"
    Titulo17.Caption = "Ingreso de Moviimentos Varios Cuenta Vendedores"
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    Opcion8.Value = 0
    Opcion9.Value = 0
    Opcion10.Value = 0
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0
    Opcion17.Value = 0
    
    
    Rem titulo2
    
    Titulo101.Caption = "Listado de Comision "
    Titulo102.Caption = "Listado Cuenta corriente por Vendedor"
    Titulo103.Caption = "Consulta Cuenta Corriente por Vendedor"
    Titulo104.Caption = "Listado Movimiento por Vendedor"
    Titulo105.Caption = "Consulta de Cuenta Corriente de Clientes por pantalla"
    Titulo106.Caption = "Listado de Cuenta Corriente de Clientes"
    Titulo107.Caption = "Listado de Saldos de Cuenta Corriente de Clientes"
    Titulo108.Caption = "Listado de Ventas"
    Titulo109.Caption = "Listado Comparativo de Ventas   "
    Titulo110.Caption = "Listado de Pedido Pendientes"
    Titulo111.Caption = "Listado de Valores en Cartera"
    Titulo112.Caption = "Listado de Recibos"
    Titulo113.Caption = "Listado de Pago"
    Titulo114.Caption = "Listado de Caja Diaria"
    Titulo115.Caption = "Listado de Diferencia de Cambio"
    Titulo116.Caption = "Listado de Pagos por Concepto"
    Titulo117.Caption = "Listado de Cobranzas por Concepto"
    Titulo118.Caption = "Listado de Imputaciones por Concepto"
    Titulo119.Caption = "Estadistica de Ventas por Vendedor y Linea"
    Titulo120.Caption = "Estadistica de Ventas por Linea y Producto (Ind.)"
    Titulo121.Caption = "Estadistica de Ventas por Linea y Producto"
    Titulo122.Caption = "Estadistica de Ventas por Vendedor, Cliente y Linea"
    Titulo123.Caption = "Estadistica de Ventas por Cliente"
    Titulo124.Caption = "Estadistica de Ventas por Vendedor"
    Titulo125.Caption = "Estadistica de Ventas por Producto"
    Titulo126.Caption = "Estadistica de Ventas Anules por Producto"
    Titulo127.Caption = "Estadistica de Ventas Anules por Cliente"
    Titulo128.Caption = "Estadistica de Ventas InterAnules por Producto"
    Titulo129.Caption = "Estadistica de Ventas InterAnules por Producto y Cliente"
    Titulo130.Caption = "Resumen de Estadisticas de Venta"
    Titulo131.Caption = "Ranking de Ventas de Colorantes y DW Consolidado"
    Titulo132.Caption = "Estadistica de Ventas de Colorantes y DW por Cliente"
 
    Opcion101.Value = 0
    Opcion102.Value = 0
    Opcion103.Value = 0
    Opcion104.Value = 0
    Opcion105.Value = 0
    Opcion106.Value = 0
    Opcion107.Value = 0
    Opcion108.Value = 0
    Opcion109.Value = 0
    Opcion110.Value = 0
    Opcion111.Value = 0
    Opcion112.Value = 0
    Opcion113.Value = 0
    Opcion114.Value = 0
    Opcion115.Value = 0
    Opcion116.Value = 0
    Opcion117.Value = 0
    Opcion118.Value = 0
    Opcion119.Value = 0
    Opcion120.Value = 0
    Opcion121.Value = 0
    Opcion122.Value = 0
    Opcion123.Value = 0
    Opcion124.Value = 0
    Opcion125.Value = 0
    Opcion126.Value = 0
    Opcion127.Value = 0
    Opcion128.Value = 0
    Opcion129.Value = 0
    Opcion130.Value = 0
    Opcion131.Value = 0
    Opcion132.Value = 0
    
    
    Rem titulo3
    
    Titulo201.Caption = "Grabacion de Comisiones (Ventas)"
    Titulo202.Caption = "Grabacion de Comisiones (Dif. de Cambio)"
    Titulo203.Caption = "Borrado de Saldos de Cuenta Corriente (menor a $ 0.10)"
    
    opcion201.Value = 0
    Opcion202.Value = 0
    Opcion203.Value = 0
    
End Sub


Private Sub Graba_Click()

    XParam = "'" + Operador.Text + "','" _
                 + "7" + "'"
    spAtributos = "BorrarAtributos " + XParam
    Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
    
    WAtributo1 = ""
    WAtributo2 = ""
    WAtributo3 = ""
    WAtributo4 = ""
    WAtributo5 = ""
    WAtributo6 = ""
    WAtributo7 = ""
    WAtributo8 = ""
    WAtributo9 = ""
    WAtributo10 = ""
    
    
    
    If Opcion1.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion2.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion3.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion4.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion5.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion6.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion7.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion8.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion9.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion10.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion11.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion12.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion13.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion14.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion15.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion16.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion17.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    
    
    
    If Opcion101.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion102.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion103.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion104.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion105.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion106.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion107.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion108.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion109.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion110.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion111.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion112.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion113.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion114.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion115.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion116.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion117.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion118.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion119.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion120.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion121.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion122.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion123.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion124.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion125.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion126.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion127.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion128.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion129.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion130.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion131.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion132.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    
    
    
    
    If opcion201.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion202.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion203.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    
    WProceso = "7"
                                       
    XParam = "'" + Operador.Text + "','" _
                 + WProceso + "','" _
                 + WAtributo1 + "','" _
                 + WAtributo2 + "','" _
                 + WAtributo3 + "','" _
                 + WAtributo4 + "','" _
                 + WAtributo5 + "','" _
                 + WAtributo6 + "','" _
                 + WAtributo7 + "','" _
                 + WAtributo8 + "','" _
                 + WAtributo9 + "','" _
                 + WAtributo10 + "'"
                    
    spAtributos = "AltaAtributos " + XParam
    Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
    
    Operador.Text = ""
    DesOperador.Caption = ""
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    Opcion8.Value = 0
    Opcion9.Value = 0
    Opcion10.Value = 0
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0
    Opcion17.Value = 0
                    
    Opcion101.Value = 0
    Opcion102.Value = 0
    Opcion103.Value = 0
    Opcion104.Value = 0
    Opcion105.Value = 0
    Opcion106.Value = 0
    Opcion107.Value = 0
    Opcion108.Value = 0
    Opcion109.Value = 0
    Opcion110.Value = 0
    Opcion111.Value = 0
    Opcion112.Value = 0
    Opcion113.Value = 0
    Opcion114.Value = 0
    Opcion115.Value = 0
    Opcion116.Value = 0
    Opcion117.Value = 0
    Opcion118.Value = 0
    Opcion119.Value = 0
    Opcion120.Value = 0
    Opcion121.Value = 0
    Opcion122.Value = 0
    Opcion123.Value = 0
    Opcion124.Value = 0
    Opcion125.Value = 0
    Opcion126.Value = 0
    Opcion127.Value = 0
    Opcion128.Value = 0
    Opcion129.Value = 0
    Opcion130.Value = 0
    Opcion131.Value = 0
    Opcion132.Value = 0

    opcion201.Value = 0
    Opcion202.Value = 0
    Opcion203.Value = 0
    
    Operador.SetFocus
    
    Tablas.Tab = 0

End Sub

Private Sub Operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Operador.Text <> "" Then
        
            spOperador = "ConsultaOperador " + "'" + Operador.Text + "'"
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                DesOperador.Caption = rstOperador!Descripcion
                rstOperador.Close
                
                Opcion1.Value = 0
                Opcion2.Value = 0
                Opcion3.Value = 0
                Opcion4.Value = 0
                Opcion5.Value = 0
                Opcion6.Value = 0
                Opcion7.Value = 0
                Opcion8.Value = 0
                Opcion9.Value = 0
                Opcion10.Value = 0
                Opcion11.Value = 0
                Opcion12.Value = 0
                Opcion13.Value = 0
                Opcion14.Value = 0
                Opcion15.Value = 0
                Opcion16.Value = 0
                Opcion17.Value = 0
                    
                Opcion101.Value = 0
                Opcion102.Value = 0
                Opcion103.Value = 0
                Opcion104.Value = 0
                Opcion105.Value = 0
                Opcion106.Value = 0
                Opcion107.Value = 0
                Opcion108.Value = 0
                Opcion109.Value = 0
                Opcion110.Value = 0
                Opcion111.Value = 0
                Opcion112.Value = 0
                Opcion113.Value = 0
                Opcion114.Value = 0
                Opcion115.Value = 0
                Opcion116.Value = 0
                Opcion117.Value = 0
                Opcion118.Value = 0
                Opcion119.Value = 0
                Opcion120.Value = 0
                Opcion121.Value = 0
                Opcion122.Value = 0
                Opcion123.Value = 0
                Opcion124.Value = 0
                Opcion125.Value = 0
                Opcion126.Value = 0
                Opcion127.Value = 0
                Opcion128.Value = 0
                Opcion129.Value = 0
                Opcion130.Value = 0
                Opcion131.Value = 0
                Opcion132.Value = 0
        
                opcion201.Value = 0
                Opcion202.Value = 0
                Opcion203.Value = 0
                
                
                XParam = "'" + Operador.Text + "','" _
                             + "7" + "'"
                spAtributos = "ConsultaAtributo " + XParam
                Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
                If rstAtributos.RecordCount > 0 Then
                
                    Opcion1.Value = Val(Mid$(rstAtributos!atributo1, 1, 1))
                    Opcion2.Value = Val(Mid$(rstAtributos!atributo1, 2, 1))
                    Opcion3.Value = Val(Mid$(rstAtributos!atributo1, 3, 1))
                    Opcion4.Value = Val(Mid$(rstAtributos!atributo1, 4, 1))
                    Opcion5.Value = Val(Mid$(rstAtributos!atributo1, 5, 1))
                    Opcion6.Value = Val(Mid$(rstAtributos!atributo1, 6, 1))
                    Opcion7.Value = Val(Mid$(rstAtributos!atributo1, 7, 1))
                    Opcion8.Value = Val(Mid$(rstAtributos!atributo1, 8, 1))
                    Opcion9.Value = Val(Mid$(rstAtributos!atributo1, 9, 1))
                    Opcion10.Value = Val(Mid$(rstAtributos!atributo1, 10, 1))
                    Opcion11.Value = Val(Mid$(rstAtributos!atributo1, 11, 1))
                    Opcion12.Value = Val(Mid$(rstAtributos!atributo1, 12, 1))
                    Opcion13.Value = Val(Mid$(rstAtributos!atributo1, 13, 1))
                    Opcion14.Value = Val(Mid$(rstAtributos!atributo1, 14, 1))
                    Opcion15.Value = Val(Mid$(rstAtributos!atributo1, 15, 1))
                    Opcion16.Value = Val(Mid$(rstAtributos!atributo1, 16, 1))
                    Opcion17.Value = Val(Mid$(rstAtributos!atributo1, 17, 1))
                    
                    Opcion101.Value = Val(Mid$(rstAtributos!atributo2, 1, 1))
                    Opcion102.Value = Val(Mid$(rstAtributos!atributo2, 2, 1))
                    Opcion103.Value = Val(Mid$(rstAtributos!atributo2, 3, 1))
                    Opcion104.Value = Val(Mid$(rstAtributos!atributo2, 4, 1))
                    Opcion105.Value = Val(Mid$(rstAtributos!atributo2, 5, 1))
                    Opcion106.Value = Val(Mid$(rstAtributos!atributo2, 6, 1))
                    Opcion107.Value = Val(Mid$(rstAtributos!atributo2, 7, 1))
                    Opcion108.Value = Val(Mid$(rstAtributos!atributo2, 8, 1))
                    Opcion109.Value = Val(Mid$(rstAtributos!atributo2, 9, 1))
                    Opcion110.Value = Val(Mid$(rstAtributos!atributo2, 10, 1))
                    Opcion111.Value = Val(Mid$(rstAtributos!atributo2, 11, 1))
                    Opcion112.Value = Val(Mid$(rstAtributos!atributo2, 12, 1))
                    Opcion113.Value = Val(Mid$(rstAtributos!atributo2, 13, 1))
                    Opcion114.Value = Val(Mid$(rstAtributos!atributo2, 14, 1))
                    Opcion115.Value = Val(Mid$(rstAtributos!atributo2, 15, 1))
                    Opcion116.Value = Val(Mid$(rstAtributos!atributo2, 16, 1))
                    Opcion117.Value = Val(Mid$(rstAtributos!atributo2, 17, 1))
                    Opcion118.Value = Val(Mid$(rstAtributos!atributo2, 18, 1))
                    Opcion119.Value = Val(Mid$(rstAtributos!atributo2, 19, 1))
                    Opcion120.Value = Val(Mid$(rstAtributos!atributo2, 20, 1))
                    Opcion121.Value = Val(Mid$(rstAtributos!atributo2, 21, 1))
                    Opcion122.Value = Val(Mid$(rstAtributos!atributo2, 22, 1))
                    Opcion123.Value = Val(Mid$(rstAtributos!atributo2, 23, 1))
                    Opcion124.Value = Val(Mid$(rstAtributos!atributo2, 24, 1))
                    Opcion125.Value = Val(Mid$(rstAtributos!atributo2, 25, 1))
                    Opcion126.Value = Val(Mid$(rstAtributos!atributo2, 26, 1))
                    Opcion127.Value = Val(Mid$(rstAtributos!atributo2, 27, 1))
                    Opcion128.Value = Val(Mid$(rstAtributos!atributo2, 28, 1))
                    Opcion129.Value = Val(Mid$(rstAtributos!atributo2, 29, 1))
                    Opcion130.Value = Val(Mid$(rstAtributos!atributo2, 30, 1))
                    Opcion131.Value = Val(Mid$(rstAtributos!atributo2, 31, 1))
                    Opcion132.Value = Val(Mid$(rstAtributos!atributo2, 32, 1))
                    
                    opcion201.Value = Val(Mid$(rstAtributos!atributo3, 1, 1))
                    Opcion202.Value = Val(Mid$(rstAtributos!atributo3, 2, 1))
                    Opcion203.Value = Val(Mid$(rstAtributos!atributo3, 3, 1))
                    
                    rstAtributos.Close
                    
                End If
                
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salida_Click()
    Operador.SetFocus
    PrgConfigVende.Hide
    Unload Me
    Menu.Show
End Sub

