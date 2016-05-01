VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgConfigVentas 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreos de Atributos de Ventas"
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
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "configventas.frx":0000
      Tab(0).ControlCount=   30
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Titulo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Titulo2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Titulo3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Titulo4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Titulo8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Titulo7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Titulo6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Titulo5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Titulo012"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Titulo011"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Titulo010"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Titulo9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Titulo013"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Titulo014"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Titulo015"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Opcion1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Opcion2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Opcion3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Opcion4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Opcion8"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Opcion7"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Opcion6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Opcion5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Opcion012"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Opcion011"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Opcion010"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Opcion9"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Opcion013"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Opcion014"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Opcion015"
      Tab(0).Control(29).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "configventas.frx":001C
      Tab(1).ControlCount=   44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Titulo11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Titulo12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Titulo13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Titulo14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Titulo18"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Titulo17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Titulo16"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Titulo15"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Titulo112"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Titulo111"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Titulo110"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Titulo19"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Titulo115"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Titulo114"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Titulo113"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Titulo116"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Titulo117"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Titulo118"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Titulo119"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Titulo120"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Titulo121"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Titulo122"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Opcion11"
      Tab(1).Control(22).Enabled=   -1  'True
      Tab(1).Control(23)=   "Opcion12"
      Tab(1).Control(23).Enabled=   -1  'True
      Tab(1).Control(24)=   "Opcion13"
      Tab(1).Control(24).Enabled=   -1  'True
      Tab(1).Control(25)=   "Opcion14"
      Tab(1).Control(25).Enabled=   -1  'True
      Tab(1).Control(26)=   "Opcion18"
      Tab(1).Control(26).Enabled=   -1  'True
      Tab(1).Control(27)=   "Opcion17"
      Tab(1).Control(27).Enabled=   -1  'True
      Tab(1).Control(28)=   "Opcion16"
      Tab(1).Control(28).Enabled=   -1  'True
      Tab(1).Control(29)=   "Opcion15"
      Tab(1).Control(29).Enabled=   -1  'True
      Tab(1).Control(30)=   "Opcion112"
      Tab(1).Control(30).Enabled=   -1  'True
      Tab(1).Control(31)=   "Opcion111"
      Tab(1).Control(31).Enabled=   -1  'True
      Tab(1).Control(32)=   "Opcion110"
      Tab(1).Control(32).Enabled=   -1  'True
      Tab(1).Control(33)=   "Opcion19"
      Tab(1).Control(33).Enabled=   -1  'True
      Tab(1).Control(34)=   "Opcion115"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Opcion114"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Opcion113"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Opcion116"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Opcion117"
      Tab(1).Control(38).Enabled=   -1  'True
      Tab(1).Control(39)=   "Opcion118"
      Tab(1).Control(39).Enabled=   -1  'True
      Tab(1).Control(40)=   "Opcion119"
      Tab(1).Control(40).Enabled=   -1  'True
      Tab(1).Control(41)=   "Opcion120"
      Tab(1).Control(41).Enabled=   -1  'True
      Tab(1).Control(42)=   "Opcion121"
      Tab(1).Control(42).Enabled=   -1  'True
      Tab(1).Control(43)=   "Opcion122"
      Tab(1).Control(43).Enabled=   -1  'True
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "configventas.frx":0038
      Tab(2).ControlCount=   48
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Opcion223"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "Opcion224"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "Opcion219"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "Opcion220"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "Opcion221"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "Opcion222"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "Opcion216"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "Opcion215"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "Opcion214"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Opcion213"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "Opcion218"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "Opcion217"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "Opcion29"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "Opcion210"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "Opcion211"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "Opcion212"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "Opcion25"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "Opcion26"
      Tab(2).Control(17).Enabled=   -1  'True
      Tab(2).Control(18)=   "Opcion27"
      Tab(2).Control(18).Enabled=   -1  'True
      Tab(2).Control(19)=   "Opcion28"
      Tab(2).Control(19).Enabled=   -1  'True
      Tab(2).Control(20)=   "opcion21"
      Tab(2).Control(20).Enabled=   -1  'True
      Tab(2).Control(21)=   "Opcion22"
      Tab(2).Control(21).Enabled=   -1  'True
      Tab(2).Control(22)=   "Opcion23"
      Tab(2).Control(22).Enabled=   -1  'True
      Tab(2).Control(23)=   "Opcion24"
      Tab(2).Control(23).Enabled=   -1  'True
      Tab(2).Control(24)=   "Titulo223"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Titulo224"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Titulo219"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Titulo220"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Titulo221"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Titulo222"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Titulo216"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Titulo215"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Titulo214"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Titulo213"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Titulo218"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Titulo217"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Titulo29"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Titulo210"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Titulo211"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Titulo212"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Titulo25"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Titulo26"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Titulo27"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Titulo28"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Titulo21"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Titulo22"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Titulo23"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Titulo24"
      Tab(2).Control(47).Enabled=   0   'False
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "configventas.frx":0054
      Tab(3).ControlCount=   10
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Titulo31"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Titulo32"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Titulo33"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Titulo34"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Titulo35"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Opcion31"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Opcion32"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Opcion33"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Opcion34"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Opcion35"
      Tab(3).Control(9).Enabled=   0   'False
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
         Left            =   -64800
         TabIndex        =   136
         Top             =   3960
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
         Left            =   -64800
         TabIndex        =   134
         Top             =   3600
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
         Left            =   -64800
         TabIndex        =   132
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Opcion35 
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
         Left            =   -69840
         TabIndex        =   130
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Opcion223 
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
         Left            =   -64800
         TabIndex        =   123
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Opcion224 
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
         Left            =   -64800
         TabIndex        =   122
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox Opcion219 
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
         Left            =   -64800
         TabIndex        =   121
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Opcion220 
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
         Left            =   -64800
         TabIndex        =   120
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox Opcion221 
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
         Left            =   -64800
         TabIndex        =   119
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Opcion222 
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
         Left            =   -64800
         TabIndex        =   118
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox Opcion015 
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
         TabIndex        =   114
         Top             =   5760
         Width           =   615
      End
      Begin VB.CheckBox Opcion014 
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
         TabIndex        =   113
         Top             =   5400
         Width           =   615
      End
      Begin VB.CheckBox Opcion013 
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
         TabIndex        =   112
         Top             =   5040
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
         Left            =   -64800
         TabIndex        =   110
         Top             =   2880
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
         TabIndex        =   105
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Opcion010 
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
         TabIndex        =   104
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox Opcion011 
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
         TabIndex        =   103
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Opcion012 
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
         TabIndex        =   102
         Top             =   4680
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
         TabIndex        =   97
         Top             =   2160
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
         TabIndex        =   96
         Top             =   2520
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
         TabIndex        =   95
         Top             =   2880
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
         TabIndex        =   94
         Top             =   3240
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
         Left            =   -64800
         TabIndex        =   92
         Top             =   2460
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
         Left            =   -64800
         TabIndex        =   90
         Top             =   2100
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
         Left            =   -64800
         TabIndex        =   88
         Top             =   1740
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
         Left            =   -64800
         TabIndex        =   84
         Top             =   720
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
         Left            =   -64800
         TabIndex        =   83
         Top             =   1080
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
         Left            =   -64800
         TabIndex        =   82
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion34 
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
         Left            =   -69840
         TabIndex        =   77
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion33 
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
         Left            =   -69840
         TabIndex        =   76
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion32 
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
         Left            =   -69840
         TabIndex        =   75
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Opcion31 
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
         Left            =   -69840
         TabIndex        =   74
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Opcion216 
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
         Left            =   -64800
         TabIndex        =   67
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion215 
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
         Left            =   -64800
         TabIndex        =   66
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion214 
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
         Left            =   -64800
         TabIndex        =   65
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Opcion213 
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
         Left            =   -64800
         TabIndex        =   64
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Opcion218 
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
         Left            =   -64800
         TabIndex        =   63
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Opcion217 
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
         Left            =   -64800
         TabIndex        =   62
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Opcion29 
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
         TabIndex        =   57
         Top             =   3600
         Width           =   615
      End
      Begin VB.CheckBox Opcion210 
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
         TabIndex        =   56
         Top             =   3960
         Width           =   615
      End
      Begin VB.CheckBox Opcion211 
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
         TabIndex        =   55
         Top             =   4320
         Width           =   615
      End
      Begin VB.CheckBox Opcion212 
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
         Top             =   4680
         Width           =   615
      End
      Begin VB.CheckBox Opcion25 
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
         TabIndex        =   49
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox Opcion26 
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
         TabIndex        =   48
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Opcion27 
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
         TabIndex        =   47
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Opcion28 
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
         TabIndex        =   46
         Top             =   3240
         Width           =   615
      End
      Begin VB.CheckBox opcion21 
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
         TabIndex        =   41
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Opcion22 
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
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Opcion23 
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
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Opcion24 
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
         TabIndex        =   38
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion19 
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
         TabIndex        =   33
         Top             =   3600
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
         TabIndex        =   32
         Top             =   3960
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
         TabIndex        =   31
         Top             =   4320
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
         TabIndex        =   30
         Top             =   4680
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
         Left            =   -70320
         TabIndex        =   25
         Top             =   2160
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
         Left            =   -70320
         TabIndex        =   24
         Top             =   2520
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
         Left            =   -70320
         TabIndex        =   23
         Top             =   2880
         Width           =   615
      End
      Begin VB.CheckBox Opcion18 
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
         Top             =   3240
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
         TabIndex        =   21
         Top             =   1800
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
         TabIndex        =   19
         Top             =   1440
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
         TabIndex        =   17
         Top             =   1080
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
         TabIndex        =   15
         Top             =   720
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
         Left            =   -70320
         TabIndex        =   13
         Top             =   1800
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
         Left            =   -70320
         TabIndex        =   11
         Top             =   1440
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
         Left            =   -70320
         TabIndex        =   9
         Top             =   1080
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
         Left            =   -70320
         TabIndex        =   7
         Top             =   720
         Width           =   615
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
         Left            =   -69720
         TabIndex        =   137
         Top             =   3960
         Width           =   4695
      End
      Begin VB.Label Titulo121 
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
         Left            =   -69720
         TabIndex        =   135
         Top             =   3600
         Width           =   4695
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
         Left            =   -69720
         TabIndex        =   133
         Top             =   3240
         Width           =   4695
      End
      Begin VB.Label Titulo35 
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
         Height          =   495
         Left            =   -74520
         TabIndex        =   131
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label Titulo223 
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
         Left            =   -69240
         TabIndex        =   129
         Top             =   4320
         Width           =   4335
      End
      Begin VB.Label Titulo224 
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
         Left            =   -69240
         TabIndex        =   128
         Top             =   4680
         Width           =   4335
      End
      Begin VB.Label Titulo219 
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
         Left            =   -69240
         TabIndex        =   127
         Top             =   2880
         Width           =   4335
      End
      Begin VB.Label Titulo220 
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
         Left            =   -69240
         TabIndex        =   126
         Top             =   3240
         Width           =   4335
      End
      Begin VB.Label Titulo221 
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
         Left            =   -69240
         TabIndex        =   125
         Top             =   3600
         Width           =   4335
      End
      Begin VB.Label Titulo222 
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
         Left            =   -69240
         TabIndex        =   124
         Top             =   3960
         Width           =   4335
      End
      Begin VB.Label Titulo015 
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
         Height          =   615
         Left            =   480
         TabIndex        =   117
         Top             =   5760
         Width           =   3975
      End
      Begin VB.Label Titulo014 
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
         TabIndex        =   116
         Top             =   5400
         Width           =   3855
      End
      Begin VB.Label Titulo013 
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
         TabIndex        =   115
         Top             =   5040
         Width           =   3975
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
         Left            =   -69720
         TabIndex        =   111
         Top             =   2880
         Width           =   4935
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
         TabIndex        =   109
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label Titulo010 
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
         TabIndex        =   108
         Top             =   3960
         Width           =   3975
      End
      Begin VB.Label Titulo011 
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
         TabIndex        =   107
         Top             =   4320
         Width           =   3855
      End
      Begin VB.Label Titulo012 
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
         Height          =   615
         Left            =   480
         TabIndex        =   106
         Top             =   4680
         Width           =   3975
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
         TabIndex        =   101
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Titulo6 
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
         TabIndex        =   100
         Top             =   2520
         Width           =   3975
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
         TabIndex        =   99
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Titulo8 
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
         TabIndex        =   98
         Top             =   3240
         Width           =   3975
      End
      Begin VB.Label Titulo118 
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
         Left            =   -69720
         TabIndex        =   93
         Top             =   2460
         Width           =   4935
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
         Left            =   -69720
         TabIndex        =   91
         Top             =   2100
         Width           =   4935
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
         Left            =   -69720
         TabIndex        =   89
         Top             =   1740
         Width           =   4935
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
         Height          =   375
         Left            =   -69720
         TabIndex        =   87
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Titulo114 
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
         Height          =   375
         Left            =   -69720
         TabIndex        =   86
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Titulo115 
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
         Left            =   -69720
         TabIndex        =   85
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label Titulo34 
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
         Height          =   495
         Left            =   -74520
         TabIndex        =   81
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label Titulo33 
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
         TabIndex        =   80
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label Titulo32 
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
         TabIndex        =   79
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Titulo31 
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
         TabIndex        =   78
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Titulo216 
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
         Left            =   -69240
         TabIndex        =   73
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Titulo215 
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
         Left            =   -69240
         TabIndex        =   72
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Titulo214 
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
         Left            =   -69240
         TabIndex        =   71
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Titulo213 
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
         Left            =   -69240
         TabIndex        =   70
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Titulo218 
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
         Left            =   -69240
         TabIndex        =   69
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Titulo217 
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
         Left            =   -69240
         TabIndex        =   68
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Titulo29 
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
         TabIndex        =   61
         Top             =   3600
         Width           =   3975
      End
      Begin VB.Label Titulo210 
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
         TabIndex        =   60
         Top             =   3960
         Width           =   4095
      End
      Begin VB.Label Titulo211 
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
         TabIndex        =   59
         Top             =   4320
         Width           =   4095
      End
      Begin VB.Label Titulo212 
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
         TabIndex        =   58
         Top             =   4680
         Width           =   3975
      End
      Begin VB.Label Titulo25 
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
         TabIndex        =   53
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label Titulo26 
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
         TabIndex        =   52
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label Titulo27 
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
         TabIndex        =   51
         Top             =   2880
         Width           =   4095
      End
      Begin VB.Label Titulo28 
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
         Height          =   375
         Left            =   -74520
         TabIndex        =   50
         Top             =   3240
         Width           =   4095
      End
      Begin VB.Label Titulo21 
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
         TabIndex        =   45
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Titulo22 
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
         TabIndex        =   44
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Titulo23 
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
         TabIndex        =   43
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Titulo24 
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
         TabIndex        =   42
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Label Titulo19 
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
         TabIndex        =   37
         Top             =   3600
         Width           =   4095
      End
      Begin VB.Label Titulo110 
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
         TabIndex        =   36
         Top             =   3960
         Width           =   4095
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
         TabIndex        =   35
         Top             =   4320
         Width           =   3855
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
         Height          =   615
         Left            =   -74520
         TabIndex        =   34
         Top             =   4680
         Width           =   3975
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
         Left            =   -74520
         TabIndex        =   29
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Titulo16 
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
         TabIndex        =   28
         Top             =   2520
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
         Left            =   -74520
         TabIndex        =   27
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label Titulo18 
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
         TabIndex        =   26
         Top             =   3240
         Width           =   3975
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
         TabIndex        =   20
         Top             =   1800
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
         TabIndex        =   18
         Top             =   1440
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
         TabIndex        =   16
         Top             =   1080
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
         TabIndex        =   14
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Titulo14 
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
         TabIndex        =   12
         Top             =   1800
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
         Left            =   -74520
         TabIndex        =   10
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Titulo12 
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
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Titulo11 
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
         Top             =   720
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
Attribute VB_Name = "PrgConfigVentas"
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
    
    Tablas.TabCaption(0) = "Maestros"
    Tablas.TabCaption(1) = "Novedades"
    Tablas.TabCaption(2) = "Listados"
    Tablas.TabCaption(3) = "Procesos"
    
    Rem titulo1
    
    Titulo1.Caption = "Ingreso de Rubros"
    Titulo2.Caption = "Ingreso de Vendedores"
    Titulo3.Caption = "Ingreso de Condiciones de Pago"
    Titulo4.Caption = "Ingreso de Cambios"
    Titulo5.Caption = "Ingreso de Lineas de Venta"
    Titulo6.Caption = "Ingreso de Familias de Dy"
    Titulo7.Caption = "Ingreso de Envases"
    Titulo8.Caption = "Ingreso de Clientes"
    Titulo9.Caption = "Ingreso de Precios por Cliente"
    Titulo010.Caption = "Ingreso de Composicion de Productos"
    Titulo011.Caption = "Modificacion de Precios"
    Titulo012.Caption = "Ingreso de Conceptos de Gastos de Importacion"
    Titulo013.Caption = "Consulta de Versiones de Composion de Productos Terminados"
    Titulo014.Caption = "Consulta de Revisiones de Ensayos"
    Titulo015.Caption = "Ingreso de Listas de Precios"
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    Opcion8.Value = 0
    Opcion9.Value = 0
    Opcion010.Value = 0
    Opcion011.Value = 0
    Opcion012.Value = 0
    Opcion013.Value = 0
    Opcion014.Value = 0
    Opcion015.Value = 0
    
    
    Rem titulo2
    
    Titulo11.Caption = "Ingreso de Pedidos"
    Titulo12.Caption = "Emision de Factura/Remito U$S"
    Titulo13.Caption = "Emision de Factura/Remito $"
    Titulo14.Caption = "Emision de Factura de Consignacion"
    Titulo15.Caption = "Ingreso de Devolucion"
    Titulo16.Caption = "Ingreso de Comprobantes Varios"
    Titulo17.Caption = "Ingreso de Remitos en Consignacion"
    Titulo18.Caption = "Devolucion de Mercaderia en Consignacion"
    Titulo19.Caption = "Ingreso de Factura de Exportacion Provisoria"
    Titulo110.Caption = "Ingreso de Factura de Exportacion"
    Titulo111.Caption = "Ingreso de Factura de Exportacion por Conceptos Varios"
    Titulo112.Caption = "Emision de Notas de Debito/Credito por Diferencia de Cambio"
    Titulo113.Caption = "Emision de Notas de Debito/Credito por Diferencia de Cambio (Acreditacion)"
    Titulo114.Caption = "Autorizacion de Pedidos"
    Titulo115.Caption = "Actualizacion de Pedidos"
    Titulo116.Caption = "Ingreso de Gastos de Importacion"
    Titulo117.Caption = "Ingreso de Solicitud de Devolucion de Mercaderia"
    Titulo118.Caption = "Ingreso de Gastos de Importacion Parciales"
    Titulo119.Caption = "Ingreso de Ordenes de Compra de Importacion"
    Titulo120.Caption = "Ingreso de Pedidos de Desarrollo"
    Titulo121.Caption = "Consulta de Pedidos de Desarrollo"
    Titulo122.Caption = "Autorizacion de Desarrollo"
    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0
    Opcion17.Value = 0
    Opcion18.Value = 0
    Opcion19.Value = 0
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
    
    
    Rem titulo3
    
    Titulo21.Caption = "Consulta de Cuenta Corriente de Clientes por Pantalla"
    Titulo22.Caption = "Cuenta Corriente de Clientes"
    Titulo23.Caption = "Saldos de Cuenta Corriente de Clientes"
    Titulo24.Caption = "Subdiario de Iva Ventas"
    Titulo25.Caption = "Listado de Pedidos Pendientes"
    Titulo26.Caption = "Listado de Cash Flow"
    Titulo27.Caption = "Listado de Ventas por Provincia"
    Titulo28.Caption = "Listado de Mercaderia en Consignacion por Cliente"
    Titulo29.Caption = "Listado de Mercaderia en Consignacion pendiente"
    Titulo210.Caption = "Listado de Ventas fuera de Fecha"
    Titulo211.Caption = "Listado de Cuenta Corriente a Fecha"
    Titulo212.Caption = "Listado de Gastos de Importacion por Carpeta"
    Titulo213.Caption = "Listado del Calculo de Costo de Importacion por Carpeta"
    Titulo214.Caption = "Listado de Ventas Diarias"
    Titulo215.Caption = "Listado de Notas de Debito por Diferencia de Cambio Pendientes"
    Titulo216.Caption = "Listado de Precepciones de Ingresos Brutos"
    Titulo217.Caption = "Analisis de Cumplimiento de Pedidos de Venta"
    Titulo218.Caption = "Listado de Clientes por Vendedor"
    Titulo219.Caption = "Listado de Ordenes Compras Pendientes de Dy (Articulo)"
    Titulo220.Caption = "Listado del Calculo de Costo de Nacionalizacion de Mercaderia"
    Titulo221.Caption = "Listado de Precios (Grupo)"
    Titulo222.Caption = "Listado de Precios Comparativo (Grupo)"
    Titulo223.Caption = "Listado de Precios Comparativo (Cliente)"
    Titulo224.Caption = "Listado General de Productos"
    
    opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion210.Value = 0
    Opcion211.Value = 0
    Opcion212.Value = 0
    Opcion213.Value = 0
    Opcion214.Value = 0
    Opcion215.Value = 0
    Opcion216.Value = 0
    Opcion217.Value = 0
    Opcion218.Value = 0
    Opcion219.Value = 0
    Opcion220.Value = 0
    Opcion221.Value = 0
    Opcion222.Value = 0
    Opcion223.Value = 0
    Opcion224.Value = 0
    
    
    
    Rem titulo4
    
    Titulo31.Caption = "Cierre de Mes"
    Titulo32.Caption = "Verificacion de Formulas"
    Titulo33.Caption = "Grabacion de Datos Electronicamente"
    Titulo34.Caption = "Actualizacion de Costos de Importacion"
    Titulo35.Caption = "Fin del Sistema"
    
    Opcion31.Value = 0
    Opcion32.Value = 0
    Opcion33.Value = 0
    Opcion34.Value = 0
    Opcion35.Value = 0
   
    

End Sub


Private Sub Graba_Click()

    XParam = "'" + Operador.Text + "','" _
                 + "2" + "'"
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
    If Opcion010.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion011.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion012.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion013.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion014.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion015.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    
    
    
    If Opcion11.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion12.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion13.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion14.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion15.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion16.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion17.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion18.Value = 0 Then
        WAtributo2 = WAtributo2 + "0"
            Else
        WAtributo2 = WAtributo2 + "1"
    End If
    If Opcion19.Value = 0 Then
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
    
    
    
    
    If opcion21.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion22.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion23.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion24.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion25.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion26.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion27.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion28.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion29.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion210.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion211.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion212.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion213.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion214.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion215.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion216.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion217.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion218.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion219.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion220.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion221.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion222.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion223.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    If Opcion224.Value = 0 Then
        WAtributo3 = WAtributo3 + "0"
            Else
        WAtributo3 = WAtributo3 + "1"
    End If
    
    
    
    
    If Opcion31.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion32.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion33.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion34.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    If Opcion35.Value = 0 Then
        WAtributo4 = WAtributo4 + "0"
            Else
        WAtributo4 = WAtributo4 + "1"
    End If
    
    WProceso = "2"
                                       
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
    Opcion010.Value = 0
    Opcion011.Value = 0
    Opcion012.Value = 0
    Opcion013.Value = 0
    Opcion014.Value = 0
    Opcion015.Value = 0
                    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    Opcion14.Value = 0
    Opcion15.Value = 0
    Opcion16.Value = 0
    Opcion17.Value = 0
    Opcion18.Value = 0
    Opcion19.Value = 0
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

    opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion210.Value = 0
    Opcion211.Value = 0
    Opcion212.Value = 0
    Opcion213.Value = 0
    Opcion214.Value = 0
    Opcion215.Value = 0
    Opcion216.Value = 0
    Opcion217.Value = 0
    Opcion218.Value = 0
    Opcion219.Value = 0
    Opcion220.Value = 0
    Opcion221.Value = 0
    Opcion222.Value = 0
    Opcion223.Value = 0
    Opcion224.Value = 0
                    
    Opcion31.Value = 0
    Opcion32.Value = 0
    Opcion33.Value = 0
    Opcion34.Value = 0
    Opcion35.Value = 0
    
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
                Opcion010.Value = 0
                Opcion011.Value = 0
                Opcion012.Value = 0
                Opcion013.Value = 0
                Opcion014.Value = 0
                Opcion015.Value = 0
                    
                Opcion11.Value = 0
                Opcion12.Value = 0
                Opcion13.Value = 0
                Opcion14.Value = 0
                Opcion15.Value = 0
                Opcion16.Value = 0
                Opcion17.Value = 0
                Opcion18.Value = 0
                Opcion19.Value = 0
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
        
                opcion21.Value = 0
                Opcion22.Value = 0
                Opcion23.Value = 0
                Opcion24.Value = 0
                Opcion25.Value = 0
                Opcion26.Value = 0
                Opcion27.Value = 0
                Opcion28.Value = 0
                Opcion29.Value = 0
                Opcion210.Value = 0
                Opcion211.Value = 0
                Opcion212.Value = 0
                Opcion213.Value = 0
                Opcion214.Value = 0
                Opcion215.Value = 0
                Opcion216.Value = 0
                Opcion217.Value = 0
                Opcion218.Value = 0
                Opcion219.Value = 0
                Opcion220.Value = 0
                Opcion221.Value = 0
                Opcion222.Value = 0
                Opcion223.Value = 0
                Opcion224.Value = 0
                    
                Opcion31.Value = 0
                Opcion32.Value = 0
                Opcion33.Value = 0
                Opcion34.Value = 0
                Opcion35.Value = 0
                
                
                XParam = "'" + Operador.Text + "','" _
                             + "2" + "'"
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
                    Opcion010.Value = Val(Mid$(rstAtributos!atributo1, 10, 1))
                    Opcion011.Value = Val(Mid$(rstAtributos!atributo1, 11, 1))
                    Opcion012.Value = Val(Mid$(rstAtributos!atributo1, 12, 1))
                    Opcion013.Value = Val(Mid$(rstAtributos!atributo1, 13, 1))
                    Opcion014.Value = Val(Mid$(rstAtributos!atributo1, 14, 1))
                    Opcion015.Value = Val(Mid$(rstAtributos!atributo1, 15, 1))
                    
                    Opcion11.Value = Val(Mid$(rstAtributos!atributo2, 1, 1))
                    Opcion12.Value = Val(Mid$(rstAtributos!atributo2, 2, 1))
                    Opcion13.Value = Val(Mid$(rstAtributos!atributo2, 3, 1))
                    Opcion14.Value = Val(Mid$(rstAtributos!atributo2, 4, 1))
                    Opcion15.Value = Val(Mid$(rstAtributos!atributo2, 5, 1))
                    Opcion16.Value = Val(Mid$(rstAtributos!atributo2, 6, 1))
                    Opcion17.Value = Val(Mid$(rstAtributos!atributo2, 7, 1))
                    Opcion18.Value = Val(Mid$(rstAtributos!atributo2, 8, 1))
                    Opcion19.Value = Val(Mid$(rstAtributos!atributo2, 9, 1))
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
                    
                    opcion21.Value = Val(Mid$(rstAtributos!atributo3, 1, 1))
                    Opcion22.Value = Val(Mid$(rstAtributos!atributo3, 2, 1))
                    Opcion23.Value = Val(Mid$(rstAtributos!atributo3, 3, 1))
                    Opcion24.Value = Val(Mid$(rstAtributos!atributo3, 4, 1))
                    Opcion25.Value = Val(Mid$(rstAtributos!atributo3, 5, 1))
                    Opcion26.Value = Val(Mid$(rstAtributos!atributo3, 6, 1))
                    Opcion27.Value = Val(Mid$(rstAtributos!atributo3, 7, 1))
                    Opcion28.Value = Val(Mid$(rstAtributos!atributo3, 8, 1))
                    Opcion29.Value = Val(Mid$(rstAtributos!atributo3, 9, 1))
                    Opcion210.Value = Val(Mid$(rstAtributos!atributo3, 10, 1))
                    Opcion211.Value = Val(Mid$(rstAtributos!atributo3, 11, 1))
                    Opcion212.Value = Val(Mid$(rstAtributos!atributo3, 12, 1))
                    Opcion213.Value = Val(Mid$(rstAtributos!atributo3, 13, 1))
                    Opcion214.Value = Val(Mid$(rstAtributos!atributo3, 14, 1))
                    Opcion215.Value = Val(Mid$(rstAtributos!atributo3, 15, 1))
                    Opcion216.Value = Val(Mid$(rstAtributos!atributo3, 16, 1))
                    Opcion217.Value = Val(Mid$(rstAtributos!atributo3, 17, 1))
                    Opcion218.Value = Val(Mid$(rstAtributos!atributo3, 18, 1))
                    Opcion219.Value = Val(Mid$(rstAtributos!atributo3, 19, 1))
                    Opcion220.Value = Val(Mid$(rstAtributos!atributo3, 20, 1))
                    Opcion221.Value = Val(Mid$(rstAtributos!atributo3, 21, 1))
                    Opcion222.Value = Val(Mid$(rstAtributos!atributo3, 22, 1))
                    Opcion223.Value = Val(Mid$(rstAtributos!atributo3, 23, 1))
                    Opcion224.Value = Val(Mid$(rstAtributos!atributo3, 24, 1))
                    
                    Opcion31.Value = Val(Mid$(rstAtributos!atributo4, 1, 1))
                    Opcion32.Value = Val(Mid$(rstAtributos!atributo4, 2, 1))
                    Opcion33.Value = Val(Mid$(rstAtributos!atributo4, 3, 1))
                    Opcion34.Value = Val(Mid$(rstAtributos!atributo4, 4, 1))
                    Opcion35.Value = Val(Mid$(rstAtributos!atributo4, 5, 1))
                    
                    rstAtributos.Close
                End If
                
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salida_Click()
    Operador.SetFocus
    PrgConfigVentas.Hide
    Unload Me
    Menu.Show
End Sub




