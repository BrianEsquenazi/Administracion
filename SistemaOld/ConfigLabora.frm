VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgConfigLabora 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreos de Atributos de Desarrollo"
   ClientHeight    =   8400
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   8400
   ScaleWidth      =   11790
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
      Left            =   6360
      TabIndex        =   3
      Top             =   7080
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
      Top             =   7080
      Width           =   1215
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   10821
      _Version        =   327680
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "ConfigLabora.frx":0000
      Tab(0).ControlCount=   60
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Titulo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Titulo2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Titulo4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Titulo3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Titulo5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Titulo10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Titulo8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Titulo9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Titulo7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Titulo6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Titulo15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Titulo13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Titulo14"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Titulo12"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Titulo11"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Titulo20"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Titulo18"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Titulo19"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Titulo17"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Titulo16"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Titulo25"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Titulo23"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Titulo24"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Titulo22"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Titulo21"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Titulo30"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Titulo28"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Titulo29"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Titulo27"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Titulo26"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Opcion1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Opcion2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Opcion4"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Opcion3"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Opcion5"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Opcion10"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Opcion8"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Opcion9"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Opcion7"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Opcion6"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Opcion15"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Opcion13"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Opcion14"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Opcion12"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Opcion11"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Opcion20"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Opcion18"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Opcion19"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Opcion17"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Opcion16"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Opcion25"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Opcion23"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Opcion24"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Opcion22"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Opcion21"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Opcion30"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Opcion28"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Opcion29"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Opcion27"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Opcion26"
      Tab(0).Control(59).Enabled=   0   'False
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
         Left            =   10920
         TabIndex        =   60
         Top             =   3780
         Width           =   255
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
         Left            =   10920
         TabIndex        =   59
         Top             =   4140
         Width           =   375
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
         Left            =   10920
         TabIndex        =   58
         Top             =   4860
         Width           =   375
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
         Left            =   10920
         TabIndex        =   57
         Top             =   4500
         Width           =   375
      End
      Begin VB.CheckBox Opcion30 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   56
         Top             =   5160
         Width           =   375
      End
      Begin VB.CheckBox Opcion21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   50
         Top             =   2100
         Width           =   375
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
         Left            =   10920
         TabIndex        =   49
         Top             =   2460
         Width           =   375
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
         Left            =   10920
         TabIndex        =   48
         Top             =   3180
         Width           =   375
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
         Left            =   10920
         TabIndex        =   47
         Top             =   2820
         Width           =   375
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
         Left            =   10920
         TabIndex        =   46
         Top             =   3480
         Width           =   375
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
         Left            =   10920
         TabIndex        =   40
         Top             =   420
         Width           =   375
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
         Left            =   10920
         TabIndex        =   39
         Top             =   780
         Width           =   255
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
         Left            =   10920
         TabIndex        =   38
         Top             =   1500
         Width           =   375
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
         Left            =   10920
         TabIndex        =   37
         Top             =   1140
         Width           =   375
      End
      Begin VB.CheckBox Opcion20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   36
         Top             =   1800
         Width           =   375
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
         TabIndex        =   30
         Top             =   3780
         Width           =   375
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
         TabIndex        =   29
         Top             =   4140
         Width           =   255
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
         TabIndex        =   28
         Top             =   4860
         Width           =   375
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
         TabIndex        =   27
         Top             =   4500
         Width           =   255
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
         TabIndex        =   26
         Top             =   5160
         Width           =   375
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
         TabIndex        =   20
         Top             =   2100
         Width           =   375
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
         TabIndex        =   19
         Top             =   2460
         Width           =   375
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
         TabIndex        =   18
         Top             =   3180
         Width           =   375
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
         TabIndex        =   17
         Top             =   2820
         Width           =   375
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
         TabIndex        =   16
         Top             =   3480
         Width           =   375
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
         TabIndex        =   14
         Top             =   1800
         Width           =   375
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
         TabIndex        =   11
         Top             =   1140
         Width           =   375
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
         TabIndex        =   10
         Top             =   1500
         Width           =   375
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
         TabIndex        =   9
         Top             =   780
         Width           =   375
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
         TabIndex        =   7
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Titulo26 
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
         Left            =   5160
         TabIndex        =   65
         Top             =   3795
         Width           =   5295
      End
      Begin VB.Label Titulo27 
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
         Left            =   5160
         TabIndex        =   64
         Top             =   4140
         Width           =   5295
      End
      Begin VB.Label Titulo29 
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
         Left            =   5160
         TabIndex        =   63
         Top             =   4800
         Width           =   5295
      End
      Begin VB.Label Titulo28 
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
         Left            =   5160
         TabIndex        =   62
         Top             =   4440
         Width           =   5295
      End
      Begin VB.Label Titulo30 
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
         Left            =   5160
         TabIndex        =   61
         Top             =   5160
         Width           =   5415
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
         Height          =   255
         Left            =   5160
         TabIndex        =   55
         Top             =   2100
         Width           =   5415
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
         Left            =   5160
         TabIndex        =   54
         Top             =   2460
         Width           =   5295
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
         Left            =   5160
         TabIndex        =   53
         Top             =   3105
         Width           =   5655
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
         Left            =   5160
         TabIndex        =   52
         Top             =   2760
         Width           =   5295
      End
      Begin VB.Label Titulo25 
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
         Left            =   5160
         TabIndex        =   51
         Top             =   3480
         Width           =   5535
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
         Left            =   5160
         TabIndex        =   45
         Top             =   405
         Width           =   5295
      End
      Begin VB.Label Titulo17 
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
         Left            =   5160
         TabIndex        =   44
         Top             =   780
         Width           =   5175
      End
      Begin VB.Label Titulo19 
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
         Left            =   5160
         TabIndex        =   43
         Top             =   1455
         Width           =   5535
      End
      Begin VB.Label Titulo18 
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
         Left            =   5160
         TabIndex        =   42
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label Titulo20 
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
         Left            =   5160
         TabIndex        =   41
         Top             =   1800
         Width           =   5295
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
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   3840
         Width           =   4335
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
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   4200
         Width           =   4335
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
         Left            =   240
         TabIndex        =   33
         Top             =   4860
         Width           =   4335
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
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label Titulo15 
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
         Left            =   240
         TabIndex        =   31
         Top             =   5160
         Width           =   4335
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
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2100
         Width           =   4335
      End
      Begin VB.Label Titulo7 
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
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   2460
         Width           =   4335
      End
      Begin VB.Label Titulo9 
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
         Left            =   240
         TabIndex        =   23
         Top             =   3105
         Width           =   4335
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
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   4335
      End
      Begin VB.Label Titulo10 
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
         Left            =   240
         TabIndex        =   21
         Top             =   3480
         Width           =   4335
      End
      Begin VB.Label Titulo5 
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   4335
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
         Height          =   360
         Left            =   240
         TabIndex        =   13
         Top             =   1095
         Width           =   4335
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
         Height          =   360
         Left            =   240
         TabIndex        =   12
         Top             =   1455
         Width           =   4335
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   780
         Width           =   4335
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
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4455
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
Attribute VB_Name = "PrgConfigLabora"
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
    
    Tablas.TabCaption(0) = "SISTEMA"
    
    Rem titulo1
    
    Titulo1.Caption = "Ingreso de Ensayos"
    Titulo2.Caption = "Consulta de Especificaciones (Historico)"
    Titulo3.Caption = "Consulta de Especificaciones por Version"
    Titulo4.Caption = "Especificaciones (Unificado)"
    Titulo5.Caption = "Controles de M.P."
    Titulo6.Caption = "Listado de Ensayos en Materia Prima"
    Titulo7.Caption = "Homologacion de Muestras de Materias Primas"
    Titulo8.Caption = "Verificacion de Vencimientos de Materia Prima"
    Titulo9.Caption = "Listado de Especificaciones de Materia Prima por Fecha"
    Titulo10.Caption = "Etiquetas de Muestra Simple"
    Titulo11.Caption = "Consulta de Especificaciones (Historico)"
    Titulo12.Caption = "Consulta de Especificaciones por Version"
    Titulo13.Caption = "Especificaciones (Unificado)"
    Titulo14.Caption = "Controles de P.T."
    Titulo15.Caption = "Devolucion de NK o RE"
    Titulo16.Caption = "Listado de Ensayos de Producto Terminado"
    Titulo17.Caption = "Carga de Ensayos a Imprimir en los Certificados de Analisis"
    Titulo18.Caption = "Emision de Certificado de Analisis"
    Titulo19.Caption = "Listado de Especificaciones de Producto Terminado por Fecha"
    Titulo20.Caption = "Listado de Productos Terminados Vencidos "
    Titulo21.Caption = "Especificaciones (Farma)"
    Titulo22.Caption = "Controles (Farma)"
    Titulo23.Caption = "Ingreso y Actualizacion de  Hojas de Produccion"
    Titulo24.Caption = "Ingreso Hoja de Produccion Planta III y V"
    Titulo25.Caption = "Movimientos Varios de Stock"
    Titulo26.Caption = "Liberacion de Productos Devueltos a Verificar"
    Titulo27.Caption = "Listado de Productos Pendientes de Liberar"
    Titulo28.Caption = "Verificacion de Pedidos de Desarrollo"
    Titulo29.Caption = "Ingreso de Informe de Recepcion de Drogas de Laboratorio"
    Titulo30.Caption = "Verificacion de Lotes Inactivos"
    
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
    Opcion18.Value = 0
    Opcion19.Value = 0
    Opcion20.Value = 0
    Opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion30.Value = 0

End Sub


Private Sub Graba_Click()

    XParam = "'" + Operador.Text + "','" _
                 + "6" + "'"
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
    If Opcion18.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion19.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion20.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion21.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion22.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion23.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion24.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion25.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion26.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion27.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion28.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion29.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion30.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    
    WProceso = "6"
                                       
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
    Opcion18.Value = 0
    Opcion19.Value = 0
    Opcion20.Value = 0
    Opcion21.Value = 0
    Opcion22.Value = 0
    Opcion23.Value = 0
    Opcion24.Value = 0
    Opcion25.Value = 0
    Opcion26.Value = 0
    Opcion27.Value = 0
    Opcion28.Value = 0
    Opcion29.Value = 0
    Opcion30.Value = 0
    
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
                Opcion18.Value = 0
                Opcion19.Value = 0
                Opcion20.Value = 0
                Opcion21.Value = 0
                Opcion22.Value = 0
                Opcion23.Value = 0
                Opcion24.Value = 0
                Opcion25.Value = 0
                Opcion26.Value = 0
                Opcion27.Value = 0
                Opcion28.Value = 0
                Opcion29.Value = 0
                Opcion30.Value = 0
                
                XParam = "'" + Operador.Text + "','" _
                             + "6" + "'"
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
                    Opcion18.Value = Val(Mid$(rstAtributos!atributo1, 18, 1))
                    Opcion19.Value = Val(Mid$(rstAtributos!atributo1, 19, 1))
                    Opcion20.Value = Val(Mid$(rstAtributos!atributo1, 20, 1))
                    Opcion21.Value = Val(Mid$(rstAtributos!atributo1, 21, 1))
                    Opcion22.Value = Val(Mid$(rstAtributos!atributo1, 22, 1))
                    Opcion23.Value = Val(Mid$(rstAtributos!atributo1, 23, 1))
                    Opcion24.Value = Val(Mid$(rstAtributos!atributo1, 24, 1))
                    Opcion25.Value = Val(Mid$(rstAtributos!atributo1, 25, 1))
                    Opcion26.Value = Val(Mid$(rstAtributos!atributo1, 26, 1))
                    Opcion27.Value = Val(Mid$(rstAtributos!atributo1, 27, 1))
                    Opcion28.Value = Val(Mid$(rstAtributos!atributo1, 28, 1))
                    Opcion29.Value = Val(Mid$(rstAtributos!atributo1, 29, 1))
                    Opcion30.Value = Val(Mid$(rstAtributos!atributo1, 30, 1))
                    rstAtributos.Close
                End If
                
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salida_Click()
    PrgConfigLabora.Hide
    Unload Me
    Menu.Show
End Sub

