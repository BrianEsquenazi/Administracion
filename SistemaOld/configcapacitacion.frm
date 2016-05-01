VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgConfigCapacitacion 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreos de Atributos de Capacitacion"
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
      TabPicture(0)   =   "configcapacitacion.frx":0000
      Tab(0).ControlCount=   14
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
      Tab(0).Control(7)=   "Opcion1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Opcion2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Opcion3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Opcion4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Opcion5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Opcion6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Opcion7"
      Tab(0).Control(13).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "configcapacitacion.frx":001C
      Tab(1).ControlCount=   6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Opcion13"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Opcion12"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "Opcion11"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "Titulo13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Titulo12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Titulo11"
      Tab(1).Control(5).Enabled=   0   'False
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "configcapacitacion.frx":0038
      Tab(2).ControlCount=   28
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Opcion214"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "Opcion213"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "Opcion210"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "Opcion29"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "Opcion28"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "Opcion27"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "Opcion212"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "Opcion211"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "Opcion25"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Opcion26"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "opcion21"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "Opcion22"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "Opcion23"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "Opcion24"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "Titulo214"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Titulo213"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Titulo210"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Titulo29"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Titulo28"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Titulo27"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Titulo212"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Titulo211"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Titulo25"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Titulo26"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Titulo21"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Titulo22"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Titulo23"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Titulo24"
      Tab(2).Control(27).Enabled=   0   'False
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "configcapacitacion.frx":0054
      Tab(3).ControlCount=   4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Opcion32"
      Tab(3).Control(0).Enabled=   -1  'True
      Tab(3).Control(1)=   "Opcion31"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "Titulo32"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Titulo31"
      Tab(3).Control(3).Enabled=   0   'False
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
         TabIndex        =   56
         Top             =   2880
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
         TabIndex        =   54
         Top             =   1080
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
         Left            =   -70320
         TabIndex        =   51
         Top             =   5400
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
         Left            =   -70320
         TabIndex        =   50
         Top             =   5040
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
         TabIndex        =   43
         Top             =   3960
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
         TabIndex        =   42
         Top             =   3600
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
         TabIndex        =   41
         Top             =   3240
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
         TabIndex        =   40
         Top             =   2880
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
         TabIndex        =   39
         Top             =   4680
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
         TabIndex        =   38
         Top             =   4320
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
         TabIndex        =   36
         Top             =   2520
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
         TabIndex        =   34
         Top             =   2160
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
         TabIndex        =   32
         Top             =   720
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   2520
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1800
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
         TabIndex        =   17
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
         TabIndex        =   15
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
         TabIndex        =   13
         Top             =   720
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
         TabIndex        =   57
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Titulo32 
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
         Top             =   1080
         Width           =   4575
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
         Left            =   -74520
         TabIndex        =   53
         Top             =   5400
         Width           =   3975
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
         Height          =   255
         Left            =   -74520
         TabIndex        =   52
         Top             =   5040
         Width           =   4095
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
         TabIndex        =   49
         Top             =   3960
         Width           =   4095
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
         Height          =   255
         Left            =   -74520
         TabIndex        =   48
         Top             =   3600
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
         Height          =   255
         Left            =   -74520
         TabIndex        =   47
         Top             =   3240
         Width           =   4095
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
         Height          =   375
         Left            =   -74520
         TabIndex        =   46
         Top             =   2880
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
         TabIndex        =   45
         Top             =   4680
         Width           =   3975
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
         TabIndex        =   44
         Top             =   4320
         Width           =   4095
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
         TabIndex        =   37
         Top             =   2520
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
         TabIndex        =   35
         Top             =   2160
         Width           =   3855
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
         TabIndex        =   33
         Top             =   720
         Width           =   4575
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   2520
         Width           =   3975
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   1800
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
         TabIndex        =   16
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
         TabIndex        =   14
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
         TabIndex        =   12
         Top             =   720
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
Attribute VB_Name = "PrgConfigCapacitacion"
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
    
    Titulo1.Caption = "Ingreso de Sectores"
    Titulo2.Caption = "Ingreso de Tema"
    Titulo3.Caption = "Ingreso de Cursos"
    Titulo4.Caption = "Ingreso de Perfiles"
    Titulo5.Caption = "Ingreso de Legajos"
    Titulo6.Caption = "Consulta de Verison de Legajos"
    Titulo7.Caption = "Consulta de Version de Perfiles"
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    
    
    Rem titulo2
    
    Titulo11.Caption = "Ingreso de Planificacion Anual de Capacitacion por Legajo"
    Titulo12.Caption = "Ingreso de Cronograma de Capacitacion"
    Titulo13.Caption = "Ingreso de Cursos Realizados"
    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0
    
    
    Rem titulo3
    
    Titulo21.Caption = "Perfil de Puesto"
    Titulo22.Caption = "Informe de Competencia y Necesidades de Capacitacion"
    Titulo23.Caption = "Listado de Temas Realizados por Legajos"
    Titulo24.Caption = "Listado de Cursos Realizados por Tema"
    Titulo25.Caption = "Listado de Evolucion de Temas Programados"
    Titulo26.Caption = "Listado de Legajos con Necesidades Pendientes por IC y NC vigente"
    Titulo27.Caption = "Listado de Temas Realizados por Sector"
    Titulo28.Caption = "Plan de Capacitacion Anual"
    Titulo29.Caption = "Listado de Legajos por Perfil"
    Titulo210.Caption = "Planilla de Temas no Programados"
    Titulo211.Caption = "Listado de Temas Realizados y No Realizados"
    Titulo212.Caption = "Listado de Temas Realizados y No Realizados"
    Titulo213.Caption = "Listado de Temas Realizados por Legajos (Consolidado)"
    Titulo214.Caption = "ListaInactivo"
    
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
    
    
    
    Rem titulo4
    
    Titulo31.Caption = "Fin del Sistema"
    
    Opcion31.Value = 0
   
    

End Sub


Private Sub Graba_Click()

    XParam = "'" + Operador.Text + "','" _
                 + "3" + "'"
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
    
    WProceso = "3"
                                       
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
                    
    Opcion11.Value = 0
    Opcion12.Value = 0
    Opcion13.Value = 0

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
                    
    Opcion31.Value = 0
    Opcion32.Value = 0
    
    Operador.SetFocus
    
    Tablas.Tab = 0

End Sub


Private Sub Label1_Click()

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
                    
                Opcion11.Value = 0
                Opcion12.Value = 0
                Opcion13.Value = 0
        
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
                    
                Opcion31.Value = 0
                Opcion32.Value = 0
                
                
                XParam = "'" + Operador.Text + "','" _
                             + "3" + "'"
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
                    
                    Opcion11.Value = Val(Mid$(rstAtributos!atributo2, 1, 1))
                    Opcion12.Value = Val(Mid$(rstAtributos!atributo2, 2, 1))
                    Opcion13.Value = Val(Mid$(rstAtributos!atributo2, 3, 1))
                    
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
                    
                    Opcion31.Value = Val(Mid$(rstAtributos!atributo4, 1, 1))
                    Opcion32.Value = Val(Mid$(rstAtributos!atributo4, 2, 1))
                    
                    rstAtributos.Close
                End If
                
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salida_Click()
    Operador.SetFocus
    PrgConfigCapacitacion.Hide
    Unload Me
    Menu.Show
End Sub

