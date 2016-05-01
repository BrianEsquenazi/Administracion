VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaSacAdicional 
   Caption         =   "Carga de SAC - Datos Adicionales"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11775
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   27
      Top             =   2040
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7858
      _Version        =   327680
      Tabs            =   10
      Tab             =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Comentario 1"
      TabPicture(0)   =   "cargasacadicional.frx":0000
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Dato1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Comentario 2"
      TabPicture(1)   =   "cargasacadicional.frx":001C
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Dato2"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Comentario 3"
      TabPicture(2)   =   "cargasacadicional.frx":0038
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Dato3"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Comentario 4"
      TabPicture(3)   =   "cargasacadicional.frx":0054
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Dato4"
      Tab(3).Control(0).Enabled=   -1  'True
      TabCaption(4)   =   "Comentario 5"
      TabPicture(4)   =   "cargasacadicional.frx":0070
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Dato5"
      Tab(4).Control(0).Enabled=   -1  'True
      TabCaption(5)   =   "Foto 1"
      TabPicture(5)   =   "cargasacadicional.frx":008C
      Tab(5).ControlCount=   5
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "MuestraFoto1"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Foto1"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "File1"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Dir1"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Drive1"
      Tab(5).Control(4).Enabled=   0   'False
      TabCaption(6)   =   "Foto 2"
      TabPicture(6)   =   "cargasacadicional.frx":00A8
      Tab(6).ControlCount=   5
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "MuestraFoto2"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Foto2"
      Tab(6).Control(1).Enabled=   -1  'True
      Tab(6).Control(2)=   "Dir2"
      Tab(6).Control(2).Enabled=   -1  'True
      Tab(6).Control(3)=   "File2"
      Tab(6).Control(3).Enabled=   -1  'True
      Tab(6).Control(4)=   "Drive2"
      Tab(6).Control(4).Enabled=   -1  'True
      TabCaption(7)   =   "Foto 3"
      TabPicture(7)   =   "cargasacadicional.frx":00C4
      Tab(7).ControlCount=   5
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "MuestraFoto3"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Foto3"
      Tab(7).Control(1).Enabled=   -1  'True
      Tab(7).Control(2)=   "Dir3"
      Tab(7).Control(2).Enabled=   -1  'True
      Tab(7).Control(3)=   "File3"
      Tab(7).Control(3).Enabled=   -1  'True
      Tab(7).Control(4)=   "Drive3"
      Tab(7).Control(4).Enabled=   -1  'True
      TabCaption(8)   =   "Foto 4"
      TabPicture(8)   =   "cargasacadicional.frx":00E0
      Tab(8).ControlCount=   5
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "MuestraFoto4"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Foto4"
      Tab(8).Control(1).Enabled=   -1  'True
      Tab(8).Control(2)=   "Dir4"
      Tab(8).Control(2).Enabled=   -1  'True
      Tab(8).Control(3)=   "File4"
      Tab(8).Control(3).Enabled=   -1  'True
      Tab(8).Control(4)=   "Drive4"
      Tab(8).Control(4).Enabled=   -1  'True
      TabCaption(9)   =   "Foto 5"
      TabPicture(9)   =   "cargasacadicional.frx":00FC
      Tab(9).ControlCount=   4
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "MuestraFoto5"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Foto5"
      Tab(9).Control(1).Enabled=   -1  'True
      Tab(9).Control(2)=   "Dir5"
      Tab(9).Control(2).Enabled=   -1  'True
      Tab(9).Control(3)=   "File5"
      Tab(9).Control(3).Enabled=   -1  'True
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   -71520
         TabIndex        =   51
         Top             =   1080
         Width           =   3975
      End
      Begin VB.DriveListBox Drive4 
         Height          =   315
         Left            =   -71640
         TabIndex        =   43
         Top             =   960
         Width           =   3975
      End
      Begin VB.DriveListBox Drive3 
         Height          =   315
         Left            =   -71520
         TabIndex        =   40
         Top             =   960
         Width           =   3975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   3480
         TabIndex        =   35
         Top             =   960
         Width           =   3975
      End
      Begin VB.FileListBox File5 
         Height          =   2040
         Left            =   -67440
         TabIndex        =   50
         Top             =   960
         Width           =   3495
      End
      Begin VB.FileListBox File4 
         Height          =   2040
         Left            =   -67560
         TabIndex        =   49
         Top             =   960
         Width           =   3495
      End
      Begin VB.FileListBox File3 
         Height          =   2040
         Left            =   -67440
         TabIndex        =   48
         Top             =   960
         Width           =   3495
      End
      Begin VB.FileListBox File2 
         Height          =   2040
         Left            =   -67440
         TabIndex        =   47
         Top             =   1080
         Width           =   3495
      End
      Begin VB.DirListBox Dir5 
         Height          =   1665
         Left            =   -71520
         TabIndex        =   46
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox Foto5 
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
         Top             =   3480
         Width           =   9615
      End
      Begin VB.DirListBox Dir4 
         Height          =   1665
         Left            =   -71640
         TabIndex        =   44
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox Foto4 
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
         Left            =   -74880
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         Text            =   " "
         Top             =   3480
         Width           =   9615
      End
      Begin VB.DirListBox Dir3 
         Height          =   1665
         Left            =   -71520
         TabIndex        =   41
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox Foto3 
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
         TabIndex        =   39
         Text            =   " "
         Top             =   3480
         Width           =   9615
      End
      Begin VB.DirListBox Dir2 
         Height          =   1665
         Left            =   -71520
         TabIndex        =   38
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox Foto2 
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
         TabIndex        =   37
         Text            =   " "
         Top             =   3600
         Width           =   9615
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   3480
         TabIndex        =   36
         Top             =   1560
         Width           =   3975
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   7680
         TabIndex        =   34
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox Foto1 
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         Text            =   " "
         Top             =   3480
         Width           =   9615
      End
      Begin VB.TextBox Dato5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   840
         Width           =   11055
      End
      Begin VB.TextBox Dato4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   720
         Width           =   11055
      End
      Begin VB.TextBox Dato2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   840
         Width           =   11055
      End
      Begin VB.TextBox Dato3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   11055
      End
      Begin VB.TextBox Dato1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   840
         Width           =   11055
      End
      Begin VB.Image MuestraFoto5 
         Height          =   2295
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2895
      End
      Begin VB.Image MuestraFoto4 
         Height          =   2295
         Left            =   -74880
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2895
      End
      Begin VB.Image MuestraFoto3 
         Height          =   2295
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2895
      End
      Begin VB.Image MuestraFoto2 
         Height          =   2295
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Image MuestraFoto1 
         Height          =   2295
         Left            =   240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2895
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   25
      Text            =   " "
      Top             =   1200
      Width           =   10455
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
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
      Left            =   4200
      MaxLength       =   6
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   735
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   9
      Text            =   " "
      Top             =   1560
      Width           =   10455
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
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   8
      Text            =   " "
      Top             =   840
      Width           =   855
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   7
      Text            =   " "
      Top             =   840
      Width           =   855
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2535
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
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   2775
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
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   3
      Text            =   " "
      Top             =   120
      Width           =   855
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
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   120
      Width           =   1095
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
      Left            =   1680
      TabIndex        =   12
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
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
   Begin VB.Label Label11 
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
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   975
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
      Left            =   2040
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
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
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
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
      Left            =   5160
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label13 
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
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   975
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
      Left            =   7440
      TabIndex        =   20
      Top             =   840
      Width           =   3135
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
      Left            =   2160
      TabIndex        =   19
      Top             =   840
      Width           =   3135
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
      Left            =   5400
      TabIndex        =   18
      Top             =   840
      Width           =   1215
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
      Left            =   3960
      TabIndex        =   17
      Top             =   480
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
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   975
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
      Left            =   7920
      TabIndex        =   15
      Top             =   480
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
      Left            =   9240
      TabIndex        =   14
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
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
      Left            =   7200
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6600
      MouseIcon       =   "cargasacadicional.frx":0118
      MousePointer    =   99  'Custom
      Picture         =   "cargasacadicional.frx":0422
      ToolTipText     =   "Salida"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "cargasacadicional.frx":0C64
      MousePointer    =   99  'Custom
      Picture         =   "cargasacadicional.frx":0F6E
      ToolTipText     =   "Elimina el Registro"
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "cargasacadicional.frx":17B0
      MousePointer    =   99  'Custom
      Picture         =   "cargasacadicional.frx":1ABA
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5640
      MouseIcon       =   "cargasacadicional.frx":22FC
      MousePointer    =   99  'Custom
      Picture         =   "cargasacadicional.frx":2606
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "PrgCargaSacAdicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String
Dim rstCargaSacAdicional As Recordset
Dim spCargaSacAdicional As String

Dim XParam As String
Dim ZZLugar As Integer

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

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
    
    Call CargaFotos
    
End Sub

Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

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
        
        rstCargaSac.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSacAdicional"
    ZSql = ZSql + " Where CargaSacAdicional.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSacAdicional.Numero = " + "'" + Numero.Text + "'"
    spCargaSacAdicional = ZSql
    Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSacAdicional.RecordCount > 0 Then
    
        Dato1.Text = rstCargaSacAdicional!Dato1
        Dato2.Text = rstCargaSacAdicional!Dato2
        Dato3.Text = rstCargaSacAdicional!Dato3
        Dato4.Text = rstCargaSacAdicional!Dato4
        Dato5.Text = rstCargaSacAdicional!Dato5
        
        Foto1.Text = rstCargaSacAdicional!Foto1
        Foto2.Text = rstCargaSacAdicional!Foto2
        Foto3.Text = rstCargaSacAdicional!Foto3
        Foto4.Text = rstCargaSacAdicional!Foto4
        Foto5.Text = rstCargaSacAdicional!Foto5
        
        
        rstCargaSacAdicional.Close
    End If
    
    Call Imprime_Descripcion
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub cmdAdd_Click()

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
        
        Auxi3 = Tipo.Text
        Auxi1 = Ano.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi3, 4)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 6)
        WClave = Auxi3 + Auxi1 + Auxi2
        
        For Ciclo = 1 To 5
        
            Select Case Ciclo
                Case 1
                    ZLargo = Len(Dato1.Text)
                    ZDato = Dato1.Text
                Case 2
                    ZLargo = Len(Dato2.Text)
                    ZDato = Dato2.Text
                Case 3
                    ZLargo = Len(Dato3.Text)
                    ZDato = Dato3.Text
                Case 4
                    ZLargo = Len(Dato4.Text)
                    ZDato = Dato4.Text
                Case Else
                    ZLargo = Len(Dato5.Text)
                    ZDato = Dato5.Text
            End Select
            
            For CicloII = 1 To ZLargo
                If Mid$(ZDato, CicloII, 1) = Chr$(39) Then
                    ZDato = Left$(ZDato, CicloII - 1) + " " + Mid$(ZDato, CicloII + 1, ZLargo)
                End If
            Next CicloII
        
            Select Case Ciclo
                Case 1
                    Dato1.Text = ZDato
                Case 2
                    Dato2.Text = ZDato
                Case 3
                    Dato3.Text = ZDato
                Case 4
                    Dato4.Text = ZDato
                Case Else
                    Dato5.Text = ZDato
            End Select
        
        Next Ciclo
        
    
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
            ZSql = ZSql + " Dato1 = " + "'" + Dato1.Text + "',"
            ZSql = ZSql + " Dato2 = " + "'" + Dato2.Text + "',"
            ZSql = ZSql + " Dato3 = " + "'" + Dato3.Text + "',"
            ZSql = ZSql + " Dato4 = " + "'" + Dato4.Text + "',"
            ZSql = ZSql + " Dato5 = " + "'" + Dato5.Text + "',"
            ZSql = ZSql + " Foto1 = " + "'" + Foto1.Text + "',"
            ZSql = ZSql + " Foto2 = " + "'" + Foto2.Text + "',"
            ZSql = ZSql + " Foto3 = " + "'" + Foto3.Text + "',"
            ZSql = ZSql + " Foto4 = " + "'" + Foto4.Text + "',"
            ZSql = ZSql + " Foto5 = " + "'" + Foto5.Text + "'"
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
            ZSql = ZSql + "'" + Dato1.Text + "',"
            ZSql = ZSql + "'" + Dato2.Text + "',"
            ZSql = ZSql + "'" + Dato3.Text + "',"
            ZSql = ZSql + "'" + Dato4.Text + "',"
            ZSql = ZSql + "'" + Dato5.Text + "',"
            ZSql = ZSql + "'" + Foto1.Text + "',"
            ZSql = ZSql + "'" + Foto2.Text + "',"
            ZSql = ZSql + "'" + Foto3.Text + "',"
            ZSql = ZSql + "'" + Foto4.Text + "',"
            ZSql = ZSql + "'" + Foto5.Text + "')"
            
            spCargaSacAdicional = ZSql
            Set rstCargaSacAdicional = db.OpenRecordset(spCargaSacAdicional, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        Call CmdLimpiar_Click
        Tipo.SetFocus
        
    End If
End Sub


Private Sub CmdLimpiar_Click()
    
    
    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
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
    
    
    
    Dato1.Text = ""
    Dato2.Text = ""
    Dato3.Text = ""
    Dato4.Text = ""
    Dato5.Text = ""
    Foto1.Text = ""
    Foto2.Text = ""
    Foto3.Text = ""
    Foto4.Text = ""
    Foto5.Text = ""
    
    SSTab1.Tab = 0
    
    
    Tipo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgCargaSacAdicional.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    Tipo.Text = "1"
    DesTipo.Caption = "SAC"
    Ano.Text = "2010"
    Numero.Text = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    Fecha.Text = "  /  /    "
    ResponsableEmisor.Text = ""
    ResponsableDestino.Text = ""
    Referencia.Text = ""
    Titulo.Text = ""
    DesResponsableEmisor.Caption = ""
    DesResponsableDestino.Caption = ""
    
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

    Dato1.Text = ""
    Dato2.Text = ""
    Dato3.Text = ""
    Dato4.Text = ""
    Dato5.Text = ""
    Foto1.Text = ""
    Foto2.Text = ""
    Foto3.Text = ""
    Foto4.Text = ""
    Foto5.Text = ""
    
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Imagen3_Click()

End Sub

Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
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
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        DesTipo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            
                rstCargaSac.Close
                Call Imprime_Datos
                
                    Else
                    
                WTipo = Tipo.Text
                WAno = Ano.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Tipo.Text = WTipo
                Ano.Text = WAno
                Numero.Text = WNumero
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
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub





Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub

Private Sub File1_dblClick()

    On Error GoTo WError

    WDrive = Drive1.Drive
    WDir = Dir1.Path
    WLon = Len(WDir)
    If Right$(WDir, 1) = "\" Then
        WDir = Mid(WDir, 1, WLon - 1)
    End If
    XNombre = WDir + "\"
    
    WPasoUnifica = XNombre + File1.filename
    
    Foto1.Text = WPasoUnifica
    
    Exit Sub
    
WError:
    MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub

Private Sub File1_Click()

    On Error GoTo WError
    
    WDrive = Drive1.Drive
    WDir = Dir1.Path
    WLon = Len(WDir)
    If Right$(WDir, 1) = "\" Then
        WDir = Mid(WDir, 1, WLon - 1)
    End If
    XNombre = WDir + "\"
    
    WPasoUnifica = XNombre + File1.filename
    
    MuestraFoto1.Picture = LoadPicture("")
    MuestraFoto1.Picture = LoadPicture(WPasoUnifica)
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub







Private Sub CargaFotos()

    On Error GoTo WError
    
    MuestraFoto1.Picture = LoadPicture("")
    MuestraFoto1.Picture = LoadPicture(Foto1.Text)
    
    MuestraFoto2.Picture = LoadPicture("")
    MuestraFoto2.Picture = LoadPicture(Foto2.Text)
    
    MuestraFoto3.Picture = LoadPicture("")
    MuestraFoto3.Picture = LoadPicture(Foto3.Text)
    
    MuestraFoto4.Picture = LoadPicture("")
    MuestraFoto4.Picture = LoadPicture(Foto4.Text)
    
    MuestraFoto5.Picture = LoadPicture("")
    MuestraFoto5.Picture = LoadPicture(Foto5.Text)
    
    MuestraFoto5.Picture = LoadPicture("")
    MuestraFoto5.Picture = LoadPicture(Foto6.Text)
    
    Exit Sub
    
WError:
    Rem MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    Resume Next
    
End Sub

