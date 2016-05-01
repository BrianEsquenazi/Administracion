VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdenTrabajo 
   Caption         =   "Ingreso de Orden de Trabajo"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11625
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   10440
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"ordentrabajo.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   1980
      ItemData        =   "ordentrabajo.frx":007C
      Left            =   1680
      List            =   "ordentrabajo.frx":0083
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
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
      Left            =   2880
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
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
      Left            =   1200
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Observaciones 
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
      MaxLength       =   50
      TabIndex        =   9
      Text            =   " "
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox Cliente 
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
      MaxLength       =   6
      TabIndex        =   6
      Text            =   " "
      Top             =   600
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSMask.MaskEdBox FechaEntrega 
      Height          =   285
      Left            =   8640
      TabIndex        =   4
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSMask.MaskEdBox Orden 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####"
      PromptChar      =   " "
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5895
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   10398
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion de la Orden"
      TabPicture(0)   =   "ordentrabajo.frx":0091
      Tab(0).ControlCount=   18
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Encargado"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ObservacionesI"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ObservacionesIII"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ObservacionesII"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DescripcionV"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DescripcionIV"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DescripcionIII"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DescripcionII"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "DescripcionI"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Uso"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Muestra"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Material"
      Tab(0).Control(17).Enabled=   0   'False
      TabCaption(1)   =   "Requisitos"
      TabPicture(1)   =   "ordentrabajo.frx":00AD
      Tab(1).ControlCount=   20
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ReferenciaII"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "ReferenciaI"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "RequisitoVI"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "RequisitoV"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "RequisitoIV"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "RequisitoIII"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "RequisitoII"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "RequisitoI"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Fin_Nota_Estabilidad"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "NotaEstabilidad"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Fin_Nota_Aplicacion"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "NotaAplicacion"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Estabilidad"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "Aplicacion"
      Tab(1).Control(19).Enabled=   -1  'True
      Begin VB.ComboBox Aplicacion 
         Height          =   315
         Left            =   -72960
         TabIndex        =   52
         Top             =   3900
         Width           =   2055
      End
      Begin VB.ComboBox Estabilidad 
         Height          =   315
         Left            =   -72960
         TabIndex        =   51
         Top             =   4380
         Width           =   2055
      End
      Begin VB.CommandButton NotaAplicacion 
         Caption         =   "Notas Aplicacion"
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
         Left            =   -70440
         TabIndex        =   50
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton Fin_Nota_Aplicacion 
         Caption         =   "Cierre Nota"
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
         Left            =   -65160
         TabIndex        =   49
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton NotaEstabilidad 
         Caption         =   "Notas Estabilidad"
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
         Left            =   -70440
         TabIndex        =   48
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton Fin_Nota_Estabilidad 
         Caption         =   "Cierre Nota"
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
         Left            =   -65160
         TabIndex        =   47
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Material 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   36
         Text            =   " "
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox Muestra 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   35
         Text            =   " "
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox Uso 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   34
         Text            =   " "
         Top             =   1200
         Width           =   5895
      End
      Begin VB.TextBox DescripcionI 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   33
         Text            =   " "
         Top             =   1800
         Width           =   5895
      End
      Begin VB.TextBox DescripcionII 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   32
         Text            =   " "
         Top             =   2160
         Width           =   5895
      End
      Begin VB.TextBox DescripcionIII 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   31
         Text            =   " "
         Top             =   2520
         Width           =   5895
      End
      Begin VB.TextBox DescripcionIV 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   30
         Text            =   " "
         Top             =   2880
         Width           =   5895
      End
      Begin VB.TextBox DescripcionV 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   29
         Text            =   " "
         Top             =   3240
         Width           =   5895
      End
      Begin VB.TextBox RequisitoI 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   28
         Text            =   " "
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox RequisitoII 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   27
         Text            =   " "
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox RequisitoIII 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   26
         Text            =   " "
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox RequisitoIV 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   25
         Text            =   " "
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox RequisitoV 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   24
         Text            =   " "
         Top             =   2160
         Width           =   5895
      End
      Begin VB.TextBox RequisitoVI 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   23
         Text            =   " "
         Top             =   2520
         Width           =   5895
      End
      Begin VB.TextBox ReferenciaI 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   22
         Text            =   " "
         Top             =   3000
         Width           =   5895
      End
      Begin VB.TextBox ReferenciaII 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   21
         Text            =   " "
         Top             =   3360
         Width           =   5895
      End
      Begin VB.TextBox ObservacionesII 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   20
         Text            =   " "
         Top             =   4260
         Width           =   5895
      End
      Begin VB.TextBox ObservacionesIII 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   19
         Text            =   " "
         Top             =   4620
         Width           =   5895
      End
      Begin VB.TextBox ObservacionesI 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   18
         Text            =   " "
         Top             =   3900
         Width           =   5895
      End
      Begin VB.TextBox Encargado 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   17
         Text            =   " "
         Top             =   5220
         Width           =   5895
      End
      Begin VB.Label Label15 
         Caption         =   "Aplicacion"
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
         Left            =   -74640
         TabIndex        =   54
         Top             =   3900
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Estabilidad"
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
         Left            =   -74640
         TabIndex        =   53
         Top             =   4380
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Material Provisto por el Cliente"
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
         Left            =   480
         TabIndex        =   46
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Muestra"
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
         Left            =   480
         TabIndex        =   45
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Uso"
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
         Left            =   480
         TabIndex        =   44
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Descripcion del Trabajo"
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
         Left            =   480
         TabIndex        =   43
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Requisitos Funcionales"
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
         Left            =   -74640
         TabIndex        =   42
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "Otros Requisitos"
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
         Left            =   -74640
         TabIndex        =   41
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label12 
         Caption         =   "Requisitos Legales / Normas / Regulaciones"
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
         Height          =   615
         Left            =   -74640
         TabIndex        =   40
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label13 
         Caption         =   "Referencias "
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
         Left            =   -74640
         TabIndex        =   39
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label18 
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
         Height          =   375
         Left            =   480
         TabIndex        =   38
         Top             =   3900
         Width           =   2895
      End
      Begin VB.Label Label19 
         Caption         =   "Encargado del Proyecto"
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
         Left            =   480
         TabIndex        =   37
         Top             =   5220
         Width           =   2895
      End
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   6240
      MouseIcon       =   "ordentrabajo.frx":00C9
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajo.frx":03D3
      ToolTipText     =   "Consulta de Datos"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6960
      MouseIcon       =   "ordentrabajo.frx":0C15
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajo.frx":0F1F
      ToolTipText     =   "Salida"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   4680
      MouseIcon       =   "ordentrabajo.frx":1761
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajo.frx":1A6B
      ToolTipText     =   "Elimina el Registro"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "ordentrabajo.frx":22AD
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajo.frx":25B7
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "ordentrabajo.frx":2DF9
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajo.frx":3103
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label10 
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
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente"
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
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label DesCliente 
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
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Comrpometida"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   240
      Width           =   1935
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
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Trabajo"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "PrgOrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrdenTrabajo As Recordset
Dim spOrdenTrabajo As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim EmpresaActual As String
Private XEmpresa As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM OrdenTrabajo"
    ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + Orden.Text + "'"
    spOrdenTrabajo = ZSql
    Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrdenTrabajo.RecordCount > 0 Then
        Fecha.Text = rstOrdenTrabajo!Fecha
        FechaEntrega.Text = rstOrdenTrabajo!FechaEntrega
        Cliente.Text = rstOrdenTrabajo!Cliente
        Observaciones.Text = Trim(rstOrdenTrabajo!Observaciones)
        Material.Text = Trim(rstOrdenTrabajo!Material)
        Muestra.Text = Trim(rstOrdenTrabajo!Muestra)
        Uso.Text = Trim(rstOrdenTrabajo!Uso)
        DescripcionI.Text = Trim(rstOrdenTrabajo!DescripcionI)
        DescripcionII.Text = Trim(rstOrdenTrabajo!DescripcionII)
        DescripcionIII.Text = Trim(rstOrdenTrabajo!DescripcionIII)
        DescripcionIV.Text = Trim(rstOrdenTrabajo!DescripcionIV)
        DescripcionV.Text = Trim(rstOrdenTrabajo!DescripcionV)
        ObservacionesI.Text = Trim(rstOrdenTrabajo!ObservacionesI)
        ObservacionesII.Text = Trim(rstOrdenTrabajo!ObservacionesII)
        ObservacionesIII.Text = Trim(rstOrdenTrabajo!ObservacionesIII)
        RequisitoI.Text = Trim(rstOrdenTrabajo!RequisitoI)
        RequisitoII.Text = Trim(rstOrdenTrabajo!RequisitoII)
        RequisitoIII.Text = Trim(rstOrdenTrabajo!RequisitoIII)
        RequisitoIV.Text = Trim(rstOrdenTrabajo!RequisitoIV)
        RequisitoV.Text = Trim(rstOrdenTrabajo!RequisitoV)
        RequisitoVI.Text = Trim(rstOrdenTrabajo!RequisitoVI)
        ReferenciaI.Text = Trim(rstOrdenTrabajo!ReferenciaI)
        ReferenciaII.Text = Trim(rstOrdenTrabajo!ReferenciaII)
        Aplicacion.ListIndex = rstOrdenTrabajo!Aplicacion
        Estabilidad.ListIndex = rstOrdenTrabajo!Estabilidad
        Encargado.Text = Trim(rstOrdenTrabajo!Encargado)
        rstOrdenTrabajo.Close
    End If
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!razon
        rstCliente.Close
            Else
        DesCliente.Caption = ""
    End If
    
End Sub



Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub




Private Sub cmdAdd_Click()
 Call Conecta_Empresa
 
    
    
    
    
    If Orden.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenTrabajo"
        ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + Orden.Text + "'"
        spOrdenTrabajo = ZSql
        Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenTrabajo.RecordCount > 0 Then
            rstOrdenTrabajo.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE OrdenTrabajo SET "
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " FechaEntrega = " + "'" + FechaEntrega.Text + "',"
            ZSql = ZSql + " Cliente = " + "'" + Cliente.Text + "',"
            ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " Material = " + "'" + Material.Text + "',"
            ZSql = ZSql + " Muestra = " + "'" + Muestra.Text + "',"
            ZSql = ZSql + " Uso = " + "'" + Uso.Text + "',"
            ZSql = ZSql + " DescripcionI = " + "'" + DescripcionI.Text + "',"
            ZSql = ZSql + " DescripcionII = " + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + " DescripcionIII = " + "'" + DescripcionIII.Text + "',"
            ZSql = ZSql + " DescripcionIV = " + "'" + DescripcionIV.Text + "',"
            ZSql = ZSql + " DescripcionV = " + "'" + DescripcionV.Text + "',"
            ZSql = ZSql + " ObservacionesI = " + "'" + ObservacionesI.Text + "',"
            ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesII.Text + "',"
            ZSql = ZSql + " ObservacionesIII = " + "'" + ObservacionesIII.Text + "',"
            ZSql = ZSql + " Encargado = " + "'" + Encargado.Text + "',"
            ZSql = ZSql + " RequisitoI = " + "'" + RequisitoI.Text + "',"
            ZSql = ZSql + " RequisitoII = " + "'" + RequisitoII.Text + "',"
            ZSql = ZSql + " RequisitoIII = " + "'" + RequisitoIII.Text + "',"
            ZSql = ZSql + " RequisitoIV = " + "'" + RequisitoIV.Text + "',"
            ZSql = ZSql + " RequisitoV = " + "'" + RequisitoV.Text + "',"
            ZSql = ZSql + " RequisitoVI = " + "'" + RequisitoVI.Text + "',"
            ZSql = ZSql + " ReferenciaI = " + "'" + ReferenciaI.Text + "',"
            ZSql = ZSql + " ReferenciaII = " + "'" + ReferenciaII.Text + "',"
            ZSql = ZSql + " Aplicacion = " + "'" + Str$(Aplicacion.ListIndex) + "',"
            ZSql = ZSql + " Estabilidad = " + "'" + Str$(Estabilidad.ListIndex) + "'"
            ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
            spOrdenTrabajo = ZSql
            Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO OrdenTrabajo ("
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "FechaEntrega ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Material ,"
            ZSql = ZSql + "Muestra ,"
            ZSql = ZSql + "Uso ,"
            ZSql = ZSql + "DescripcionI ,"
            ZSql = ZSql + "DescripcionII ,"
            ZSql = ZSql + "DescripcionIII ,"
            ZSql = ZSql + "DescripcionIV ,"
            ZSql = ZSql + "DescripcionV ,"
            ZSql = ZSql + "ObservacionesI ,"
            ZSql = ZSql + "ObservacionesII ,"
            ZSql = ZSql + "ObservacionesIII ,"
            ZSql = ZSql + "Encargado ,"
            ZSql = ZSql + "RequisitoI ,"
            ZSql = ZSql + "RequisitoII ,"
            ZSql = ZSql + "RequisitoIII ,"
            ZSql = ZSql + "RequisitoIV ,"
            ZSql = ZSql + "RequisitoV ,"
            ZSql = ZSql + "RequisitoVI ,"
            ZSql = ZSql + "ReferenciaI ,"
            ZSql = ZSql + "ReferenciaII ,"
            ZSql = ZSql + "Aplicacion ,"
            ZSql = ZSql + "Estabilidad )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Orden.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + FechaEntrega.Text + "',"
            ZSql = ZSql + "'" + Cliente.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + Material.Text + "',"
            ZSql = ZSql + "'" + Muestra.Text + "',"
            ZSql = ZSql + "'" + Uso.Text + "',"
            ZSql = ZSql + "'" + DescripcionI.Text + "',"
            ZSql = ZSql + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + "'" + DescripcionIII.Text + "',"
            ZSql = ZSql + "'" + DescripcionIV.Text + "',"
            ZSql = ZSql + "'" + DescripcionV.Text + "',"
            ZSql = ZSql + "'" + ObservacionesI.Text + "',"
            ZSql = ZSql + "'" + ObservacionesII.Text + "',"
            ZSql = ZSql + "'" + ObservacionesIII.Text + "',"
            ZSql = ZSql + "'" + Encargado.Text + "',"
            ZSql = ZSql + "'" + RequisitoI.Text + "',"
            ZSql = ZSql + "'" + RequisitoII.Text + "',"
            ZSql = ZSql + "'" + RequisitoIII.Text + "',"
            ZSql = ZSql + "'" + RequisitoIV.Text + "',"
            ZSql = ZSql + "'" + RequisitoV.Text + "',"
            ZSql = ZSql + "'" + RequisitoVI.Text + "',"
            ZSql = ZSql + "'" + ReferenciaI.Text + "',"
            ZSql = ZSql + "'" + ReferenciaII.Text + "',"
            ZSql = ZSql + "'" + Str$(Aplicacion.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estabilidad.ListIndex) + "')"
            spOrdenTrabajo = ZSql
            Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call CmdLimpiar_Click
        Orden.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Orden.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenTrabajo"
        ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + Orden.Text + "'"
        spOrdenTrabajo = ZSql
        Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenTrabajo.RecordCount > 0 Then
            rstOrdenTrabajo.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                zSql1 = ZSql + "DELETE OrdenTrabajo"
                zSql2 = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
                spOrdenTrabajo = ZSql
                Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    Orden.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Orden.Text = "  -     "
    Fecha.Text = "  /  /    "
    FechaEntrega.Text = "  /  /    "
    Cliente.Text = ""
    Observaciones.Text = ""
    Material.Text = ""
    Muestra.Text = ""
    Uso.Text = ""
    DescripcionI.Text = ""
    DescripcionII.Text = ""
    DescripcionIII.Text = ""
    DescripcionIV.Text = ""
    DescripcionV.Text = ""
    ObservacionesI.Text = ""
    ObservacionesII.Text = ""
    ObservacionesIII.Text = ""
    Encargado.Text = ""
    RequisitoI.Text = ""
    RequisitoII.Text = ""
    RequisitoIII.Text = ""
    RequisitoIV.Text = ""
    RequisitoV.Text = ""
    RequisitoVI.Text = ""
    ReferenciaI.Text = ""
    ReferenciaII.Text = ""
    
    Aplicacion.ListIndex = 0
    Estabilidad.ListIndex = 0
    
    DesCliente.Caption = ""
    Tablas.Tab = 0

    Orden.SetFocus
    
End Sub

Private Sub CmdClose_Click()

    Call CmdLimpiar_Click
    PrgOrdenTrabajo.Hide
    Unload Me
    Menu.Show
    
End Sub



Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            FechaEntrega.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub FechaEntrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaEntrega.Text, Auxi)
        If Auxi = "S" Or FechaEntrega.Text = "  /  /    " Then
            Cliente.SetFocus
                Else
            FechaEntrega.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaEntrega.Text = "  /  /    "
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
                Cliente.Text = UCase(Cliente.Text)
                XEmpresa = WEmpresa
                txtOdbc = "Empresa" + "01"
                WEmpresa = "0001"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
        
        
        If Cliente.Text <> "" Then
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!razon
                rstCliente.Close
                Observaciones.SetFocus
                    Else
                Cliente.SetFocus
            End If
                Else
            DesCliente.Caption = ""
            Observaciones.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
Call Conecta_Empresa
End Sub

Private Sub NotaAplicacion_Click()

    On Error GoTo WError
        
    If Orden.Text <> "  -     " Then
    
        Fin_Nota_Aplicacion.Visible = True

        Agenda.LoadFile "blanco.rtf", 0
        Agenda.LoadFile "A" + Orden.Text + ".rtf", 0
        Agenda.Visible = True
        Agenda.Height = 6700
        Agenda.Left = 650
        Agenda.Top = 720
        Agenda.Width = 9375
        Agenda.SetFocus
        
    End If
    
WError:
    Resume Next
    
End Sub

Private Sub Fin_Nota_Aplicacion_Click()

    Agenda.SaveFile "A" + Orden.Text + ".rtf", 0
    Agenda.Visible = False
    Fin_Nota_Aplicacion.Visible = False
    Orden.SetFocus

End Sub

Private Sub NotaEstabilidad_Click()

    On Error GoTo WError

    If Orden.Text <> "  -     " Then
    
        Fin_Nota_Estabilidad.Visible = True
    
        Agenda.LoadFile "blanco.rtf", 0
        Agenda.LoadFile "E" + Orden.Text + ".rtf", 0
        Agenda.Visible = True
        Agenda.Height = 6700
        Agenda.Left = 650
        Agenda.Top = 720
        Agenda.Width = 9375
        Agenda.SetFocus
        
    End If
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Fin_Nota_Estabilidad_Click()

    Agenda.SaveFile "E" + Orden.Text + ".rtf", 0
    Agenda.Visible = False
    Fin_Nota_Estabilidad.Visible = False
    Orden.SetFocus

End Sub


Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tablas.Tab = 0
        Material.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub





Private Sub Material_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Muestra.SetFocus
    End If
    If KeyAscii = 27 Then
        Material.Text = ""
    End If
End Sub

Private Sub Muestra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Uso.SetFocus
    End If
    If KeyAscii = 27 Then
        Muestra.Text = ""
    End If
End Sub



Private Sub Uso_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionI.SetFocus
    End If
    If KeyAscii = 27 Then
        Uso.Text = ""
    End If
End Sub

Private Sub DescripcionI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionII.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionI.Text = ""
    End If
End Sub

Private Sub DescripcionII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionIII.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionII.Text = ""
    End If
End Sub

Private Sub DescripcionIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionIV.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionIII.Text = ""
    End If
End Sub

Private Sub DescripcionIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionV.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionIV.Text = ""
    End If
End Sub

Private Sub DescripcionV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesI.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionV.Text = ""
    End If
End Sub

Private Sub ObservacionesI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesI.Text = ""
    End If
End Sub

Private Sub ObservacionesII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesII.Text = ""
    End If
End Sub

Private Sub ObservacionesIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Encargado.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesIII.Text = ""
    End If
End Sub

Private Sub Encargado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Material.SetFocus
    End If
    If KeyAscii = 27 Then
        Encargado.Text = ""
    End If
End Sub







Private Sub RequisitoI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoII.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoI.Text = ""
    End If
End Sub

Private Sub RequisitoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoIII.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoII.Text = ""
    End If
End Sub

Private Sub RequisitoIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoIV.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoIII.Text = ""
    End If
End Sub

Private Sub RequisitoIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoV.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoIV.Text = ""
    End If
End Sub

Private Sub RequisitoV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ReferenciaI.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoV.Text = ""
    End If
End Sub

Private Sub ReferenciaI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ReferenciaII.SetFocus
    End If
    If KeyAscii = 27 Then
        ReferenciaI.Text = ""
    End If
End Sub

Private Sub ReferenciaII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoI.SetFocus
    End If
    If KeyAscii = 27 Then
        ReferenciaII.Text = ""
    End If
End Sub













Private Sub Aplicacion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estabilidad.SetFocus
    End If
End Sub

Private Sub Estabilidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesIV.SetFocus
    End If
End Sub

Private Sub ObservacionesIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesIV.Text = ""
    End If
End Sub

Private Sub ObservacionesV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesVI.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesV.Text = ""
    End If
End Sub

Private Sub ObservacionesVI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesIV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesVI.Text = ""
    End If
End Sub






Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Orden.Text <> "" Then
        
            Orden.Text = UCase(Orden.Text)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM OrdenTrabajo"
            ZSql = ZSql + " Where OrdenTrabajo.Orden = " + "'" + Orden.Text + "'"
            spOrdenTrabajo = ZSql
            Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrdenTrabajo.RecordCount > 0 Then
                rstOrdenTrabajo.Close
                Call Imprime_Datos
                    Else
                WOrden = Orden.Text
                CmdLimpiar_Click
                Orden.Text = WOrden
            End If
            
        End If
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Orden.Text = ""
    End If
End Sub

Sub Form_Load()

    Aplicacion.Clear
    
    Aplicacion.AddItem ""
    Aplicacion.AddItem "Si"
    Aplicacion.AddItem "No"
    
    Aplicacion.ListIndex = 0


    Estabilidad.Clear
    
    Estabilidad.AddItem ""
    Estabilidad.AddItem "Si"
    Estabilidad.AddItem "No"
    
    Estabilidad.ListIndex = 0

    Orden.Text = "  -     "
    Fecha.Text = "  /  /    "
    FechaEntrega.Text = "  /  /    "
    Cliente.Text = ""
    Observaciones.Text = ""
    Material.Text = ""
    Muestra.Text = ""
    Uso.Text = ""
    DescripcionI.Text = ""
    DescripcionII.Text = ""
    DescripcionIII.Text = ""
    DescripcionIV.Text = ""
    DescripcionV.Text = ""
    ObservacionesI.Text = ""
    ObservacionesII.Text = ""
    ObservacionesIII.Text = ""
    Encargado.Text = ""
    RequisitoI.Text = ""
    RequisitoII.Text = ""
    RequisitoIII.Text = ""
    RequisitoIV.Text = ""
    RequisitoV.Text = ""
    RequisitoVI.Text = ""
    ReferenciaI.Text = ""
    ReferenciaII.Text = ""
    
    DesCliente.Caption = ""
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    Select Case Tablas.Tab
        Case 0
            Material.SetFocus
        Case 1
            RequisitoI.SetFocus
        Case Else
    End Select
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"
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
    
    
     
     
     Rem BY NAN 12-9-11
    
    XEmpresa = WEmpresa
    
    txtOdbc = "Empresa" + "01"
    WEmpresa = "0001"
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
       
       
       spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
       
         
       
       Rem     ZSql = ""
        Rem    ZSql = ZSql + "Select *"
        Rem    ZSql = ZSql + " FROM Cliente"
        Rem    ZSql = ZSql + " Order by Cliente"
        Rem    spCliente = ZSql
         Rem   Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
         Rem   If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    
              
                    
                    Do
                 
                         If .EOF = False Then
                            
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
   Rem BY NAN
    Rem WEmpresa = oldempresa
    Call Conecta_Empresa
    
    Rem END BY NAN
    
    
    Exit Sub
    
WError:
    Resume Next

End Sub


Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
Ayuda.Text = UCase(Ayuda)
        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        XIndice = Opcion.ListIndex
    
        Select Case XIndice
            Case 0
                Rem by nan 12-9-11
                XEmpresa = WEmpresa
                txtOdbc = "Empresa" + "01"
                WEmpresa = "0001"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                Rem fin by nan
      Rem BY NAN 7-10
            spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
                
                
       Rem FIN BY NAN
                
       Rem     ZSql = ""
       Rem     ZSql = ZSql + "Select *"
       Rem     ZSql = ZSql + " FROM Cliente"
       Rem     ZSql = ZSql + " Order by Cliente"
       Rem     spCliente = ZSql
       Rem     Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
       Rem     If rstCliente.RecordCount > 0 Then
       
    
                    With rstCliente
                        .MoveFirst
                        Do
                            If .EOF = False Then
            
       
            
            
            
            
                                DA = Len(rstCliente!razon) - WEspacios
                
                                For Aaa = 1 To DA + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstCliente!razon, Aaa, WEspacios) Then
                                        IngresaItem = rstCliente!Cliente + " " + rstCliente!razon
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstCliente!Cliente
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next Aaa
                                .MoveNext
                    
                                    Else
                        
                                Exit Do
                
                            End If
                        Loop
                    End With
    
                    rstCliente.Close
    
                End If
                
            Case Else
        End Select
    
    End If
    
    Call Conecta_Empresa
    
End Sub










