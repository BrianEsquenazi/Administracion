VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgpedidoOrdenTrabajo 
   Caption         =   "Ingreso de Pedido de Desarrollo"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11625
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
      Left            =   720
      TabIndex        =   38
      Top             =   5520
      Visible         =   0   'False
      Width           =   4215
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
      ItemData        =   "pedidoordentrabajo.frx":0000
      Left            =   480
      List            =   "pedidoordentrabajo.frx":0007
      TabIndex        =   41
      Top             =   5400
      Visible         =   0   'False
      Width           =   6855
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
      Left            =   480
      TabIndex        =   51
      Top             =   5160
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Pedido 
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
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Vendedor 
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
      MaxLength       =   6
      TabIndex        =   42
      Text            =   " "
      Top             =   960
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   10440
      TabIndex        =   40
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
      TextRTF         =   $"pedidoordentrabajo.frx":0015
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
   Begin TabDlg.SSTab Tablas 
      Height          =   5535
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   9763
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion del Pedido"
      TabPicture(0)   =   "pedidoordentrabajo.frx":0091
      Tab(0).ControlCount=   20
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label21"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Material"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Muestra"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Uso"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DescripcionI"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DescripcionII"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DescripcionIII"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DescripcionIV"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "DescripcionV"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ObservacionesII"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ObservacionesIII"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "ObservacionesI"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Volumen"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Costo"
      Tab(0).Control(19).Enabled=   0   'False
      TabCaption(1)   =   "Especificaciones"
      TabPicture(1)   =   "pedidoordentrabajo.frx":00AD
      Tab(1).ControlCount=   16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "RequisitoI"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "RequisitoII"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "RequisitoIII"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "RequisitoIV"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "RequisitoV"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "RequisitoVI"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "ReferenciaI"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "ReferenciaII"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "ReferenciaIII"
      Tab(1).Control(15).Enabled=   -1  'True
      TabCaption(2)   =   "Respuesta"
      TabPicture(2)   =   "pedidoordentrabajo.frx":00C9
      Tab(2).ControlCount=   14
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fghg"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label23"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label24"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "CodigoOrden"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Respuesta"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "RespuestaI"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "RespuestaIII"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "RespuestaII"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Destino"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "RespuestaV"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "RespuestaVI"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "RespuestaIV"
      Tab(2).Control(13).Enabled=   0   'False
      Begin VB.TextBox RespuestaIV 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   64
         Text            =   " "
         Top             =   4080
         Width           =   5895
      End
      Begin VB.TextBox RespuestaVI 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   63
         Text            =   " "
         Top             =   4800
         Width           =   5895
      End
      Begin VB.TextBox RespuestaV 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   62
         Text            =   " "
         Top             =   4440
         Width           =   5895
      End
      Begin VB.ComboBox Destino 
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
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Costo 
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
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   58
         Text            =   " "
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox Volumen 
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
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   56
         Text            =   " "
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox ReferenciaIII 
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
         TabIndex        =   53
         Text            =   " "
         Top             =   4200
         Width           =   5895
      End
      Begin VB.TextBox RespuestaII 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Text            =   " "
         Top             =   1920
         Width           =   5895
      End
      Begin VB.TextBox RespuestaIII 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Text            =   " "
         Top             =   2280
         Width           =   5895
      End
      Begin VB.TextBox RespuestaI 
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
         Left            =   -71400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   1560
         Width           =   5895
      End
      Begin VB.ComboBox Respuesta 
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
         Left            =   -71400
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   840
         Width           =   3375
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   36
         Text            =   " "
         Top             =   3540
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   35
         Text            =   " "
         Top             =   4260
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   34
         Text            =   " "
         Top             =   3900
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
         TabIndex        =   33
         Text            =   " "
         Top             =   3840
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
         TabIndex        =   31
         Text            =   " "
         Top             =   3480
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
         TabIndex        =   30
         Text            =   " "
         Top             =   2640
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
         TabIndex        =   28
         Text            =   " "
         Top             =   2280
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
         TabIndex        =   27
         Text            =   " "
         Top             =   1800
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
         TabIndex        =   25
         Text            =   " "
         Top             =   1440
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
         TabIndex        =   24
         Text            =   " "
         Top             =   960
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
         TabIndex        =   22
         Text            =   " "
         Top             =   600
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   21
         Text            =   " "
         Top             =   3120
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   20
         Text            =   " "
         Top             =   2760
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   19
         Text            =   " "
         Top             =   2400
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   18
         Text            =   " "
         Top             =   2040
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   1680
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   14
         Text            =   " "
         Top             =   1320
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   960
         Width           =   5895
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
         Left            =   3360
         MaxLength       =   50
         TabIndex        =   10
         Text            =   " "
         Top             =   600
         Width           =   5895
      End
      Begin MSMask.MaskEdBox CodigoOrden 
         Height          =   285
         Left            =   -71400
         TabIndex        =   67
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
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
      Begin VB.Label Label24 
         Caption         =   "Codigo de Desarrollo"
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
         TabIndex        =   66
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label23 
         Caption         =   "Observaciones  Laboratorio"
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
         TabIndex        =   65
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label fghg 
         Caption         =   "Destino"
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
         TabIndex        =   61
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label Label21 
         Caption         =   "Costo Maximo"
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
         Left            =   240
         TabIndex        =   59
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Volumen Estimado de Venta"
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
         Left            =   240
         TabIndex        =   57
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label19 
         Caption         =   "Referencias"
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
         TabIndex        =   55
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label17 
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
         Height          =   375
         Left            =   -74640
         TabIndex        =   54
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "Hoja Seguridad"
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
         TabIndex        =   52
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label15 
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
         Left            =   -74640
         TabIndex        =   50
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label14 
         Caption         =   "Respuesta"
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
         TabIndex        =   45
         Top             =   840
         Width           =   3015
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
         Left            =   240
         TabIndex        =   37
         Top             =   3540
         Width           =   2895
      End
      Begin VB.Label Label13 
         Caption         =   "Hoja Tecnica"
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
         TabIndex        =   32
         Top             =   3480
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
         TabIndex        =   29
         Top             =   2280
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
         TabIndex        =   26
         Top             =   1440
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
         TabIndex        =   23
         Top             =   600
         Width           =   2895
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
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   2895
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
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1815
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
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   3015
      End
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
      TabIndex        =   7
      Text            =   " "
      Top             =   1320
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
      TabIndex        =   4
      Text            =   " "
      Top             =   600
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4320
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6120
      TabIndex        =   39
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label DesVendedor 
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
      TabIndex        =   44
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Vendedor"
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
      TabIndex        =   43
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   6240
      MouseIcon       =   "pedidoordentrabajo.frx":00E5
      MousePointer    =   99  'Custom
      Picture         =   "pedidoordentrabajo.frx":03EF
      ToolTipText     =   "Consulta de Datos"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6960
      MouseIcon       =   "pedidoordentrabajo.frx":0C31
      MousePointer    =   99  'Custom
      Picture         =   "pedidoordentrabajo.frx":0F3B
      ToolTipText     =   "Salida"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   4680
      MouseIcon       =   "pedidoordentrabajo.frx":177D
      MousePointer    =   99  'Custom
      Picture         =   "pedidoordentrabajo.frx":1A87
      ToolTipText     =   "Elimina el Registro"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "pedidoordentrabajo.frx":22C9
      MousePointer    =   99  'Custom
      Picture         =   "pedidoordentrabajo.frx":25D3
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "pedidoordentrabajo.frx":2E15
      MousePointer    =   99  'Custom
      Picture         =   "pedidoordentrabajo.frx":311F
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label10 
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
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1320
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   600
      Width           =   4695
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
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Pedido Nro."
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
Attribute VB_Name = "PrgpedidoOrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPedidoOrdenTrabajo As Recordset
Dim spPedidoOrdenTrabajo As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim XParam As String
Dim EmpresaActual As String

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PedidoOrdenTrabajo"
    ZSql = ZSql + " Where PedidoOrdenTrabajo.Pedido = " + "'" + Pedido.Text + "'"
    spPedidoOrdenTrabajo = ZSql
    Set rsPedidoOrdenTrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rsPedidoOrdenTrabajo.RecordCount > 0 Then
        Fecha.Text = rsPedidoOrdenTrabajo!Fecha
        Cliente.Text = rsPedidoOrdenTrabajo!Cliente
        Observaciones.Text = Trim(rsPedidoOrdenTrabajo!Observaciones)
        vendedor.Text = rsPedidoOrdenTrabajo!vendedor
        Material.Text = Trim(rsPedidoOrdenTrabajo!Material)
        Muestra.Text = Trim(rsPedidoOrdenTrabajo!Muestra)
        Uso.Text = Trim(rsPedidoOrdenTrabajo!Uso)
        DescripcionI.Text = Trim(rsPedidoOrdenTrabajo!DescripcionI)
        DescripcionII.Text = Trim(rsPedidoOrdenTrabajo!DescripcionII)
        DescripcionIII.Text = Trim(rsPedidoOrdenTrabajo!DescripcionIII)
        DescripcionIV.Text = Trim(rsPedidoOrdenTrabajo!DescripcionIV)
        DescripcionV.Text = Trim(rsPedidoOrdenTrabajo!DescripcionV)
        ObservacionesI.Text = Trim(rsPedidoOrdenTrabajo!ObservacionesI)
        ObservacionesII.Text = Trim(rsPedidoOrdenTrabajo!ObservacionesII)
        ObservacionesIII.Text = Trim(rsPedidoOrdenTrabajo!ObservacionesIII)
        Volumen.Text = Str$(rsPedidoOrdenTrabajo!Volumen)
        Costo.Text = Str$(rsPedidoOrdenTrabajo!Costo)
        RequisitoI.Text = Trim(rsPedidoOrdenTrabajo!RequisitoI)
        RequisitoII.Text = Trim(rsPedidoOrdenTrabajo!RequisitoII)
        RequisitoIII.Text = Trim(rsPedidoOrdenTrabajo!RequisitoIII)
        RequisitoIV.Text = Trim(rsPedidoOrdenTrabajo!RequisitoIV)
        RequisitoV.Text = Trim(rsPedidoOrdenTrabajo!RequisitoV)
        RequisitoVI.Text = Trim(rsPedidoOrdenTrabajo!RequisitoVI)
        ReferenciaI.Text = Trim(rsPedidoOrdenTrabajo!ReferenciaI)
        ReferenciaII.Text = Trim(rsPedidoOrdenTrabajo!ReferenciaII)
        ReferenciaIII.Text = Trim(rsPedidoOrdenTrabajo!ReferenciaIII)
        Destino.ListIndex = rsPedidoOrdenTrabajo!Destino
        CodigoOrden.Text = rsPedidoOrdenTrabajo!CodigoOrden
        Respuesta.ListIndex = rsPedidoOrdenTrabajo!Respuesta
        RespuestaI.Text = Trim(rsPedidoOrdenTrabajo!RespuestaI)
        RespuestaII.Text = Trim(rsPedidoOrdenTrabajo!RespuestaII)
        RespuestaIII.Text = Trim(rsPedidoOrdenTrabajo!RespuestaIII)
        RespuestaIV.Text = Trim(rsPedidoOrdenTrabajo!RespuestaVI)
        RespuestaV.Text = Trim(rsPedidoOrdenTrabajo!RespuestaV)
        RespuestaVI.Text = Trim(rsPedidoOrdenTrabajo!RespuestaVI)
        Destino.ListIndex = rsPedidoOrdenTrabajo!Destino
        CodigoOrden.Text = rsPedidoOrdenTrabajo!CodigoOrden
        rsPedidoOrdenTrabajo.Close
    End If
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
            Else
        DesCliente.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Vendedor"
    ZSql = ZSql + " Where Vendedor.Vendedor = " + "'" + vendedor.Text + "'"
    spVendedor = ZSql
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstVendedor.RecordCount > 0 Then
        DesVendedor.Caption = rstVendedor!Nombre
        rstVendedor.Close
    End If
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdAdd_Click()

    If Respuesta.ListIndex < 0 Then
        Respuesta.ListIndex = 0
    End If

    If Val(Pedido.Text) = 0 Then
    
        Sql1 = "Select Max(Pedido) as [PedidoMayor]"
        Sql2 = " FROM PedidoOrdenTrabajo"
        spPedidoOrdenTrabajo = Sql1 + Sql2
        Set rstPedidoordentrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoordentrabajo.RecordCount > 0 Then
            rstPedidoordentrabajo.MoveLast
            WPedidoMayor = IIf(IsNull(rstPedidoordentrabajo!PedidoMayor), "0", rstPedidoordentrabajo!PedidoMayor)
            WPedido = Mid$(Str$(WPedidoMayor + 1), 2, 8)
            rstPedidoordentrabajo.Close
                Else
            WPedido = "1"
        End If
        
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WEstadoLabora = ""
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO PedidoOrdenTrabajo ("
        ZSql = ZSql + "Pedido ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Vendedor ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Material ,"
        ZSql = ZSql + "Muestra ,"
        ZSql = ZSql + "Uso ,"
        ZSql = ZSql + "DescripcionI ,"
        ZSql = ZSql + "DescripcionII ,"
        ZSql = ZSql + "DescripcionIII ,"
        ZSql = ZSql + "DescripcionIV ,"
        ZSql = ZSql + "DescripcionV ,"
        ZSql = ZSql + "Respuesta ,"
        ZSql = ZSql + "RespuestaI ,"
        ZSql = ZSql + "RespuestaII ,"
        ZSql = ZSql + "RespuestaIII ,"
        ZSql = ZSql + "RespuestaIV ,"
        ZSql = ZSql + "RespuestaV ,"
        ZSql = ZSql + "RespuestaVI ,"
        ZSql = ZSql + "ObservacionesI ,"
        ZSql = ZSql + "ObservacionesII ,"
        ZSql = ZSql + "ObservacionesIII ,"
        ZSql = ZSql + "Volumen ,"
        ZSql = ZSql + "Costo ,"
        ZSql = ZSql + "Destino ,"
        ZSql = ZSql + "CodigoOrden ,"
        ZSql = ZSql + "EstadoLabora ,"
        ZSql = ZSql + "RequisitoI ,"
        ZSql = ZSql + "RequisitoII ,"
        ZSql = ZSql + "RequisitoIII ,"
        ZSql = ZSql + "RequisitoIV ,"
        ZSql = ZSql + "RequisitoV ,"
        ZSql = ZSql + "RequisitoVI ,"
        ZSql = ZSql + "ReferenciaI ,"
        ZSql = ZSql + "ReferenciaII ,"
        ZSql = ZSql + "ReferenciaIII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WPedido + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + WOrdFecha + "',"
        ZSql = ZSql + "'" + Cliente.Text + "',"
        ZSql = ZSql + "'" + vendedor.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + Material.Text + "',"
        ZSql = ZSql + "'" + Muestra.Text + "',"
        ZSql = ZSql + "'" + Uso.Text + "',"
        ZSql = ZSql + "'" + DescripcionI.Text + "',"
        ZSql = ZSql + "'" + DescripcionII.Text + "',"
        ZSql = ZSql + "'" + DescripcionIII.Text + "',"
        ZSql = ZSql + "'" + DescripcionIV.Text + "',"
        ZSql = ZSql + "'" + DescripcionV.Text + "',"
        ZSql = ZSql + "'" + Str$(Respuesta.ListIndex) + "',"
        ZSql = ZSql + "'" + RespuestaI.Text + "',"
        ZSql = ZSql + "'" + RespuestaII.Text + "',"
        ZSql = ZSql + "'" + RespuestaIII.Text + "',"
        ZSql = ZSql + "'" + RespuestaIV.Text + "',"
        ZSql = ZSql + "'" + RespuestaV.Text + "',"
        ZSql = ZSql + "'" + RespuestaVI.Text + "',"
        ZSql = ZSql + "'" + ObservacionesI.Text + "',"
        ZSql = ZSql + "'" + ObservacionesII.Text + "',"
        ZSql = ZSql + "'" + ObservacionesIII.Text + "',"
        ZSql = ZSql + "'" + Volumen.Text + "',"
        ZSql = ZSql + "'" + Costo.Text + "',"
        ZSql = ZSql + "'" + Str$(Destino.ListIndex) + "',"
        ZSql = ZSql + "'" + CodigoOrden.Text + "',"
        ZSql = ZSql + "'" + WEstadoLabora + "',"
        ZSql = ZSql + "'" + RequisitoI.Text + "',"
        ZSql = ZSql + "'" + RequisitoII.Text + "',"
        ZSql = ZSql + "'" + RequisitoIII.Text + "',"
        ZSql = ZSql + "'" + RequisitoIV.Text + "',"
        ZSql = ZSql + "'" + RequisitoV.Text + "',"
        ZSql = ZSql + "'" + RequisitoVI.Text + "',"
        ZSql = ZSql + "'" + ReferenciaI.Text + "',"
        ZSql = ZSql + "'" + ReferenciaII.Text + "',"
        ZSql = ZSql + "'" + ReferenciaIII.Text + "')"
        spPedidoOrdenTrabajo = ZSql
        Set rsPedidoOrdenTrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        
        m$ = "Se a asignado el numero de pedido " + WPedido
        a% = MsgBox(m$, 0, "Pedidos de Desarrollos")
        
            Else
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WEstadoLabora = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoOrdenTrabajo"
        ZSql = ZSql + " Where PedidoOrdenTrabajo.Pedido = " + "'" + Pedido.Text + "'"
        spPedidoOrdenTrabajo = ZSql
        Set rsPedidoOrdenTrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        If rsPedidoOrdenTrabajo.RecordCount > 0 Then
            rsPedidoOrdenTrabajo.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE PedidoOrdenTrabajo SET "
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
            ZSql = ZSql + " Cliente = " + "'" + Cliente.Text + "',"
            ZSql = ZSql + " Vendedor = " + "'" + vendedor.Text + "',"
            ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " Material = " + "'" + Material.Text + "',"
            ZSql = ZSql + " Muestra = " + "'" + Muestra.Text + "',"
            ZSql = ZSql + " Uso = " + "'" + Uso.Text + "',"
            ZSql = ZSql + " DescripcionI = " + "'" + DescripcionI.Text + "',"
            ZSql = ZSql + " DescripcionII = " + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + " DescripcionIII = " + "'" + DescripcionIII.Text + "',"
            ZSql = ZSql + " DescripcionIV = " + "'" + DescripcionIV.Text + "',"
            ZSql = ZSql + " DescripcionV = " + "'" + DescripcionV.Text + "',"
            ZSql = ZSql + " Respuesta = " + "'" + Str$(Respuesta.ListIndex) + "',"
            ZSql = ZSql + " RespuestaI = " + "'" + RespuestaI.Text + "',"
            ZSql = ZSql + " RespuestaII = " + "'" + RespuestaII.Text + "',"
            ZSql = ZSql + " RespuestaIII = " + "'" + RespuestaIII.Text + "',"
            ZSql = ZSql + " RespuestaIV = " + "'" + RespuestaIV.Text + "',"
            ZSql = ZSql + " RespuestaV = " + "'" + RespuestaV.Text + "',"
            ZSql = ZSql + " RespuestaVI = " + "'" + RespuestaVI.Text + "',"
            ZSql = ZSql + " ObservacionesI = " + "'" + ObservacionesI.Text + "',"
            ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesII.Text + "',"
            ZSql = ZSql + " ObservacionesIII = " + "'" + ObservacionesIII.Text + "',"
            ZSql = ZSql + " Volumen = " + "'" + Volumen.Text + "',"
            ZSql = ZSql + " Costo = " + "'" + Costo.Text + "',"
            ZSql = ZSql + " Destino = " + "'" + Str$(Destino.ListIndex) + "',"
            ZSql = ZSql + " CodigoOrden = " + "'" + CodigoOrden.Text + "',"
            ZSql = ZSql + " EstadoLabora = " + "'" + WEstadoLabora + "',"
            ZSql = ZSql + " RequisitoI = " + "'" + RequisitoI.Text + "',"
            ZSql = ZSql + " RequisitoII = " + "'" + RequisitoII.Text + "',"
            ZSql = ZSql + " RequisitoIII = " + "'" + RequisitoIII.Text + "',"
            ZSql = ZSql + " RequisitoIV = " + "'" + RequisitoIV.Text + "',"
            ZSql = ZSql + " RequisitoV = " + "'" + RequisitoV.Text + "',"
            ZSql = ZSql + " RequisitoVI = " + "'" + RequisitoVI.Text + "',"
            ZSql = ZSql + " ReferenciaI = " + "'" + ReferenciaI.Text + "',"
            ZSql = ZSql + " ReferenciaII = " + "'" + ReferenciaII.Text + "',"
            ZSql = ZSql + " ReferenciaIII = " + "'" + ReferenciaIII.Text + "'"
            ZSql = ZSql + " Where Pedido = " + "'" + Pedido.Text + "'"
            spPedidoOrdenTrabajo = ZSql
            Set rsPedidoOrdenTrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    End If
    
    Call CmdLimpiar_Click
    Pedido.SetFocus
    
End Sub

Private Sub cmdDelete_Click()

    If Pedido.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoOrdenTrabajo"
        ZSql = ZSql + " Where PedidoOrdenTrabajo.Pedido = " + "'" + Pedido.Text + "'"
        spPedidoOrdenTrabajo = ZSql
        Set rsPedidoOrdenTrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        If rsPedidoOrdenTrabajo.RecordCount > 0 Then
            rsPedidoOrdenTrabajo.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            zRespuesta% = MsgBox(m$, 32 + 4, T$)
            If zRespuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE PedidoOrdenTrabajo"
                ZSql = ZSql + " Where Pedido = " + "'" + Pedido.Text + "'"
                spPedidoOrdenTrabajo = ZSql
                Set rsPedidoOrdenTrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    Pedido.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Pedido.Text = ""
    Fecha.Text = "  /  /    "
    Cliente.Text = ""
    vendedor.Text = ""
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
    Volumen.Text = ""
    Costo.Text = ""
    RequisitoI.Text = ""
    RequisitoII.Text = ""
    RequisitoIII.Text = ""
    RequisitoIV.Text = ""
    RequisitoV.Text = ""
    RequisitoVI.Text = ""
    ReferenciaI.Text = ""
    ReferenciaII.Text = ""
    ReferenciaIII.Text = ""
    
    RespuestaI.Text = ""
    RespuestaII.Text = ""
    RespuestaIII.Text = ""
    RespuestaIV.Text = ""
    RespuestaV.Text = ""
    RespuestaVI.Text = ""
    
    CodigoOrden.Text = "  -     "
    
    Respuesta.ListIndex = 0
    Destino.ListIndex = 0
    
    DesCliente.Caption = ""
    DesVendedor.Caption = ""
    Tablas.Tab = 0

    Pedido.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgpedidoOrdenTrabajo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cliente.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
                vendedor.SetFocus
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
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If vendedor.Text <> "" Then
            spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                vendedor.Text = rstVendedor!vendedor
                DesVendedor.Caption = rstVendedor!Nombre
                rstVendedor.Close
                Observaciones.SetFocus
                    Else
                vendedor.SetFocus
            End If
                Else
            DesVendedor.Caption = ""
            Observaciones.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        vendedor.Text = ""
        DesVendedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        Volumen.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesIII.Text = ""
    End If
End Sub

Private Sub Volumen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo.SetFocus
    End If
    If KeyAscii = 27 Then
        Volumen.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Costo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Material.SetFocus
    End If
    If KeyAscii = 27 Then
        Costo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        RequisitoVI.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoV.Text = ""
    End If
End Sub

Private Sub RequisitoVI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ReferenciaI.SetFocus
    End If
    If KeyAscii = 27 Then
        RequisitoVI.Text = ""
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
        ReferenciaIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ReferenciaII.Text = ""
    End If
End Sub

Private Sub ReferenciaIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RequisitoI.SetFocus
    End If
    If KeyAscii = 27 Then
        ReferenciaIII.Text = ""
    End If
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Pedido.Text <> "" Then
        
            Pedido.Text = UCase(Pedido.Text)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTrabajo"
            ZSql = ZSql + " Where PedidoOrdenTrabajo.Pedido = " + "'" + Pedido.Text + "'"
            spPedidoOrdenTrabajo = ZSql
            Set rsPedidoOrdenTrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
            If rsPedidoOrdenTrabajo.RecordCount > 0 Then
                rsPedidoOrdenTrabajo.Close
                Call Imprime_Datos
                    Else
                WPedido = Pedido.Text
                CmdLimpiar_Click
                Pedido.Text = WPedido
            End If
            
        End If
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Pedido.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Respuesta.Clear
    
    Respuesta.AddItem ""
    Respuesta.AddItem "Aceptada"
    Respuesta.AddItem "Rechazada"
    
    Destino.Clear
    
    Destino.AddItem ""
    Destino.AddItem "Desarrollo"
    Destino.AddItem "Laboratorio"


    Pedido.Text = ""
    Fecha.Text = "  /  /    "
    Cliente.Text = ""
    vendedor.Text = ""
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
    Volumen.Text = ""
    Costo.Text = ""
    RequisitoI.Text = ""
    RequisitoII.Text = ""
    RequisitoIII.Text = ""
    RequisitoIV.Text = ""
    RequisitoV.Text = ""
    RequisitoVI.Text = ""
    ReferenciaI.Text = ""
    ReferenciaII.Text = ""
    ReferenciaIII.Text = ""
    
    
    DesCliente.Caption = ""
    DesVendedor.Caption = ""
    
    Respuesta.ListIndex = 0
    Destino.ListIndex = 0
    
    RespuestaI.Text = ""
    RespuestaII.Text = ""
    RespuestaIII.Text = ""
    RespuestaIV.Text = ""
    RespuestaV.Text = ""
    RespuestaVI.Text = ""
    
    CodigoOrden.Text = "  -     "
    
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
     Opcion.AddItem "Vendedor"
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
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Order by Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
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
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Order by Vendedor"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstVendedor!vendedor) + " " + rstVendedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstVendedor!vendedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
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
            
        Case 1
            Indice = Pantalla.ListIndex
            vendedor.Text = WIndice.List(Indice)
            Call Vendedor_KeyPress(13)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        XIndice = Opcion.ListIndex
    
        Select Case XIndice
            Case 0
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cliente"
                ZSql = ZSql + " Order by Cliente"
                spCliente = ZSql
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
    
                    With rstCliente
                        .MoveFirst
                        Do
                            If .EOF = False Then
            
                                DA = Len(rstCliente!Razon) - WEspacios
                
                                For Aaa = 1 To DA + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstCliente!Razon, Aaa, WEspacios) Then
                                        IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
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
                
            Case 1
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Vendedor"
                ZSql = ZSql + " Order by Vendedor"
                spVendedor = ZSql
                Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstVendedor.RecordCount > 0 Then
    
                    With rstVendedor
                        .MoveFirst
                        Do
                            If .EOF = False Then
            
                                DA = Len(rstVendedor!Nombre) - WEspacios
                
                                For Aaa = 1 To DA + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstVendedor!Nombre, Aaa, WEspacios) Then
                                        IngresaItem = Str$(rstVendedor!vendedor) + " " + rstVendedor!Nombre
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstVendedor!vendedor
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
    
                    rstVendedor.Close
    
                End If
                
                
            Case Else
        End Select
    
    End If

End Sub



