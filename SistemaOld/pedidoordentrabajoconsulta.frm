VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgpedidoOrdenTrabajoConsulta 
   Caption         =   "Ingreso de Pedido de Desarrollo"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11625
   Begin VB.Frame PantaConsultaEnsayo 
      Height          =   855
      Left            =   8760
      TabIndex        =   65
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      Begin VB.TextBox WTexto14 
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
         Left            =   2040
         TabIndex        =   68
         Top             =   1080
         Width           =   375
      End
      Begin VB.ComboBox WCombo14 
         Height          =   315
         Left            =   1440
         TabIndex        =   67
         Top             =   1680
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto24 
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
         Left            =   1440
         TabIndex        =   66
         Top             =   1080
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto34 
         Height          =   285
         Left            =   2640
         TabIndex        =   69
         Top             =   1080
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
      Begin MSFlexGridLib.MSFlexGrid WVector4 
         Height          =   5535
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   9763
         _Version        =   327680
         BackColor       =   16777152
      End
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
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   40
      Text            =   " "
      Top             =   960
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   10440
      TabIndex        =   39
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
      TextRTF         =   $"pedidoordentrabajoconsulta.frx":0000
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
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   9763
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion del Pedido"
      TabPicture(0)   =   "pedidoordentrabajoconsulta.frx":007C
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
      TabPicture(1)   =   "pedidoordentrabajoconsulta.frx":0098
      Tab(1).ControlCount=   16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ReferenciaIII"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "ReferenciaII"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "ReferenciaI"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "RequisitoVI"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "RequisitoV"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "RequisitoIV"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "RequisitoIII"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "RequisitoII"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "RequisitoI"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Label19"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label17"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label16"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label13"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label11"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label9"
      Tab(1).Control(15).Enabled=   0   'False
      TabCaption(2)   =   "Respuesta"
      TabPicture(2)   =   "pedidoordentrabajoconsulta.frx":00B4
      Tab(2).ControlCount=   14
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CodigoOrden"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "RespuestaIV"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "RespuestaVI"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "RespuestaV"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "Destino"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "RespuestaII"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "RespuestaIII"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "RespuestaI"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "Respuesta"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Label24"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label23"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "fghg"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label15"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label14"
      Tab(2).Control(13).Enabled=   0   'False
      Begin VB.TextBox CodigoOrden 
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
         MaxLength       =   10
         TabIndex        =   64
         Text            =   " "
         Top             =   3480
         Width           =   1095
      End
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   57
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
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   55
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
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   53
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
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
         MaxLength       =   50
         TabIndex        =   47
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
         MaxLength       =   50
         TabIndex        =   46
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
         MaxLength       =   50
         TabIndex        =   45
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
         TabIndex        =   44
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Text            =   " "
         Top             =   600
         Width           =   5895
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   58
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
         Height          =   255
         Left            =   240
         TabIndex        =   56
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
         TabIndex        =   54
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   43
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6120
      TabIndex        =   38
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
      TabIndex        =   42
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
      TabIndex        =   41
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6120
      MouseIcon       =   "pedidoordentrabajoconsulta.frx":00D0
      MousePointer    =   99  'Custom
      Picture         =   "pedidoordentrabajoconsulta.frx":03DA
      ToolTipText     =   "Salida"
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
Attribute VB_Name = "PrgpedidoOrdenTrabajoConsulta"
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

Rem para el vector IV

Dim WBorraIV(1000, 20) As String
Dim WParametrosIV(10, 20) As Double
Dim WFormatoIV(20) As String
Dim WControlIV As String

Private Sub cmdClose_Click()
    PrgpedidoOrdenTrabajoConsulta.Hide
    Unload Me
    PrgConsultaDesarrollo.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub CodigoOrden_DblClick()

    PantaConsultaEnsayo.Visible = True
    PantaConsultaEnsayo.Height = 6500
    PantaConsultaEnsayo.Left = 100
    PantaConsultaEnsayo.Top = 1400
    PantaConsultaEnsayo.Width = 11300
    
    Call Limpia_VectorIV
    WRenglon = 0
    
    XEmpresa = WEmpresa
        
    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaEnsayoV"
    ZSql = ZSql + " Where CargaEnsayoV.Orden = " + "'" + CodigoOrden.Text + "'"
    ZSql = ZSql + " Order by CargaEnsayoV.Clave"
    spCargaEnsayoV = ZSql
    Set rstCargaEnsayoV = db.OpenRecordset(spCargaEnsayoV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaEnsayoV.RecordCount > 0 Then
        With rstCargaEnsayoV
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector4.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector4.Col = 1
                    WVector4.Text = Trim(rstCargaEnsayoV!Version)
                    
                    WVector4.Col = 2
                    WVector4.Text = Trim(rstCargaEnsayoV!Etapa)
            
                    WVector4.Col = 3
                    WVector4.Text = Trim(rstCargaEnsayoV!Fecha)
            
                    WVector4.Col = 4
                    WVector4.Text = Trim(rstCargaEnsayoV!Participantes)
                    
                    WVector4.Col = 5
                    WVector4.Text = Trim(rstCargaEnsayoV!Resultados)
                    
                    WVector4.Col = 6
                    WVector4.Text = Trim(rstCargaEnsayoV!Acciones)
                    
                    WVector4.Col = 7
                    WVector4.Text = Trim(rstCargaEnsayoV!Responsables)
            
                    WVector4.Col = 8
                    WVector4.Text = Trim(rstCargaEnsayoV!Estado)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaEnsayoV.Close
    End If
    
    Call Conecta_Empresa
    
End Sub

Private Sub WVector4_DblClick()
    PantaConsultaEnsayo.Visible = False
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
    
    Pedido.Text = WXPed
    
    XEmpresa = WEmpresa
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
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
        Respuesta.ListIndex = rsPedidoOrdenTrabajo!Respuesta
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
    
    Call Conecta_Empresa
    
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



Rem
Rem Controles de la WVector4
Rem

Private Sub GridEditTextIV(ByVal KeyAscii As Integer)

    XColumna = WVector4.Col
    XTipoDato = WParametrosIV(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto14.Left = WVector4.CellLeft + WVector4.Left
            WTexto14.Top = WVector4.CellTop + WVector4.Top
            WTexto14.Width = WVector4.CellWidth
            WTexto14.Height = WVector4.CellHeight
            WTexto14.MaxLength = WParametrosIV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto14.Text = WVector4.Text
                    WTexto14.SelStart = Len(WTexto14.Text)
                Case Else
                    WTexto14.Text = Chr$(KeyAscii)
                    WTexto14.SelStart = 1
            End Select
            WTexto14.Visible = True
            WTexto14.SetFocus
        Case 1
            WTexto24.Left = WVector4.CellLeft + WVector4.Left
            WTexto24.Top = WVector4.CellTop + WVector4.Top
            WTexto24.Width = WVector4.CellWidth
            WTexto24.Height = WVector4.CellHeight
            WTexto24.MaxLength = WParametrosIV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto24.Text = WVector4.Text
                    Rem WTexto24.SelStart = Len(WTexto24.Text)
                    WTexto24.SelStart = 0
                Case Else
                    WTexto24.Text = Chr$(KeyAscii)
                    WTexto24.SelStart = 1
            End Select
            WTexto24.Visible = True
            WTexto24.SetFocus
        Case 2
            WTexto34.Left = WVector4.CellLeft + WVector4.Left
            WTexto34.Top = WVector4.CellTop + WVector4.Top
            WTexto34.Width = WVector4.CellWidth
            WTexto34.Height = WVector4.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector4.Text) = 10 Then
                        WTexto34.Text = WVector4.Text
                            Else
                        WTexto34.Text = "  /  /    "
                    End If
                    WTexto34.SelStart = 0
                Case Else
                    WTexto34.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto34.SelStart = 1
            End Select
            WTexto34.Visible = True
            WTexto34.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditIV()
    Pasa = 0
    If WCombo14.Visible Then
        Pasa = 0
        WVector4.Text = WCombo14.Text
        WCombo14.Visible = False
            Else
        If WTexto14.Visible Then
            Pasa = 1
            WVector4.Text = WTexto14.Text
            WTexto14.Visible = False
                Else
            If WTexto24.Visible Then
                Pasa = 1
                WVector4.Text = WTexto24.Text
                WTexto24.Visible = False
                    Else
                If WTexto34.Visible Then
                    Pasa = 1
                    WVector4.Text = WTexto34.Text
                    WTexto34.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoIV(WVector4.Col) <> "" Then
            WVector4.Text = Pusing(WFormatoIV(WVector4.Col), WVector4.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboIV()
    ' Position the ComboBox over the cell.
    WCombo14.Left = WVector4.CellLeft + WVector4.Left
    WCombo14.Top = WVector4.CellTop + WVector4.Top
    WCombo14.Width = WVector4.CellWidth
    WCombo14.Visible = True
    WCombo14.SetFocus
End Sub

Private Sub WTexto14_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto14.Text = ""
            
        Rem F1
        Case 113
            WTexto14.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            DoEvents
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV

        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 123
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Col > 1 Then
                WVector4.Col = WVector4.Col - 1
            End If
            Call StartEditIV

    End Select
End Sub

Private Sub WTexto24_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto24.Text = ""
            
        Rem F1
        Case 113
            WTexto24.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            DoEvents
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                Rem End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                Rem End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV

    End Select
End Sub

Private Sub WTexto34_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto34.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto34.Text = WVector4.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector4.SetFocus
            Call Control_CampoIV
            If WControlIV = "S" Then
                Call Control_WVectorIV
            End If
            Call StartEditIV

        Case vbKeyDown
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row < WVector4.Rows - 1 Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row + 1
                End If
            End If
            Call StartEditIV

        Case vbKeyUp
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.Row > WVector4.FixedRows Then
                Call Control_CampoIV
                If WControlIV = "S" Then
                    WVector4.Row = WVector4.Row - 1
                End If
            End If
            Call StartEditIV
        Case 34
            ' Move down 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow < WVector4.Rows - 12 Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow + 12
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV
            
        Case 33
            ' Move up 1 row.
            WVector4.SetFocus
            DoEvents
            If WVector4.TopRow - 12 > WVector4.FixedRows Then
                Rem Call Control_CampoIV
                Rem If WControlIV = "S" Then
                    WVector4.TopRow = WVector4.TopRow - 12
                    WVector4.Row = WVector4.TopRow
                        Else
                    WVector4.TopRow = 1
                    WVector4.Row = WVector4.TopRow
                Rem End If
            End If
            Call StartEditIV

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto14_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto24_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto34_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo14_Click()
    WVector4.SetFocus
End Sub


Private Sub WVector4_Click()
    StartEditIV
End Sub

Private Sub WVector4_LeaveCell()
    EndEditIV
End Sub

Private Sub WVector4_GotFocus()
    EndEditIV
End Sub

Private Sub WVector4_KeyPress(KeyAscii As Integer)
    XColumna = WVector4.Col
    Select Case WParametrosIV(4, WVector4.Col)
        Case 1
        Case Else
            If WParametrosIV(2, XColumna) = 0 Then
                GridEditTextIV KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditIV()
    Select Case WParametrosIV(4, WVector4.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo14.Clear
            WCombo14.AddItem "Campo1"
            WCombo14.AddItem "Campo2"
            On Error Resume Next
            WCombo14.Text = WVector4.Text
            On Error GoTo 0
            GridEditComboIV
        Case Else
            If WParametrosIV(2, WVector4.Col) = 0 Then
                GridEditTextIV Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVectorIV()
    Select Case WVector4.Col
        Case 8
            If WVector4.Row < WVector4.Rows - 1 Then
                WVector4.Row = WVector4.Row + 1
            End If
            WVector4.Col = 1
        Case Else
            If WVector4.Col < WVector4.Cols - 1 Then
                WVector4.Col = WVector4.Col + 1
            End If
    End Select
    WVector4.SetFocus
    GridEditTextIV KeyAscii
End Sub

Private Sub Control_CampoIV()
    XColumna = WVector4.Col
    XFila = WVector4.Row
    WControlIV = "S"
End Sub




Private Sub AgregaRenglonIV_Click()

    Hasta = WVector4.Row

    For iRow = 1000 To Hasta Step -1
        WVector4.TextMatrix(iRow, 0) = WVector4.TextMatrix(iRow - 1, 0)
        WVector4.TextMatrix(iRow, 1) = WVector4.TextMatrix(iRow - 1, 1)
        WVector4.TextMatrix(iRow, 2) = WVector4.TextMatrix(iRow - 1, 2)
        WVector4.TextMatrix(iRow, 3) = WVector4.TextMatrix(iRow - 1, 3)
        WVector4.TextMatrix(iRow, 4) = WVector4.TextMatrix(iRow - 1, 4)
        WVector4.TextMatrix(iRow, 5) = WVector4.TextMatrix(iRow - 1, 5)
        WVector4.TextMatrix(iRow, 6) = WVector4.TextMatrix(iRow - 1, 6)
        WVector4.TextMatrix(iRow, 7) = WVector4.TextMatrix(iRow - 1, 7)
        WVector4.TextMatrix(iRow, 8) = WVector4.TextMatrix(iRow - 1, 8)
    Next iRow

    WVector4.TextMatrix(Hasta, 0) = ""
    WVector4.TextMatrix(Hasta, 1) = ""
    WVector4.TextMatrix(Hasta, 2) = ""
    WVector4.TextMatrix(Hasta, 3) = ""
    WVector4.TextMatrix(Hasta, 4) = ""
    WVector4.TextMatrix(Hasta, 5) = ""
    WVector4.TextMatrix(Hasta, 6) = ""
    WVector4.TextMatrix(Hasta, 7) = ""
    WVector4.TextMatrix(Hasta, 8) = ""
    
    WTexto14.Text = ""
    WTexto24.Text = ""

End Sub




Private Sub Limpia_VectorIV()

    WVector4.Clear

    Rem ponga la WVector4 en negritas
    WVector4.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto14.FontName = WVector4.FontName
    WTexto14.FontSize = WVector4.FontSize
    WTexto14.Visible = False
    WTexto24.FontName = WVector4.FontName
    WTexto24.FontSize = WVector4.FontSize
    WTexto24.Visible = False
    WTexto34.FontName = WVector4.FontName
    WTexto34.FontSize = WVector4.FontSize
    WTexto34.Visible = False
    WCombo14.FontName = WVector4.FontName
    WCombo14.FontSize = WVector4.FontSize
    WCombo14.Visible = False

    ' Establesco loa Valores de la WVector4
    
    WVector4.FixedCols = 1
    WVector4.Cols = 9
    WVector4.FixedRows = 1
    WVector4.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector4.Text = "Articulo"
    
    Rem Longitud
    Rem WVector4.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector4.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosIV(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosIV(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosIV(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosIV(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector4.ColWidth(0) = 200
    WVector4.Row = 0
    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector4.Text = "Version"
                WVector4.ColWidth(Ciclo) = 800
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 10
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 2
                WVector4.Text = "Etapa"
                WVector4.ColWidth(Ciclo) = 900
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 3
                WVector4.Text = "Fecha"
                WVector4.ColWidth(Ciclo) = 1000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 10
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 4
                WVector4.Text = "Participantes"
                WVector4.ColWidth(Ciclo) = 2000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 5
                WVector4.Text = "Resultados"
                WVector4.ColWidth(Ciclo) = 4000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 6
                WVector4.Text = "Acciones"
                WVector4.ColWidth(Ciclo) = 4000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 50
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 7
                WVector4.Text = "Responsables"
                WVector4.ColWidth(Ciclo) = 2000
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
            Case 8
                WVector4.Text = "Estado"
                WVector4.ColWidth(Ciclo) = 1200
                WVector4.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIV(1, Ciclo) = 20
                WParametrosIV(2, Ciclo) = 0
                WParametrosIV(3, Ciclo) = 0
                WParametrosIV(4, Ciclo) = 0
                WFormatoIV(Ciclo) = ""
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector4.Row = 0
    For Ciclo = 1 To WVector4.Cols - 1
        WVector4.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector4.Text
        Rem WTitulo(Ciclo).Left = WVector4.CellLeft + WVector4.Left
        Rem WTitulo(Ciclo).Top = WVector4.CellTop + WVector4.Top
        Rem WTitulo(Ciclo).Width = WVector4.CellWidth
        Rem WTitulo(Ciclo).Height = WVector4.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector4
    
    WAncho = 400
    For Ciclo = 0 To WVector4.Cols - 1
        WAncho = WAncho + WVector4.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector4.Width = WAncho

    ' Size the columns.
    Font.Name = WVector4.Font.Name
    Font.Size = WVector4.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector4.AllowUserResizing = flexResizeBoth
    
    WVector4.Col = 1
    WVector4.Row = 1
    
End Sub

Private Sub WVector4_Scroll()
    WTexto14.Visible = False
    WTexto24.Visible = False
    WTexto34.Visible = False
End Sub




