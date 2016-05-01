VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdenTrabajoAnterior 
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
      TabIndex        =   61
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
      TextRTF         =   $"ordentrabajoAnterior.frx":0000
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
      ItemData        =   "ordentrabajoAnterior.frx":007C
      Left            =   480
      List            =   "ordentrabajoAnterior.frx":0083
      TabIndex        =   54
      Top             =   5160
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6120
      TabIndex        =   53
      Top             =   5400
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
      Left            =   2520
      TabIndex        =   52
      Top             =   5280
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
      Left            =   480
      TabIndex        =   51
      Top             =   4800
      Visible         =   0   'False
      Width           =   6855
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5895
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   10398
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion de la Orden"
      TabPicture(0)   =   "ordentrabajoAnterior.frx":0091
      Tab(0).ControlCount=   18
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
      Tab(0).Control(5)=   "Label19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Material"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Muestra"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Uso"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DescripcionI"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DescripcionII"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DescripcionIII"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DescripcionIV"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "DescripcionV"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ObservacionesII"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ObservacionesIII"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "ObservacionesI"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Encargado"
      Tab(0).Control(17).Enabled=   0   'False
      TabCaption(1)   =   "Requisitos"
      TabPicture(1)   =   "ordentrabajoAnterior.frx":00AD
      Tab(1).ControlCount=   12
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "RequisitoI"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "RequisitoII"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "RequisitoIII"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "RequisitoIV"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "RequisitoV"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "RequisitoVI"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "ReferenciaI"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "ReferenciaII"
      Tab(1).Control(11).Enabled=   -1  'True
      TabCaption(2)   =   "Especificaciones"
      TabPicture(2)   =   "ordentrabajoAnterior.frx":00C9
      Tab(2).ControlCount=   18
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label16"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label17"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label22"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "WVector1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "WTexto3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Aplicacion"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "Estabilidad"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "ObservacionesIV"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "ObservacionesV"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "ObservacionesVI"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "WTexto1"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "WCombo1"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "WTexto2"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "NotaAplicacion"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "Fin_Nota_Aplicacion"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "NotaEstabilidad"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "Fin_Nota_Estabilidad"
      Tab(2).Control(17).Enabled=   -1  'True
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
         TabIndex        =   64
         Top             =   960
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
         Left            =   -68520
         TabIndex        =   63
         Top             =   960
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
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   -68520
         TabIndex        =   60
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox WTexto2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   -73920
         TabIndex        =   57
         Top             =   2760
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   -71880
         TabIndex        =   56
         Top             =   2640
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   -73320
         TabIndex        =   55
         Top             =   2760
         Width           =   375
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   49
         Text            =   " "
         Top             =   5280
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   3960
         Width           =   5895
      End
      Begin VB.TextBox ObservacionesVI 
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
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   5280
         Width           =   5895
      End
      Begin VB.TextBox ObservacionesV 
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
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   45
         Text            =   " "
         Top             =   4920
         Width           =   5895
      End
      Begin VB.TextBox ObservacionesIV 
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
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   43
         Text            =   " "
         Top             =   4560
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   41
         Text            =   " "
         Top             =   4680
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   40
         Text            =   " "
         Top             =   4320
         Width           =   5895
      End
      Begin VB.ComboBox Estabilidad 
         Height          =   315
         Left            =   -71040
         TabIndex        =   39
         Top             =   1020
         Width           =   2055
      End
      Begin VB.ComboBox Aplicacion 
         Height          =   315
         Left            =   -71040
         TabIndex        =   38
         Top             =   540
         Width           =   2055
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
         TabIndex        =   35
         Text            =   " "
         Top             =   3420
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
         TabIndex        =   33
         Text            =   " "
         Top             =   3060
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
         TabIndex        =   32
         Text            =   " "
         Top             =   2580
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
         TabIndex        =   30
         Text            =   " "
         Top             =   2220
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
         TabIndex        =   29
         Text            =   " "
         Top             =   1740
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
         TabIndex        =   27
         Text            =   " "
         Top             =   1380
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
         TabIndex        =   26
         Text            =   " "
         Top             =   900
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
         TabIndex        =   24
         Text            =   " "
         Top             =   540
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   23
         Text            =   " "
         Top             =   3300
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   22
         Text            =   " "
         Top             =   2940
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   21
         Text            =   " "
         Top             =   2580
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   20
         Text            =   " "
         Top             =   2220
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   18
         Text            =   " "
         Top             =   1860
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   1260
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   14
         Text            =   " "
         Top             =   900
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   540
         Width           =   5895
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   -72720
         TabIndex        =   58
         Top             =   2760
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
         Height          =   2535
         Left            =   -74880
         TabIndex        =   59
         Top             =   1800
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4471
         _Version        =   327680
         BackColor       =   16777152
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
         Left            =   360
         TabIndex        =   50
         Top             =   5280
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
         Left            =   360
         TabIndex        =   48
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label22 
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
         Left            =   -74760
         TabIndex        =   44
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "ESPECIFICACIONES"
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
         Left            =   -74520
         TabIndex        =   42
         Top             =   1440
         Width           =   6975
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
         Left            =   -72720
         TabIndex        =   37
         Top             =   1020
         Width           =   1575
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
         Left            =   -72720
         TabIndex        =   36
         Top             =   540
         Width           =   1575
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
         TabIndex        =   34
         Top             =   3060
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
         TabIndex        =   31
         Top             =   2220
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
         TabIndex        =   28
         Top             =   1380
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
         TabIndex        =   25
         Top             =   540
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
         Left            =   360
         TabIndex        =   19
         Top             =   1860
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
         Left            =   360
         TabIndex        =   17
         Top             =   1260
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
         Left            =   360
         TabIndex        =   15
         Top             =   900
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
         Left            =   360
         TabIndex        =   13
         Top             =   540
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
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   6240
      MouseIcon       =   "ordentrabajoAnterior.frx":00E5
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajoAnterior.frx":03EF
      ToolTipText     =   "Consulta de Datos"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6960
      MouseIcon       =   "ordentrabajoAnterior.frx":0C31
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajoAnterior.frx":0F3B
      ToolTipText     =   "Salida"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   4680
      MouseIcon       =   "ordentrabajoAnterior.frx":177D
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajoAnterior.frx":1A87
      ToolTipText     =   "Elimina el Registro"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "ordentrabajoAnterior.frx":22C9
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajoAnterior.frx":25D3
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "ordentrabajoAnterior.frx":2E15
      MousePointer    =   99  'Custom
      Picture         =   "ordentrabajoAnterior.frx":311F
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
Attribute VB_Name = "PrgOrdenTrabajoAnterior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrdenTrabajo As Recordset
Dim spOrdenTrabajo As String
Dim rstOrdenTrabajoII As Recordset
Dim spOrdenTrabajoII As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
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
        ObservacionesIV.Text = Trim(rstOrdenTrabajo!ObservacionesIV)
        ObservacionesV.Text = Trim(rstOrdenTrabajo!ObservacionesV)
        ObservacionesVI.Text = Trim(rstOrdenTrabajo!ObservacionesVI)
        Encargado.Text = Trim(rstOrdenTrabajo!Encargado)
        rstOrdenTrabajo.Close
    End If
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
            Else
        DesCliente.Caption = ""
    End If
    
    Call Limpia_Vector
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM OrdenTrabajoII"
    ZSql = ZSql + " Where OrdenTrabajoII.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " Order by OrdenTrabajoII.Clave"
    spOrdenTrabajoII = ZSql
    Set rstOrdenTrabajoII = db.OpenRecordset(spOrdenTrabajoII, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrdenTrabajoII.RecordCount > 0 Then
        With rstOrdenTrabajoII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstOrdenTrabajoII!Ensayo)
                    If Val(WVector1.Text) = 0 Then
                        WVector1.Text = ""
                    End If
                    
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstOrdenTrabajoII!Descripcion)
            
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstOrdenTrabajoII!Resultado)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrdenTrabajoII.Close
    End If
    
    For Ciclo = 1 To WRenglon
        If Val(WVector1.TextMatrix(Ciclo, 1)) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Ensayos"
            Sql3 = " Where Ensayos.Codigo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
            spEnsayo = Sql1 + Sql2 + Sql3
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                WVector1.TextMatrix(Ciclo, 2) = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        End If
    Next Ciclo
    
End Sub



Private Sub cmdAdd_Click()
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
            ZSql = ZSql + " Estabilidad = " + "'" + Str$(Estabilidad.ListIndex) + "',"
            ZSql = ZSql + " ObservacionesIV = " + "'" + ObservacionesIV.Text + "',"
            ZSql = ZSql + " ObservacionesV = " + "'" + ObservacionesV.Text + "',"
            ZSql = ZSql + " ObservacionesVI = " + "'" + ObservacionesVI.Text + "'"
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
            ZSql = ZSql + "Estabilidad ,"
            ZSql = ZSql + "ObservacionesIV ,"
            ZSql = ZSql + "ObservacionesV ,"
            ZSql = ZSql + "ObservacionesVI )"
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
            ZSql = ZSql + "'" + Str$(Estabilidad.ListIndex) + "',"
            ZSql = ZSql + "'" + ObservacionesIV.Text + "',"
            ZSql = ZSql + "'" + ObservacionesV.Text + "',"
            ZSql = ZSql + "'" + ObservacionesVI.Text + "')"
            spOrdenTrabajo = ZSql
            Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Sql1 = "DELETE OrdenTrabajoII"
        Sql2 = " Where Orden = " + "'" + Orden.Text + "'"
        spOrdenTrabajoII = Sql1 + Sql2
        Set rstOrdenTrabajoII = db.OpenRecordset(spOrdenTrabajoII, dbOpenSnapshot, dbSQLPassThrough)
        
        
        WRenglon = 0
        For iRow = 1 To 99
    
            ZEnsayo = WVector1.TextMatrix(iRow, 1)
            ZDescripcion = WVector1.TextMatrix(iRow, 2)
            ZResultado = WVector1.TextMatrix(iRow, 3)
            
            If Val(ZEnsayo) <> 0 Or ZDescripcion <> "" Or ZResultado <> "" Then
            
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                WClave = Orden.Text + Auxi
        
                ZSql = ""
                ZSql = ZSql + "INSERT INTO OrdenTrabajoII ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Orden ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Ensayo ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "Resultado )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Orden.Text + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + ZEnsayo + "',"
                ZSql = ZSql + "'" + ZDescripcion + "',"
                ZSql = ZSql + "'" + ZResultado + "')"
            
                spOrdenTrabajoII = ZSql
                Set rstOrdenTrabajoII = db.OpenRecordset(spOrdenTrabajoII, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next iRow
        
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
    
    Call Limpia_Vector

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
    ObservacionesIV.Text = ""
    ObservacionesV.Text = ""
    ObservacionesVI.Text = ""
    
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
        If Cliente.Text <> "" Then
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
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

    Call Limpia_Vector
    
    WVector1.Col = 1
    WVector1.Row = 1

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
    ObservacionesIV.Text = ""
    ObservacionesV.Text = ""
    ObservacionesVI.Text = ""
    
    Aplicacion.ListIndex = 0
    Estabilidad.ListIndex = 0
    
    DesCliente.Caption = ""
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    Select Case Tablas.Tab
        Case 0
            Material.SetFocus
        Case 1
            RequisitoI.SetFocus
        Case 2
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        Case Else
    End Select
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Ensayos"
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
      Rem BY NAN
       XEmpresa = WEmpresa
       txtOdbc = "Empresa" + "01"
       WEmpresa = "0001"
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
             
            
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
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Order by Codigo"
            spEnsayo = ZSql
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                With rstEnsayo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnsayo!Codigo) + " " + rstEnsayo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEnsayo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayo.Close
            End If
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    Call Conecta_Empresa
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
            ZEnsayo = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Ensayos"
            Sql3 = " Where Ensayos.Codigo = " + "'" + ZEnsayo + "'"
            spEnsayo = Sql1 + Sql2 + Sql3
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = Trim(rstEnsayo!Codigo)
                WVector1.Col = 2
                WVector1.Text = Trim(rstEnsayo!Descripcion)
                WVector1.Col = 3
                rstEnsayo.Close
                Call StartEdit
            End If
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
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
        Case 3
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
            If Val(WVector1.Text) <> 0 Then
            
                ZCodigo = WVector1.Text
                
                Sql1 = "Select *"
                Sql2 = " FROM Ensayos"
                Sql3 = " Where Ensayos.Codigo = " + "'" + ZCodigo + "'"
                spEnsayo = Sql1 + Sql2 + Sql3
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstEnsayo!Descripcion
                    rstEnsayo.Close
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
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        ZCodigo = WVector1.TextMatrix(iRow, 1)
        If ZCodigo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For DA = 0 To WVector1.Cols - 1
            WVector1.Col = DA
            WVector1.Text = WBorra(Ciclo, DA)
        Next DA
    Next Ciclo
    
    End If
    
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Clientes"
     Opcion.AddItem "Ensayos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
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
    WVector1.Cols = 4
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Ensayo"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Requerido"
                WVector1.ColWidth(Ciclo) = 4200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
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

