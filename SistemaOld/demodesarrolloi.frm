VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDemoDesarrolloI 
   Caption         =   "Ingreso de Orden de Trabajo"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11625
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   10398
      _Version        =   327680
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion de la Orden"
      TabPicture(0)   =   "demodesarrolloi.frx":0000
      Tab(0).ControlCount=   18
      Tab(0).ControlEnabled=   0   'False
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
      TabCaption(1)   =   "Especificaciones"
      TabPicture(1)   =   "demodesarrolloi.frx":001C
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
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "RequisitoII"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "RequisitoIII"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "RequisitoIV"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "RequisitoV"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "RequisitoVI"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ReferenciaI"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "ReferenciaII"
      Tab(1).Control(11).Enabled=   0   'False
      TabCaption(2)   =   "Controles"
      TabPicture(2)   =   "demodesarrolloi.frx":0038
      Tab(2).ControlCount=   11
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label22"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Aplicacion"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Estabilidad"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text39"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text40"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text41"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "MSFlexGrid1"
      Tab(2).Control(10).Enabled=   0   'False
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   51
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   49
         Text            =   " "
         Top             =   3960
         Width           =   5895
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2655
         Left            =   240
         TabIndex        =   48
         Top             =   1800
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4683
         _Version        =   327680
      End
      Begin VB.TextBox Text41 
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   5280
         Width           =   5895
      End
      Begin VB.TextBox Text40 
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   4920
         Width           =   5895
      End
      Begin VB.TextBox Text39 
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   44
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   42
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   41
         Text            =   " "
         Top             =   4320
         Width           =   5895
      End
      Begin VB.ComboBox Estabilidad 
         Height          =   315
         Left            =   3960
         TabIndex        =   40
         Top             =   1020
         Width           =   2055
      End
      Begin VB.ComboBox Aplicacion 
         Height          =   315
         Left            =   3960
         TabIndex        =   39
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
         Left            =   -71520
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
         Left            =   -71520
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
         Left            =   -71520
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
         Left            =   -71520
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
         Left            =   -71520
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
         Left            =   -71520
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
         Left            =   -71520
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   540
         Width           =   5895
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
         Left            =   -74640
         TabIndex        =   52
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
         Left            =   -74640
         TabIndex        =   50
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
         Left            =   240
         TabIndex        =   45
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
         Left            =   480
         TabIndex        =   43
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
         Left            =   2280
         TabIndex        =   38
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
         Left            =   2280
         TabIndex        =   37
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Verificar"
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
         Left            =   480
         TabIndex        =   36
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Referencias de Desarrollo Similares"
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
         Left            =   -74640
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
         Left            =   -74640
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
         Left            =   -74640
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
         Left            =   -74640
         TabIndex        =   13
         Top             =   540
         Width           =   3015
      End
   End
   Begin VB.TextBox observaciones 
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
   Begin VB.TextBox Orden 
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
      TabIndex        =   1
      Top             =   240
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
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6120
      MouseIcon       =   "demodesarrolloi.frx":0054
      MousePointer    =   99  'Custom
      Picture         =   "demodesarrolloi.frx":035E
      ToolTipText     =   "Salida"
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   2760
      MouseIcon       =   "demodesarrolloi.frx":0BA0
      MousePointer    =   99  'Custom
      Picture         =   "demodesarrolloi.frx":0EAA
      ToolTipText     =   "Elimina el Registro"
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   1920
      MouseIcon       =   "demodesarrolloi.frx":16EC
      MousePointer    =   99  'Custom
      Picture         =   "demodesarrolloi.frx":19F6
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   3600
      MouseIcon       =   "demodesarrolloi.frx":2238
      MousePointer    =   99  'Custom
      Picture         =   "demodesarrolloi.frx":2542
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   5280
      MouseIcon       =   "demodesarrolloi.frx":2D84
      MousePointer    =   99  'Custom
      Picture         =   "demodesarrolloi.frx":308E
      ToolTipText     =   "Impresion "
      Top             =   7680
      Visible         =   0   'False
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
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "PrgDemoDesarrolloI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEquipoFabrica As Recordset
Dim spEquipoFabrica As String

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "0"
    End If
End Sub

Sub Imprime_Datos()

    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        Descripcion.Text = Trim(rstEquipoFabrica!Descripcion)
        DescripcionII.Text = Trim(rstEquipoFabrica!DescripcionII)
        DescripcionIII.Text = Trim(rstEquipoFabrica!DescripcionIII)
        rstEquipoFabrica.Close
    End If
    
End Sub

Private Sub Acepta_Click()
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = ""
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{EquipoFabrica.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Val(Codigo.Text) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM EquipoFabrica"
        Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
        spEquipoFabrica = Sql1 + Sql2 + Sql3
        Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipoFabrica.RecordCount > 0 Then
            rstEquipoFabrica.Close
            Sql1 = "UPDATE EquipoFabrica SET "
            Sql2 = " Descripcion = " + "'" + Descripcion.Text + "',"
            Sql3 = " DescripcionII = " + "'" + DescripcionII.Text + "',"
            Sql4 = " DescripcionIII = " + "'" + DescripcionIII.Text + "'"
            Sql5 = " Where Codigo = " + "'" + Codigo.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
                Else
            Sql1 = "INSERT INTO EquipoFabrica ("
            Sql2 = "Codigo ,"
            Sql3 = "Descripcion ,"
            Sql4 = "DescripcionII ,"
            Sql5 = "DescripcionIII )"
            Sql6 = "Values ("
            Sql7 = "'" + Codigo.Text + "',"
            Sql8 = "'" + Descripcion.Text + "',"
            Sql9 = "'" + DescripcionII.Text + "',"
            Sql10 = "'" + DescripcionIII.Text + "')"
            spEquipoFabrica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Val(Codigo.Text) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM EquipoFabrica"
        Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
        spEquipoFabrica = Sql1 + Sql2 + Sql3
        Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipoFabrica.RecordCount > 0 Then
            rstEquipoFabrica.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Sql1 = "DELETE EquipoFabrica"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spEquipoFabrica = Sql1 + Sql2
                Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    DescripcionIII.Text = ""

    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgEquiposFabrica.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Lista_Click()
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionII.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
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
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionIII.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
        
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                rstEquipoFabrica.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Anterior_Click()

    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Where EquipoFabrica.Codigo < " + "'" + Codigo.Text + "'"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        With rstEquipoFabrica
            .MoveLast
            Codigo.Text = rstEquipoFabrica!Codigo
        End With
        rstEquipoFabrica.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Equipos, Control y Instrucciones de Seguridad")
    End If
    
End Sub

Private Sub Siguiente_Click()

    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Where EquipoFabrica.Codigo > " + "'" + Codigo.Text + "'"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        With rstEquipoFabrica
            .MoveFirst
            Codigo.Text = rstEquipoFabrica!Codigo
        End With
        rstEquipoFabrica.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Equipos, Control y Instrucciones de Seguridad")
    End If

End Sub

Sub Form_Load()

    Codigo.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    DescripcionIII.Text = ""
    
End Sub


