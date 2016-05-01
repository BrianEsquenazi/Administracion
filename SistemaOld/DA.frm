VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDA 
   AutoRedraw      =   -1  'True
   Caption         =   "Modificacion de Precios"
   ClientHeight    =   7275
   ClientLeft      =   1875
   ClientTop       =   840
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   8400
   Begin TabDlg.SSTab Tablas 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12303
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "DA.frx":0000
      Tab(0).ControlCount=   19
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "HastaDescri"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DesdeDescri"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DesdeProd"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "HastaProd"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Hastacliente"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DesdeCliente"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Porcentaje"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Consulta"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Cancela"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Acepta"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Ayuda"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Pantalla"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Opcion"
      Tab(0).Control(18).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "DA.frx":001C
      Tab(1).ControlCount=   19
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "HastaDescri1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DesdeDescri1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label13"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DesdeArti"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "HastaArti"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "HastaCliente1"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "DesdeCliente1"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Porcentaje1"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Consulta1"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Cancela1"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Acepta1"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Pantalla1"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "Ayuda1"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Opcion1"
      Tab(1).Control(18).Enabled=   -1  'True
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
         Height          =   1740
         Left            =   1920
         TabIndex        =   39
         Top             =   4440
         Width           =   3615
      End
      Begin VB.ListBox Pantalla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   240
         TabIndex        =   38
         Top             =   4440
         Width           =   7335
      End
      Begin VB.TextBox Opcion1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -73680
         TabIndex        =   34
         Top             =   4440
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox Ayuda1 
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
         TabIndex        =   33
         Top             =   4080
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox Pantalla1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   32
         Top             =   4440
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.CommandButton Acepta1 
         Caption         =   "Aceptar"
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
         Left            =   -70920
         TabIndex        =   31
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Cancela1 
         Caption         =   "Cancelar"
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
         Left            =   -70920
         TabIndex        =   30
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Consulta1 
         Caption         =   "Consulta"
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
         Left            =   -69480
         TabIndex        =   29
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Porcentaje1 
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
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   21
         Text            =   " "
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox DesdeCliente1 
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
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   20
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox HastaCliente1 
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
         Left            =   -72840
         MaxLength       =   6
         TabIndex        =   19
         Text            =   " "
         Top             =   1440
         Width           =   975
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
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         Left            =   4080
         TabIndex        =   17
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Left            =   4080
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
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
         Left            =   5520
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Porcentaje 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   4
         Text            =   " "
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox DesdeCliente 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Hastacliente 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   3
         Text            =   " "
         Top             =   1440
         Width           =   975
      End
      Begin MSMask.MaskEdBox HastaProd 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeProd 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox HastaArti 
         Height          =   375
         Left            =   -72840
         TabIndex        =   36
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeArti 
         Height          =   375
         Left            =   -72840
         TabIndex        =   37
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label14 
         Caption         =   "Hasta Cliente"
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
         TabIndex        =   35
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "ACTUALIZACION DE PRECIOS DE DY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73920
         TabIndex        =   28
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label Label12 
         Caption         =   "Desde Cliente"
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
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Desde Codigo"
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
         Left            =   -74640
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Hasta Codigo"
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
         Left            =   -74640
         TabIndex        =   25
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Porcentaje"
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
         Left            =   -74640
         TabIndex        =   24
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label DesdeDescri1 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   -71640
         TabIndex        =   23
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label HastaDescri1 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   -71640
         TabIndex        =   22
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta Cliente"
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "ACTUALIZACION DE PRECIOS DE PRODUCTOS TERMINADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Cliente"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Producto"
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
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Producto"
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
         TabIndex        =   10
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Porcentaje"
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
         TabIndex        =   9
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label DesdeDescri 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label HastaDescri 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   1440
         Width           =   4215
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   -120
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "listsol.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Solicitudes de Conpras Realizadas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Imprime_Descripcion1()

    WCliente = DesdeCliente1.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesdeDescri1.Caption = rstCliente!Razon
            Else
        DesdeDescri1.Caption = ""
    End If
    
    WCliente = Hastacliente2.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        HastaDescri2.Caption = rstCliente!Razon
            Else
        HastaDescri2.Caption = ""
    End If
    
End Sub

Private Sub Acepta1_Click()
    
    DesdeCliente1.Text = UCase(DesdeCliente1.Text)
    HastaCliente1.Text = UCase(HastaCliente1.Text)
    DesdeArti.Text = UCase(DesdeArti.Text)
    DesdeArti.Text = UCase(DesdeArti.Text)
                
    spPreciosMp = "ListaPreciosMP"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    
    Renglon = 0
    Erase Vector
    
    With rstPreciosMp
        .MoveFirst
        If .NoMatch = False Then
            Do
                If DesdeArti.Text <= rstPreciosMp!Articulo And DesdeArti.Text >= rstPreciosMp!Articulo Then
                    If DesdeCliente1.Text <= rstPreciosMp!Cliente And HastaCliente1.Text >= rstPreciosMp!Cliente Then
                        Renglon = Renglon + 1
                        Vector(Renglon) = rstPreciosMp!Clave
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    For XX = 1 To 10000
        If Vector(XX) <> "" Then
            spPreciosMp = "ConsultaPrecios " + "'" + Vector(XX) + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                WPrecio = Str$(rstPreciosMp!Precio + (rstPreciosMp!Precio * Val(Porcentaje.Text) / 100))
                WClave = rstPreciosMp!Clave
                WCliente = rstPreciosMp!Cliente
                WArticulo = rstPreciosMp!Articulo
                WDescripcion = rstPreciosMp!Descripcion
                WDate = Date$
                                     
                XParam = "'" + WClave + "','" + WPrecio + "','" + WDate + "'"
                Set rstPreciosMp = db.OpenRecordset("ModificaPrecios3 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    Next XX
    
    Call Cancela1_click
    
End Sub

Private Sub Cancela1_click()

    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeArti.Text = "  -     -   "
    DesdeArti.Text = "  -     -   "
    Porcentaje.Text = ""
    DesdeDescri.Caption = ""
    HastaDescri.Caption = ""
    Opcion1.Visible = False
    Pantalla1.Visible = False
    
    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeArti.Text = "  -   -   "
    HastaArti.Text = "  -   -   "
    Porcentaje1.Text = ""
    DesdeDescri1.Caption = ""
    HastaDescri1.Caption = ""
    Opcion1.Visible = False
    Pantalla1.Visible = False

    DesdeCliente1.SetFocus
    PrgModif.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub DesdeCliente1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente1.Text = UCase(DesdeCliente1.Text)
        Call Imprime_Descripcion1
        HastaCliente1.SetFocus
    End If
End Sub

Private Sub HastaCliente1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCliente1.Text = UCase(HastaCliente1.Text)
        Call Imprime_Descripcion1
        DesdeArti.SetFocus
    End If
End Sub

Private Sub DesdeArti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArti.Text = UCase(DesdeArti.Text)
        DesdeArti.SetFocus
    End If
End Sub

Private Sub DesdeArti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArti.Text = UCase(DesdeArti.Text)
        Porcentaje1.SetFocus
    End If
End Sub

Private Sub Porcentaje1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Porcentaje1.Text = Pusing("###,###.##", Str$(Val(Porcentaje1.Text)))
        DesdeCliente1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta1_Click()

    Opcion1.Visible = False
    Pantalla1.Visible = False

    Opcion1.Clear

    Opcion1.AddItem "Clientes"
    Opcion1.AddItem "DY"

    Opcion1.Visible = True
    
End Sub

Private Sub Opcion1_Click()

    Opcion1.Visible = False
     
    Dim IngresaItem As String

    Pantalla1.Clear
    WIndice.Clear

    XIndice = Opcion1.ListIndex
    
    Select Case XIndice
        Case 0
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                        Pantalla1.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCliente.Close
            
        Case 1
            spArticulo = "ListaTerminado"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Left$(rstArticulo!XCodigo, 2) = "DY" Then
                                IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                Pantalla1.AddItem IngresaItem
                                IngresaItem = rstArticulo!Codigo
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case Else
    End Select
            
    Pantalla1.Visible = True
    Ayuda1.Text = ""
    Ayuda1.Visible = True
    Ayuda1.SetFocus

End Sub

Private Sub Ayuda1_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla1.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda1.Text)
    
    Select Case XIndice
        Case 0
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            da = Len(rstCliente!Razon) - WEspacios
                
                            For aa = 1 To da
                                If Left$(Ayuda1.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                                    Auxi = rstCliente!Cliente
                                    IngresaItem = Auxi + "    " + rstCliente!Razon
                                    Pantalla1.AddItem IngresaItem
                                    IngresaItem = rstCliente!Cliente
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 1
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            If Left$(rstArticulo!Codigo2) = "DY" Then
                                da = Len(rstArticulo!Descripcion) - WEspacios
                                For aa = 1 To da
                                    If Left$(Ayuda1.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                        Auxi = rstArticulo!Codigo
                                        IngresaItem = Auxi + "    " + rstTerminado!Descripcion
                                        Pantalla1.AddItem IngresaItem
                                        IngresaItem = rstArticulo!Codigo
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next aa
                            End If
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case Else
        
    End Select
            
    End If

End Sub

Private Sub Pantalla1_Click()
    Pantalla1.Visible = False
    Ayuda1.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla1.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesdeCliente1.Text = rstCliente!Cliente
                HastaCliente1.Text = rstCliente!Cliente
                Call Imprime_Descripcion
                        Else
                DesdeCliente1.Text = WCliente
                HastaCliente1.Text = WCliente
                Call Imprime_Descripcion
            End If
            DesdeCliente1.SetFocus
            
        Case 1
            Indice = Pantalla1.ListIndex
            WTerminado = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesdeArti.Text = rstTerminado!Codigo
                DesdeArti.Text = rstTerminado!Codigo
                        Else
                DesdeArti.Text = WTerminado
                DesdeArti.Text = WTerminado
            End If
            DesdeArti.SetFocus
            
        Case Else
    End Select
    
End Sub

