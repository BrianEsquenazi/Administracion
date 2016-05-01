VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDatosEtiquetaMp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Instrucciones de Produccion de M.P."
   ClientHeight    =   9450
   ClientLeft      =   135
   ClientTop       =   285
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9450
   ScaleWidth      =   14745
   Visible         =   0   'False
   Begin VB.CommandButton Impre 
      Caption         =   "MP CARGADOS"
      Height          =   735
      Left            =   12960
      TabIndex        =   53
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   8040
      TabIndex        =   52
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   15
      Left            =   2880
      TabIndex        =   51
      Top             =   240
      Width           =   135
   End
   Begin VB.ComboBox Palabra 
      Height          =   315
      Left            =   2400
      TabIndex        =   45
      Top             =   480
      Width           =   2055
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3240
      TabIndex        =   23
      Top             =   6600
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   26
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame XClaveII 
      Height          =   1935
      Left            =   3360
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClaveII 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGrabaII 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   9763
      _Version        =   327680
      Tabs            =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pictogramas"
      TabPicture(0)   =   "datosetiquetaMp.frx":0000
      Tab(0).ControlCount=   27
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "asd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "sdafdsf"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "sddsf"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "sdsdfsdfsd"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "sdfsdfds"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "sdfds"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "sfdsf"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "sdfsdf"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label11"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label13"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Logo1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Logo2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Logo3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Logo4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "logo5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Logo6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Logo7"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Logo8"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Logo9"
      Tab(0).Control(26).Enabled=   0   'False
      TabCaption(1)   =   "Frases H"
      TabPicture(1)   =   "datosetiquetaMp.frx":001C
      Tab(1).ControlCount=   5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WTexto3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "WVector1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WTexto1"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "WCombo1"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "WTexto2"
      Tab(1).Control(4).Enabled=   -1  'True
      TabCaption(2)   =   "Frases P"
      TabPicture(2)   =   "datosetiquetaMp.frx":0038
      Tab(2).ControlCount=   5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "WVector2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "WTexto32"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "WCombo12"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "WTexto22"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "WTexto12"
      Tab(2).Control(4).Enabled=   -1  'True
      TabCaption(3)   =   "Denominacion Componentes Peligrosos"
      TabPicture(3)   =   "datosetiquetaMp.frx":0054
      Tab(3).ControlCount=   5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "WTexto33"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "WTexto23"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "WTexto13"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "WCombo13"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "WVector3"
      Tab(3).Control(4).Enabled=   0   'False
      TabCaption(4)   =   "Datos Onu"
      TabPicture(4)   =   "datosetiquetaMp.frx":0070
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame5 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   54
         Top             =   840
         Width           =   11295
         Begin VB.TextBox Clase 
            Height          =   285
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   59
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox Intervencion 
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   58
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox Naciones 
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   57
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Embalaje 
            Height          =   285
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   56
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox Caracteristicas 
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   55
            Text            =   " "
            Top             =   1680
            Width           =   9135
         End
         Begin VB.Label Label17 
            Caption         =   "Clase"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label27 
            Caption         =   "F.Intervencion"
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   240
            TabIndex        =   63
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label29 
            Caption         =   "Nro. N.Unidas"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label30 
            Caption         =   "Grupo Embalaje"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label62 
            Caption         =   "Caracteristicas"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   240
            TabIndex        =   60
            Top             =   1680
            Width           =   1455
         End
      End
      Begin MSMask.MaskEdBox WTexto33 
         Height          =   285
         Left            =   -71280
         TabIndex        =   50
         Top             =   2160
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
      Begin VB.TextBox WTexto23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   -72120
         TabIndex        =   49
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox WTexto13 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -73080
         TabIndex        =   48
         Top             =   2040
         Width           =   375
      End
      Begin VB.ComboBox WCombo13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70440
         TabIndex        =   46
         Top             =   1980
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox Logo9 
         Height          =   315
         Left            =   7200
         TabIndex        =   44
         Top             =   4860
         Width           =   1575
      End
      Begin VB.ComboBox Logo8 
         Height          =   315
         Left            =   4920
         TabIndex        =   43
         Top             =   4860
         Width           =   1575
      End
      Begin VB.ComboBox Logo7 
         Height          =   315
         Left            =   2640
         TabIndex        =   42
         Top             =   4860
         Width           =   1575
      End
      Begin VB.ComboBox Logo6 
         Height          =   315
         Left            =   360
         TabIndex        =   41
         Top             =   4860
         Width           =   1575
      End
      Begin VB.ComboBox logo5 
         Height          =   315
         Left            =   9480
         TabIndex        =   40
         Top             =   2580
         Width           =   1575
      End
      Begin VB.ComboBox Logo4 
         Height          =   315
         Left            =   7200
         TabIndex        =   39
         Top             =   2580
         Width           =   1575
      End
      Begin VB.ComboBox Logo3 
         Height          =   315
         Left            =   4920
         TabIndex        =   38
         Top             =   2580
         Width           =   1575
      End
      Begin VB.ComboBox Logo2 
         Height          =   315
         Left            =   2640
         TabIndex        =   37
         Top             =   2580
         Width           =   1575
      End
      Begin VB.ComboBox Logo1 
         Height          =   315
         Left            =   360
         TabIndex        =   36
         Top             =   2580
         Width           =   1575
      End
      Begin VB.TextBox WTexto2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   -72000
         TabIndex        =   20
         Top             =   1980
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70560
         TabIndex        =   19
         Top             =   1920
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto1 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -72600
         TabIndex        =   18
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox WTexto12 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -73440
         TabIndex        =   15
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox WTexto22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   -72720
         TabIndex        =   14
         Top             =   2040
         Width           =   375
      End
      Begin VB.ComboBox WCombo12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71160
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   -72000
         TabIndex        =   16
         Top             =   2040
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
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   17
         Top             =   780
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   6588
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   21
         Top             =   780
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   6588
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   -71640
         TabIndex        =   22
         Top             =   1980
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
      Begin MSFlexGridLib.MSFlexGrid WVector3 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   47
         Top             =   720
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6588
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.Label Label13 
         Caption         =   "7 - Peligro"
         Height          =   255
         Left            =   2640
         TabIndex        =   35
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "6 - Toxico"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "3 - Carburante"
         Height          =   255
         Left            =   4920
         TabIndex        =   33
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "8 - Peligro P/la Salud"
         Height          =   255
         Left            =   4920
         TabIndex        =   32
         Top             =   4620
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "5 - Corrosivo"
         Height          =   255
         Left            =   9480
         TabIndex        =   31
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "2 - Inflamable"
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "9- -Medio Ambiente"
         Height          =   255
         Left            =   7200
         TabIndex        =   29
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "4 - Gases Bajo Presion"
         Height          =   255
         Left            =   7200
         TabIndex        =   28
         Top             =   2340
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "1 - Explosivo"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Image sdfsdf 
         Height          =   1500
         Left            =   2640
         Picture         =   "datosetiquetaMp.frx":008C
         Top             =   3060
         Width           =   1500
      End
      Begin VB.Image sfdsf 
         Height          =   1500
         Left            =   360
         Picture         =   "datosetiquetaMp.frx":0B93
         Top             =   3060
         Width           =   1500
      End
      Begin VB.Image sdfds 
         Height          =   1500
         Left            =   4920
         Picture         =   "datosetiquetaMp.frx":1949
         Top             =   780
         Width           =   1500
      End
      Begin VB.Image sdfsdfds 
         Height          =   1500
         Left            =   4920
         Picture         =   "datosetiquetaMp.frx":2638
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image sdsdfsdfsd 
         Height          =   1500
         Left            =   9480
         Picture         =   "datosetiquetaMp.frx":33AE
         Top             =   780
         Width           =   1500
      End
      Begin VB.Image sddsf 
         Height          =   1500
         Left            =   2640
         Picture         =   "datosetiquetaMp.frx":406B
         Top             =   720
         Width           =   1500
      End
      Begin VB.Image sdafdsf 
         Height          =   1500
         Left            =   7200
         Picture         =   "datosetiquetaMp.frx":4D5F
         Top             =   3060
         Width           =   1500
      End
      Begin VB.Image asd 
         Height          =   1500
         Left            =   7200
         Picture         =   "datosetiquetaMp.frx":5A0A
         Top             =   780
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1500
         Left            =   360
         Picture         =   "datosetiquetaMp.frx":6510
         Top             =   780
         Width           =   1500
      End
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11400
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Left            =   840
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
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
      Height          =   2220
      ItemData        =   "datosetiquetaMp.frx":7199
      Left            =   120
      List            =   "datosetiquetaMp.frx":71A0
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.Label Label8 
      Caption         =   "Palabra Advertencia"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label DesArticulo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   10320
      MouseIcon       =   "datosetiquetaMp.frx":71AE
      MousePointer    =   99  'Custom
      Picture         =   "datosetiquetaMp.frx":74B8
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7440
      MouseIcon       =   "datosetiquetaMp.frx":7CFA
      MousePointer    =   99  'Custom
      Picture         =   "datosetiquetaMp.frx":8004
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   9360
      MouseIcon       =   "datosetiquetaMp.frx":8846
      MousePointer    =   99  'Custom
      Picture         =   "datosetiquetaMp.frx":8B50
      ToolTipText     =   "Consulta de Datos"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   8400
      MouseIcon       =   "datosetiquetaMp.frx":9392
      MousePointer    =   99  'Custom
      Picture         =   "datosetiquetaMp.frx":969C
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Articulo"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgDatosEtiquetaMp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstDatosEtiquetaMp As Recordset
Dim spDatosEtiquetaMp As String
Dim rstFraseH As Recordset
Dim spFraseH As String
Dim rstFraseP As Recordset
Dim spFraseP As String


Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim Renglon As Integer

Private WGraba As String
Private WGrabaII As String
Dim CargaEmpresa(12, 2) As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Rem para el vector II

Dim WBorraII(1000, 20) As String
Dim WParametrosII(10, 20) As Double
Dim WFormatoII(20) As String
Dim WControlII As String

Rem para el vector III

Dim WBoraIII(1000, 20) As String
Dim WParametrosIII(10, 20) As Double
Dim WFormatoIII(20) As String
Dim WControlIII As String


Private Sub Command2_Click()


    Dim ZZVector(5000, 10) As String
    
    Erase ZZVector
    ZZLugar = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM DatosEtiquetaMp"
    Sql3 = " Order by Clave"
    spDatosEtiquetaMp = Sql1 + Sql2 + Sql3
    Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstDatosEtiquetaMp.RecordCount > 0 Then
        With rstDatosEtiquetaMp
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstDatosEtiquetaMp!Tipo = 1 Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstDatosEtiquetaMp!Clave
                        ZZVector(ZZLugar, 2) = rstDatosEtiquetaMp!fraseh
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstDatosEtiquetaMp.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClave = ZZVector(Ciclo, 1)
        ZZFraseH = ZZVector(Ciclo, 2)
        
        Sql1 = "Select *"
        Sql2 = " FROM FraseH"
        Sql3 = " Where FraseH.Codigo = " + "'" + ZZFraseH + "'"
        spFraseH = Sql1 + Sql2 + Sql3
        Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
        If rstFraseH.RecordCount > 0 Then
            ZZDescriI = rstFraseH!Descripcion
            ZZDescriII = rstFraseH!DescripcionII
            ZZDescriIII = rstFraseH!DescripcionIII
            rstFraseH.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "UPDATE DatosEtiquetaMp SET "
        ZSql = ZSql + " Descripcion1H = " + "'" + ZZDescriI + "',"
        ZSql = ZSql + " Descripcion2H = " + "'" + ZZDescriII + "',"
        ZSql = ZSql + " Descripcion3H = " + "'" + ZZDescriIII + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
        
        spDatosEtiquetaMp = ZSql
        Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Stop
    



End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Materias Primas"
     Opcion.AddItem "Frases H"
     Opcion.AddItem "Frases P"

     Opcion.Visible = True
     
End Sub

Private Sub Impre_Click()
    Listado.ReportFileName = "etimppeligrolist.rpt"
    Listado.Destination = 0
    Listado.Action = 1
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
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM FraseH"
            Sql3 = " Order by Codigo"
            spFraseH = Sql1 + Sql2 + Sql3
            Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
            If rstFraseH.RecordCount > 0 Then
                With rstFraseH
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstFrasesH!Codigo + " " + rstFraseH!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstFraseH!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstFrasesH.Close
            End If
            
            
        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM FraseP"
            Sql3 = " Order by Codigo"
            spFrasesP = Sql1 + Sql2 + Sql3
            Set rstFrasesP = db.OpenRecordset(spFraseP, dbOpenSnapshot, dbSQLPassThrough)
            If rstFraseP.RecordCount > 0 Then
                With rstFraseP
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstFraseP!Codigo + " " + rstFraseP!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstFraseP!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstFrasesP.Close
            End If
            
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + Naciones.Text + "'"
            ZSql = ZSql + " Order by Peligroso.Ficha, Peligroso.Descripcion"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
    
                With rstPeligroso
                
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            sal = 0
                            IngresaItem = Trim(rstPeligroso!ficha) + " " + rstPeligroso!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstPeligroso!Codigo
                            WIndice.AddItem IngresaItem
                            
                            .MoveNext
                            
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                        Loop
                    End If
                    
                End With
                rstPeligroso.Close
            End If
            
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgDatosEtiquetaMp.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    Articulo.Text = UCase(Articulo.Text)
    
    ZSql = ""
    ZSql = ZSql + "DELETE DatosEtiquetaMp"
    ZSql = ZSql + " Where Articulo = " + "'" + Articulo.Text + "'"
    spDatosEtiquetaMp = ZSql
    Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)

    For Ciclo = 1 To 99
    
        If Trim(WVector1.TextMatrix(Ciclo, 1)) <> "" Then
    
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 3)
            
            WClave = Articulo.Text + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO DatosEtiquetaMp ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Palabra ,"
            ZSql = ZSql + "Pictograma1 ,"
            ZSql = ZSql + "Pictograma2 ,"
            ZSql = ZSql + "Pictograma3 ,"
            ZSql = ZSql + "Pictograma4 ,"
            ZSql = ZSql + "Pictograma5 ,"
            ZSql = ZSql + "Pictograma6 ,"
            ZSql = ZSql + "Pictograma7 ,"
            ZSql = ZSql + "Pictograma8 ,"
            ZSql = ZSql + "Pictograma9 ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "FraseH ,"
            ZSql = ZSql + "Descripcion1H ,"
            ZSql = ZSql + "Descripcion2H ,"
            ZSql = ZSql + "Descripcion3H ,"
            ZSql = ZSql + "FraseP ,"
            ZSql = ZSql + "Descripcion1P ,"
            ZSql = ZSql + "Descripcion2P ,"
            ZSql = ZSql + "Descripcion3P ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Denominacion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Articulo.Text + "',"
            ZSql = ZSql + "'" + Auxi + "',"
            ZSql = ZSql + "'" + Str$(Palabra.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo1.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo2.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo3.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo4.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(logo5.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo6.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo7.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo8.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo9.ListIndex) + "',"
            ZSql = ZSql + "'" + "1" + "',"
            ZSql = ZSql + "'" + WVector1.TextMatrix(Ciclo, 1) + "',"
            ZSql = ZSql + "'" + WVector1.TextMatrix(Ciclo, 2) + "',"
            ZSql = ZSql + "'" + WVector1.TextMatrix(Ciclo, 3) + "',"
            ZSql = ZSql + "'" + WVector1.TextMatrix(Ciclo, 4) + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "')"
                
            spDatosEtiquetaMp = ZSql
            Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
        If Trim(WVector2.TextMatrix(Ciclo, 1)) <> "" Then
    
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 3)
            
            WClave = Articulo.Text + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO DatosEtiquetaMp ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Palabra ,"
            ZSql = ZSql + "Pictograma1 ,"
            ZSql = ZSql + "Pictograma2 ,"
            ZSql = ZSql + "Pictograma3 ,"
            ZSql = ZSql + "Pictograma4 ,"
            ZSql = ZSql + "Pictograma5 ,"
            ZSql = ZSql + "Pictograma6 ,"
            ZSql = ZSql + "Pictograma7 ,"
            ZSql = ZSql + "Pictograma8 ,"
            ZSql = ZSql + "Pictograma9 ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "FraseH ,"
            ZSql = ZSql + "Descripcion1H ,"
            ZSql = ZSql + "Descripcion2H ,"
            ZSql = ZSql + "Descripcion3H ,"
            ZSql = ZSql + "FraseP ,"
            ZSql = ZSql + "Descripcion1P ,"
            ZSql = ZSql + "Descripcion2P ,"
            ZSql = ZSql + "Descripcion3P ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Denominacion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Articulo.Text + "',"
            ZSql = ZSql + "'" + Auxi + "',"
            ZSql = ZSql + "'" + Str$(Palabra.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo1.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo2.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo3.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo4.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(logo5.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo6.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo7.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo8.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo9.ListIndex) + "',"
            ZSql = ZSql + "'" + "2" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WVector2.TextMatrix(Ciclo, 1) + "',"
            ZSql = ZSql + "'" + WVector2.TextMatrix(Ciclo, 2) + "',"
            ZSql = ZSql + "'" + WVector2.TextMatrix(Ciclo, 3) + "',"
            ZSql = ZSql + "'" + WVector2.TextMatrix(Ciclo, 4) + "',"
            ZSql = ZSql + "'" + WVector2.TextMatrix(Ciclo, 5) + "',"
            ZSql = ZSql + "'" + "" + "')"
                
            spDatosEtiquetaMp = ZSql
            Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
        If Trim(WVector3.TextMatrix(Ciclo, 1)) <> "" Then
    
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 3)
            
            WClave = Articulo.Text + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO DatosEtiquetaMp ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Palabra ,"
            ZSql = ZSql + "Pictograma1 ,"
            ZSql = ZSql + "Pictograma2 ,"
            ZSql = ZSql + "Pictograma3 ,"
            ZSql = ZSql + "Pictograma4 ,"
            ZSql = ZSql + "Pictograma5 ,"
            ZSql = ZSql + "Pictograma6 ,"
            ZSql = ZSql + "Pictograma7 ,"
            ZSql = ZSql + "Pictograma8 ,"
            ZSql = ZSql + "Pictograma9 ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "FraseH ,"
            ZSql = ZSql + "Descripcion1H ,"
            ZSql = ZSql + "Descripcion2H ,"
            ZSql = ZSql + "Descripcion3H ,"
            ZSql = ZSql + "FraseP ,"
            ZSql = ZSql + "Descripcion1P ,"
            ZSql = ZSql + "Descripcion2P ,"
            ZSql = ZSql + "Descripcion3P ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Denominacion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Articulo.Text + "',"
            ZSql = ZSql + "'" + Auxi + "',"
            ZSql = ZSql + "'" + Str$(Palabra.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo1.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo2.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo3.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo4.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(logo5.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo6.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo7.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo8.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Logo9.ListIndex) + "',"
            ZSql = ZSql + "'" + "3" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WVector3.TextMatrix(Ciclo, 1) + "')"
                
            spDatosEtiquetaMp = ZSql
            Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
        
        
    Next Ciclo
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Articulo"
    Sql3 = " Where Articulo.Codigo = " + "'" + Articulo.Text + "'"
    spArticulo = Sql1 + Sql2 + Sql3
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WWNaciones = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
        WWClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
        WWIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
        WWEmbalaje = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
        WWCaracteristicas = IIf(IsNull(rstArticulo!Descrionu), "", rstArticulo!Descrionu)
        rstArticulo.Close
    End If
    
    ZZGraba = "N"
    
    If Trim(WWNaciones) <> Trim(Naciones.Text) Then
        ZZGraba = "S"
    End If
    If Trim(WWClase) <> Trim(Clase.Text) Then
        ZZGraba = "S"
    End If
    If Trim(WWIntervencion) <> Trim(Intervencion.Text) Then
        ZZGraba = "S"
    End If
    If Trim(WWEmbalaje) <> Trim(Embalaje.Text) Then
        ZZGraba = "S"
    End If
    If Trim(WWCaracteristicas) <> Trim(Caracteristicas.Text) Then
        ZZGraba = "S"
    End If
    
    If ZZGraba = "S" Then
    
        T$ = "Datos Onu"
        m$ = "Desea actualizar los datos Onu "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            XEmpresa = Wempresa
                
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
            CargaEmpresa(2, 1) = "0003"
            CargaEmpresa(2, 2) = "Empresa03"
            CargaEmpresa(3, 1) = "0005"
            CargaEmpresa(3, 2) = "Empresa05"
            CargaEmpresa(4, 1) = "0006"
            CargaEmpresa(4, 2) = "Empresa06"
            CargaEmpresa(5, 1) = "0007"
            CargaEmpresa(5, 2) = "Empresa07"
            CargaEmpresa(6, 1) = "0010"
            CargaEmpresa(6, 2) = "Empresa10"
            CargaEmpresa(7, 1) = "0011"
            CargaEmpresa(7, 2) = "Empresa11"
            
            For Cicla = 1 To 7
            
                If CargaEmpresa(Cicla, 1) <> "" Then
            
                    Wempresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Articulo SET "
                    ZSql = ZSql & "Naciones = " + "'" + Naciones.Text + "',"
                    ZSql = ZSql & "Clase = " + "'" + Clase.Text + "',"
                    ZSql = ZSql & "Intervencion = " + "'" + Intervencion.Text + "',"
                    ZSql = ZSql & "Embalaje = " + "'" + Embalaje.Text + "',"
                    ZSql = ZSql & "DescriOnu = " + "'" + Caracteristicas.Text + "'"
                    ZSql = ZSql & " Where Codigo = " + "'" + Articulo.Text + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
                End If
                
            Next Cicla
            
            Call Conecta_Empresa
        
        End If
        
    End If
    
    
    
    
    
    Call Limpia_Click
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector
    Call Limpia_VectorII
    Call Limpia_VectorIII
    
    Tablas.Tab = 0

    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
            
    Naciones.Text = ""
    Clase.Text = ""
    Intervencion.Text = ""
    Embalaje.Text = ""
    Caracteristicas.Text = ""
    
    Palabra.ListIndex = 0
    Logo1.ListIndex = 0
    Logo2.ListIndex = 0
    Logo3.ListIndex = 0
    Logo4.ListIndex = 0
    logo5.ListIndex = 0
    Logo6.ListIndex = 0
    Logo7.ListIndex = 0
    Logo8.ListIndex = 0
    Logo9.ListIndex = 0
    
    Articulo.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Articulo.Text = WIndice.List(Indice)
            Call Articulo_Keypress(13)
            
            
        Case 3
            Indice = Pantalla.ListIndex
            WCodigo = WIndice.List(Indice)
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.Codigo = " + "'" + WCodigo + "'"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
                If Trim(rstPeligroso!Clase) <> "" Then
                    Clase.Text = rstPeligroso!Clase
                End If
                If Trim(rstPeligroso!ficha) <> "" Then
                    Intervencion.Text = rstPeligroso!ficha
                End If
                If Trim(rstPeligroso!Embalaje) <> "" Then
                    Embalaje.Text = rstPeligroso!Embalaje
                End If
                If Trim(rstPeligroso!Descripcion) <> "" Then
                    Caracteristicas.Text = Left$(rstPeligroso!Descripcion, 100)
                End If
                rstPeligroso.Close
            End If
            
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    
    


    Call Limpia_Vector
    Call Limpia_VectorII
    Call Limpia_VectorIII
    
    Logo1.Clear
    
    Logo1.AddItem ""
    Logo1.AddItem "1"
    Logo1.AddItem "2"
    Logo1.AddItem "3"
    Logo1.AddItem "4"
    Logo1.AddItem "5"
    
    Logo1.ListIndex = 0
    
    Logo2.Clear
    
    Logo2.AddItem ""
    Logo2.AddItem "1"
    Logo2.AddItem "2"
    Logo2.AddItem "3"
    Logo2.AddItem "4"
    Logo2.AddItem "5"
    
    Logo2.ListIndex = 0
    
    
    Logo3.Clear
    
    Logo3.AddItem ""
    Logo3.AddItem "1"
    Logo3.AddItem "2"
    Logo3.AddItem "3"
    Logo3.AddItem "4"
    Logo3.AddItem "5"
    
    Logo3.ListIndex = 0
    
    
    Logo4.Clear
    
    Logo4.AddItem ""
    Logo4.AddItem "1"
    Logo4.AddItem "2"
    Logo4.AddItem "3"
    Logo4.AddItem "4"
    Logo4.AddItem "5"
    
    Logo4.ListIndex = 0
    
    
    logo5.Clear
    
    logo5.AddItem ""
    logo5.AddItem "1"
    logo5.AddItem "2"
    logo5.AddItem "3"
    logo5.AddItem "4"
    logo5.AddItem "5"
    
    logo5.ListIndex = 0
    
    
    Logo6.Clear
    
    Logo6.AddItem ""
    Logo6.AddItem "1"
    Logo6.AddItem "2"
    Logo6.AddItem "3"
    Logo6.AddItem "4"
    Logo6.AddItem "5"
    
    Logo6.ListIndex = 0
    
    
    Logo7.Clear
    
    Logo7.AddItem ""
    Logo7.AddItem "1"
    Logo7.AddItem "2"
    Logo7.AddItem "3"
    Logo7.AddItem "4"
    Logo7.AddItem "5"
    
    Logo7.ListIndex = 0
    
    
    Logo8.Clear
    
    Logo8.AddItem ""
    Logo8.AddItem "1"
    Logo8.AddItem "2"
    Logo8.AddItem "3"
    Logo8.AddItem "4"
    Logo8.AddItem "5"
    
    Logo8.ListIndex = 0
    
    
    Logo9.Clear
    
    Logo9.AddItem ""
    Logo9.AddItem "1"
    Logo9.AddItem "2"
    Logo9.AddItem "3"
    Logo9.AddItem "4"
    Logo9.AddItem "5"
    
    Logo9.ListIndex = 0
    
    Palabra.Clear
    
    Palabra.AddItem ""
    Palabra.AddItem "Peligro"
    Palabra.AddItem "Atencion"
    
    Palabra.ListIndex = 0
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1

    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso()
    
    ZZLugar1 = 0
    ZZLugar2 = 0
    ZZLugar3 = 0
    
    Call Limpia_Vector
    Call Limpia_VectorII
    Call Limpia_VectorIII

    For Ciclo = 1 To 999
    
        Auxi = Ciclo
        Call Ceros(Auxi, 3)
        
        ZZClave = Articulo.Text + Auxi
    
        Sql1 = "Select *"
        Sql2 = " FROM DatosEtiquetaMP"
        Sql3 = " Where DatosEtiquetaMp.Clave = " + "'" + ZZClave + "'"
        spDatosEtiquetaMp = Sql1 + Sql2 + Sql3
        Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)
        If rstDatosEtiquetaMp.RecordCount > 0 Then
    
            Palabra.ListIndex = rstDatosEtiquetaMp!Palabra
            Logo1.ListIndex = rstDatosEtiquetaMp!pictograma1
            Logo2.ListIndex = rstDatosEtiquetaMp!pictograma2
            Logo3.ListIndex = rstDatosEtiquetaMp!pictograma3
            Logo4.ListIndex = rstDatosEtiquetaMp!pictograma4
            logo5.ListIndex = rstDatosEtiquetaMp!pictograma5
            Logo6.ListIndex = rstDatosEtiquetaMp!pictograma6
            Logo7.ListIndex = rstDatosEtiquetaMp!pictograma7
            Logo8.ListIndex = rstDatosEtiquetaMp!pictograma8
            Logo9.ListIndex = rstDatosEtiquetaMp!pictograma9
            
            Select Case rstDatosEtiquetaMp!Tipo
                Case 1
                    ZZLugar1 = ZZLugar1 + 1
                    WVector1.TextMatrix(ZZLugar1, 1) = Trim(rstDatosEtiquetaMp!fraseh)
                    WVector1.TextMatrix(ZZLugar1, 2) = Trim(rstDatosEtiquetaMp!descripcion1H)
                    WVector1.TextMatrix(ZZLugar1, 3) = Trim(rstDatosEtiquetaMp!descripcion2H)
                    WVector1.TextMatrix(ZZLugar1, 4) = Trim(rstDatosEtiquetaMp!descripcion3H)
                    
                Case 2
                    ZZLugar2 = ZZLugar2 + 1
                    WVector2.TextMatrix(ZZLugar2, 1) = Trim(rstDatosEtiquetaMp!FraseP)
                    WVector2.TextMatrix(ZZLugar2, 2) = Trim(rstDatosEtiquetaMp!descripcion1p)
                    WVector2.TextMatrix(ZZLugar2, 3) = Trim(rstDatosEtiquetaMp!descripcion2p)
                    WVector2.TextMatrix(ZZLugar2, 4) = Trim(rstDatosEtiquetaMp!descripcion3p)
                    WVector2.TextMatrix(ZZLugar2, 5) = Trim(rstDatosEtiquetaMp!Observaciones)
                Case 3
                    ZZLugar3 = ZZLugar3 + 1
                    WVector3.TextMatrix(ZZLugar3, 1) = Trim(rstDatosEtiquetaMp!denominacion)
                Case Else
            End Select
            
        End If
                
    Next Ciclo
    

    Rem For Ciclo = 1 To 99
    Rem     If Trim(WVector1.TextMatrix(Ciclo, 1)) <> "" Then
    Rem         Sql1 = "Select *"
    Rem         Sql2 = " FROM FraseH"
    Rem         Sql3 = " Where FraseH.Codigo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
    Rem         spFraseH = Sql1 + Sql2 + Sql3
    Rem         Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
    Rem         If rstFraseH.RecordCount > 0 Then
    Rem             WVector1.TextMatrix(Ciclo, 2) = rstFraseH!Descripcion
    Rem             rstFraseH.Close
    Rem         End If
    Rem     End If
    Rem Next Ciclo
    

End Sub

Private Sub Revalida_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE ProduI SET "
    ZSql = ZSql + " Autorizado = " + "'" + "S" + "'"
    ZSql = ZSql + " Where Articulo = " + "'" + Articulo.Text + "'"
                
    spProduI = ZSql
    Set rstProduI = db.OpenRecordset(spProduI, dbOpenSnapshot, dbSQLPassThrough)

    XEmpresa = Wempresa
    Erase CargaEmpresa
    
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
            CargaEmpresa(2, 1) = "0003"
            CargaEmpresa(2, 2) = "Empresa03"
            CargaEmpresa(3, 1) = "0005"
            CargaEmpresa(3, 2) = "Empresa05"
            CargaEmpresa(4, 1) = "0006"
            CargaEmpresa(4, 2) = "Empresa06"
            CargaEmpresa(5, 1) = "0007"
            CargaEmpresa(5, 2) = "Empresa07"
        Case 2, 4, 8, 9
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 10
            CargaEmpresa(1, 1) = "0010"
            CargaEmpresa(1, 2) = "Empresa10"
        Case Else
    End Select
    
    Rem dada
    Rem sacar ente renglon
    Rem para que actualiza el abm de producto
    Rem dada
    Erase CargaEmpresa
            
    For Cicla = 1 To 5
        If CargaEmpresa(Cicla, 1) <> "" Then
        
            Wempresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " EstadoI = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Articulo.Text + "'"
                
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next Cicla
    
    Call Conecta_Empresa

    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
    
    Articulo.SetFocus

End Sub

Rem Private Sub Siguiente_Click()
   Rem  If Val(Paso.Text) < 99 Then
     Rem   Paso.Text = Str$(Val(Paso.Text) + 1)
    Rem     Call Imprime_Proceso
  Rem  End If
Rem End Sub

Rem Private Sub Anterior_Click()
 Rem   If Val(Paso.Text) > 1 Then
  Rem     Call Pasa_Datos
   Rem     Paso.Text = Str$(Val(Paso.Text) - 1)
   Rem     Call Imprime_Proceso
  Rem  End If
Rem End Sub




Private Sub Articulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Articulo.Text = UCase(Articulo.Text)
        
        Sql1 = "Select *"
        Sql2 = " FROM Articulo"
        Sql3 = " Where Articulo.Codigo = " + "'" + Articulo.Text + "'"
        spArticulo = Sql1 + Sql2 + Sql3
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = Trim(rstArticulo!Descripcion)
                
            Naciones.Text = ""
            Clase.Text = ""
            Intervencion.Text = ""
            Embalaje.Text = ""
            Caracteristicas.Text = ""
                
            Naciones.Text = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
            Clase.Text = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
            Intervencion.Text = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
            Embalaje.Text = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
            Caracteristicas.Text = IIf(IsNull(rstArticulo!Descrionu), "", rstArticulo!Descrionu)
            
            rstArticulo.Close
            Call Proceso
                Else
            Articulo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Articulo.Text = "  -   -   "
        DesArticulo.Caption = ""
    End If
End Sub



Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0, 2
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Da = Len(rstArticulo!Descripcion) - WEspacios
                            For aa = 1 To Da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, aa, WEspacios) Then
                                    IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstArticulo!Codigo
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
                rstArticulo.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM MaterialAuxiliar"
            Sql3 = " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            Sql4 = " Order by Codigo"
            spMaterialAuxiliar = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMaterialAuxiliar.RecordCount > 0 Then
                With rstMaterialAuxiliar
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMaterialAuxiliar.Close
            End If
            
        Case 3
            Rem XEmpresa = WEmpresa
            Rem Select Case Val(WEmpresa)
            Rem     Case 1, 3, 5, 6, 7, 9.10
            Rem         WEmpresa = "0003"
            Rem        txtOdbc = "Empresa03"
            Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Rem     Case Else
            Rem         WEmpresa = "0004"
            Rem         txtOdbc = "Empresa04"
            Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Rem End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                With rstEnsayos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayos.Close
            End If
            
            Rem Call Conecta_Empresa
            
        Case 4
            Sql1 = "Select *"
            Sql2 = " FROM MetodoFiltrado"
            Sql3 = " Order by Codigo"
            spMetodoFiltrado = Sql1 + Sql2 + Sql3
            Set rstMetodoFiltrado = db.OpenRecordset(spMetodoFiltrado, dbOpenSnapshot, dbSQLPassThrough)
            If rstMetodoFiltrado.RecordCount > 0 Then
                With rstMetodoFiltrado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Da = Len(rstMetodoFiltrado!Descripcion) - WEspacios
                            For aa = 1 To Da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstMetodoFiltrado!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstMetodoFiltrado!Codigo) + " " + rstMetodoFiltrado!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstMetodoFiltrado!Codigo
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
                rstMetodoFiltrado.Close
            End If
            
        Case Else
    End Select
            
    End If

End Sub

Private Sub Articulo_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "materias Primas"
    Opcion.AddItem "Material Auxiliar a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

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
        Case 4
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
            Sql1 = "Select *"
            Sql2 = " FROM FraseH"
            Sql3 = " Where FraseH.Codigo = " + "'" + WVector1.Text + "'"
            spFraseH = Sql1 + Sql2 + Sql3
            Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
            If rstFraseH.RecordCount > 0 Then
            
                WVector1.Col = 2
                If Trim(WVector1.Text) = "" Then
                    WVector1.Text = Trim(rstFraseH!Descripcion)
                End If
                
                WVector1.Col = 3
                If Trim(WVector1.Text) = "" Then
                    WVector1.Text = Trim(rstFraseH!DescripcionII)
                End If
                
                WVector1.Col = 4
                If Trim(WVector1.Text) = "" Then
                    WVector1.Text = Trim(rstFraseH!DescripcionIII)
                End If
                
                WVector1.Col = 1
                
                rstFraseH.Close
                    Else
                WControl = "N"
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
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        XDescripcion = WVector1.Text
        
        If XDescripcion <> "" Then
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
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For Da = 1 To WVector1.Cols - 1
            WVector1.Col = Da
            WVector1.Text = WBorra(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Materias Primas"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos os a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    If WVector1.Col = 2 Then

    Opcion.Clear
    
     Opcion.AddItem "Materias Primas"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos os a Utilizar"

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
    WVector1.Cols = 5
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Codigo"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 8000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 100
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 8000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 100
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 8000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 100
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
    Rem WVector1.Width = WAncho

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














Rem
Rem Controles de la WVector2
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector2.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector2.CellLeft + WVector2.Left
            WTexto22.Top = WVector2.CellTop + WVector2.Top
            WTexto22.Width = WVector2.CellWidth
            WTexto22.Height = WVector2.CellHeight
            WTexto22.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector2.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector2.CellLeft + WVector2.Left
            WTexto32.Top = WVector2.CellTop + WVector2.Top
            WTexto32.Width = WVector2.CellWidth
            WTexto32.Height = WVector2.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector2.Text) = 10 Then
                        WTexto32.Text = WVector2.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector2.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector2.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector2.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector2.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormatoII(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector2.CellLeft + WVector2.Left
    WCombo12.Top = WVector2.CellTop + WVector2.Top
    WCombo12.Width = WVector2.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1
        Case 113
            WTexto12.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 123
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Col > 1 Then
                WVector2.Col = WVector2.Col - 1
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1
        Case 113
            WTexto22.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub Wtexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto32.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub Wtexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEditII
End Sub

Private Sub WVector2_LeaveCell()
    EndEditII
End Sub

Private Sub WVector2_GotFocus()
    EndEditII
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector2.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 5
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 1
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControlII = "S"
    Select Case XColumna
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM FraseP"
            Sql3 = " Where FraseP.Codigo = " + "'" + WVector2.Text + "'"
            spFraseP = Sql1 + Sql2 + Sql3
            Set rstFraseP = db.OpenRecordset(spFraseP, dbOpenSnapshot, dbSQLPassThrough)
            If rstFraseP.RecordCount > 0 Then
            
                WVector2.Col = 2
                If Trim(WVector2.Text) = "" Then
                    WVector2.Text = Trim(rstFraseP!Descripcion)
                End If
                
                WVector2.Col = 3
                If Trim(WVector2.Text) = "" Then
                    WVector2.Text = Trim(rstFraseP!DescripcionII)
                End If
                
                WVector2.Col = 4
                If Trim(WVector2.Text) = "" Then
                    WVector2.Text = Trim(rstFraseP!DescripcionIII)
                End If
                
                WVector2.Col = 1
                
                rstFraseP.Close
                    Else
                WControlII = "N"
            End If
            
        Case 3, 6, 7
            Rem If Val(WVector2.Text) <> 0 Then
            Rem     ZCodigo = Val(WVector2.Text)
            Rem     Call Ceros(ZCodigo, 4)
            Rem
            Rem     Sql1 = "Select *"
            Rem     Sql2 = " FROM EquipoFabrica"
            Rem     Sql3 = " Where EquipoFabrica.Codigo = " + "'" + ZCodigo + "'"
            Rem     spEquipoFabrica = Sql1 + Sql2 + Sql3
            Rem     Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstEquipoFabrica.RecordCount > 0 Then
            Rem         rstEquipoFabrica.Close
            Rem     End If
            Rem End If
            
        Case Else
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
    
    RenglonAuxiliar = WVector2.Row

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorraII
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        Ensayo = WVector2.TextMatrix(iRow, 1)
            
        If Ensayo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorraII(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For Da = 0 To WVector2.Cols - 1
            WVector2.Col = Da
            WVector2.Text = WBorraII(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_VectorII()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector2.FontName
    WTexto12.FontSize = WVector2.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector2.FontName
    WTexto22.FontSize = WVector2.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector2.FontName
    WTexto32.FontSize = WVector2.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector2.FontName
    WCombo12.FontSize = WVector2.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 6
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Codigo"
                WVector2.ColWidth(Ciclo) = 2000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 20
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 8000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 100
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 8000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 100
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 4
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 8000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 100
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 5
                WVector2.Text = "Observaciones"
                WVector2.ColWidth(Ciclo) = 8000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector2.Text
        Rem WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        Rem WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        Rem WTitulo(Ciclo).Width = WVector2.CellWidth
        Rem WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub

















Rem
Rem Controles de la WVector3
Rem

Private Sub GridEditTextIII(ByVal KeyAscii As Integer)

    XColumna = WVector3.Col
    XTipoDato = WParametrosIII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto13.Left = WVector3.CellLeft + WVector3.Left
            WTexto13.Top = WVector3.CellTop + WVector3.Top
            WTexto13.Width = WVector3.CellWidth
            WTexto13.Height = WVector3.CellHeight
            WTexto13.MaxLength = WParametrosIII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto13.Text = WVector3.Text
                    WTexto13.SelStart = Len(WTexto13.Text)
                Case Else
                    WTexto13.Text = Chr$(KeyAscii)
                    WTexto13.SelStart = 1
            End Select
            WTexto13.Visible = True
            WTexto13.SetFocus
        Case 1
            WTexto23.Left = WVector3.CellLeft + WVector3.Left
            WTexto23.Top = WVector3.CellTop + WVector3.Top
            WTexto23.Width = WVector3.CellWidth
            WTexto23.Height = WVector3.CellHeight
            WTexto23.MaxLength = WParametrosIII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto23.Text = WVector3.Text
                    Rem WTexto23.SelStart = Len(WTexto23.Text)
                    WTexto23.SelStart = 0
                Case Else
                    WTexto23.Text = Chr$(KeyAscii)
                    WTexto23.SelStart = 1
            End Select
            WTexto23.Visible = True
            WTexto23.SetFocus
        Case 2
            WTexto33.Left = WVector3.CellLeft + WVector3.Left
            WTexto33.Top = WVector3.CellTop + WVector3.Top
            WTexto33.Width = WVector3.CellWidth
            WTexto33.Height = WVector3.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector3.Text) = 10 Then
                        WTexto33.Text = WVector3.Text
                            Else
                        WTexto33.Text = "  /  /    "
                    End If
                    WTexto33.SelStart = 0
                Case Else
                    WTexto33.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto33.SelStart = 1
            End Select
            WTexto33.Visible = True
            WTexto33.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditIII()
    Pasa = 0
    If WCombo13.Visible Then
        Pasa = 0
        WVector3.Text = WCombo13.Text
        WCombo13.Visible = False
            Else
        If WTexto13.Visible Then
            Pasa = 1
            WVector3.Text = WTexto13.Text
            WTexto13.Visible = False
                Else
            If WTexto23.Visible Then
                Pasa = 1
                WVector3.Text = WTexto23.Text
                WTexto23.Visible = False
                    Else
                If WTexto33.Visible Then
                    Pasa = 1
                    WVector3.Text = WTexto33.Text
                    WTexto33.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoIII(WVector3.Col) <> "" Then
            WVector3.Text = Pusing(WFormatoIII(WVector3.Col), WVector3.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboIII()
    ' Position the ComboBox over the cell.
    WCombo13.Left = WVector3.CellLeft + WVector3.Left
    WCombo13.Top = WVector3.CellTop + WVector3.Top
    WCombo13.Width = WVector3.CellWidth
    WCombo13.Visible = True
    WCombo13.SetFocus
End Sub

Private Sub WTexto13_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto13.Text = ""
            
        Rem F1
        Case 113
            WTexto13.Text = WVector3.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector3.SetFocus
            DoEvents
            Call Control_CampoIII
            If WControlIII = "S" Then
                Call Control_WVector3
            End If
            Call StartEditIII

        Case vbKeyDown
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row < WVector3.Rows - 1 Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row + 1
                End If
            End If
            Call StartEditIII

        Case vbKeyUp
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row > WVector3.FixedRows Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row - 1
                End If
            End If
            Call StartEditIII
        Case 34
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow < WVector3.Rows - 12 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow + 12
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 33
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow - 12 > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow - 12
                    WVector3.Row = WVector3.TopRow
                        Else
                    WVector3.TopRow = 1
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 123
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Col > 1 Then
                WVector3.Col = WVector3.Col - 1
            End If
            Call StartEditIII

    End Select
End Sub

Private Sub WTexto23_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto23.Text = ""
            
        Rem F1
        Case 113
            WTexto23.Text = WVector3.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector3.SetFocus
            DoEvents
            Call Control_CampoIII
            If WControlIII = "S" Then
                Call Control_WVector3
            End If
            Call StartEditIII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row < WVector3.Rows - 1 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row + 1
                Rem End If
            End If
            Call StartEditIII

        Case vbKeyUp
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row - 1
                Rem End If
            End If
            Call StartEditIII
        Case 34
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow < WVector3.Rows - 12 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow + 12
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 33
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow - 12 > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow - 12
                    WVector3.Row = WVector3.TopRow
                        Else
                    WVector3.TopRow = 1
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII

    End Select
End Sub

Private Sub WTexto33_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto33.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto33.Text = WVector3.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector3.SetFocus
            Call Control_CampoIII
            If WControlIII = "S" Then
                Call Control_WVector3
            End If
            Call StartEditIII

        Case vbKeyDown
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row < WVector3.Rows - 1 Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row + 1
                End If
            End If
            Call StartEditIII

        Case vbKeyUp
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.Row > WVector3.FixedRows Then
                Call Control_CampoIII
                If WControlIII = "S" Then
                    WVector3.Row = WVector3.Row - 1
                End If
            End If
            Call StartEditIII
        Case 34
            ' Move down 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow < WVector3.Rows - 12 Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow + 12
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII
            
        Case 33
            ' Move up 1 row.
            WVector3.SetFocus
            DoEvents
            If WVector3.TopRow - 12 > WVector3.FixedRows Then
                Rem Call Control_CampoIII
                Rem If WControlIII = "S" Then
                    WVector3.TopRow = WVector3.TopRow - 12
                    WVector3.Row = WVector3.TopRow
                        Else
                    WVector3.TopRow = 1
                    WVector3.Row = WVector3.TopRow
                Rem End If
            End If
            Call StartEditIII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto13_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto23_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto33_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo13_Click()
    WVector3.SetFocus
End Sub


Private Sub WVector3_Click()
    StartEditIII
End Sub

Private Sub WVector3_LeaveCell()
    EndEditIII
End Sub

Private Sub WVector3_GotFocus()
    EndEditIII
End Sub

Private Sub WVector3_KeyPress(KeyAscii As Integer)
    XColumna = WVector3.Col
    Select Case WParametrosIII(4, WVector3.Col)
        Case 1
        Case Else
            If WParametrosIII(2, XColumna) = 0 Then
                GridEditTextIII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditIII()
    Select Case WParametrosIII(4, WVector3.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo13.Clear
            WCombo13.AddItem "Campo1"
            WCombo13.AddItem "Campo2"
            On Error Resume Next
            WCombo13.Text = WVector3.Text
            On Error GoTo 0
            GridEditComboIII
        Case Else
            If WParametrosIII(2, WVector3.Col) = 0 Then
                GridEditTextIII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector3()
    Select Case WVector3.Col
        Case 1
            If WVector3.Row < WVector3.Rows - 1 Then
                WVector3.Row = WVector3.Row + 1
            End If
            WVector3.Col = 1
        Case Else
            If WVector3.Col < WVector3.Cols - 1 Then
                WVector3.Col = WVector3.Col + 1
            End If
    End Select
    WVector3.SetFocus
    GridEditTextIII KeyAscii
End Sub

Private Sub Control_CampoIII()
    XColumna = WVector3.Col
    XFila = WVector3.Row
    WControlIII = "S"
    Select Case XColumna
            
        Case Else
            WVector3.Col = XColumna
    End Select
End Sub

Private Sub WVector3_DblClick()

    If WVector3.Col = 0 Or WVector3.Col = 1 Then
    
    WTexto13.Visible = False
    WTexto23.Visible = False
    WTexto33.Visible = False
    
    RenglonAuxiliar = WVector3.Row

    For Ciclo = 1 To WVector3.Cols - 1
        WVector3.Col = Ciclo
        WVector3.Text = ""
    Next Ciclo
    
    Erase WBorraIII
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        If WVector3.TextMatrix(iRow, 1) <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector3.Row = Ciclo
        WVector3.Col = 1
        WAuxi1 = WVector3.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector3.Cols - 1
                WVector3.Col = Ciclo1
                WBorraIII(EntraVector, Ciclo1) = WVector3.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorIII
    
    For Ciclo = 1 To EntraVector
        WVector3.Row = Ciclo
        For Da = 0 To WVector3.Cols - 1
            WVector3.Col = Da
            WVector3.Text = WBorraIII(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_VectorIII()

    WVector3.Clear

    Rem ponga la WVector3 en negritas
    WVector3.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto13.FontName = WVector3.FontName
    WTexto13.FontSize = WVector3.FontSize
    WTexto13.Visible = False
    WTexto23.FontName = WVector3.FontName
    WTexto23.FontSize = WVector3.FontSize
    WTexto23.Visible = False
    WTexto33.FontName = WVector3.FontName
    WTexto33.FontSize = WVector3.FontSize
    WTexto33.Visible = False
    WCombo13.FontName = WVector3.FontName
    WCombo13.FontSize = WVector3.FontSize
    WCombo13.Visible = False

    ' Establesco loa Valores de la WVector3
    
    WVector3.FixedCols = 1
    WVector3.Cols = 2
    WVector3.FixedRows = 1
    WVector3.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector3.Text = "Articulo"
    
    Rem Longitud
    Rem WVector3.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector3.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosIII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosIII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosIII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosIII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector3.ColWidth(0) = 200
    WVector3.Row = 0
    For Ciclo = 1 To WVector3.Cols - 1
        WVector3.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector3.Text = "Denominacion Componentes Peligrosos"
                WVector3.ColWidth(Ciclo) = 6000
                WVector3.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosIII(1, Ciclo) = 100
                WParametrosIII(2, Ciclo) = 0
                WParametrosIII(3, Ciclo) = 0
                WParametrosIII(4, Ciclo) = 0
                WFormatoIII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector3.Row = 0
    For Ciclo = 1 To WVector3.Cols - 1
        WVector3.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector3.Text
        Rem WTitulo(Ciclo).Left = WVector3.CellLeft + WVector3.Left
        Rem WTitulo(Ciclo).Top = WVector3.CellTop + WVector3.Top
        Rem WTitulo(Ciclo).Width = WVector3.CellWidth
        Rem WTitulo(Ciclo).Height = WVector3.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector3
    
    WAncho = 400
    For Ciclo = 0 To WVector3.Cols - 1
        WAncho = WAncho + WVector3.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector3.Width = WAncho

    ' Size the columns.
    Font.Name = WVector3.Font.Name
    Font.Size = WVector3.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector3.AllowUserResizing = flexResizeBoth
    
    WVector3.Col = 1
    WVector3.Row = 1
    
End Sub

Private Sub WVector3_Scroll()
    WTexto13.Visible = False
    WTexto23.Visible = False
    WTexto33.Visible = False
End Sub



































Private Sub Equipo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(Equipo.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Equipo.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                If Trim(DesEquipoI.Text) = "" And Trim(DesEquipoII.Text) = "" And Trim(DesEquipoIII.Text) = "" Then
                    DesEquipoI.Text = Trim(rstEquipoFabrica!Descripcion)
                    DesEquipoII.Text = Trim(rstEquipoFabrica!DescripcionII)
                    DesEquipoIII.Text = Trim(rstEquipoFabrica!DescripcionIII)
                End If
                rstEquipoFabrica.Close
                DesEquipoI.SetFocus
                    Else
                Equipo.SetFocus
            End If
                Else
            DesEquipoI.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Equipo.Text = ""
        DesEquipoI.Text = ""
        DesEquipoII.Text = ""
        DesEquipoIII.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesEquipoI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesEquipoII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesEquipoI.Text = ""
    End If
End Sub

Private Sub DesEquipoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesEquipoIII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesEquipoII.Text = ""
    End If
End Sub

Private Sub DesEquipoIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Seguridad.SetFocus
    End If
    If KeyAscii = 27 Then
        DesEquipoIII.Text = ""
    End If
End Sub

Private Sub Seguridad_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(Seguridad.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Seguridad.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                If Trim(DesSeguridadI.Text) = "" And Trim(DesSeguridadII.Text) = "" And Trim(DesSeguridadIII.Text) = "" Then
                    DesSeguridadI.Text = Trim(rstEquipoFabrica!Descripcion)
                    DesSeguridadII.Text = Trim(rstEquipoFabrica!DescripcionII)
                    DesSeguridadIII.Text = Trim(rstEquipoFabrica!DescripcionIII)
                End If
                rstEquipoFabrica.Close
                DesSeguridadI.SetFocus
                    Else
                Seguridad.SetFocus
            End If
                Else
            DesSeguridadI.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Seguridad.Text = ""
        DesSeguridadI.Text = ""
        DesSeguridadII.Text = ""
        DesSeguridadIII.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesSeguridadI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesSeguridadII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesSeguridadI.Text = ""
    End If
End Sub

Private Sub DesSeguridadII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesSeguridadIII.SetFocus
    End If
    If KeyAscii = 27 Then
        DesSeguridadII.Text = ""
    End If
End Sub

Private Sub DesSeguridadIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ControlI.SetFocus
    End If
    If KeyAscii = 27 Then
        DesSeguridadIII.Text = ""
    End If
End Sub








Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 1
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        Case 2
            WVector2.Col = 1
            WVector2.Row = 1
            Call StartEditII
        Case Else
            Rem DADA
    End Select
End Sub

Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub







Private Sub Naciones_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(Naciones.Text) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Peligroso"
            ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + Naciones.Text + "'"
            spPeligroso = ZSql
            Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
            If rstPeligroso.RecordCount > 0 Then
                rstPeligroso.Close
                
                ZLugar = 0
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Peligroso"
                ZSql = ZSql + " Where Peligroso.NroOnu = " + "'" + Naciones.Text + "'"
                ZSql = ZSql + " Order by Peligroso.Codigo"
                spPeligroso = ZSql
                Set rstPeligroso = db.OpenRecordset(spPeligroso, dbOpenSnapshot, dbSQLPassThrough)
                If rstPeligroso.RecordCount > 0 Then
        
                    With rstPeligroso
                    
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                                                            
                                ZLugar = ZLugar + 1
                                ZPeligrosoI = rstPeligroso!ficha
                                ZPeligrosoII = Left$(rstPeligroso!Descripcion, 100)
                                ZPeligrosoIII = rstPeligroso!Clase
                                ZPeligrosoIV = rstPeligroso!Secundario
                                ZPeligrosoV = rstPeligroso!Riesgo
                                ZPeligrosoVI = rstPeligroso!Embalaje
                                
                                .MoveNext
                                
                                If .EOF = True Then
                                    Exit Do
                                End If
                                
                            Loop
                        End If
                        
                    End With
                    rstPeligroso.Close
        
                End If
                
                If ZLugar = 1 Then
                
                    Pantalla.Visible = False
                
                    If Trim(ZPeligrosoIII) <> "" Then
                        Clase.Text = Trim(ZPeligrosoIII)
                    End If
                    If Trim(ZPeligrosoI) <> "" Then
                        Intervencion.Text = Trim(ZPeligrosoI)
                    End If
                    If Trim(ZPeligrosoVI) <> "" Then
                        Embalaje.Text = Trim(ZPeligrosoVI)
                    End If
                    If Trim(ZPeligrosoII) <> "" Then
                        Caracteristicas.Text = Left$(ZPeligrosoII, 100)
                    End If
                    
                        Else
                        
                    Opcion.Clear
                    Opcion.AddItem ""
                    Opcion.AddItem ""
                    Opcion.AddItem ""
                    Opcion.AddItem ""
                    Rem Opcion.Visible = True
                    Opcion.ListIndex = 3
                        
                End If
                
                    Else
                    
                Clase.Text = ""
                Intervencion.Text = ""
                Embalaje.Text = ""
                Caracteristicas.Text = ""
            
                m$ = "Nro de Naciones Unidas Inexistente"
                a% = MsgBox(m$, 0, "Archivo de Materias Primas")
                Exit Sub
                
            End If
            
                Else
                
            Clase.Text = ""
            Intervencion.Text = ""
            Embalaje.Text = ""
            Caracteristicas.Text = ""
            
        End If
        
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Clase_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Intervencion.SetFocus
    End If
End Sub

Private Sub Intervencion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Embalaje.SetFocus
    End If
End Sub

Private Sub Embalaje_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Caracteristicas.SetFocus
    End If
End Sub

Private Sub Caracteristicas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Clase.SetFocus
    End If
End Sub

