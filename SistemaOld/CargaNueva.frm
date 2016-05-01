VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaNueva 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Instrucciones de Produccion de P.T."
   ClientHeight    =   8625
   ClientLeft      =   135
   ClientTop       =   285
   ClientWidth     =   11685
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
   ScaleHeight     =   8625
   ScaleWidth      =   11685
   Visible         =   0   'False
   Begin VB.Frame IngresaBase 
      Height          =   1215
      Left            =   3600
      TabIndex        =   80
      Top             =   2280
      Visible         =   0   'False
      Width           =   3015
      Begin MSMask.MaskEdBox ProductoBase 
         Height          =   285
         Left            =   720
         TabIndex        =   81
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   2640
      TabIndex        =   86
      Top             =   1800
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   88
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   87
         Top             =   960
         Width           =   1815
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
         TabIndex        =   89
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame XClaveII 
      Height          =   1935
      Left            =   3360
      TabIndex        =   82
      Top             =   1800
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClaveII 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   84
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGrabaII 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   83
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
         TabIndex        =   85
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Revalida 
      Caption         =   "Revalida"
      Height          =   495
      Left            =   9600
      TabIndex        =   79
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton AgregaRenglon 
      Caption         =   "Agrega Renglon"
      Height          =   495
      Left            =   9600
      TabIndex        =   78
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Base 
      Caption         =   "Instrucciones Base"
      Height          =   495
      Left            =   7440
      TabIndex        =   77
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton BajaEtapa 
      Caption         =   "Borra Etapa"
      Height          =   855
      Left            =   8520
      TabIndex        =   76
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton AltaEtapa 
      Caption         =   "Inserta Etapa"
      Height          =   855
      Left            =   7440
      TabIndex        =   75
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox Paso 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   73
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Operador 
      Height          =   285
      Left            =   10800
      MaxLength       =   2
      TabIndex        =   72
      Text            =   " "
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Version 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   66
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox MetodoLavado 
      Height          =   285
      Left            =   7440
      MaxLength       =   2
      TabIndex        =   64
      Text            =   " "
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox MetodoFiltrado 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   61
      Text            =   " "
      Top             =   480
      Width           =   735
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   9840
      TabIndex        =   42
      Top             =   480
      Width           =   1695
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   8070
      _Version        =   327680
      Tab             =   1
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
      TabCaption(0)   =   "Procedimiento"
      TabPicture(0)   =   "CargaNueva.frx":0000
      Tab(0).ControlCount=   6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "WTexto3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "WVector1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "WTitulo(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "WTexto2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "WCombo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "WTexto1"
      Tab(0).Control(5).Enabled=   0   'False
      TabCaption(1)   =   "Controles de Calidad"
      TabPicture(1)   =   "CargaNueva.frx":001C
      Tab(1).ControlCount=   5
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "WVector2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "WTexto32"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WTexto12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "WTexto22"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "WCombo12"
      Tab(1).Control(4).Enabled=   0   'False
      TabCaption(2)   =   "Controles de Proceso"
      TabPicture(2)   =   "CargaNueva.frx":0038
      Tab(2).ControlCount=   40
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ControlV"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "DesdeV"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "HastaV"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "ControlIV"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "DesdeIV"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "HastaIV"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "ControlIII"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "DesdeIII"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "HastaIII"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "DesSeguridadIII"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "DesSeguridadII"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "Seguridad"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "DesSeguridadI"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "DesEquipoI"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "Equipo"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "DesEquipoII"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "DesEquipoIII"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "TiempoI"
      Tab(2).Control(17).Enabled=   -1  'True
      Tab(2).Control(18)=   "TiempoII"
      Tab(2).Control(18).Enabled=   -1  'True
      Tab(2).Control(19)=   "HastaI"
      Tab(2).Control(19).Enabled=   -1  'True
      Tab(2).Control(20)=   "DesdeI"
      Tab(2).Control(20).Enabled=   -1  'True
      Tab(2).Control(21)=   "ControlII"
      Tab(2).Control(21).Enabled=   -1  'True
      Tab(2).Control(22)=   "ControlI"
      Tab(2).Control(22).Enabled=   -1  'True
      Tab(2).Control(23)=   "Label7"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "lblLabels(11)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "lblLabels(10)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Label6"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "lblLabels(9)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "lblLabels(8)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Label5"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "lblLabels(7)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "lblLabels(6)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "lblLabels(5)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "lblLabels(4)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "lblLabels(3)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "lblLabels(2)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "lblLabels(1)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "lblLabels(0)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Label4"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Label3"
      Tab(2).Control(39).Enabled=   0   'False
      Begin VB.ComboBox ControlV 
         Height          =   315
         Left            =   -72720
         TabIndex        =   57
         Top             =   3720
         Width           =   3255
      End
      Begin VB.TextBox DesdeV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68280
         TabIndex        =   56
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox HastaV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66360
         TabIndex        =   55
         Top             =   3720
         Width           =   735
      End
      Begin VB.ComboBox ControlIV 
         Height          =   315
         Left            =   -72720
         TabIndex        =   51
         Top             =   3360
         Width           =   3255
      End
      Begin VB.TextBox DesdeIV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68280
         TabIndex        =   50
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox HastaIV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66360
         TabIndex        =   49
         Top             =   3360
         Width           =   735
      End
      Begin VB.ComboBox ControlIII 
         Height          =   315
         Left            =   -72720
         TabIndex        =   45
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox DesdeIII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68280
         TabIndex        =   44
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox HastaIII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66360
         TabIndex        =   43
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox DesSeguridadIII 
         Height          =   285
         Left            =   -72720
         MaxLength       =   100
         TabIndex        =   36
         Top             =   1800
         Width           =   9015
      End
      Begin VB.TextBox DesSeguridadII 
         Height          =   285
         Left            =   -72720
         MaxLength       =   100
         TabIndex        =   35
         Top             =   1560
         Width           =   9015
      End
      Begin VB.TextBox Seguridad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73680
         MaxLength       =   4
         TabIndex        =   34
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox DesSeguridadI 
         Height          =   285
         Left            =   -72720
         MaxLength       =   100
         TabIndex        =   33
         Top             =   1320
         Width           =   9015
      End
      Begin VB.TextBox DesEquipoI 
         Height          =   285
         Left            =   -72720
         MaxLength       =   100
         TabIndex        =   31
         Top             =   480
         Width           =   9015
      End
      Begin VB.TextBox Equipo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73680
         MaxLength       =   4
         TabIndex        =   30
         Text            =   " "
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox DesEquipoII 
         Height          =   285
         Left            =   -72720
         MaxLength       =   100
         TabIndex        =   29
         Top             =   720
         Width           =   9015
      End
      Begin VB.TextBox DesEquipoIII 
         Height          =   285
         Left            =   -72720
         MaxLength       =   100
         TabIndex        =   28
         Top             =   960
         Width           =   9015
      End
      Begin VB.TextBox TiempoI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64440
         TabIndex        =   27
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox TiempoII 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68280
         TabIndex        =   26
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox HastaI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -66360
         TabIndex        =   25
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox DesdeI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -68280
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox ControlII 
         Height          =   315
         Left            =   -72720
         TabIndex        =   21
         Top             =   2520
         Width           =   3255
      End
      Begin VB.ComboBox ControlI 
         Height          =   315
         Left            =   -72720
         TabIndex        =   20
         Top             =   2160
         Width           =   3255
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
         Left            =   3840
         TabIndex        =   17
         Top             =   1740
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   1740
         Width           =   375
      End
      Begin VB.TextBox WTexto12 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   1740
         Width           =   375
      End
      Begin VB.TextBox WTexto1 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   -72600
         TabIndex        =   12
         Top             =   1680
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
         TabIndex        =   11
         Top             =   1620
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         Height          =   285
         Left            =   -72000
         TabIndex        =   10
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   1
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2220
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7011
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   -71280
         TabIndex        =   14
         Top             =   1740
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
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Top             =   1740
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
         Height          =   3975
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7011
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.Label Label7 
         Caption         =   "Controla ??"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   60
         Top             =   3765
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   11
         Left            =   -69120
         TabIndex        =   59
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   10
         Left            =   -67200
         TabIndex        =   58
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Controla ??"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   54
         Top             =   3405
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   9
         Left            =   -69120
         TabIndex        =   53
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   8
         Left            =   -67200
         TabIndex        =   52
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Controla Presion"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   3045
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   -69120
         TabIndex        =   47
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   -67200
         TabIndex        =   46
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tiempo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   -69120
         TabIndex        =   41
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Durante"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   -65400
         TabIndex        =   40
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   -67200
         TabIndex        =   39
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   -69120
         TabIndex        =   38
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Seguridad"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   37
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Equipo"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   32
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Controla Rampa"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   2550
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Controla Temperatura"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   2200
         Width           =   2175
      End
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   5880
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
      Left            =   2040
      TabIndex        =   4
      Top             =   6480
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
      ItemData        =   "CargaNueva.frx":0054
      Left            =   120
      List            =   "CargaNueva.frx":005B
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
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
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   1560
      TabIndex        =   67
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
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
   Begin VB.Label Label9 
      Caption         =   "Tipo de Proceso"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8280
      TabIndex        =   74
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Siguiente 
      Height          =   480
      Left            =   9840
      MouseIcon       =   "CargaNueva.frx":0069
      MousePointer    =   99  'Custom
      Picture         =   "CargaNueva.frx":0373
      ToolTipText     =   "Registro Posterior"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Anterior 
      Height          =   480
      Left            =   9120
      MouseIcon       =   "CargaNueva.frx":07B5
      MousePointer    =   99  'Custom
      Picture         =   "CargaNueva.frx":0ABF
      ToolTipText     =   "Registro Anterior"
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label11 
      Caption         =   "Fecha"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   71
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Version"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3240
      TabIndex        =   70
      Top             =   840
      Width           =   975
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9000
      TabIndex        =   69
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Responsable"
      Height          =   255
      Left            =   7320
      TabIndex        =   68
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label40 
      Caption         =   "Metodo Lavado"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5880
      TabIndex        =   65
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label DesMetodoFiltrado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3240
      TabIndex        =   63
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Metodo de Filtrado"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   62
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Etapa"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesTerminado 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   10320
      MouseIcon       =   "CargaNueva.frx":0F01
      MousePointer    =   99  'Custom
      Picture         =   "CargaNueva.frx":120B
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7440
      MouseIcon       =   "CargaNueva.frx":1A4D
      MousePointer    =   99  'Custom
      Picture         =   "CargaNueva.frx":1D57
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   9360
      MouseIcon       =   "CargaNueva.frx":2599
      MousePointer    =   99  'Custom
      Picture         =   "CargaNueva.frx":28A3
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   8400
      MouseIcon       =   "CargaNueva.frx":30E5
      MousePointer    =   99  'Custom
      Picture         =   "CargaNueva.frx":33EF
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Producto"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgCargaNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMetodoFiltrado As Recordset
Dim spMetodoFiltrado As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayos As Recordset
Dim spEnsayos As String
Dim rstProduI As Recordset
Dim rsProduI As String
Dim rstProduII As Recordset
Dim rsProduII As String
Dim rstProduIII As Recordset
Dim rsProduIII As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim XPaso As String
Dim Renglon As Integer
Dim ZEnsayo As String
Dim ZValor As String

Dim ZDatosI(100, 100) As String
Dim ZDatosII(100, 30) As String
Dim ZDatosIII(100, 30) As String

Private WGraba As String
Private WGrabaII As String

Dim CargaEmpresa(10, 2) As String

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

Private Sub AgregaRenglon_Click()

    Hasta = WVector1.Row

    For iRow = 100 To Hasta Step -1
        WVector1.TextMatrix(iRow, 0) = WVector1.TextMatrix(iRow - 1, 0)
        WVector1.TextMatrix(iRow, 1) = WVector1.TextMatrix(iRow - 1, 1)
    Next iRow

    WVector1.TextMatrix(Hasta, 0) = ""
    WVector1.TextMatrix(Hasta, 1) = ""
    
    WTexto1.Text = ""
    WTexto2.Text = ""

End Sub

Private Sub Base_Click()

    IngresaBase.Visible = True
    
    ProductoBase.Text = "  -     -   "
    ProductoBase.SetFocus

End Sub



Private Sub ProductoBase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo = " + "'" + ProductoBase.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZZTerminado = Terminado.Text
            ZZDesTerminado = DesTerminado.Caption
            Terminado.Text = ProductoBase.Text
            rstTerminado.Close
            Call Proceso
            Terminado.Text = ZZTerminado
            DesTerminado.Caption = ZZDesTerminado
            Fecha.Text = "  /  /    "
            Version.Text = ""
            Operador.Text = ""
            DesOperador.Caption = ""
        End If
        IngresaBase.Visible = False
    End If
    If KeyAscii = 27 Then
        ProductoBase.Text = "  -     -   "
        IngresaBase.Visible = False
    End If
End Sub



Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"
     Opcion.AddItem "Ensayos"
     Opcion.AddItem "Metodos de Filtrado"

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
        Case 0, 2
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Order by Codigo"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            
        Case 1
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
            
        Case 3
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 9, 10
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Order by Codigo"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                With rstEnsayos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnsayos!Codigo) + " " + rstEnsayos!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEnsayos!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayos.Close
            End If
    
            Call Conecta_Empresa
            
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
                            IngresaItem = Str$(rstMetodoFiltrado!Codigo) + " " + rstMetodoFiltrado!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstMetodoFiltrado!Codigo
                            WIndice.AddItem IngresaItem
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
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgCargaNueva.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    Call Pasa_Datos
        
    Sql1 = "DELETE ProduI"
    Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
    rsProduI = Sql1 + Sql2
    Set rstProduI = db.OpenRecordset(rsProduI, dbOpenSnapshot, dbSQLPassThrough)

    Sql1 = "DELETE ProduII"
    Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
    rsProduII = Sql1 + Sql2
    Set rstProduII = db.OpenRecordset(rsProduII, dbOpenSnapshot, dbSQLPassThrough)

    Sql1 = "DELETE ProduIII"
    Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
    rsProduIII = Sql1 + Sql2
    Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
    
    For Ciclo = 1 To 99
    
        ZPaso = Ciclo
    
        HastaRenglon = 0
        For iRow = 100 To 1 Step -1
            XDescripcion = ZDatosI(ZPaso, iRow)
            If XDescripcion <> "" Then
                HastaRenglon = iRow
                Exit For
            End If
        Next iRow
        
        If HastaRenglon > 0 Then
    
            WRenglon = 0
            For iRow = 1 To HastaRenglon
            
                XDescripcion = ZDatosI(ZPaso, iRow)
                    
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
                
                XPaso = ZPaso
                Call Ceros(XPaso, 4)
                                
                WClave = Terminado.Text + XPaso + Auxi
                
                ZVersion = Str$(Val(Version.Text) + 1)
                ZFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZAutorizado = "S"
                ZOrdFecha = Right$(XXFechaVersion, 4) + Mid$(XXFechaVersion, 4, 2) + Left$(XXFechaVersion, 2)
                ZTipo = Str$(Tipo.ListIndex)
                    
                ZSql = ""
                ZSql = ZSql + "INSERT INTO ProduI ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Etapa ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Instruccion ,"
                ZSql = ZSql + "MetodoFiltrado ,"
                ZSql = ZSql + "Version ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "MetodoLavado ,"
                ZSql = ZSql + "Operador ,"
                ZSql = ZSql + "Autorizado ,"
                ZSql = ZSql + "Tipo )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Terminado.Text + "',"
                ZSql = ZSql + "'" + XPaso + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + XDescripcion + "',"
                ZSql = ZSql + "'" + MetodoFiltrado.Text + "',"
                ZSql = ZSql + "'" + ZVersion + "',"
                ZSql = ZSql + "'" + ZFechaVersion + "',"
                ZSql = ZSql + "'" + ZOrdFecha + "',"
                ZSql = ZSql + "'" + MetodoLavado.Text + "',"
                ZSql = ZSql + "'" + Operador.Text + "',"
                ZSql = ZSql + "'" + ZAutorizado + "',"
                ZSql = ZSql + "'" + ZTipo + "')"
                    
                rsProduI = ZSql
                Set rstProduI = db.OpenRecordset(rsProduI, dbOpenSnapshot, dbSQLPassThrough)
                    
            Next iRow
            
            
            WRenglon = 0
            For iRow = 1 To 10
            
                Select Case iRow
                    Case 1
                        ZEnsayo = ZDatosII(ZPaso, 1)
                        ZValor = ZDatosII(ZPaso, 2)
                    
                    Case 2
                        ZEnsayo = ZDatosII(ZPaso, 3)
                        ZValor = ZDatosII(ZPaso, 4)
                    
                    Case 3
                        ZEnsayo = ZDatosII(ZPaso, 5)
                        ZValor = ZDatosII(ZPaso, 6)
                    
                    Case 4
                        ZEnsayo = ZDatosII(ZPaso, 7)
                        ZValor = ZDatosII(ZPaso, 8)
                    
                    Case 5
                        ZEnsayo = ZDatosII(ZPaso, 9)
                        ZValor = ZDatosII(ZPaso, 10)
                    
                    Case 6
                        ZEnsayo = ZDatosII(ZPaso, 11)
                        ZValor = ZDatosII(ZPaso, 12)
                    
                    Case 7
                        ZEnsayo = ZDatosII(ZPaso, 13)
                        ZValor = ZDatosII(ZPaso, 14)
                    
                    Case 8
                        ZEnsayo = ZDatosII(ZPaso, 15)
                        ZValor = ZDatosII(ZPaso, 16)
                    
                    Case 9
                        ZEnsayo = ZDatosII(ZPaso, 17)
                        ZValor = ZDatosII(ZPaso, 18)
                    
                    Case Else
                        ZEnsayo = ZDatosII(ZPaso, 19)
                        ZValor = ZDatosII(ZPaso, 20)
                        
                End Select
            
                If Val(ZEnsayo) <> 0 Then
                
                    WRenglon = WRenglon + 1
                    Auxi = Str$(WRenglon)
                    Call Ceros(Auxi, 2)
                
                    XPaso = ZPaso
                    Call Ceros(XPaso, 4)
                                
                    WClave = Terminado.Text + XPaso + Auxi
                    
                    ZDesEnsayo = ""
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Ensayos"
                    ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + ZEnsayo + "'"
                    spEnsayos = ZSql
                    Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnsayos.RecordCount > 0 Then
                        ZDesEnsayo = Trim(rstEnsayos!Descripcion)
                        rstEnsayos.Close
                    End If
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO ProduII ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "Terminado ,"
                    ZSql = ZSql + "Etapa ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Ensayo ,"
                    ZSql = ZSql + "DesEnsayo ,"
                    ZSql = ZSql + "Valor )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WClave + "',"
                    ZSql = ZSql + "'" + Terminado.Text + "',"
                    ZSql = ZSql + "'" + XPaso + "',"
                    ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                    ZSql = ZSql + "'" + ZEnsayo + "',"
                    ZSql = ZSql + "'" + ZDesEnsayo + "',"
                    ZSql = ZSql + "'" + ZValor + "')"
                    
                    rsProduII = ZSql
                    Set rstProduII = db.OpenRecordset(rsProduII, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                    
            Next iRow
            
            
            
            XPaso = ZPaso
            Call Ceros(XPaso, 4)
                                
            WClave = Terminado.Text + XPaso
            
            ZZEquipo = ZDatosIII(ZPaso, 1)
            ZZSeguridad = ZDatosIII(ZPaso, 2)
            ZZControlI = ZDatosIII(ZPaso, 3)
            ZZDesdeI = ZDatosIII(ZPaso, 4)
            ZZHastaI = ZDatosIII(ZPaso, 5)
            ZZTiempoI = ZDatosIII(ZPaso, 6)
            ZZControlII = ZDatosIII(ZPaso, 7)
            ZZTiempoII = ZDatosIII(ZPaso, 8)
            ZZDesEquipoI = ZDatosIII(ZPaso, 9)
            ZZDesEquipoII = ZDatosIII(ZPaso, 10)
            ZZDesEquipoIII = ZDatosIII(ZPaso, 11)
            ZZDesSeguridadI = ZDatosIII(ZPaso, 12)
            ZZDesSeguridadII = ZDatosIII(ZPaso, 13)
            ZZDesSeguridadIII = ZDatosIII(ZPaso, 14)
            ZZControlIII = ZDatosIII(ZPaso, 15)
            ZZDesdeIII = ZDatosIII(ZPaso, 16)
            ZZHastaIII = ZDatosIII(ZPaso, 17)
            ZZControlIV = ZDatosIII(ZPaso, 18)
            ZZDesdeIV = ZDatosIII(ZPaso, 19)
            ZZHastaIV = ZDatosIII(ZPaso, 20)
            ZZControlV = ZDatosIII(ZPaso, 21)
            ZZDesdeV = ZDatosIII(ZPaso, 22)
            ZZHastaV = ZDatosIII(ZPaso, 23)
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ProduIII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Etapa ,"
            ZSql = ZSql + "Equipo ,"
            ZSql = ZSql + "Seguridad ,"
            ZSql = ZSql + "ControlI ,"
            ZSql = ZSql + "TemperaturaI ,"
            ZSql = ZSql + "TemperaturaII ,"
            ZSql = ZSql + "Tiempo ,"
            ZSql = ZSql + "ControlII ,"
            ZSql = ZSql + "TiempoII ,"
            ZSql = ZSql + "DesEquipoI ,"
            ZSql = ZSql + "DesEquipoII ,"
            ZSql = ZSql + "DesEquipoIII ,"
            ZSql = ZSql + "DesSeguridadI ,"
            ZSql = ZSql + "DesSeguridadII ,"
            ZSql = ZSql + "DesSeguridadIII ,"
            ZSql = ZSql + "ControlIII ,"
            ZSql = ZSql + "DesdeIII ,"
            ZSql = ZSql + "HastaIII ,"
            ZSql = ZSql + "ControlIV ,"
            ZSql = ZSql + "DesdeIV ,"
            ZSql = ZSql + "HastaIV ,"
            ZSql = ZSql + "ControlV ,"
            ZSql = ZSql + "DesdeV ,"
            ZSql = ZSql + "HastaV )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Terminado.Text + "',"
            ZSql = ZSql + "'" + XPaso + "',"
            ZSql = ZSql + "'" + ZZEquipo + "',"
            ZSql = ZSql + "'" + ZZSeguridad + "',"
            ZSql = ZSql + "'" + ZZControlI + "',"
            ZSql = ZSql + "'" + ZZDesdeI + "',"
            ZSql = ZSql + "'" + ZZHastaI + "',"
            ZSql = ZSql + "'" + ZZTiempoI + "',"
            ZSql = ZSql + "'" + ZZControlII + "',"
            ZSql = ZSql + "'" + ZZTiempoII + "',"
            ZSql = ZSql + "'" + ZZDesEquipoI + "',"
            ZSql = ZSql + "'" + ZZDesEquipoII + "',"
            ZSql = ZSql + "'" + ZZDesEquipoIII + "',"
            ZSql = ZSql + "'" + ZZDesSeguridadI + "',"
            ZSql = ZSql + "'" + ZZDesSeguridadII + "',"
            ZSql = ZSql + "'" + ZZDesSeguridadIII + "',"
            ZSql = ZSql + "'" + ZZControlIII + "',"
            ZSql = ZSql + "'" + ZZDesdeIII + "',"
            ZSql = ZSql + "'" + ZZHastaIII + "',"
            ZSql = ZSql + "'" + ZZControlIV + "',"
            ZSql = ZSql + "'" + ZZDesdeIV + "',"
            ZSql = ZSql + "'" + ZZHastaIV + "',"
            ZSql = ZSql + "'" + ZZControlV + "',"
            ZSql = ZSql + "'" + ZZDesdeV + "',"
            ZSql = ZSql + "'" + ZZHastaV + "')"
                    
            rsProduIII = ZSql
            Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next Ciclo
    
    XEmpresa = WEmpresa
    Erase CargaEmpresa
    
    Select Case Val(WEmpresa)
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
        
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            XXVersion = Str$(Val(Version.Text) + 1)
            XXFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Terminado SET "
            ZSql = ZSql + " Metodo = " + "'" + MetodoLavado.Text + "',"
            ZSql = ZSql + " VersionI = " + "'" + XXVersion + "',"
            ZSql = ZSql + " FechaVersionI = " + "'" + XXFechaVersion + "',"
            ZSql = ZSql + " EstadoI = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Terminado.Text + "'"
                
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    Call Limpia_Click

    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1

    Tablas.Tab = 0
        
    Terminado.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector
    Call Limpia_VectorII
    
    Erase ZDatosI
    Erase ZDatosII
    Erase ZDatosIII
    
    Tablas.Tab = 0

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    MetodoFiltrado.Text = ""
    DesMetodoFiltrado.Caption = ""
    Paso.Text = ""
    MetodoLavado.Text = ""
    Operador.Text = ""
    DesOperador.Caption = ""
    Fecha.Text = "  /  /    "
    Version.Text = ""
    
    Equipo.Text = ""
    DesEquipoI.Text = ""
    DesEquipoII.Text = ""
    DesEquipoIII.Text = ""
    
    Seguridad.Text = ""
    DesSeguridadI.Text = ""
    DesSeguridadII.Text = ""
    DesSeguridadIII.Text = ""
    
    DesdeI.Text = ""
    HastaI.Text = ""
    TiempoI.Text = ""
    TiempoII.Text = ""
    
    DesdeIII.Text = ""
    HastaIII.Text = ""
    DesdeIV.Text = ""
    HastaIV.Text = ""
    DesdeV.Text = ""
    HastaV.Text = ""
    
    ControlI.ListIndex = 0
    ControlII.ListIndex = 0
    ControlIII.ListIndex = 0
    ControlIV.ListIndex = 0
    ControlV.ListIndex = 0
    Tipo.ListIndex = 0
    
    Renglon = 0
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1
    
    ControlI.ListIndex = 0
    ControlII.ListIndex = 0
    ControlIII.ListIndex = 0
    ControlIV.ListIndex = 0
    ControlV.ListIndex = 0
    
    Terminado.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Terminado.Text = WIndice.List(Indice)
            Call Terminado_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = Trim(rstArticulo!Codigo)
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstArticulo!Descripcion)
                End If
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case 2
            Indice = Pantalla.ListIndex
            WPTerminado = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + WPTerminado + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = Trim(rstTerminado!Codigo)
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstTerminado!Descripcion)
                End If
                WVector1.Col = 3
                rstTerminado.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case 3
            Indice = Pantalla.ListIndex
            WEnsayo = WIndice.List(Indice)
            
            WTexto12.Visible = False
            WTexto22.Visible = False
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WEnsayo + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                WVector2.Col = 1
                WVector2.Text = Trim(rstEnsayos!Codigo)
                WVector2.Col = 2
                WVector2.Text = Trim(rstEnsayos!Descripcion)
                WVector2.Col = 3
                rstEnsayos.Close
                Call StartEditII
            End If
            Rem Ayuda.Visible = False
            
        Case 4
            Indice = Pantalla.ListIndex
            MetodoFiltrado.Text = WIndice.List(Indice)
            Call MetodoFiltrado_KeyPress(13)
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    MetodoFiltrado.Text = ""
    DesMetodoFiltrado.Caption = ""
    Paso.Text = ""
    MetodoLavado.Text = ""
    Operador.Text = ""
    DesOperador.Caption = ""
    Fecha.Text = "  /  /    "
    Version.Text = ""
    
    ControlI.Clear
    
    ControlI.AddItem ""
    ControlI.AddItem "Controla"
    
    ControlI.ListIndex = 0
    
    ControlII.Clear
    
    ControlII.AddItem ""
    ControlII.AddItem "Controla"
    
    ControlII.ListIndex = 0
    
    ControlIII.Clear
    
    ControlIII.AddItem ""
    ControlIII.AddItem "Controla"
    
    ControlIII.ListIndex = 0
    
    ControlIV.Clear
    
    ControlIV.AddItem ""
    ControlIV.AddItem "Controla"
    
    ControlIV.ListIndex = 0
    
    ControlV.Clear
    
    ControlV.AddItem ""
    ControlV.AddItem "Controla"
    
    ControlV.ListIndex = 0
    
    
    Tipo.Clear
    
    Tipo.AddItem "Fabricacion"
    Tipo.AddItem "Control Calidad"
    Tipo.AddItem "Envasamiento"
    
    Tipo.ListIndex = 0
    
    Equipo.Text = ""
    DesEquipoI.Text = ""
    DesEquipoII.Text = ""
    DesEquipoIII.Text = ""
    
    Seguridad.Text = ""
    DesSeguridadI.Text = ""
    DesSeguridadII.Text = ""
    DesSeguridadIII.Text = ""
    
    DesdeI.Text = ""
    HastaI.Text = ""
    TiempoI.Text = ""
    TiempoII.Text = ""
    
    DesdeIII.Text = ""
    HastaIII.Text = ""
    DesdeIV.Text = ""
    HastaIV.Text = ""
    DesdeV.Text = ""
    HastaV.Text = ""
    
    ControlI.ListIndex = 0
    ControlII.ListIndex = 0
    ControlIII.ListIndex = 0
    ControlIV.ListIndex = 0
    ControlV.ListIndex = 0
    
    Renglon = 0
    
End Sub

Private Sub Proceso()
    
    Erase ZDatosI
    Erase ZDatosII
    Erase ZDatosIII
    ZAlta = "N"
    
    ZSql = " "
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ProduI"
    ZSql = ZSql + " Where ProduI.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by ProduI.Clave"
    
    rsProduI = ZSql
    Set rstProduI = db.OpenRecordset(rsProduI, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduI.RecordCount > 0 Then
        ZAlta = "S"
        MetodoFiltrado.Text = rstProduI!MetodoFiltrado
        MetodoLavado.Text = rstProduI!MetodoLavado
        Operador.Text = rstProduI!Operador
        Fecha.Text = rstProduI!Fecha
        Version.Text = rstProduI!Version
        rstProduI.Close
    End If
    
    
    ZPasa = 0
    ZSql = " "
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ProduI"
    ZSql = ZSql + " Where ProduI.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by ProduI.Clave"
    rsProduI = ZSql
    Set rstProduI = db.OpenRecordset(rsProduI, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduI.RecordCount > 0 Then
        With rstProduI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZPaso = Val(rstProduI!etapa)
                    
                    If ZPasa = 0 Then
                        WRenglon = 0
                        ZCorte = ZPaso
                        ZPasa = 1
                    End If
                    
                    If ZCorte <> ZPaso Then
                        WRenglon = 0
                        ZCorte = ZPaso
                    End If
                
                    WRenglon = WRenglon + 1
                    
                    ZDatosI(ZPaso, WRenglon) = Trim(rstProduI!instruccion)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProduI.Close
    End If
    
    
    
    ZPasa = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ProduII"
    ZSql = ZSql + " Where ProduII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by ProduII.Clave"
    rsProduII = ZSql
    Set rstProduII = db.OpenRecordset(rsProduII, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduII.RecordCount > 0 Then
        With rstProduII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZPaso = rstProduII!etapa
                    
                    If ZPasa = 0 Then
                        WRenglon = 0
                        ZCorte = ZPaso
                        ZPasa = 1
                    End If
                    
                    If ZCorte <> ZPaso Then
                        WRenglon = 0
                        ZCorte = ZPaso
                    End If
                
                    WRenglon = WRenglon + 1
                    Select Case WRenglon
                        Case 1
                            ZDatosII(ZPaso, 1) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 2) = Trim(rstProduII!Valor)
                        Case 2
                            ZDatosII(ZPaso, 3) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 4) = Trim(rstProduII!Valor)
                        Case 3
                            ZDatosII(ZPaso, 5) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 6) = Trim(rstProduII!Valor)
                        Case 4
                            ZDatosII(ZPaso, 7) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 8) = Trim(rstProduII!Valor)
                        Case 5
                            ZDatosII(ZPaso, 9) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 10) = Trim(rstProduII!Valor)
                        Case 6
                            ZDatosII(ZPaso, 11) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 12) = Trim(rstProduII!Valor)
                        Case 7
                            ZDatosII(ZPaso, 13) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 14) = Trim(rstProduII!Valor)
                        Case 8
                            ZDatosII(ZPaso, 15) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 16) = Trim(rstProduII!Valor)
                        Case 9
                            ZDatosII(ZPaso, 17) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 18) = Trim(rstProduII!Valor)
                        Case Else
                            ZDatosII(ZPaso, 19) = Trim(Str$(rstProduII!Ensayo))
                            ZDatosII(ZPaso, 20) = Trim(rstProduII!Valor)
                    End Select
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProduII.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ProduIII"
    ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by ProduIII.Clave"
    rsProduIII = ZSql
    Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduIII.RecordCount > 0 Then
        With rstProduIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZPaso = rstProduIII!etapa
                    
                    ZDatosIII(ZPaso, 1) = rstProduIII!Equipo
                    ZDatosIII(ZPaso, 2) = rstProduIII!Seguridad
                    ZDatosIII(ZPaso, 3) = rstProduIII!ControlI
                    ZDatosIII(ZPaso, 4) = rstProduIII!TemperaturaI
                    ZDatosIII(ZPaso, 5) = rstProduIII!TemperaturaII
                    ZDatosIII(ZPaso, 6) = rstProduIII!Tiempo
                    ZDatosIII(ZPaso, 7) = rstProduIII!ControlII
                    ZDatosIII(ZPaso, 8) = rstProduIII!TiempoII
                    ZDatosIII(ZPaso, 9) = Trim(rstProduIII!DesEquipoI)
                    ZDatosIII(ZPaso, 10) = Trim(rstProduIII!DesEquipoII)
                    ZDatosIII(ZPaso, 11) = Trim(rstProduIII!DesEquipoIII)
                    ZDatosIII(ZPaso, 12) = Trim(rstProduIII!DesSeguridadI)
                    ZDatosIII(ZPaso, 13) = Trim(rstProduIII!DesSeguridadII)
                    ZDatosIII(ZPaso, 14) = Trim(rstProduIII!DesSeguridadIII)
                    ZDatosIII(ZPaso, 15) = rstProduIII!ControlIII
                    ZDatosIII(ZPaso, 16) = rstProduIII!DesdeIII
                    ZDatosIII(ZPaso, 17) = rstProduIII!HastaIII
                    ZDatosIII(ZPaso, 18) = rstProduIII!ControlIV
                    ZDatosIII(ZPaso, 19) = rstProduIII!DesdeIV
                    ZDatosIII(ZPaso, 20) = rstProduIII!HastaIV
                    ZDatosIII(ZPaso, 21) = rstProduIII!ControlV
                    ZDatosIII(ZPaso, 22) = rstProduIII!DesdeV
                    ZDatosIII(ZPaso, 23) = rstProduIII!HastaV
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProduIII.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Terminado"
    Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = Sql1 + Sql2 + Sql3
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = Trim(rstTerminado!Descripcion)
        rstTerminado.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM MetodoFiltrado"
    Sql3 = " Where MetodoFiltrado.Codigo = " + "'" + MetodoFiltrado.Text + "'"
    spMetodoFiltrado = Sql1 + Sql2 + Sql3
    Set rstMetodoFiltrado = db.OpenRecordset(spMetodoFiltrado, dbOpenSnapshot, dbSQLPassThrough)
    If rstMetodoFiltrado.RecordCount > 0 Then
        DesMetodoFiltrado.Caption = Trim(rstMetodoFiltrado!Descripcion)
        rstMetodoFiltrado.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Operador"
    ZSql = ZSql + " Where Operador.Operador = " + "'" + Operador.Text + "'"
    spOperador = ZSql
    Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
    If rstOperador.RecordCount > 0 Then
        DesOperador.Caption = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
        rstOperador.Close
    End If
    
    Paso.Text = "1"
    If ZAlta = "S" Then
        Call Imprime_Proceso
            Else
        MetodoFiltrado.SetFocus
    End If

End Sub

Private Sub Revalida_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE ProduI SET "
    ZSql = ZSql + " Autorizado = " + "'" + "S" + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
                
    spProduI = ZSql
    Set rstProduI = db.OpenRecordset(spProduI, dbOpenSnapshot, dbSQLPassThrough)

    XEmpresa = WEmpresa
    Erase CargaEmpresa
    
    Select Case Val(WEmpresa)
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
        
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Terminado SET "
            ZSql = ZSql + " EstadoI = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Terminado.Text + "'"
                
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    Next Cicla
    
    Call Conecta_Empresa

    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
    
    Terminado.SetFocus

End Sub

Private Sub Siguiente_Click()
    If Val(Paso.Text) < 99 Then
        Paso.Text = Str$(Val(Paso.Text) + 1)
        Call Imprime_Proceso
    End If
End Sub

Private Sub Anterior_Click()
    If Val(Paso.Text) > 1 Then
        Call Pasa_Datos
        Paso.Text = Str$(Val(Paso.Text) - 1)
        Call Imprime_Proceso
    End If
End Sub

Private Sub Paso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Pasa_Datos
        Call Imprime_Proceso
    End If
    If KeyAscii = 27 Then
        Paso.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pasa_Datos()

    ZPaso = Val(Paso.Text)
    
    For Ciclo = 1 To 99
        ZDatosI(ZPaso, Ciclo) = WVector1.TextMatrix(Ciclo, 1)
    Next Ciclo
    
    For Ciclo = 1 To 10
    
        Select Case Ciclo
            Case 1
                ZDatosII(ZPaso, 1) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 2) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 2
                ZDatosII(ZPaso, 3) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 4) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 3
                ZDatosII(ZPaso, 5) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 6) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 4
                ZDatosII(ZPaso, 7) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 8) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 5
                ZDatosII(ZPaso, 9) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 10) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 6
                ZDatosII(ZPaso, 11) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 12) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 7
                ZDatosII(ZPaso, 13) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 14) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 8
                ZDatosII(ZPaso, 15) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 16) = WVector2.TextMatrix(Ciclo, 3)
            
            Case 9
                ZDatosII(ZPaso, 17) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 18) = WVector2.TextMatrix(Ciclo, 3)
            
            Case Else
                ZDatosII(ZPaso, 19) = WVector2.TextMatrix(Ciclo, 1)
                ZDatosII(ZPaso, 20) = WVector2.TextMatrix(Ciclo, 3)
                
        End Select
        
    Next Ciclo
          
    ZDatosIII(ZPaso, 1) = Equipo.Text
    ZDatosIII(ZPaso, 2) = Seguridad.Text
    ZDatosIII(ZPaso, 3) = ControlI.ListIndex
    ZDatosIII(ZPaso, 4) = DesdeI.Text
    ZDatosIII(ZPaso, 5) = HastaI.Text
    ZDatosIII(ZPaso, 6) = TiempoI.Text
    ZDatosIII(ZPaso, 7) = ControlII.ListIndex
    ZDatosIII(ZPaso, 8) = TiempoII.Text
    ZDatosIII(ZPaso, 9) = DesEquipoI.Text
    ZDatosIII(ZPaso, 10) = DesEquipoII.Text
    ZDatosIII(ZPaso, 11) = DesEquipoIII.Text
    ZDatosIII(ZPaso, 12) = DesSeguridadI.Text
    ZDatosIII(ZPaso, 13) = DesSeguridadII.Text
    ZDatosIII(ZPaso, 14) = DesSeguridadIII.Text
    ZDatosIII(ZPaso, 15) = ControlIII.ListIndex
    ZDatosIII(ZPaso, 16) = DesdeIII.Text
    ZDatosIII(ZPaso, 17) = HastaIII.Text
    ZDatosIII(ZPaso, 18) = ControlIV.ListIndex
    ZDatosIII(ZPaso, 19) = DesdeIV.Text
    ZDatosIII(ZPaso, 20) = HastaIV.Text
    ZDatosIII(ZPaso, 21) = ControlV.ListIndex
    ZDatosIII(ZPaso, 22) = DesdeV.Text
    ZDatosIII(ZPaso, 23) = HastaV.Text

End Sub

Private Sub Imprime_Proceso()

    ZPaso = Val(Paso.Text)
    
    For Ciclo = 1 To 99
        WVector1.TextMatrix(Ciclo, 1) = ZDatosI(ZPaso, Ciclo)
    Next Ciclo
    
    For Ciclo = 1 To 10
    
        Select Case Ciclo
            Case 1
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 1)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 2)
            
            Case 2
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 3)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 4)
            
            Case 3
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 5)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 6)
            
            Case 4
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 7)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 8)
            
            Case 5
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 9)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 10)
            
            Case 6
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 11)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 12)
            
            Case 7
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 13)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 14)
            
            Case 8
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 15)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 16)
            
            Case 9
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 17)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 18)
            
            Case Else
                WVector2.TextMatrix(Ciclo, 1) = ZDatosII(ZPaso, 19)
                WVector2.TextMatrix(Ciclo, 3) = ZDatosII(ZPaso, 20)
                
        End Select
        
    Next Ciclo
          
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 9, 10
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    For Ciclo = 1 To 10
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ensayos"
        ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector2.TextMatrix(Ciclo, 1) + "'"
        spEnsayos = ZSql
        Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayos.RecordCount > 0 Then
            WVector2.TextMatrix(Ciclo, 2) = Trim(rstEnsayos!Descripcion)
            rstEnsayos.Close
        End If
        
    Next Ciclo
    
    Call Conecta_Empresa
    
    If Val(ZDatosIII(ZPaso, 1)) <> 0 Then
        Equipo.Text = ZDatosIII(ZPaso, 1)
            Else
        Equipo.Text = ""
    End If
    If Val(ZDatosIII(ZPaso, 2)) <> 0 Then
        Seguridad.Text = ZDatosIII(ZPaso, 2)
            Else
        Seguridad.Text = ""
   End If
    
    ControlI.ListIndex = Val(ZDatosIII(ZPaso, 3))
    If Val(ZDatosIII(ZPaso, 4)) <> 0 Then
        DesdeI.Text = ZDatosIII(ZPaso, 4)
            Else
        DesdeI.Text = ""
    End If
    If Val(ZDatosIII(ZPaso, 5)) <> 0 Then
        HastaI.Text = ZDatosIII(ZPaso, 5)
            Else
        HastaI.Text = ""
    End If
    If Val(ZDatosIII(ZPaso, 6)) <> 0 Then
        TiempoI.Text = ZDatosIII(ZPaso, 6)
            Else
        TiempoI.Text = ""
    End If
    
    ControlII.ListIndex = Val(ZDatosIII(ZPaso, 7))
    If Val(ZDatosIII(ZPaso, 8)) <> 0 Then
        TiempoII.Text = ZDatosIII(ZPaso, 8)
            Else
        TiempoII.Text = ""
    End If
    
    DesEquipoI.Text = ZDatosIII(ZPaso, 9)
    DesEquipoII.Text = ZDatosIII(ZPaso, 10)
    DesEquipoIII.Text = ZDatosIII(ZPaso, 11)
    DesSeguridadI.Text = ZDatosIII(ZPaso, 12)
    DesSeguridadII.Text = ZDatosIII(ZPaso, 13)
    DesSeguridadIII.Text = ZDatosIII(ZPaso, 14)
    
    ControlIII.ListIndex = Val(ZDatosIII(ZPaso, 15))
    If Val(ZDatosIII(ZPaso, 16)) <> 0 Then
        DesdeIII.Text = ZDatosIII(ZPaso, 16)
            Else
        DesdeIII.Text = ""
    End If
    If Val(ZDatosIII(ZPaso, 17)) <> 0 Then
        HastaIII.Text = ZDatosIII(ZPaso, 17)
            Else
        HastaIII.Text = ""
    End If
    
    ControlIV.ListIndex = Val(ZDatosIII(ZPaso, 18))
    If Val(ZDatosIII(ZPaso, 19)) <> 0 Then
        DesdeIV.Text = ZDatosIII(ZPaso, 19)
            Else
        DesdeIV.Text = ""
    End If
    If Val(ZDatosIII(ZPaso, 20)) <> 0 Then
        HastaIV.Text = ZDatosIII(ZPaso, 20)
            Else
        HastaIV.Text = ""
    End If
    
    ControlV.ListIndex = Val(ZDatosIII(ZPaso, 21))
    If Val(ZDatosIII(ZPaso, 22)) <> 0 Then
        DesdeV.Text = ZDatosIII(ZPaso, 22)
            Else
        DesdeV.Text = ""
    End If
    If Val(ZDatosIII(ZPaso, 23)) <> 0 Then
        HastaV.Text = ZDatosIII(ZPaso, 23)
            Else
        HastaV.Text = ""
    End If
    
    Rem Tablas.Tab = 0
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1
    
    Select Case Tablas.Tab
        Case 0
            Call StartEdit
        Case 1
            Call StartEditII
        Case Else
            Equipo.SetFocus
    End Select

End Sub



Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = Trim(rstTerminado!Descripcion)
            rstTerminado.Close
            Call Proceso
                Else
            Terminado.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Terminado.Text = "  -     -   "
        DesTerminado.Caption = ""
    End If
End Sub


Private Sub MetodoFiltrado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM MetodoFiltrado"
        Sql3 = " Where MetodoFiltrado.Codigo = " + "'" + MetodoFiltrado.Text + "'"
        spMetodoFiltrado = Sql1 + Sql2 + Sql3
        Set rstMetodoFiltrado = db.OpenRecordset(spMetodoFiltrado, dbOpenSnapshot, dbSQLPassThrough)
        If rstMetodoFiltrado.RecordCount > 0 Then
            DesMetodoFiltrado.Caption = Trim(rstMetodoFiltrado!Descripcion)
            rstMetodoFiltrado.Close
            Tablas.Tab = 0
            WVector1.TopRow = 1
            WVector1.Col = 1
            WVector1.Row = 1
            WVector2.TopRow = 1
            WVector2.Col = 1
            WVector2.Row = 1
            Call StartEdit
                Else
            MetodoFiltrado.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        MetodoFiltrado.Text = "  -     -   "
        DesMetodoFiltrado.Caption = ""
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
            Sql2 = " FROM Terminado"
            Sql3 = " Order by Codigo"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstTerminado!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, aa, WEspacios) Then
                                    IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTerminado!Codigo
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
                rstTerminado.Close
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
                            da = Len(rstMetodoFiltrado!Descripcion) - WEspacios
                            For aa = 1 To da + 1
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

Private Sub Terminado_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Productos Terminados"
    Opcion.AddItem "Material Auxiliar a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub MetodoFiltrado_DblClick()

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"
     Opcion.AddItem "Ensayos"
     Opcion.AddItem "Metodos de Filtrado"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 4
    
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
        Case 1
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
        Case 99
            
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
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    If WVector1.Col = 2 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

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
    WVector1.Cols = 2
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
                WVector1.Text = "Intrucciones"
                WVector1.ColWidth(Ciclo) = 10000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 70
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
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
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
        Case 3
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
            Rem XEmpresa = WEmpresa
            Rem Select Case Val(WEmpresa)
            Rem     Case 1, 3, 5, 6, 7, 9,10
            Rem         WEmpresa = "0003"
            Rem         txtOdbc = "Empresa03"
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
            ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector2.Text + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                WVector2.Col = 2
                WVector2.Text = Trim(rstEnsayos!Descripcion)
                rstEnsayos.Close
                    Else
                WControlII = "N"
            End If
    
            Rem Call Conecta_Empresa
            
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
        For da = 0 To WVector2.Cols - 1
            WVector2.Col = da
            WVector2.Text = WBorraII(Ciclo, da)
        Next da
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
    WVector2.Cols = 4
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
                WVector2.Text = "Ensayos"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 4500
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Valor"
                WVector2.ColWidth(Ciclo) = 5000
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
















Private Sub Equipo_Keypress(KeyAscii As Integer)

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

Private Sub Seguridad_KeyPress(KeyAscii As Integer)

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
        Case 0
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        Case 1
            WVector2.Col = 1
            WVector2.Row = 1
            Call StartEditII
        Case 2
            Equipo.SetFocus
        Case 3
        Case Else
    End Select
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

