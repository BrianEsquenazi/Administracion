VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRevalida 
   BackColor       =   &H00C0C000&
   Caption         =   "Revalida de Fecha de Vencimiento de Materias Primas"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11880
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   4800
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label9 
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
         TabIndex        =   25
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Rechazo 
      Caption         =   "Rechazo"
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
      TabIndex        =   21
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox Revalida 
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
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   17
      Text            =   " "
      Top             =   120
      Width           =   975
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
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
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
      Left            =   6600
      TabIndex        =   15
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox Responsable 
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   5400
      Width           =   3255
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   5040
      Width           =   5295
   End
   Begin VB.TextBox Resultado 
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   1
      Text            =   " "
      Top             =   4680
      Width           =   5295
   End
   Begin VB.TextBox Lote 
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
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   10
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSMask.MaskEdBox Articulo 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   9600
      TabIndex        =   0
      Top             =   480
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox VtoAnterior 
      Height          =   285
      Left            =   9600
      TabIndex        =   19
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
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3300
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5821
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Especificaciones 1 - 10"
      TabPicture(0)   =   "prgrevalida.frx":0000
      Tab(0).ControlCount=   44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ValorOri1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ValorOri2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ValorOri3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ValorOri4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ValorOri5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ValorOri6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ValorOri7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ValorOri8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ValorOri9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ValorOri10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Descri1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Descri2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Descri3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Descri4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Descri5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Descri6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Descri7"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Descri8"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Descri9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Descri10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label90"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Std1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Std2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Std3"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Std4"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Std5"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Std6"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Std7"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Std8"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Std9"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Std10"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblresultado"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label92"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Valor1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Valor2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Valor3"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Valor4"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Valor5"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Valor6"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Valor7"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Valor8"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Valor9"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Valor10"
      Tab(0).Control(43).Enabled=   0   'False
      TabCaption(1)   =   "Especificaciones 11 - 20"
      TabPicture(1)   =   "prgrevalida.frx":001C
      Tab(1).ControlCount=   44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ValorOri11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ValorOri12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ValorOri13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ValorOri14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ValorOri15"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ValorOri16"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ValorOri17"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ValorOri18"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ValorOri19"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "ValorOri20"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label22"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Descri11"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Descri12"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Descri13"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Descri14"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Descri15"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Descri16"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Descri17"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Descri18"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Descri19"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Descri20"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label104"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Std11"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Std12"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Std13"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Std14"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Std15"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Std16"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Std17"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Std18"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Std19"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Std20"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "lblresultadoII"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label116"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Valor11"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Valor12"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Valor13"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Valor14"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Valor15"
      Tab(1).Control(38).Enabled=   -1  'True
      Tab(1).Control(39)=   "Valor16"
      Tab(1).Control(39).Enabled=   -1  'True
      Tab(1).Control(40)=   "Valor17"
      Tab(1).Control(40).Enabled=   -1  'True
      Tab(1).Control(41)=   "Valor18"
      Tab(1).Control(41).Enabled=   -1  'True
      Tab(1).Control(42)=   "Valor19"
      Tab(1).Control(42).Enabled=   -1  'True
      Tab(1).Control(43)=   "Valor20"
      Tab(1).Control(43).Enabled=   -1  'True
      TabCaption(2)   =   "Especificaciones 21 - 30"
      TabPicture(2)   =   "prgrevalida.frx":0038
      Tab(2).ControlCount=   44
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Valor30"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "Valor29"
      Tab(2).Control(1).Enabled=   -1  'True
      Tab(2).Control(2)=   "Valor28"
      Tab(2).Control(2).Enabled=   -1  'True
      Tab(2).Control(3)=   "Valor27"
      Tab(2).Control(3).Enabled=   -1  'True
      Tab(2).Control(4)=   "Valor26"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "Valor25"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "Valor24"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "Valor23"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "Valor22"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Valor21"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "Label45"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "ValorOri30"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "ValorOri29"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "ValorOri28"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "ValorOri27"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "ValorOri26"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "ValorOri25"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "ValorOri24"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "ValorOri23"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "ValorOri22"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "ValorOri21"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label29"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label30"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Std30"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Std29"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Std28"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Std27"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Std26"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Std25"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Std24"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Std23"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Std22"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Std21"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Label41"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Descri30"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Descri29"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Descri28"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Descri27"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Descri26"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Descri25"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Descri24"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Descri23"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Descri22"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Descri21"
      Tab(2).Control(43).Enabled=   0   'False
      Begin VB.TextBox Valor21 
         Height          =   285
         Left            =   -65220
         TabIndex        =   56
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor22 
         Height          =   285
         Left            =   -65220
         TabIndex        =   55
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor23 
         Height          =   285
         Left            =   -65220
         TabIndex        =   54
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor24 
         Height          =   285
         Left            =   -65220
         TabIndex        =   53
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor25 
         Height          =   285
         Left            =   -65220
         TabIndex        =   52
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor26 
         Height          =   285
         Left            =   -65220
         TabIndex        =   51
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor27 
         Height          =   285
         Left            =   -65220
         TabIndex        =   50
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor28 
         Height          =   285
         Left            =   -65220
         TabIndex        =   49
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor29 
         Height          =   285
         Left            =   -65220
         TabIndex        =   48
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor30 
         Height          =   285
         Left            =   -65220
         TabIndex        =   47
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor20 
         Height          =   285
         Left            =   -65220
         TabIndex        =   46
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor19 
         Height          =   285
         Left            =   -65220
         TabIndex        =   45
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor18 
         Height          =   285
         Left            =   -65220
         TabIndex        =   44
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor17 
         Height          =   285
         Left            =   -65220
         TabIndex        =   43
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor16 
         Height          =   285
         Left            =   -65220
         TabIndex        =   42
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor15 
         Height          =   285
         Left            =   -65220
         TabIndex        =   41
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor14 
         Height          =   285
         Left            =   -65220
         TabIndex        =   40
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor13 
         Height          =   285
         Left            =   -65220
         TabIndex        =   39
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor12 
         Height          =   285
         Left            =   -65220
         TabIndex        =   38
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor11 
         Height          =   285
         Left            =   -65220
         TabIndex        =   37
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor10 
         Height          =   285
         Left            =   9780
         TabIndex        =   36
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor9 
         Height          =   285
         Left            =   9780
         TabIndex        =   35
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor8 
         Height          =   285
         Left            =   9780
         TabIndex        =   34
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor7 
         Height          =   285
         Left            =   9780
         TabIndex        =   33
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor6 
         Height          =   285
         Left            =   9780
         TabIndex        =   32
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor5 
         Height          =   285
         Left            =   9780
         TabIndex        =   31
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor4 
         Height          =   285
         Left            =   9780
         TabIndex        =   30
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   9780
         TabIndex        =   29
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor2 
         Height          =   285
         Left            =   9780
         TabIndex        =   28
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   9780
         TabIndex        =   27
         Top             =   720
         Width           =   1700
      End
      Begin VB.Label Descri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   158
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Descri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   157
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   156
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   155
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   154
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   153
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   152
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   151
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   150
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   149
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   -74880
         TabIndex        =   148
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   147
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Std22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   146
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   145
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   144
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   143
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   142
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   141
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   140
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   139
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   138
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -71650
         TabIndex        =   137
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Registrado"
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
         Left            =   -65220
         TabIndex        =   136
         Top             =   480
         Width           =   1700
      End
      Begin VB.Label Label116 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Registrado"
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
         Left            =   -65220
         TabIndex        =   135
         Top             =   480
         Width           =   1700
      End
      Begin VB.Label lblresultadoII 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -71650
         TabIndex        =   134
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   133
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Std19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   132
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   131
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   130
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   129
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   128
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   127
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   126
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   125
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   124
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Label104 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   -74880
         TabIndex        =   123
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Descri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   122
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Descri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   121
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   120
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   119
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   118
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   117
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   116
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   115
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   114
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   113
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label Label92 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Registrado"
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
         Left            =   9780
         TabIndex        =   112
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   3350
         TabIndex        =   111
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   110
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Std9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   109
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   108
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   107
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   106
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   105
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   104
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   103
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   102
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   101
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Label90 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   120
         TabIndex        =   100
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   6550
         TabIndex        =   89
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label ValorOri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   88
         Top             =   2880
         Width           =   3195
      End
      Begin VB.Label ValorOri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   87
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label ValorOri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   86
         Top             =   2400
         Width           =   3195
      End
      Begin VB.Label ValorOri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   85
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label ValorOri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   84
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label ValorOri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   83
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Label ValorOri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   82
         Top             =   1440
         Width           =   3195
      End
      Begin VB.Label ValorOri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   81
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label ValorOri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   80
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label ValorOri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   79
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -68450
         TabIndex        =   78
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label ValorOri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   77
         Top             =   2880
         Width           =   3195
      End
      Begin VB.Label ValorOri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   76
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label ValorOri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   75
         Top             =   2400
         Width           =   3195
      End
      Begin VB.Label ValorOri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   74
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label ValorOri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   73
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label ValorOri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   72
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Label ValorOri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   71
         Top             =   1440
         Width           =   3195
      End
      Begin VB.Label ValorOri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   70
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label ValorOri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   69
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label ValorOri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   68
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label ValorOri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   67
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label ValorOri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   66
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label ValorOri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   65
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label ValorOri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   64
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label ValorOri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   63
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label ValorOri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   62
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label ValorOri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   61
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label ValorOri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   60
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label ValorOri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   59
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label ValorOri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   58
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -68450
         TabIndex        =   57
         Top             =   480
         Width           =   3200
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vencimiento"
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
      Left            =   7560
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Revalida"
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
      Left            =   4800
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Vencimiento"
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
      Left            =   7560
      TabIndex        =   14
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Responsable"
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
      TabIndex        =   13
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado"
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
      TabIndex        =   11
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label DesArticulo 
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
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo de M.P."
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
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lote"
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
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgRevalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WGraba As String
Dim ZEnsayo(30) As Integer
Dim ZEnsayoActual(30) As Integer
Dim ZEnsayoActualGraba(30) As Integer
Dim ZValor(30) As String
Dim WRevalida As Integer
Dim WProceso As Integer

Private Sub Cancela_click()
    PrgRevalida.Hide
    Unload Me
    If Val(ZProgramaOrigen) = 0 Then
        PrgPruart.Show
            Else
        PrgVerificaLoteArti.Show
    End If
End Sub

Private Sub Form_Activate()
    If Val(WEmpresaRevalida) <> 0 Then
        XEmpresa = WEmpresaRevalida
        Call Conecta_Empresa
    End If
End Sub

Private Sub Form_Load()
    
    If Val(WEmpresaRevalida) <> 0 Then
        XEmpresa = WEmpresaRevalida
        Call Conecta_Empresa
    End If

    Lote.Text = ZLoteRevalida
    Fecha.Text = ZFechaRevalida
    Articulo.Text = ZArticuloRevalida
    DesArticulo.Caption = ZDesArticuloRevalida
    WGraba = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Prueart"
    ZSql = ZSql + " Where Lote = " + "'" + Lote.Text + "'"
    spPrueart = ZSql
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueart.RecordCount > 0 Then
        ZValor(1) = rstPrueart!Valor1
        ZValor(2) = rstPrueart!Valor2
        ZValor(3) = rstPrueart!Valor3
        ZValor(4) = rstPrueart!Valor4
        ZValor(5) = rstPrueart!Valor5
        ZValor(6) = rstPrueart!Valor6
        ZValor(7) = rstPrueart!Valor7
        ZValor(8) = rstPrueart!Valor8
        ZValor(9) = rstPrueart!Valor9
        ZValor(10) = rstPrueart!Valor10
        ZValor(11) = IIf(IsNull(rstPrueart!Valor11), "", rstPrueart!Valor11)
        ZValor(12) = IIf(IsNull(rstPrueart!Valor12), "", rstPrueart!Valor12)
        ZValor(13) = IIf(IsNull(rstPrueart!Valor13), "", rstPrueart!Valor13)
        ZValor(14) = IIf(IsNull(rstPrueart!Valor14), "", rstPrueart!Valor14)
        ZValor(15) = IIf(IsNull(rstPrueart!Valor15), "", rstPrueart!Valor15)
        ZValor(16) = IIf(IsNull(rstPrueart!Valor16), "", rstPrueart!Valor16)
        ZValor(17) = IIf(IsNull(rstPrueart!Valor17), "", rstPrueart!Valor17)
        ZValor(18) = IIf(IsNull(rstPrueart!Valor18), "", rstPrueart!Valor18)
        ZValor(19) = IIf(IsNull(rstPrueart!Valor19), "", rstPrueart!Valor19)
        ZValor(20) = IIf(IsNull(rstPrueart!Valor20), "", rstPrueart!Valor20)
        ZValor(21) = IIf(IsNull(rstPrueart!Valor21), "", rstPrueart!Valor21)
        ZValor(22) = IIf(IsNull(rstPrueart!Valor22), "", rstPrueart!Valor22)
        ZValor(23) = IIf(IsNull(rstPrueart!Valor23), "", rstPrueart!Valor23)
        ZValor(24) = IIf(IsNull(rstPrueart!Valor24), "", rstPrueart!Valor24)
        ZValor(25) = IIf(IsNull(rstPrueart!Valor25), "", rstPrueart!Valor25)
        ZValor(26) = IIf(IsNull(rstPrueart!Valor26), "", rstPrueart!Valor26)
        ZValor(27) = IIf(IsNull(rstPrueart!Valor27), "", rstPrueart!Valor27)
        ZValor(28) = IIf(IsNull(rstPrueart!Valor28), "", rstPrueart!Valor28)
        ZValor(29) = IIf(IsNull(rstPrueart!Valor29), "", rstPrueart!Valor29)
        ZValor(30) = IIf(IsNull(rstPrueart!Valor30), "", rstPrueart!Valor30)
        rstPrueart.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo = " + "'" + Lote.Text + "'"
    ZSql = ZSql + " and Articulo = " + "'" + Articulo.Text + "'"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        WVto = IIf(IsNull(rstLaudo!fechavencimiento), "  /  /    ", rstLaudo!fechavencimiento)
        VtoAnterior.Text = WVto
        WRevalida = IIf(IsNull(rstLaudo!Revalida), "0", rstLaudo!Revalida)
        Revalida.Text = Str$(WRevalida + 1)
        WFecha = rstLaudo!Fecha
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        rstLaudo.Close
    End If
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
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
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Articulo.Text + "'"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        ZEnsayoActual(1) = rstEspecificacionesUnifica!Ensayo1
        ZEnsayoActual(2) = rstEspecificacionesUnifica!Ensayo2
        ZEnsayoActual(3) = rstEspecificacionesUnifica!Ensayo3
        ZEnsayoActual(4) = rstEspecificacionesUnifica!Ensayo4
        ZEnsayoActual(5) = rstEspecificacionesUnifica!Ensayo5
        ZEnsayoActual(6) = rstEspecificacionesUnifica!Ensayo6
        ZEnsayoActual(7) = rstEspecificacionesUnifica!Ensayo7
        ZEnsayoActual(8) = rstEspecificacionesUnifica!Ensayo8
        ZEnsayoActual(9) = rstEspecificacionesUnifica!Ensayo9
        ZEnsayoActual(10) = rstEspecificacionesUnifica!Ensayo10
        ZEnsayoActual(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "0", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayoActual(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "0", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayoActual(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "0", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayoActual(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "0", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayoActual(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "0", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayoActual(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "0", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayoActual(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "0", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayoActual(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "0", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayoActual(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "0", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayoActual(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "0", rstEspecificacionesUnifica!Ensayo20)
        
        
        ZEnsayoActualGraba(1) = rstEspecificacionesUnifica!Ensayo1
        ZEnsayoActualGraba(2) = rstEspecificacionesUnifica!Ensayo2
        ZEnsayoActualGraba(3) = rstEspecificacionesUnifica!Ensayo3
        ZEnsayoActualGraba(4) = rstEspecificacionesUnifica!Ensayo4
        ZEnsayoActualGraba(5) = rstEspecificacionesUnifica!Ensayo5
        ZEnsayoActualGraba(6) = rstEspecificacionesUnifica!Ensayo6
        ZEnsayoActualGraba(7) = rstEspecificacionesUnifica!Ensayo7
        ZEnsayoActualGraba(8) = rstEspecificacionesUnifica!Ensayo8
        ZEnsayoActualGraba(9) = rstEspecificacionesUnifica!Ensayo9
        ZEnsayoActualGraba(10) = rstEspecificacionesUnifica!Ensayo10
        ZEnsayoActualGraba(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "0", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayoActualGraba(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "0", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayoActualGraba(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "0", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayoActualGraba(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "0", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayoActualGraba(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "0", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayoActualGraba(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "0", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayoActualGraba(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "0", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayoActualGraba(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "0", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayoActualGraba(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "0", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayoActualGraba(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "0", rstEspecificacionesUnifica!Ensayo20)
        
        Std1.Caption = rstEspecificacionesUnifica!Valor1
        Std2.Caption = rstEspecificacionesUnifica!Valor2
        Std3.Caption = rstEspecificacionesUnifica!Valor3
        Std4.Caption = rstEspecificacionesUnifica!Valor4
        Std5.Caption = rstEspecificacionesUnifica!Valor5
        Std6.Caption = rstEspecificacionesUnifica!Valor6
        Std7.Caption = rstEspecificacionesUnifica!Valor7
        Std8.Caption = rstEspecificacionesUnifica!Valor8
        Std9.Caption = rstEspecificacionesUnifica!Valor9
        Std10.Caption = rstEspecificacionesUnifica!Valor10
        Std10.Caption = rstEspecificacionesUnifica!Valor10
        Std11.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
        Std12.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
        Std13.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
        Std14.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
        Std15.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
        Std16.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
        Std17.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
        Std18.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
        Std19.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
        Std20.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
        
        rstEspecificacionesUnifica.Close
    End If
    
    

    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaIII"
    Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Articulo.Text + "'"
    spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
        ZEnsayoActualGraba(11) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "0", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayoActualGraba(11) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "0", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayoActualGraba(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "0", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayoActualGraba(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "0", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayoActualGraba(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "0", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayoActualGraba(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "0", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayoActualGraba(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "0", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayoActualGraba(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "0", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayoActualGraba(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "0", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayoActualGraba(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "0", rstEspecificacionesUnificaIII!Ensayo30)
    
        ZEnsayoActual(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "0", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayoActual(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "0", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayoActual(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "0", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayoActual(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "0", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayoActual(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "0", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayoActual(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "0", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayoActual(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "0", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayoActual(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "0", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayoActual(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "0", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayoActual(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "0", rstEspecificacionesUnificaIII!Ensayo30)
    
        ZEnsayoActualGraba(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "0", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayoActualGraba(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "0", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayoActualGraba(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "0", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayoActualGraba(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "0", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayoActualGraba(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "0", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayoActualGraba(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "0", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayoActualGraba(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "0", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayoActualGraba(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "0", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayoActualGraba(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "0", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayoActualGraba(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "0", rstEspecificacionesUnificaIII!Ensayo30)
        
        Std21.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor21), "", rstEspecificacionesUnificaIII!Valor21)
        Std22.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor22), "", rstEspecificacionesUnificaIII!Valor22)
        Std23.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor23), "", rstEspecificacionesUnificaIII!Valor23)
        Std24.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor24), "", rstEspecificacionesUnificaIII!Valor24)
        Std25.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor25), "", rstEspecificacionesUnificaIII!Valor25)
        Std26.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor26), "", rstEspecificacionesUnificaIII!Valor26)
        Std27.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor27), "", rstEspecificacionesUnificaIII!Valor27)
        Std28.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor28), "", rstEspecificacionesUnificaIII!Valor28)
        Std29.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor29), "", rstEspecificacionesUnificaIII!Valor29)
        Std30.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor30), "", rstEspecificacionesUnificaIII!Valor30)
        
        rstEspecificacionesUnificaIII.Close
    End If
    
    
    
    
    
    
    
    
    
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(1)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(1)
        Call Ceros(Auxi1, 4)
        Descri1.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(2)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(2)
        Call Ceros(Auxi1, 4)
        Descri2.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(3)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(3)
        Call Ceros(Auxi1, 4)
        Descri3.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(4)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(4)
        Call Ceros(Auxi1, 4)
        Descri4.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(5)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(5)
        Call Ceros(Auxi1, 4)
        Descri5.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(6)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(6)
        Call Ceros(Auxi1, 4)
        Descri6.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(7)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(7)
        Call Ceros(Auxi1, 4)
        Descri7.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(8)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(8)
        Call Ceros(Auxi1, 4)
        Descri8.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(9)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(9)
        Call Ceros(Auxi1, 4)
        Descri9.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(10)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(10)
        Call Ceros(Auxi1, 4)
        Descri10.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(11)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(11)
        Call Ceros(Auxi1, 4)
        Descri11.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri11.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(12)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(12)
        Call Ceros(Auxi1, 4)
        Descri12.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri12.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(13)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(13)
        Call Ceros(Auxi1, 4)
        Descri13.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri13.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(14)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(14)
        Call Ceros(Auxi1, 4)
        Descri14.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri14.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(15)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(15)
        Call Ceros(Auxi1, 4)
        Descri15.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri15.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(16)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(16)
        Call Ceros(Auxi1, 4)
        Descri16.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri16.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(17)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(17)
        Call Ceros(Auxi1, 4)
        Descri17.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri17.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(18)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(18)
        Call Ceros(Auxi1, 4)
        Descri18.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri18.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(19)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(19)
        Call Ceros(Auxi1, 4)
        Descri19.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri19.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(20)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(20)
        Call Ceros(Auxi1, 4)
        Descri20.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri20.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(21)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(21)
        Call Ceros(Auxi1, 4)
        Descri21.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri21.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(22)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(22)
        Call Ceros(Auxi1, 4)
        Descri22.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri22.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(23)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(23)
        Call Ceros(Auxi1, 4)
        Descri23.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri23.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(24)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(24)
        Call Ceros(Auxi1, 4)
        Descri24.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri24.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(25)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(25)
        Call Ceros(Auxi1, 4)
        Descri25.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri25.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(26)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(26)
        Call Ceros(Auxi1, 4)
        Descri26.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri26.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(27)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(27)
        Call Ceros(Auxi1, 4)
        Descri27.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri27.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(28)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(28)
        Call Ceros(Auxi1, 4)
        Descri28.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri28.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(29)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(29)
        Call Ceros(Auxi1, 4)
        Descri29.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri29.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(30)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Auxi1 = ZEnsayoActual(30)
        Call Ceros(Auxi1, 4)
        Descri30.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri30.Caption = ""
    End If
    

    If Val(ZEnsayoActual(1)) = 0 Then
        Descri1.Caption = ""
    End If
    If Val(ZEnsayoActual(2)) = 0 Then
        Descri2.Caption = ""
    End If
    If Val(ZEnsayoActual(3)) = 0 Then
        Descri3.Caption = ""
    End If
    If Val(ZEnsayoActual(4)) = 0 Then
        Descri4.Caption = ""
    End If
    If Val(ZEnsayoActual(5)) = 0 Then
        Descri5.Caption = ""
    End If
    If Val(ZEnsayoActual(6)) = 0 Then
        Descri6.Caption = ""
    End If
    If Val(ZEnsayoActual(7)) = 0 Then
        Descri7.Caption = ""
    End If
    If Val(ZEnsayoActual(8)) = 0 Then
        Descri8.Caption = ""
    End If
    If Val(ZEnsayoActual(9)) = 0 Then
        Descri9.Caption = ""
    End If
    If Val(ZEnsayoActual(10)) = 0 Then
        Descri10.Caption = ""
    End If
    If Val(ZEnsayoActual(11)) = 0 Then
        Descri11.Caption = ""
    End If
    If Val(ZEnsayoActual(12)) = 0 Then
        Descri12.Caption = ""
    End If
    If Val(ZEnsayoActual(13)) = 0 Then
        Descri13.Caption = ""
    End If
    If Val(ZEnsayoActual(14)) = 0 Then
        Descri14.Caption = ""
    End If
    If Val(ZEnsayoActual(15)) = 0 Then
        Descri15.Caption = ""
    End If
    If Val(ZEnsayoActual(16)) = 0 Then
        Descri16.Caption = ""
    End If
    If Val(ZEnsayoActual(17)) = 0 Then
        Descri17.Caption = ""
    End If
    If Val(ZEnsayoActual(18)) = 0 Then
        Descri18.Caption = ""
    End If
    If Val(ZEnsayoActual(19)) = 0 Then
        Descri19.Caption = ""
    End If
    If Val(ZEnsayoActual(20)) = 0 Then
        Descri20.Caption = ""
    End If
    If Val(ZEnsayoActual(21)) = 0 Then
        Descri21.Caption = ""
    End If
    If Val(ZEnsayoActual(22)) = 0 Then
        Descri22.Caption = ""
    End If
    If Val(ZEnsayoActual(23)) = 0 Then
        Descri23.Caption = ""
    End If
    If Val(ZEnsayoActual(24)) = 0 Then
        Descri24.Caption = ""
    End If
    If Val(ZEnsayoActual(25)) = 0 Then
        Descri25.Caption = ""
    End If
    If Val(ZEnsayoActual(26)) = 0 Then
        Descri26.Caption = ""
    End If
    If Val(ZEnsayoActual(27)) = 0 Then
        Descri27.Caption = ""
    End If
    If Val(ZEnsayoActual(28)) = 0 Then
        Descri28.Caption = ""
    End If
    If Val(ZEnsayoActual(29)) = 0 Then
        Descri29.Caption = ""
    End If
    If Val(ZEnsayoActual(30)) = 0 Then
        Descri30.Caption = ""
    End If
    
    
    
    
    
    
    
    
    LlamaImprime = "N"
    
    Erase ZEnsayo
                
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecificacionesUnificaVersion"
    ZSql = ZSql + " Where EspecificacionesUnificaVersion.Producto = " + "'" + Articulo.Text + "'"
    ZSql = ZSql + " Order by EspecificacionesUnificaVersion.Producto, EspecificacionesUnificaVersion.Version"
                
    spEspecificacionesUnificaVersion = ZSql
    Set rstEspecificacionesUnificaVersion = db.OpenRecordset(spEspecificacionesUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaVersion.RecordCount > 0 Then
        With rstEspecificacionesUnificaVersion
            .MoveFirst
            Do
                If .EOF = False Then
                            
                    WDesde = Right$(rstEspecificacionesUnificaVersion!FechaInicio, 4) + Mid$(rstEspecificacionesUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecificacionesUnificaVersion!FechaInicio, 2)
                    WHasta = Right$(rstEspecificacionesUnificaVersion!FechaFinal, 4) + Mid$(rstEspecificacionesUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecificacionesUnificaVersion!FechaFinal, 2)
                                
                    If WDesde <= WFechaord And WHasta >= WFechaord Then
                        ZEnsayo(1) = rstEspecificacionesUnificaVersion!Ensayo1
                        ZEnsayo(2) = rstEspecificacionesUnificaVersion!Ensayo2
                        ZEnsayo(3) = rstEspecificacionesUnificaVersion!Ensayo3
                        ZEnsayo(4) = rstEspecificacionesUnificaVersion!Ensayo4
                        ZEnsayo(5) = rstEspecificacionesUnificaVersion!Ensayo5
                        ZEnsayo(6) = rstEspecificacionesUnificaVersion!Ensayo6
                        ZEnsayo(7) = rstEspecificacionesUnificaVersion!Ensayo7
                        ZEnsayo(8) = rstEspecificacionesUnificaVersion!Ensayo8
                        ZEnsayo(9) = rstEspecificacionesUnificaVersion!Ensayo9
                        ZEnsayo(10) = rstEspecificacionesUnificaVersion!Ensayo10
                        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "0", rstEspecificacionesUnificaVersion!Ensayo11)
                        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "0", rstEspecificacionesUnificaVersion!Ensayo12)
                        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "0", rstEspecificacionesUnificaVersion!Ensayo13)
                        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "0", rstEspecificacionesUnificaVersion!Ensayo14)
                        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "0", rstEspecificacionesUnificaVersion!Ensayo15)
                        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "0", rstEspecificacionesUnificaVersion!Ensayo16)
                        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "0", rstEspecificacionesUnificaVersion!Ensayo17)
                        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "0", rstEspecificacionesUnificaVersion!Ensayo18)
                        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "0", rstEspecificacionesUnificaVersion!Ensayo19)
                        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "0", rstEspecificacionesUnificaVersion!Ensayo20)
                        ZZClaveVersion = rstEspecificacionesUnificaVersion!Clave
                        LlamaImprime = "S"
                    End If
                                
                    If WDesde > WFechaord And LlamaImprime = "N" Then
                        ZEnsayo(1) = rstEspecificacionesUnificaVersion!Ensayo1
                        ZEnsayo(2) = rstEspecificacionesUnificaVersion!Ensayo2
                        ZEnsayo(3) = rstEspecificacionesUnificaVersion!Ensayo3
                        ZEnsayo(4) = rstEspecificacionesUnificaVersion!Ensayo4
                        ZEnsayo(5) = rstEspecificacionesUnificaVersion!Ensayo5
                        ZEnsayo(6) = rstEspecificacionesUnificaVersion!Ensayo6
                        ZEnsayo(7) = rstEspecificacionesUnificaVersion!Ensayo7
                        ZEnsayo(8) = rstEspecificacionesUnificaVersion!Ensayo8
                        ZEnsayo(9) = rstEspecificacionesUnificaVersion!Ensayo9
                        ZEnsayo(10) = rstEspecificacionesUnificaVersion!Ensayo10
                        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "0", rstEspecificacionesUnificaVersion!Ensayo11)
                        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "0", rstEspecificacionesUnificaVersion!Ensayo12)
                        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "0", rstEspecificacionesUnificaVersion!Ensayo13)
                        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "0", rstEspecificacionesUnificaVersion!Ensayo14)
                        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "0", rstEspecificacionesUnificaVersion!Ensayo15)
                        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "0", rstEspecificacionesUnificaVersion!Ensayo16)
                        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "0", rstEspecificacionesUnificaVersion!Ensayo17)
                        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "0", rstEspecificacionesUnificaVersion!Ensayo18)
                        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "0", rstEspecificacionesUnificaVersion!Ensayo19)
                        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "0", rstEspecificacionesUnificaVersion!Ensayo20)
                        ZZClaveVersion = rstEspecificacionesUnificaVersion!Clave
                        LlamaImprime = "S"
                    End If
                                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecificacionesUnificaVersion.Close
    End If
                
    If LlamaImprime = "N" Then
                
        Sql1 = "Select *"
        Sql2 = " FROM EspecificacionesUnifica"
        Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Articulo.Text + "'"
        spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnifica.RecordCount > 0 Then
            ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
            ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
            ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
            ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
            ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
            ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
            ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
            ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
            ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
            ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
            ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "0", rstEspecificacionesUnifica!Ensayo11)
            ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "0", rstEspecificacionesUnifica!Ensayo12)
            ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "0", rstEspecificacionesUnifica!Ensayo13)
            ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "0", rstEspecificacionesUnifica!Ensayo14)
            ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "0", rstEspecificacionesUnifica!Ensayo15)
            ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "0", rstEspecificacionesUnifica!Ensayo16)
            ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "0", rstEspecificacionesUnifica!Ensayo17)
            ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "0", rstEspecificacionesUnifica!Ensayo18)
            ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "0", rstEspecificacionesUnifica!Ensayo19)
            ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "0", rstEspecificacionesUnifica!Ensayo20)
            rstEspecificacionesUnifica.Close
            LlamaImprime = "S"
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM EspecificacionesUnificaIII"
        Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Articulo.Text + "'"
        spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
        Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnificaIII.RecordCount > 0 Then
        
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "0", rstEspecificacionesUnificaIII!Ensayo21)
            ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "0", rstEspecificacionesUnificaIII!Ensayo22)
            ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "0", rstEspecificacionesUnificaIII!Ensayo23)
            ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "0", rstEspecificacionesUnificaIII!Ensayo24)
            ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "0", rstEspecificacionesUnificaIII!Ensayo25)
            ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "0", rstEspecificacionesUnificaIII!Ensayo26)
            ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "0", rstEspecificacionesUnificaIII!Ensayo27)
            ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "0", rstEspecificacionesUnificaIII!Ensayo28)
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "0", rstEspecificacionesUnificaIII!Ensayo29)
            ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "0", rstEspecificacionesUnificaIII!Ensayo30)
            
            rstEspecificacionesUnificaIII.Close
        End If
                        
            Else
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecificacionesUnificaVersionII"
        ZSql = ZSql + " Where EspecificacionesUnificaVersionII.Clave = " + "'" + ZZClaveVersion + "'"
        spEspecificacionesUnificaVersionII = ZSql
        Set rstEspecificacionesUnificaVersionII = db.OpenRecordset(spEspecificacionesUnificaVersionII, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnificaVersionII.RecordCount > 0 Then
        
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo21), "", rstEspecificacionesUnificaVersionII!Ensayo21)
            ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo22), "", rstEspecificacionesUnificaVersionII!Ensayo22)
            ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo23), "", rstEspecificacionesUnificaVersionII!Ensayo23)
            ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo24), "", rstEspecificacionesUnificaVersionII!Ensayo24)
            ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo25), "", rstEspecificacionesUnificaVersionII!Ensayo25)
            ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo26), "", rstEspecificacionesUnificaVersionII!Ensayo26)
            ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo27), "", rstEspecificacionesUnificaVersionII!Ensayo27)
            ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo28), "", rstEspecificacionesUnificaVersionII!Ensayo28)
            ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo29), "", rstEspecificacionesUnificaVersionII!Ensayo29)
            ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo30), "", rstEspecificacionesUnificaVersionII!Ensayo30)
                            
            rstEspecificacionesUnificaVersionII.Close
        End If

    End If
    
    For ZCicloI = 1 To 30
    
        Entra = "N"
        For ZCicloII = 1 To 30
            If ZEnsayoActual(ZCicloII) <> 0 Then
                If ZEnsayo(ZCicloI) = ZEnsayoActual(ZCicloII) Then
                    ZEnsayoActual(ZCicloII) = 0
                    Entra = "S"
                    ZLugar = ZCicloII
                    Exit For
                End If
            End If
        Next ZCicloII
        
        If Entra = "S" Then
            Select Case ZLugar
                Case 1
                    ValorOri1.Caption = ZValor(ZCicloI)
                Case 2
                    ValorOri2.Caption = ZValor(ZCicloI)
                Case 3
                    ValorOri3.Caption = ZValor(ZCicloI)
                Case 4
                    ValorOri4.Caption = ZValor(ZCicloI)
                Case 5
                    ValorOri5.Caption = ZValor(ZCicloI)
                Case 6
                    ValorOri6.Caption = ZValor(ZCicloI)
                Case 7
                    ValorOri7.Caption = ZValor(ZCicloI)
                Case 8
                    ValorOri8.Caption = ZValor(ZCicloI)
                Case 9
                    ValorOri9.Caption = ZValor(ZCicloI)
                Case 10
                    ValorOri10.Caption = ZValor(ZCicloI)
                Case 11
                    ValorOri11.Caption = ZValor(ZCicloI)
                Case 12
                    ValorOri12.Caption = ZValor(ZCicloI)
                Case 13
                    ValorOri13.Caption = ZValor(ZCicloI)
                Case 14
                    ValorOri14.Caption = ZValor(ZCicloI)
                Case 15
                    ValorOri15.Caption = ZValor(ZCicloI)
                Case 16
                    ValorOri16.Caption = ZValor(ZCicloI)
                Case 17
                    ValorOri17.Caption = ZValor(ZCicloI)
                Case 18
                    ValorOri18.Caption = ZValor(ZCicloI)
                Case 19
                    ValorOri19.Caption = ZValor(ZCicloI)
                Case 20
                    ValorOri20.Caption = ZValor(ZCicloI)
                Case 21
                    ValorOri21.Caption = ZValor(ZCicloI)
                Case 22
                    ValorOri22.Caption = ZValor(ZCicloI)
                Case 23
                    ValorOri23.Caption = ZValor(ZCicloI)
                Case 24
                    ValorOri24.Caption = ZValor(ZCicloI)
                Case 25
                    ValorOri25.Caption = ZValor(ZCicloI)
                Case 26
                    ValorOri26.Caption = ZValor(ZCicloI)
                Case 27
                    ValorOri27.Caption = ZValor(ZCicloI)
                Case 28
                    ValorOri28.Caption = ZValor(ZCicloI)
                Case 29
                    ValorOri29.Caption = ZValor(ZCicloI)
                Case 30
                    ValorOri30.Caption = ZValor(ZCicloI)
                Case Else
            End Select
        End If
    Next ZCicloI
    
    Call Conecta_Empresa
    
End Sub

Private Sub Graba_Click()
    
    If WGraba <> "S" Then
    
        ZProceso = 0
        Call Ingresa_clave

               Else

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo = " + "'" + Lote.Text + "'"
        ZSql = ZSql + " and Articulo = " + "'" + Articulo.Text + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            WWFecha = rstLaudo!Fecha
            WWFechaOrd = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
            ZOrdVencimiento = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
            
            rstLaudo.Close
            
            If WWFechaOrd > ZOrdVencimiento Then
                m$ = "Fecha de revalida inferior a la fecha de laudo"
                A% = MsgBox(m$, 0, "Revalida de Productos")
                Exit Sub
            End If
            
        End If

        Sql1 = "Select Max(Codigo) as [CodigoMayor]"
        Sql2 = " FROM Revalida"
        spRevalida = Sql1 + Sql2
        Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
        If rstRevalida.RecordCount > 0 Then
            rstRevalida.MoveLast
            WCodigoMayor = IIf(IsNull(rstRevalida!CodigoMayor), "0", rstRevalida!CodigoMayor)
            WCodigo = Str$(WCodigoMayor + 1)
            rstRevalida.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Revalida ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Resultado ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Responsable ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Codigo1 ,"
        ZSql = ZSql + "Codigo2 ,"
        ZSql = ZSql + "Codigo3 ,"
        ZSql = ZSql + "Codigo4 ,"
        ZSql = ZSql + "Codigo5 ,"
        ZSql = ZSql + "Codigo6 ,"
        ZSql = ZSql + "Codigo7 ,"
        ZSql = ZSql + "Codigo8 ,"
        ZSql = ZSql + "Codigo9 ,"
        ZSql = ZSql + "Codigo10 ,"
        ZSql = ZSql + "Codigo11 ,"
        ZSql = ZSql + "Codigo12 ,"
        ZSql = ZSql + "Codigo13 ,"
        ZSql = ZSql + "Codigo14 ,"
        ZSql = ZSql + "Codigo15 ,"
        ZSql = ZSql + "Codigo16 ,"
        ZSql = ZSql + "Codigo17 ,"
        ZSql = ZSql + "Codigo18 ,"
        ZSql = ZSql + "Codigo19 ,"
        ZSql = ZSql + "Codigo20 ,"
        ZSql = ZSql + "Codigo21 ,"
        ZSql = ZSql + "Codigo22 ,"
        ZSql = ZSql + "Codigo23 ,"
        ZSql = ZSql + "Codigo24 ,"
        ZSql = ZSql + "Codigo25 ,"
        ZSql = ZSql + "Codigo26 ,"
        ZSql = ZSql + "Codigo27 ,"
        ZSql = ZSql + "Codigo28 ,"
        ZSql = ZSql + "Codigo29 ,"
        ZSql = ZSql + "Codigo30 ,"
        ZSql = ZSql + "Std1 ,"
        ZSql = ZSql + "Std2 ,"
        ZSql = ZSql + "Std3 ,"
        ZSql = ZSql + "Std4 ,"
        ZSql = ZSql + "Std5 ,"
        ZSql = ZSql + "Std6 ,"
        ZSql = ZSql + "Std7 ,"
        ZSql = ZSql + "Std8 ,"
        ZSql = ZSql + "Std9 ,"
        ZSql = ZSql + "Std10 ,"
        ZSql = ZSql + "Std11 ,"
        ZSql = ZSql + "Std12 ,"
        ZSql = ZSql + "Std13 ,"
        ZSql = ZSql + "Std14 ,"
        ZSql = ZSql + "Std15 ,"
        ZSql = ZSql + "Std16 ,"
        ZSql = ZSql + "Std17 ,"
        ZSql = ZSql + "Std18 ,"
        ZSql = ZSql + "Std19 ,"
        ZSql = ZSql + "Std20 ,"
        ZSql = ZSql + "Std21 ,"
        ZSql = ZSql + "Std22 ,"
        ZSql = ZSql + "Std23 ,"
        ZSql = ZSql + "Std24 ,"
        ZSql = ZSql + "Std25 ,"
        ZSql = ZSql + "Std26 ,"
        ZSql = ZSql + "Std27 ,"
        ZSql = ZSql + "Std28 ,"
        ZSql = ZSql + "Std29 ,"
        ZSql = ZSql + "Std30 ,"
        ZSql = ZSql + "Valor1 ,"
        ZSql = ZSql + "Valor2 ,"
        ZSql = ZSql + "Valor3 ,"
        ZSql = ZSql + "Valor4 ,"
        ZSql = ZSql + "Valor5 ,"
        ZSql = ZSql + "Valor6 ,"
        ZSql = ZSql + "Valor7 ,"
        ZSql = ZSql + "Valor8 ,"
        ZSql = ZSql + "Valor9 ,"
        ZSql = ZSql + "Valor10 ,"
        ZSql = ZSql + "Valor11 ,"
        ZSql = ZSql + "Valor12 ,"
        ZSql = ZSql + "Valor13 ,"
        ZSql = ZSql + "Valor14 ,"
        ZSql = ZSql + "Valor15 ,"
        ZSql = ZSql + "Valor16 ,"
        ZSql = ZSql + "Valor17 ,"
        ZSql = ZSql + "Valor18 ,"
        ZSql = ZSql + "Valor19 ,"
        ZSql = ZSql + "Valor20 ,"
        ZSql = ZSql + "Valor21 ,"
        ZSql = ZSql + "Valor22 ,"
        ZSql = ZSql + "Valor23 ,"
        ZSql = ZSql + "Valor24 ,"
        ZSql = ZSql + "Valor25 ,"
        ZSql = ZSql + "Valor26 ,"
        ZSql = ZSql + "Valor27 ,"
        ZSql = ZSql + "Valor28 ,"
        ZSql = ZSql + "Valor29 ,"
        ZSql = ZSql + "Valor30 ,"
        ZSql = ZSql + "Revalida )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + Lote.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Articulo.Text + "',"
        ZSql = ZSql + "'" + Resultado.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + Responsable.Text + "',"
        ZSql = ZSql + "'" + Vencimiento.Text + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(1)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(2)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(3)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(4)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(5)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(6)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(7)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(8)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(9)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(10)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(11)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(12)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(13)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(14)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(15)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(16)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(17)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(18)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(19)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(20)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(21)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(22)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(23)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(24)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(25)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(26)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(27)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(28)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(29)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActualGraba(30)) + "',"
        ZSql = ZSql + "'" + Std1.Caption + "',"
        ZSql = ZSql + "'" + Std2.Caption + "',"
        ZSql = ZSql + "'" + Std3.Caption + "',"
        ZSql = ZSql + "'" + Std4.Caption + "',"
        ZSql = ZSql + "'" + Std5.Caption + "',"
        ZSql = ZSql + "'" + Std6.Caption + "',"
        ZSql = ZSql + "'" + Std7.Caption + "',"
        ZSql = ZSql + "'" + Std8.Caption + "',"
        ZSql = ZSql + "'" + Std9.Caption + "',"
        ZSql = ZSql + "'" + Std10.Caption + "',"
        ZSql = ZSql + "'" + Std11.Caption + "',"
        ZSql = ZSql + "'" + Std12.Caption + "',"
        ZSql = ZSql + "'" + Std13.Caption + "',"
        ZSql = ZSql + "'" + Std14.Caption + "',"
        ZSql = ZSql + "'" + Std15.Caption + "',"
        ZSql = ZSql + "'" + Std16.Caption + "',"
        ZSql = ZSql + "'" + Std17.Caption + "',"
        ZSql = ZSql + "'" + Std18.Caption + "',"
        ZSql = ZSql + "'" + Std19.Caption + "',"
        ZSql = ZSql + "'" + Std20.Caption + "',"
        ZSql = ZSql + "'" + Std21.Caption + "',"
        ZSql = ZSql + "'" + Std22.Caption + "',"
        ZSql = ZSql + "'" + Std23.Caption + "',"
        ZSql = ZSql + "'" + Std24.Caption + "',"
        ZSql = ZSql + "'" + Std25.Caption + "',"
        ZSql = ZSql + "'" + Std26.Caption + "',"
        ZSql = ZSql + "'" + Std27.Caption + "',"
        ZSql = ZSql + "'" + Std28.Caption + "',"
        ZSql = ZSql + "'" + Std29.Caption + "',"
        ZSql = ZSql + "'" + Std30.Caption + "',"
        ZSql = ZSql + "'" + Valor1.Text + "',"
        ZSql = ZSql + "'" + Valor2.Text + "',"
        ZSql = ZSql + "'" + Valor3.Text + "',"
        ZSql = ZSql + "'" + Valor4.Text + "',"
        ZSql = ZSql + "'" + Valor5.Text + "',"
        ZSql = ZSql + "'" + Valor6.Text + "',"
        ZSql = ZSql + "'" + Valor7.Text + "',"
        ZSql = ZSql + "'" + Valor8.Text + "',"
        ZSql = ZSql + "'" + Valor9.Text + "',"
        ZSql = ZSql + "'" + Valor10.Text + "',"
        ZSql = ZSql + "'" + Valor11.Text + "',"
        ZSql = ZSql + "'" + Valor12.Text + "',"
        ZSql = ZSql + "'" + Valor13.Text + "',"
        ZSql = ZSql + "'" + Valor14.Text + "',"
        ZSql = ZSql + "'" + Valor15.Text + "',"
        ZSql = ZSql + "'" + Valor16.Text + "',"
        ZSql = ZSql + "'" + Valor17.Text + "',"
        ZSql = ZSql + "'" + Valor18.Text + "',"
        ZSql = ZSql + "'" + Valor19.Text + "',"
        ZSql = ZSql + "'" + Valor20.Text + "',"
        ZSql = ZSql + "'" + Valor21.Text + "',"
        ZSql = ZSql + "'" + Valor22.Text + "',"
        ZSql = ZSql + "'" + Valor23.Text + "',"
        ZSql = ZSql + "'" + Valor24.Text + "',"
        ZSql = ZSql + "'" + Valor25.Text + "',"
        ZSql = ZSql + "'" + Valor26.Text + "',"
        ZSql = ZSql + "'" + Valor27.Text + "',"
        ZSql = ZSql + "'" + Valor28.Text + "',"
        ZSql = ZSql + "'" + Valor29.Text + "',"
        ZSql = ZSql + "'" + Valor30.Text + "',"
        ZSql = ZSql + "'" + Revalida.Text + "')"
            
        spRevalida = ZSql
        Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZOrdVencimiento = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Laudo SET "
        ZSql = ZSql + "FechaVencimiento = " + "'" + Vencimiento.Text + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "',"
        ZSql = ZSql + "Revalida = " + "'" + Revalida.Text + "'"
        ZSql = ZSql + " Where Laudo = " + "'" + Lote.Text + "'"
        
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        PrgPruart.NroRevalida.Text = Revalida.Text
        PrgPruart.Vto.Text = Vencimiento.Text
        
        Call Cancela_click
    
    End If

End Sub


Private Sub Rechazo_Click()

    If WGraba <> "S" Then
    
        ZProceso = 1
        Call Ingresa_clave

               Else

        Revalida.Text = "99"
    
        Sql1 = "Select Max(Codigo) as [CodigoMayor]"
        Sql2 = " FROM Revalida"
        spRevalida = Sql1 + Sql2
        Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
        If rstRevalida.RecordCount > 0 Then
            rstRevalida.MoveLast
            WCodigoMayor = IIf(IsNull(rstRevalida!CodigoMayor), "0", rstRevalida!CodigoMayor)
            WCodigo = Str$(WCodigoMayor + 1)
            rstRevalida.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Revalida ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Resultado ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Responsable ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Codigo1 ,"
        ZSql = ZSql + "Codigo2 ,"
        ZSql = ZSql + "Codigo3 ,"
        ZSql = ZSql + "Codigo4 ,"
        ZSql = ZSql + "Codigo5 ,"
        ZSql = ZSql + "Codigo6 ,"
        ZSql = ZSql + "Codigo7 ,"
        ZSql = ZSql + "Codigo8 ,"
        ZSql = ZSql + "Codigo9 ,"
        ZSql = ZSql + "Codigo10 ,"
        ZSql = ZSql + "Codigo11 ,"
        ZSql = ZSql + "Codigo12 ,"
        ZSql = ZSql + "Codigo13 ,"
        ZSql = ZSql + "Codigo14 ,"
        ZSql = ZSql + "Codigo15 ,"
        ZSql = ZSql + "Codigo16 ,"
        ZSql = ZSql + "Codigo17 ,"
        ZSql = ZSql + "Codigo18 ,"
        ZSql = ZSql + "Codigo19 ,"
        ZSql = ZSql + "Codigo20 ,"
        ZSql = ZSql + "Codigo21 ,"
        ZSql = ZSql + "Codigo22 ,"
        ZSql = ZSql + "Codigo23 ,"
        ZSql = ZSql + "Codigo24 ,"
        ZSql = ZSql + "Codigo25 ,"
        ZSql = ZSql + "Codigo26 ,"
        ZSql = ZSql + "Codigo27 ,"
        ZSql = ZSql + "Codigo28 ,"
        ZSql = ZSql + "Codigo29 ,"
        ZSql = ZSql + "Codigo30 ,"
        ZSql = ZSql + "Std1 ,"
        ZSql = ZSql + "Std2 ,"
        ZSql = ZSql + "Std3 ,"
        ZSql = ZSql + "Std4 ,"
        ZSql = ZSql + "Std5 ,"
        ZSql = ZSql + "Std6 ,"
        ZSql = ZSql + "Std7 ,"
        ZSql = ZSql + "Std8 ,"
        ZSql = ZSql + "Std9 ,"
        ZSql = ZSql + "Std10 ,"
        ZSql = ZSql + "Std11 ,"
        ZSql = ZSql + "Std12 ,"
        ZSql = ZSql + "Std13 ,"
        ZSql = ZSql + "Std14 ,"
        ZSql = ZSql + "Std15 ,"
        ZSql = ZSql + "Std16 ,"
        ZSql = ZSql + "Std17 ,"
        ZSql = ZSql + "Std18 ,"
        ZSql = ZSql + "Std19 ,"
        ZSql = ZSql + "Std20 ,"
        ZSql = ZSql + "Std21 ,"
        ZSql = ZSql + "Std22 ,"
        ZSql = ZSql + "Std23 ,"
        ZSql = ZSql + "Std24 ,"
        ZSql = ZSql + "Std25 ,"
        ZSql = ZSql + "Std26 ,"
        ZSql = ZSql + "Std27 ,"
        ZSql = ZSql + "Std28 ,"
        ZSql = ZSql + "Std29 ,"
        ZSql = ZSql + "Std30 ,"
        ZSql = ZSql + "Valor1 ,"
        ZSql = ZSql + "Valor2 ,"
        ZSql = ZSql + "Valor3 ,"
        ZSql = ZSql + "Valor4 ,"
        ZSql = ZSql + "Valor5 ,"
        ZSql = ZSql + "Valor6 ,"
        ZSql = ZSql + "Valor7 ,"
        ZSql = ZSql + "Valor8 ,"
        ZSql = ZSql + "Valor9 ,"
        ZSql = ZSql + "Valor10 ,"
        ZSql = ZSql + "Valor11 ,"
        ZSql = ZSql + "Valor12 ,"
        ZSql = ZSql + "Valor13 ,"
        ZSql = ZSql + "Valor14 ,"
        ZSql = ZSql + "Valor15 ,"
        ZSql = ZSql + "Valor16 ,"
        ZSql = ZSql + "Valor17 ,"
        ZSql = ZSql + "Valor18 ,"
        ZSql = ZSql + "Valor19 ,"
        ZSql = ZSql + "Valor20 ,"
        ZSql = ZSql + "Valor21 ,"
        ZSql = ZSql + "Valor22 ,"
        ZSql = ZSql + "Valor23 ,"
        ZSql = ZSql + "Valor24 ,"
        ZSql = ZSql + "Valor25 ,"
        ZSql = ZSql + "Valor26 ,"
        ZSql = ZSql + "Valor27 ,"
        ZSql = ZSql + "Valor28 ,"
        ZSql = ZSql + "Valor29 ,"
        ZSql = ZSql + "Valor30 ,"
        ZSql = ZSql + "Revalida )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + Lote.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + Articulo.Text + "',"
        ZSql = ZSql + "'" + Resultado.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + Responsable.Text + "',"
        ZSql = ZSql + "'" + Vencimiento.Text + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(1)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(2)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(3)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(4)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(5)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(6)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(7)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(8)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(9)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(10)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(11)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(12)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(13)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(14)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(15)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(16)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(17)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(18)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(19)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(20)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(21)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(22)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(23)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(24)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(25)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(26)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(27)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(28)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(29)) + "',"
        ZSql = ZSql + "'" + Str$(ZEnsayoActual(30)) + "',"
        ZSql = ZSql + "'" + Std1.Caption + "',"
        ZSql = ZSql + "'" + Std2.Caption + "',"
        ZSql = ZSql + "'" + Std3.Caption + "',"
        ZSql = ZSql + "'" + Std4.Caption + "',"
        ZSql = ZSql + "'" + Std5.Caption + "',"
        ZSql = ZSql + "'" + Std6.Caption + "',"
        ZSql = ZSql + "'" + Std7.Caption + "',"
        ZSql = ZSql + "'" + Std8.Caption + "',"
        ZSql = ZSql + "'" + Std9.Caption + "',"
        ZSql = ZSql + "'" + Std10.Caption + "',"
        ZSql = ZSql + "'" + Std11.Caption + "',"
        ZSql = ZSql + "'" + Std12.Caption + "',"
        ZSql = ZSql + "'" + Std13.Caption + "',"
        ZSql = ZSql + "'" + Std14.Caption + "',"
        ZSql = ZSql + "'" + Std15.Caption + "',"
        ZSql = ZSql + "'" + Std16.Caption + "',"
        ZSql = ZSql + "'" + Std17.Caption + "',"
        ZSql = ZSql + "'" + Std18.Caption + "',"
        ZSql = ZSql + "'" + Std19.Caption + "',"
        ZSql = ZSql + "'" + Std20.Caption + "',"
        ZSql = ZSql + "'" + Std21.Caption + "',"
        ZSql = ZSql + "'" + Std22.Caption + "',"
        ZSql = ZSql + "'" + Std23.Caption + "',"
        ZSql = ZSql + "'" + Std24.Caption + "',"
        ZSql = ZSql + "'" + Std25.Caption + "',"
        ZSql = ZSql + "'" + Std26.Caption + "',"
        ZSql = ZSql + "'" + Std27.Caption + "',"
        ZSql = ZSql + "'" + Std28.Caption + "',"
        ZSql = ZSql + "'" + Std29.Caption + "',"
        ZSql = ZSql + "'" + Std30.Caption + "',"
        ZSql = ZSql + "'" + Valor1.Text + "',"
        ZSql = ZSql + "'" + Valor2.Text + "',"
        ZSql = ZSql + "'" + Valor3.Text + "',"
        ZSql = ZSql + "'" + Valor4.Text + "',"
        ZSql = ZSql + "'" + Valor5.Text + "',"
        ZSql = ZSql + "'" + Valor6.Text + "',"
        ZSql = ZSql + "'" + Valor7.Text + "',"
        ZSql = ZSql + "'" + Valor8.Text + "',"
        ZSql = ZSql + "'" + Valor9.Text + "',"
        ZSql = ZSql + "'" + Valor10.Text + "',"
        ZSql = ZSql + "'" + Valor11.Text + "',"
        ZSql = ZSql + "'" + Valor12.Text + "',"
        ZSql = ZSql + "'" + Valor13.Text + "',"
        ZSql = ZSql + "'" + Valor14.Text + "',"
        ZSql = ZSql + "'" + Valor15.Text + "',"
        ZSql = ZSql + "'" + Valor16.Text + "',"
        ZSql = ZSql + "'" + Valor17.Text + "',"
        ZSql = ZSql + "'" + Valor18.Text + "',"
        ZSql = ZSql + "'" + Valor19.Text + "',"
        ZSql = ZSql + "'" + Valor20.Text + "',"
        ZSql = ZSql + "'" + Valor21.Text + "',"
        ZSql = ZSql + "'" + Valor22.Text + "',"
        ZSql = ZSql + "'" + Valor23.Text + "',"
        ZSql = ZSql + "'" + Valor24.Text + "',"
        ZSql = ZSql + "'" + Valor25.Text + "',"
        ZSql = ZSql + "'" + Valor26.Text + "',"
        ZSql = ZSql + "'" + Valor27.Text + "',"
        ZSql = ZSql + "'" + Valor28.Text + "',"
        ZSql = ZSql + "'" + Valor29.Text + "',"
        ZSql = ZSql + "'" + Valor30.Text + "',"
        ZSql = ZSql + "'" + Revalida.Text + "')"
            
        spRevalida = ZSql
        Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZVto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZOrdVencimiento = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Laudo SET "
        ZSql = ZSql + "FechaVencimiento = " + "'" + ZVto + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "',"
        ZSql = ZSql + "Revalida = " + "'" + Revalida.Text + "'"
        ZSql = ZSql + " Where Laudo = " + "'" + Lote.Text + "'"
        
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        PrgPruart.NroRevalida.Text = Revalida.Text
        Rem PrgPruart.Vto.Text = Vencimiento.Text
        
        Call Cancela_click

    End If

End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Valor1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vencimiento.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor1.SelStart = 0
        Valor2.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor1.Text = ""
    End If
End Sub

Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor2.SelStart = 0
        Valor3.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor2.Text = ""
    End If
End Sub

Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor3.SelStart = 0
        Valor4.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor3.Text = ""
    End If
End Sub

Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor4.SelStart = 0
        Valor5.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor4.Text = ""
    End If
End Sub

Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor5.SelStart = 0
        Valor6.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor5.Text = ""
    End If
End Sub

Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor6.SelStart = 0
        Valor7.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor6.Text = ""
    End If
End Sub

Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor7.SelStart = 0
        Valor8.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor7.Text = ""
    End If
End Sub

Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor8.SelStart = 0
        Valor9.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor8.Text = ""
    End If
End Sub

Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor9.SelStart = 0
        Valor10.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor9.Text = ""
    End If
End Sub

Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor10.SelStart = 0
        Resultado.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor10.Text = ""
    End If
End Sub




Private Sub Valor11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor11.SelStart = 0
        Valor12.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor11.Text = ""
    End If
End Sub


Private Sub Valor12_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor12.SelStart = 0
        Valor13.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor12.Text = ""
    End If
End Sub


Private Sub Valor13_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor13.SelStart = 0
        Valor14.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor13.Text = ""
    End If
End Sub


Private Sub Valor14_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor14.SelStart = 0
        Valor15.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor14.Text = ""
    End If
End Sub


Private Sub Valor15_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor15.SelStart = 0
        Valor16.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor15.Text = ""
    End If
End Sub


Private Sub Valor16_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor16.SelStart = 0
        Valor17.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor16.Text = ""
    End If
End Sub


Private Sub Valor17_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor17.SelStart = 0
        Valor18.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor17.Text = ""
    End If
End Sub


Private Sub Valor18_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor18.SelStart = 0
        Valor19.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor18.Text = ""
    End If
End Sub


Private Sub Valor19_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor19.SelStart = 0
        Valor20.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor10.Text = ""
    End If
End Sub


Private Sub Valor20_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor20.SelStart = 0
        Resultado.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor20.Text = ""
    End If
End Sub





Private Sub Valor21_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor21.SelStart = 0
        Valor22.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor21.Text = ""
    End If
End Sub

Private Sub Valor22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor22.SelStart = 0
        Valor23.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor22.Text = ""
    End If
End Sub

Private Sub Valor23_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor23.SelStart = 0
        Valor24.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor23.Text = ""
    End If
End Sub

Private Sub Valor24_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor24.SelStart = 0
        Valor25.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor24.Text = ""
    End If
End Sub

Private Sub Valor25_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor25.SelStart = 0
        Valor26.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor25.Text = ""
    End If
End Sub

Private Sub Valor26_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor26.SelStart = 0
        Valor27.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor26.Text = ""
    End If
End Sub

Private Sub Valor27_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor27.SelStart = 0
        Valor28.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor27.Text = ""
    End If
End Sub

Private Sub Valor28_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor28.SelStart = 0
        Valor29.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor28.Text = ""
    End If
End Sub

Private Sub Valor29_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor29.SelStart = 0
        Valor30.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor29.Text = ""
    End If
End Sub

Private Sub Valor30_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor30.SelStart = 0
        Resultado.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor30.Text = ""
    End If
End Sub






Private Sub Resultado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Resultado.Text = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Responsable_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Vencimiento.SetFocus
    End If
    If KeyAscii = 27 Then
        Responsable.Text = ""
    End If
End Sub


Sub Ingresa_clave()

    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    XClave.Visible = False

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGraba = "N"
        WGrabaI = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            WGrabaI = IIf(IsNull(rstOperador!GrabaI), "", rstOperador!GrabaI)
            rstOperador.Close
        End If
        
        If WGrabaI = "S" Then
            WGraba = "S"
            XClave.Visible = False
            If WProceso = 0 Then
                Call Graba_Click
                    Else
                Call Rechazo_Click
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Composicion de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub

