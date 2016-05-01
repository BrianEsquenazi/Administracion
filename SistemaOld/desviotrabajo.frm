VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgtrabajoDesvio 
   BackColor       =   &H00C0C000&
   Caption         =   "Traspaso a Desvio de Materias Primas"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11880
   Begin VB.TextBox Desvio 
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
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   4560
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   16
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
         TabIndex        =   18
         Top             =   240
         Width           =   2895
      End
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
      TabIndex        =   14
      Top             =   7440
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
      TabIndex        =   13
      Top             =   7440
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
      Top             =   6960
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
      Top             =   6600
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
      Top             =   6240
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
      TabIndex        =   9
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   5640
      TabIndex        =   0
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
      TabIndex        =   4
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
      Left            =   8880
      TabIndex        =   21
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
      TabIndex        =   23
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5821
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Especificaciones 1 - 10"
      TabPicture(0)   =   "desviotrabajo.frx":0000
      Tab(0).ControlCount=   44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label92"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblresultado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Std10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Std9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Std8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Std7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Std6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Std5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Std4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Std3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Std2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Std1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label90"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Descri10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Descri9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Descri8"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Descri7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Descri6"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Descri5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Descri4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Descri3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Descri2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Descri1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label7"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "ValorOri10"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "ValorOri9"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "ValorOri8"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ValorOri7"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "ValorOri6"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ValorOri5"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "ValorOri4"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "ValorOri3"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "ValorOri2"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "ValorOri1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Valor10"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Valor9"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Valor8"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Valor7"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Valor6"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Valor5"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Valor4"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Valor3"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Valor2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Valor1"
      Tab(0).Control(43).Enabled=   0   'False
      TabCaption(1)   =   "Especificaciones 11 - 20"
      TabPicture(1)   =   "desviotrabajo.frx":001C
      Tab(1).ControlCount=   44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label116"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblresultadoII"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Std20"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Std19"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Std18"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Std17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Std16"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Std15"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Std14"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Std13"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Std12"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Std11"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label104"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Descri20"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Descri19"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Descri18"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Descri17"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Descri16"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Descri15"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Descri14"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Descri13"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Descri12"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Descri11"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label22"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "ValorOri20"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "ValorOri19"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "ValorOri18"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "ValorOri17"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "ValorOri16"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "ValorOri15"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "ValorOri14"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "ValorOri13"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "ValorOri12"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "ValorOri11"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Valor20"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Valor19"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Valor18"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Valor17"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Valor16"
      Tab(1).Control(38).Enabled=   -1  'True
      Tab(1).Control(39)=   "Valor15"
      Tab(1).Control(39).Enabled=   -1  'True
      Tab(1).Control(40)=   "Valor14"
      Tab(1).Control(40).Enabled=   -1  'True
      Tab(1).Control(41)=   "Valor13"
      Tab(1).Control(41).Enabled=   -1  'True
      Tab(1).Control(42)=   "Valor12"
      Tab(1).Control(42).Enabled=   -1  'True
      Tab(1).Control(43)=   "Valor11"
      Tab(1).Control(43).Enabled=   -1  'True
      TabCaption(2)   =   "Especificaciones 21 - 30"
      TabPicture(2)   =   "desviotrabajo.frx":0038
      Tab(2).ControlCount=   44
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Descri21"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Descri22"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Descri23"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Descri24"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Descri25"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Descri26"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Descri27"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Descri28"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Descri29"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Descri30"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label41"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Std21"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Std22"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Std23"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Std24"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Std25"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Std26"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Std27"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Std28"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Std29"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Std30"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label30"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label29"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "ValorOri21"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "ValorOri22"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "ValorOri23"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "ValorOri24"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "ValorOri25"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "ValorOri26"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "ValorOri27"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "ValorOri28"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "ValorOri29"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "ValorOri30"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Label45"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Valor21"
      Tab(2).Control(34).Enabled=   -1  'True
      Tab(2).Control(35)=   "Valor22"
      Tab(2).Control(35).Enabled=   -1  'True
      Tab(2).Control(36)=   "Valor23"
      Tab(2).Control(36).Enabled=   -1  'True
      Tab(2).Control(37)=   "Valor24"
      Tab(2).Control(37).Enabled=   -1  'True
      Tab(2).Control(38)=   "Valor25"
      Tab(2).Control(38).Enabled=   -1  'True
      Tab(2).Control(39)=   "Valor26"
      Tab(2).Control(39).Enabled=   -1  'True
      Tab(2).Control(40)=   "Valor27"
      Tab(2).Control(40).Enabled=   -1  'True
      Tab(2).Control(41)=   "Valor28"
      Tab(2).Control(41).Enabled=   -1  'True
      Tab(2).Control(42)=   "Valor29"
      Tab(2).Control(42).Enabled=   -1  'True
      Tab(2).Control(43)=   "Valor30"
      Tab(2).Control(43).Enabled=   -1  'True
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   9780
         TabIndex        =   53
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor2 
         Height          =   285
         Left            =   9780
         TabIndex        =   52
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   9780
         TabIndex        =   51
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor4 
         Height          =   285
         Left            =   9780
         TabIndex        =   50
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor5 
         Height          =   285
         Left            =   9780
         TabIndex        =   49
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor6 
         Height          =   285
         Left            =   9780
         TabIndex        =   48
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor7 
         Height          =   285
         Left            =   9780
         TabIndex        =   47
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor8 
         Height          =   285
         Left            =   9780
         TabIndex        =   46
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor9 
         Height          =   285
         Left            =   9780
         TabIndex        =   45
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor10 
         Height          =   285
         Left            =   9780
         TabIndex        =   44
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor11 
         Height          =   285
         Left            =   -65220
         TabIndex        =   43
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor12 
         Height          =   285
         Left            =   -65220
         TabIndex        =   42
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor13 
         Height          =   285
         Left            =   -65220
         TabIndex        =   41
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor14 
         Height          =   285
         Left            =   -65220
         TabIndex        =   40
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor15 
         Height          =   285
         Left            =   -65220
         TabIndex        =   39
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor16 
         Height          =   285
         Left            =   -65220
         TabIndex        =   38
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor17 
         Height          =   285
         Left            =   -65220
         TabIndex        =   37
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor18 
         Height          =   285
         Left            =   -65220
         TabIndex        =   36
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor19 
         Height          =   285
         Left            =   -65220
         TabIndex        =   35
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor20 
         Height          =   285
         Left            =   -65220
         TabIndex        =   34
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor30 
         Height          =   285
         Left            =   -65220
         TabIndex        =   33
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor29 
         Height          =   285
         Left            =   -65220
         TabIndex        =   32
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor28 
         Height          =   285
         Left            =   -65220
         TabIndex        =   31
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor27 
         Height          =   285
         Left            =   -65220
         TabIndex        =   30
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor26 
         Height          =   285
         Left            =   -65220
         TabIndex        =   29
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor25 
         Height          =   285
         Left            =   -65220
         TabIndex        =   28
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor24 
         Height          =   285
         Left            =   -65220
         TabIndex        =   27
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor23 
         Height          =   285
         Left            =   -65220
         TabIndex        =   26
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor22 
         Height          =   285
         Left            =   -65220
         TabIndex        =   25
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor21 
         Height          =   285
         Left            =   -65220
         TabIndex        =   24
         Top             =   720
         Width           =   1700
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
         TabIndex        =   155
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label ValorOri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   154
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label ValorOri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   153
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label ValorOri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   152
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label ValorOri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   151
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label ValorOri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   150
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label ValorOri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   149
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label ValorOri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   148
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label ValorOri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   147
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label ValorOri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   146
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label ValorOri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   145
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label ValorOri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   144
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label ValorOri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   143
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label ValorOri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   142
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label ValorOri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   141
         Top             =   1440
         Width           =   3195
      End
      Begin VB.Label ValorOri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   140
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Label ValorOri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   139
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label ValorOri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   138
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label ValorOri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   137
         Top             =   2400
         Width           =   3195
      End
      Begin VB.Label ValorOri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   136
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label ValorOri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   135
         Top             =   2880
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
         TabIndex        =   134
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label ValorOri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   133
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label ValorOri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   132
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label ValorOri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   131
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label ValorOri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   130
         Top             =   1440
         Width           =   3195
      End
      Begin VB.Label ValorOri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   129
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Label ValorOri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   128
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label ValorOri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   127
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label ValorOri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   126
         Top             =   2400
         Width           =   3195
      End
      Begin VB.Label ValorOri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   125
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label ValorOri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   124
         Top             =   2880
         Width           =   3195
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
         TabIndex        =   123
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   2880
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
         TabIndex        =   112
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   111
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Std2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   110
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   109
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   108
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   107
         Top             =   1680
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
      Begin VB.Label Std7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   105
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   104
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   103
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   102
         Top             =   2880
         Width           =   3200
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
         TabIndex        =   101
         Top             =   480
         Width           =   3200
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
         TabIndex        =   100
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Descri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   99
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Descri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   98
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   97
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   96
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   95
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   94
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   93
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   91
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   90
         Top             =   2880
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
         TabIndex        =   89
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   88
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Std12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   87
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   86
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   85
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   84
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   83
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   82
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   81
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   80
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   79
         Top             =   2880
         Width           =   3200
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
         TabIndex        =   78
         Top             =   480
         Width           =   3200
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
         TabIndex        =   77
         Top             =   480
         Width           =   1700
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
         TabIndex        =   76
         Top             =   480
         Width           =   1700
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
         TabIndex        =   75
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   74
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Std29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   73
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   72
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   71
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   70
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   69
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   68
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   67
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   66
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   65
         Top             =   720
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
         TabIndex        =   64
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Descri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   63
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Descri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   61
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   60
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   58
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   57
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   56
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   55
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   720
         Width           =   3200
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vencimiemto"
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
      Left            =   7320
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desvio"
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
      TabIndex        =   20
      Top             =   120
      Width           =   855
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
      TabIndex        =   12
      Top             =   6960
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
      TabIndex        =   11
      Top             =   6600
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
      TabIndex        =   10
      Top             =   6240
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      Left            =   4560
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgtrabajoDesvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WGraba As String
Dim ZEnsayo(30) As Integer
Dim ZEnsayoActual(30) As Integer
Dim ZValor(30) As String
Dim WRevalida As Integer
Dim WProceso As Integer

Dim WInforme As String
Dim WLote As String
Dim WPrueba As String
Dim WProducto As String
Dim WFecha As String
Dim WOrden As String

Dim WValor1 As String
Dim WValor2 As String
Dim WValor3 As String
Dim WValor4 As String
Dim WValor5 As String
Dim WValor6 As String
Dim WValor7 As String
Dim WValor8 As String
Dim WValor9 As String
Dim WValor10 As String
Dim WValor11 As String
Dim WValor12 As String
Dim WValor13 As String
Dim WValor14 As String
Dim WValor15 As String
Dim WValor16 As String
Dim WValor17 As String
Dim WValor18 As String
Dim WValor19 As String
Dim WValor20 As String
Dim WValor21 As String
Dim WValor22 As String
Dim WValor23 As String
Dim WValor24 As String
Dim WValor25 As String
Dim WValor26 As String
Dim WValor27 As String
Dim WValor28 As String
Dim WValor29 As String
Dim WValor30 As String

Dim WValorNumero1 As String
Dim WValorNumero2 As String
Dim WValorNumero3 As String
Dim WValorNumero4 As String
Dim WValorNumero5 As String
Dim WValorNumero6 As String
Dim WValorNumero7 As String
Dim WValorNumero8 As String
Dim WValorNumero9 As String
Dim WValorNumero10 As String
Dim WValorNumero11 As String
Dim WValorNumero12 As String
Dim WValorNumero13 As String
Dim WValorNumero14 As String
Dim WValorNumero15 As String
Dim WValorNumero16 As String
Dim WValorNumero17 As String
Dim WValorNumero18 As String
Dim WValorNumero19 As String
Dim WValorNumero20 As String
Dim WValorNumero21 As String
Dim WValorNumero22 As String
Dim WValorNumero23 As String
Dim WValorNumero24 As String
Dim WValorNumero25 As String
Dim WValorNumero26 As String
Dim WValorNumero27 As String
Dim WValorNumero28 As String
Dim WValorNumero29 As String
Dim WValorNumero30 As String

Dim WEnsayo As String
Dim WAspecto As String
Dim WObservaciones As String
Dim WConfecciono As String
Dim WLiberada As String
Dim WDevuelta As String
Dim WRechazo As String
Dim WNueva As String
Dim WFechaord As String
Dim WDate As String
Dim WObserva2 As String
Dim WTipomov As String

Dim CargaEmpresa(12, 2) As String
Dim ZSaldoPlanta(10) As Double

Private Sub Cancela_click()
    PrgDesvio.Hide
    Unload Me
    PrgPruart.Show
End Sub

Private Sub Form_Load()

    Lote.Text = ZLoteRevalida
    Desvio.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = ZFechaVencimiento
    Articulo.Text = ZArticuloRevalida
    DesArticulo.Caption = ZDesArticuloRevalida
    WGraba = ""
    
    Select Case Val(WEmpresa)
        Case 1
            ZZDesde = 190000
            ZZHasta = 194999
        Case 2
            ZZDesde = 690000
            ZZHasta = 694999
        Case 3
            ZZDesde = 290000
            ZZHasta = 294999
        Case 4
            ZZDesde = 790000
            ZZHasta = 794999
        Case 5
            ZZDesde = 390000
            ZZHasta = 394999
        Case 6
            ZZDesde = 490000
            ZZHasta = 494999
        Case 7
            ZZDesde = 590000
            ZZHasta = 594999
        Case 8
            ZZDesde = 790000
            ZZHasta = 794999
        Case Else
            ZZDesde = 890000
            ZZHasta = 894999
    End Select
    
    Desvio.Text = ""
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo >= " + "'" + Str$(ZZDesde) + "'"
    ZSql = ZSql + " and Laudo <= " + "'" + Str$(ZZHasta) + "'"
    ZSql = ZSql + " Order by Laudo"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveLast
            Desvio.Text = Trim(Str$(rstLaudo!Laudo + 1))
        End With
        rstLaudo.Close
    End If
    
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
        Descri1.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(2)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri2.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(3)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(4)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(5)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(6)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(7)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(8)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(9)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(10)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(11)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri11.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri11.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(12)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri12.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri12.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(13)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri13.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri13.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(14)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri14.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri14.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(15)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri15.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri15.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(16)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri16.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri16.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(17)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri17.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri17.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(18)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri18.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri18.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(19)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri19.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri19.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(20)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri20.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri20.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(21)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri21.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri21.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(22)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri22.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri22.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(23)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri23.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri23.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(24)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri24.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri24.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(25)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri25.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri25.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(26)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri26.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri26.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(27)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri27.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri27.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(28)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri28.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri28.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(29)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri29.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri29.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Str$(ZEnsayoActual(30)) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri30.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
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
                        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "", rstEspecificacionesUnificaVersion!Ensayo11)
                        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "", rstEspecificacionesUnificaVersion!Ensayo12)
                        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "", rstEspecificacionesUnificaVersion!Ensayo13)
                        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "", rstEspecificacionesUnificaVersion!Ensayo14)
                        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "", rstEspecificacionesUnificaVersion!Ensayo15)
                        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "", rstEspecificacionesUnificaVersion!Ensayo16)
                        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "", rstEspecificacionesUnificaVersion!Ensayo17)
                        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "", rstEspecificacionesUnificaVersion!Ensayo18)
                        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "", rstEspecificacionesUnificaVersion!Ensayo19)
                        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "", rstEspecificacionesUnificaVersion!Ensayo20)
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
                        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "", rstEspecificacionesUnificaVersion!Ensayo11)
                        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "", rstEspecificacionesUnificaVersion!Ensayo12)
                        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "", rstEspecificacionesUnificaVersion!Ensayo13)
                        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "", rstEspecificacionesUnificaVersion!Ensayo14)
                        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "", rstEspecificacionesUnificaVersion!Ensayo15)
                        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "", rstEspecificacionesUnificaVersion!Ensayo16)
                        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "", rstEspecificacionesUnificaVersion!Ensayo17)
                        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "", rstEspecificacionesUnificaVersion!Ensayo18)
                        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "", rstEspecificacionesUnificaVersion!Ensayo19)
                        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "", rstEspecificacionesUnificaVersion!Ensayo20)
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
            ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
            ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
            ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
            ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
            ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
            ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
            ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
            ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
            ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
            ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
            rstEspecificacionesUnifica.Close
            LlamaImprime = "S"
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM EspecificacionesUnificaIII"
        Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Articulo.Text + "'"
        spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
        Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnificaIII.RecordCount > 0 Then
        
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo21), "", rstEspecificacionesUnifica!Ensayo21)
            ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo22), "", rstEspecificacionesUnifica!Ensayo22)
            ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo23), "", rstEspecificacionesUnifica!Ensayo23)
            ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo24), "", rstEspecificacionesUnifica!Ensayo24)
            ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo25), "", rstEspecificacionesUnifica!Ensayo25)
            ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo26), "", rstEspecificacionesUnifica!Ensayo26)
            ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo27), "", rstEspecificacionesUnifica!Ensayo27)
            ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo28), "", rstEspecificacionesUnifica!Ensayo28)
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo29), "", rstEspecificacionesUnifica!Ensayo29)
            ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo30), "", rstEspecificacionesUnifica!Ensayo30)
            
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
            ZOrden = rstLaudo!Orden
            ZInforme = rstLaudo!Informe
            ZPartiOri = rstLaudo!partiori
            ZEnvase = rstLaudo!Envase
            ZNroDespacho = rstLaudo!NroDespacho
            ZProcedencia = rstLaudo!Procedencia
            ZOrigen = rstLaudo!Origen
            rstLaudo.Close
        End If
        
        For Ciclo = 1 To 999999
        
            Auxi = Str$(Ciclo)
            Call Ceros(Auxi, 6)
            ZZClave = "0" + Auxi + "01"
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.Clave = " + "'" + ZZClave + "'"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                rstMovguia.Close
                    Else
                ZCodigo = Ciclo
                Exit For
            End If
        
        Next Ciclo
        
        
        
        ZSaldo = 0
        ZSaldoII = 0
        Erase ZSaldoPlanta
        XEmpresa = WEmpresa
        
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
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
            Case Else
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
        End Select
            
        For Cicla = 1 To 7
        
            If CargaEmpresa(Cicla, 1) <> "" Then
        
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

                WEntra = "N"
                XParam = "'" + Lote.Text + "','" _
                             + Articulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    ZSaldoPlanta(Cicla) = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    ZSaldo = ZSaldo + ZSaldoPlanta(Cicla)
                    If Val(XEmpresa) = Val(WEmpresa) Then
                        ZSaldoII = ZSaldoII + ZSaldoPlanta(Cicla)
                    End If
                    ZOrigenSaldo = 1
                    ZClaveSaldo = rstLaudo!Clave
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + Articulo.Text + "','" _
                            + Lote.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        ZSaldoPlanta(Cicla) = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        ZSaldo = ZSaldo + ZSaldoPlanta(Cicla)
                        If Val(XEmpresa) = Val(WEmpresa) Then
                            ZSaldoII = ZSaldoII + ZSaldoPlanta(Cicla)
                        End If
                        ZOrigenSaldo = 2
                        ZClaveSaldo = rstMovguia!Clave
                        rstMovguia.Close
                    End If
                End If
                
                If ZSaldoPlanta(Cicla) <> 0 Then
                
                    spMovvar = "ListaMovvarNumero"
                    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovvar.RecordCount > 0 Then
                        With rstMovvar
                            .MoveLast
                            ZVarios = rstMovvar!Codigo + 1
                        End With
                        rstMovvar.Close
                            Else
                        ZVarios = 1
                    End If
                    
                    Tipo = "M"
                    Terminado = "  -     -   "
                    Articulo = Articulo.Text
                    Cantidad = Str$(ZSaldoPlanta(Cicla))
                    Movi = "S"
                    Lote = Lote.Text
                    
                    Renglon = 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                            
                    Auxi1 = Str$(ZVarios)
                    Call Ceros(Auxi1, 6)
                    
                    WCodigo = Trim(Str$(ZVarios))
                    WRenglon = Str$(Renglon)
                    WFecha = Fecha.Text
                    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    WTipo = Tipo
                    WArticulo = Articulo
                    WTerminado = Terminado
                    WCantidad = Cantidad
                    WMovi = Movi
                    WTipomov = "1"
                    WObservaciones = "Traspaso a Desvio Partida " + Desvio.Text
                    WClave = Auxi1 + Auxi
                    WDate = Date$
                    WMarca = ""
                    WLote = Lote
                    
                    XParam = "'" + WClave + "','" _
                             + WCodigo + "','" _
                             + WRenglon + "','" _
                             + WFecha + "','" _
                             + WTipo + "','" _
                             + WArticulo + "','" _
                             + WTerminado + "','" _
                             + WCantidad + "','" _
                             + WFechaord + "','" _
                             + WMovi + "','" _
                             + WTipomov + "','" _
                             + WObservaciones + "','" _
                             + WDate + "','" _
                             + WMarca + "','" _
                             + WLote + "'"
                             
                    spMovvar = "AltaMovvar " + XParam
                    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    Select Case ZOrigenSaldo
                        Case 1
                            WDate = Date$
                            XParam = "'" + ZClaveSaldo + "','" _
                                + WDate + "','" _
                                + "0" + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                        Case Else
                            XParam = "'" + ZClaveSaldo + "','" _
                                + WDate + "','" _
                                + "0" + "'"
                            spMovguia = "ModificaMovguiaSaldo " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    End Select
                    
                    ZCodigo = 3
                    
                    If Val(XEmpresa) <> Val(WEmpresa) Then
                    
                        Auxi = "1"
                        Call Ceros(Auxi, 2)
                        Auxi1 = ZCodigo
                        Call Ceros(Auxi1, 6)
                        
                        WTipomov = Str$(Val(XEmpresa))
                        Call Ceros(WTipomov, 1)
                        WClave = WTipomov + Auxi1 + Auxi
                        WCodigo = Str$(ZCodigo)
                        WRenglon = "1"
                        WFecha = Fecha.Text
                        WTipo = "M"
                        WArticulo = Articulo.Text
                        WTerminado = "0"
                        WCantidad = Str$(ZSaldoPlanta(Cicla))
                        WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        WMovi = "E"
                        WObservaciones = ""
                        WMarca = ""
                        WDestino = "0"
                        WLote = Desvio.Text
                        WSaldo = Str$(ZSaldoPlanta(Cicla))
                        WPartida = ""
                        WPartiOri = ""
                        WTransito = ""
                    
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Guia ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMov ,"
                        ZSql = ZSql + "Codigo ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Tipo ,"
                        ZSql = ZSql + "Articulo ,"
                        ZSql = ZSql + "Terminado ,"
                        ZSql = ZSql + "Cantidad ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Movi,"
                        ZSql = ZSql + "Observaciones,"
                        ZSql = ZSql + "Marca,"
                        ZSql = ZSql + "Destino,"
                        ZSql = ZSql + "Lote,"
                        ZSql = ZSql + "Saldo,"
                        ZSql = ZSql + "Partida,"
                        ZSql = ZSql + "PartiOri,"
                        ZSql = ZSql + "Transito )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WClave + "',"
                        ZSql = ZSql + "'" + WTipomov + "',"
                        ZSql = ZSql + "'" + WCodigo + "',"
                        ZSql = ZSql + "'" + WRenglon + "',"
                        ZSql = ZSql + "'" + WFecha + "',"
                        ZSql = ZSql + "'" + WTipo + "',"
                        ZSql = ZSql + "'" + WArticulo + "',"
                        ZSql = ZSql + "'" + WTerminado + "',"
                        ZSql = ZSql + "'" + WCantidad + "',"
                        ZSql = ZSql + "'" + WFechaord + "',"
                        ZSql = ZSql + "'" + WMovi + "',"
                        ZSql = ZSql + "'" + WObservaciones + "',"
                        ZSql = ZSql + "'" + WMarca + "',"
                        ZSql = ZSql + "'" + WDestino + "',"
                        ZSql = ZSql + "'" + WLote + "',"
                        ZSql = ZSql + "'" + WSaldo + "',"
                        ZSql = ZSql + "'" + WPartida + "',"
                        ZSql = ZSql + "'" + WPartiOri + "',"
                        ZSql = ZSql + "'" + WTransito + "')"
        
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    
                        Call Conecta_Empresa
                        
                        Auxi = "1"
                        Call Ceros(Auxi, 2)
                        Auxi1 = ZCodigo
                        Call Ceros(Auxi1, 6)
                        
                        WTipomov = "0"
                        WClave = WTipomov + Auxi1 + Auxi
                        WCodigo = Str$(ZCodigo)
                        WRenglon = "1"
                        WFecha = Fecha.Text
                        WTipo = "M"
                        WArticulo = Articulo.Text
                        WTerminado = "0"
                        WCantidad = Str$(ZSaldoPlanta(Cicla))
                        WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        WMovi = "S"
                        WObservaciones = ""
                        WMarca = ""
                        WDestino = Trim(Str$(Val(CargaEmpresa(Cicla, 1))))
                        WLote = ""
                        WSaldo = Str$(ZSaldoPlanta(Cicla))
                        WPartida = Desvio.Text
                        WPartiOri = ""
                        WTransito = ""
                    
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Guia ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMov ,"
                        ZSql = ZSql + "Codigo ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Tipo ,"
                        ZSql = ZSql + "Articulo ,"
                        ZSql = ZSql + "Terminado ,"
                        ZSql = ZSql + "Cantidad ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Movi,"
                        ZSql = ZSql + "Observaciones,"
                        ZSql = ZSql + "Marca,"
                        ZSql = ZSql + "Destino,"
                        ZSql = ZSql + "Lote,"
                        ZSql = ZSql + "Saldo,"
                        ZSql = ZSql + "Partida,"
                        ZSql = ZSql + "PartiOri,"
                        ZSql = ZSql + "Transito )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WClave + "',"
                        ZSql = ZSql + "'" + WTipomov + "',"
                        ZSql = ZSql + "'" + WCodigo + "',"
                        ZSql = ZSql + "'" + WRenglon + "',"
                        ZSql = ZSql + "'" + WFecha + "',"
                        ZSql = ZSql + "'" + WTipo + "',"
                        ZSql = ZSql + "'" + WArticulo + "',"
                        ZSql = ZSql + "'" + WTerminado + "',"
                        ZSql = ZSql + "'" + WCantidad + "',"
                        ZSql = ZSql + "'" + WFechaord + "',"
                        ZSql = ZSql + "'" + WMovi + "',"
                        ZSql = ZSql + "'" + WObservaciones + "',"
                        ZSql = ZSql + "'" + WMarca + "',"
                        ZSql = ZSql + "'" + WDestino + "',"
                        ZSql = ZSql + "'" + WLote + "',"
                        ZSql = ZSql + "'" + WSaldo + "',"
                        ZSql = ZSql + "'" + WPartida + "',"
                        ZSql = ZSql + "'" + WPartiOri + "',"
                        ZSql = ZSql + "'" + WTransito + "')"
                
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                End If
                
            End If
                
        Next Cicla
        
        Call Conecta_Empresa
        
        
        
        
        Auxi = "1"
    
        WLote = Desvio.Text
        Call Ceros(WLote, 6)
        WPrueba = Auxi + WLote
        WProducto = Articulo.Text
        WFecha = Fecha.Text
        WInforme = ZInforme
        WOrden = ZOrden
        
        WValor1 = Valor1.Text
        WValor2 = Valor2.Text
        WValor3 = Valor3.Text
        WValor4 = Valor4.Text
        WValor5 = Valor5.Text
        WValor6 = Valor6.Text
        WValor7 = Valor7.Text
        WValor8 = Valor8.Text
        WValor9 = Valor9.Text
        WValor10 = Valor10.Text
        WValor11 = Valor11.Text
        WValor12 = Valor12.Text
        WValor13 = Valor13.Text
        WValor14 = Valor14.Text
        WValor15 = Valor15.Text
        WValor16 = Valor16.Text
        WValor17 = Valor17.Text
        WValor18 = Valor18.Text
        WValor19 = Valor19.Text
        WValor20 = Valor20.Text
        WValor21 = Valor21.Text
        WValor22 = Valor22.Text
        WValor23 = Valor23.Text
        WValor24 = Valor24.Text
        WValor25 = Valor25.Text
        WValor26 = Valor26.Text
        WValor27 = Valor27.Text
        WValor28 = Valor28.Text
        WValor29 = Valor29.Text
        WValor30 = Valor30.Text
        
        WValorNumero1 = ""
        WValorNumero2 = ""
        WValorNumero3 = ""
        WValorNumero4 = ""
        WValorNumero5 = ""
        WValorNumero6 = ""
        WValorNumero7 = ""
        WValorNumero8 = ""
        WValorNumero9 = ""
        WValorNumero10 = ""
        WValorNumero11 = ""
        WValorNumero12 = ""
        WValorNumero13 = ""
        WValorNumero14 = ""
        WValorNumero15 = ""
        WValorNumero16 = ""
        WValorNumero17 = ""
        WValorNumero18 = ""
        WValorNumero19 = ""
        WValorNumero20 = ""
        WValorNumero21 = ""
        WValorNumero22 = ""
        WValorNumero23 = ""
        WValorNumero24 = ""
        WValorNumero25 = ""
        WValorNumero26 = ""
        WValorNumero27 = ""
        WValorNumero28 = ""
        WValorNumero29 = ""
        WValorNumero30 = ""
        
        WEnsayo = ""
        WAspecto = Resultado.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Responsable.Text
        WLiberada = Str$(ZSaldo)
        WDevuelta = "0"
        WRechazo = ""
        WNueva = "N"
        WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WDate = Date$
        WObserva2 = ""
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO PrueArt ("
        ZSql = ZSql + "Prueba ,"
        ZSql = ZSql + "Producto ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Orden ,"
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
        ZSql = ZSql + "ValorNumero1 ,"
        ZSql = ZSql + "ValorNumero2 ,"
        ZSql = ZSql + "ValorNumero3 ,"
        ZSql = ZSql + "ValorNumero4 ,"
        ZSql = ZSql + "ValorNumero5 ,"
        ZSql = ZSql + "ValorNumero6 ,"
        ZSql = ZSql + "ValorNumero7 ,"
        ZSql = ZSql + "ValorNumero8 ,"
        ZSql = ZSql + "ValorNumero9 ,"
        ZSql = ZSql + "ValorNumero10 ,"
        ZSql = ZSql + "ValorNumero11 ,"
        ZSql = ZSql + "ValorNumero12 ,"
        ZSql = ZSql + "ValorNumero13 ,"
        ZSql = ZSql + "ValorNumero14 ,"
        ZSql = ZSql + "ValorNumero15 ,"
        ZSql = ZSql + "ValorNumero16 ,"
        ZSql = ZSql + "ValorNumero17 ,"
        ZSql = ZSql + "ValorNumero18 ,"
        ZSql = ZSql + "ValorNumero19 ,"
        ZSql = ZSql + "ValorNumero20 ,"
        ZSql = ZSql + "ValorNumero21 ,"
        ZSql = ZSql + "ValorNumero22 ,"
        ZSql = ZSql + "ValorNumero23 ,"
        ZSql = ZSql + "ValorNumero24 ,"
        ZSql = ZSql + "ValorNumero25 ,"
        ZSql = ZSql + "ValorNumero26 ,"
        ZSql = ZSql + "ValorNumero27 ,"
        ZSql = ZSql + "ValorNumero28 ,"
        ZSql = ZSql + "ValorNumero29 ,"
        ZSql = ZSql + "ValorNumero30 ,"
        ZSql = ZSql + "Ensayo ,"
        ZSql = ZSql + "Aspecto ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Observa2 ,"
        ZSql = ZSql + "Confecciono ,"
        ZSql = ZSql + "Liberada ,"
        ZSql = ZSql + "Devuelta ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Rechazo ,"
        ZSql = ZSql + "Nueva ,"
        ZSql = ZSql + "FechaOrd ,"
        ZSql = ZSql + "WDate )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WPrueba + "',"
        ZSql = ZSql + "'" + WProducto + "',"
        ZSql = ZSql + "'" + WFecha + "',"
        ZSql = ZSql + "'" + WOrden + "',"
        ZSql = ZSql + "'" + WValor1 + "',"
        ZSql = ZSql + "'" + WValor2 + "',"
        ZSql = ZSql + "'" + WValor3 + "',"
        ZSql = ZSql + "'" + WValor4 + "',"
        ZSql = ZSql + "'" + WValor5 + "',"
        ZSql = ZSql + "'" + WValor6 + "',"
        ZSql = ZSql + "'" + WValor7 + "',"
        ZSql = ZSql + "'" + WValor8 + "',"
        ZSql = ZSql + "'" + WValor9 + "',"
        ZSql = ZSql + "'" + WValor10 + "',"
        ZSql = ZSql + "'" + WValor11 + "',"
        ZSql = ZSql + "'" + WValor12 + "',"
        ZSql = ZSql + "'" + WValor13 + "',"
        ZSql = ZSql + "'" + WValor14 + "',"
        ZSql = ZSql + "'" + WValor15 + "',"
        ZSql = ZSql + "'" + WValor16 + "',"
        ZSql = ZSql + "'" + WValor17 + "',"
        ZSql = ZSql + "'" + WValor18 + "',"
        ZSql = ZSql + "'" + WValor19 + "',"
        ZSql = ZSql + "'" + WValor20 + "',"
        ZSql = ZSql + "'" + WValor21 + "',"
        ZSql = ZSql + "'" + WValor22 + "',"
        ZSql = ZSql + "'" + WValor23 + "',"
        ZSql = ZSql + "'" + WValor24 + "',"
        ZSql = ZSql + "'" + WValor25 + "',"
        ZSql = ZSql + "'" + WValor26 + "',"
        ZSql = ZSql + "'" + WValor27 + "',"
        ZSql = ZSql + "'" + WValor28 + "',"
        ZSql = ZSql + "'" + WValor29 + "',"
        ZSql = ZSql + "'" + WValor30 + "',"
        ZSql = ZSql + "'" + WValorNumero1 + "',"
        ZSql = ZSql + "'" + WValorNumero2 + "',"
        ZSql = ZSql + "'" + WValorNumero3 + "',"
        ZSql = ZSql + "'" + WValorNumero4 + "',"
        ZSql = ZSql + "'" + WValorNumero5 + "',"
        ZSql = ZSql + "'" + WValorNumero6 + "',"
        ZSql = ZSql + "'" + WValorNumero7 + "',"
        ZSql = ZSql + "'" + WValorNumero8 + "',"
        ZSql = ZSql + "'" + WValorNumero9 + "',"
        ZSql = ZSql + "'" + WValorNumero10 + "',"
        ZSql = ZSql + "'" + WValorNumero11 + "',"
        ZSql = ZSql + "'" + WValorNumero12 + "',"
        ZSql = ZSql + "'" + WValorNumero13 + "',"
        ZSql = ZSql + "'" + WValorNumero14 + "',"
        ZSql = ZSql + "'" + WValorNumero15 + "',"
        ZSql = ZSql + "'" + WValorNumero16 + "',"
        ZSql = ZSql + "'" + WValorNumero17 + "',"
        ZSql = ZSql + "'" + WValorNumero18 + "',"
        ZSql = ZSql + "'" + WValorNumero19 + "',"
        ZSql = ZSql + "'" + WValorNumero20 + "',"
        ZSql = ZSql + "'" + WValorNumero21 + "',"
        ZSql = ZSql + "'" + WValorNumero22 + "',"
        ZSql = ZSql + "'" + WValorNumero23 + "',"
        ZSql = ZSql + "'" + WValorNumero24 + "',"
        ZSql = ZSql + "'" + WValorNumero25 + "',"
        ZSql = ZSql + "'" + WValorNumero26 + "',"
        ZSql = ZSql + "'" + WValorNumero27 + "',"
        ZSql = ZSql + "'" + WValorNumero28 + "',"
        ZSql = ZSql + "'" + WValorNumero29 + "',"
        ZSql = ZSql + "'" + WValorNumero30 + "',"
        ZSql = ZSql + "'" + WEnsayo + "',"
        ZSql = ZSql + "'" + WAspecto + "',"
        ZSql = ZSql + "'" + WObservaciones + "',"
        ZSql = ZSql + "'" + WObserva2 + "',"
        ZSql = ZSql + "'" + WConfecciono + "',"
        ZSql = ZSql + "'" + WLiberada + "',"
        ZSql = ZSql + "'" + WDevuelta + "',"
        ZSql = ZSql + "'" + WLote + "',"
        ZSql = ZSql + "'" + WRechazo + "',"
        ZSql = ZSql + "'" + WNuevo + "',"
        ZSql = ZSql + "'" + WFechaord + "',"
        ZSql = ZSql + "'" + WDate + "')"
        
        spPrueart = ZSql
        Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
        
    
    
    
        WLaudo = Desvio.Text
        WRenglon = "1"
        WFecha = Fecha.Text
        WOrden = ZOrden
        WArticulo = Articulo.Text
        WLiberada = Str$(ZSaldo)
        WDevuelta = "0"
        WLote = Desvio.Text
        WRechazo = ""
        WActualiza = "N"
        WMarca = ""
        WInforme = ZInforme
        WSaldo = Str$(ZSaldoII)
        WOrigen = ZOrigen
        WPartiOri = ZPartiOri
        WEnvase = ZEnvase
            
        Auxi1 = Str$(WLaudo)
        Call Ceros(Auxi1, 6)
        Auxi2 = Str$(WRenglon)
        Call Ceros(Auxi2, 2)
            
        WClave = Auxi1 + Auxi2
        WDate = Date$
        
        XParam = "'" + WClave + "','" _
                + WLaudo + "','" _
                + WRenglon + "','" _
                + WFecha + "','" _
                + WArticulo + "','" _
                + WLiberada + "','" _
                + WDevuelta + "','" _
                + WOrden + "','" _
                + WMarca + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WInforme + "','" _
                + WActualiza + "','" _
                + WDate + "','" _
                + WSaldo + "','" _
                + WOrigen + "','" _
                + WPartiOri + "','" _
                + Str$(WEnvase) + "'"
                
        Set rstLaudo = db.OpenRecordset("AltaLaudo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WLaudo + "','" _
                     + WFechaord + "'"
                     
        Set rstLaudo = db.OpenRecordset("ModificaLaudoFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        ZVencimiento = Vencimiento.Text
        ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Laudo SET "
        ZSql = ZSql + "NroDespacho = " + "'" + ZNroDespacho + "',"
        ZSql = ZSql + "Procedencia = " + "'" + ZProcedencia + "',"
        ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
    
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        m$ = "Se ha generado el traspaso al lote de desvio nro. " + Desvio.Text + " el stock de la materia prima"
        G% = MsgBox(m$, 0, "Traspaso de Desvio de Materia Prima")
        
        Call Cancela_click
    
    End If

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
        Valor19.Text = ""
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
        Valor1.SetFocus
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
        WGrabaI = "S"
        
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
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Composicion de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub

