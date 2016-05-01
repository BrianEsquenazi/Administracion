VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRevalidaConsulta 
   BackColor       =   &H00C0C000&
   Caption         =   "Revalida de Fecha de Vencimiento de Materias Primas"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11880
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
      TabIndex        =   16
      Text            =   " "
      Top             =   120
      Width           =   975
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
      Left            =   4560
      TabIndex        =   0
      Top             =   5760
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   5280
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Text            =   " "
      Top             =   4920
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   4560
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
      TabIndex        =   11
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   5
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
      TabIndex        =   6
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
      TabIndex        =   1
      Top             =   480
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3300
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5821
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Especificaciones 1 - 10"
      TabPicture(0)   =   "PrgRevalidaConsulta.frx":0000
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
      TabPicture(1)   =   "PrgRevalidaConsulta.frx":001C
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
      TabPicture(2)   =   "PrgRevalidaConsulta.frx":0038
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
         TabIndex        =   48
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor2 
         Height          =   285
         Left            =   9780
         TabIndex        =   47
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   9780
         TabIndex        =   46
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor4 
         Height          =   285
         Left            =   9780
         TabIndex        =   45
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor5 
         Height          =   285
         Left            =   9780
         TabIndex        =   44
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor6 
         Height          =   285
         Left            =   9780
         TabIndex        =   43
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor7 
         Height          =   285
         Left            =   9780
         TabIndex        =   42
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor8 
         Height          =   285
         Left            =   9780
         TabIndex        =   41
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor9 
         Height          =   285
         Left            =   9780
         TabIndex        =   40
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor10 
         Height          =   285
         Left            =   9780
         TabIndex        =   39
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor11 
         Height          =   285
         Left            =   -65220
         TabIndex        =   38
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor12 
         Height          =   285
         Left            =   -65220
         TabIndex        =   37
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor13 
         Height          =   285
         Left            =   -65220
         TabIndex        =   36
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor14 
         Height          =   285
         Left            =   -65220
         TabIndex        =   35
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor15 
         Height          =   285
         Left            =   -65220
         TabIndex        =   34
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor16 
         Height          =   285
         Left            =   -65220
         TabIndex        =   33
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor17 
         Height          =   285
         Left            =   -65220
         TabIndex        =   32
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor18 
         Height          =   285
         Left            =   -65220
         TabIndex        =   31
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor19 
         Height          =   285
         Left            =   -65220
         TabIndex        =   30
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor20 
         Height          =   285
         Left            =   -65220
         TabIndex        =   29
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor30 
         Height          =   285
         Left            =   -65220
         TabIndex        =   28
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor29 
         Height          =   285
         Left            =   -65220
         TabIndex        =   27
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor28 
         Height          =   285
         Left            =   -65220
         TabIndex        =   26
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor27 
         Height          =   285
         Left            =   -65220
         TabIndex        =   25
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor26 
         Height          =   285
         Left            =   -65220
         TabIndex        =   24
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor25 
         Height          =   285
         Left            =   -65220
         TabIndex        =   23
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor24 
         Height          =   285
         Left            =   -65220
         TabIndex        =   22
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor23 
         Height          =   285
         Left            =   -65220
         TabIndex        =   21
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor22 
         Height          =   285
         Left            =   -65220
         TabIndex        =   20
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor21 
         Height          =   285
         Left            =   -65220
         TabIndex        =   19
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
         TabIndex        =   150
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label ValorOri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   149
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label ValorOri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   148
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label ValorOri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   147
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label ValorOri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   146
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label ValorOri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   145
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label ValorOri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   144
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label ValorOri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   143
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label ValorOri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   142
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label ValorOri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   141
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label ValorOri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   140
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label ValorOri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   139
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label ValorOri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   138
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label ValorOri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   137
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label ValorOri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   136
         Top             =   1440
         Width           =   3195
      End
      Begin VB.Label ValorOri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   135
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Label ValorOri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   134
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label ValorOri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   133
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label ValorOri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   132
         Top             =   2400
         Width           =   3195
      End
      Begin VB.Label ValorOri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   131
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label ValorOri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -68450
         TabIndex        =   130
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
         TabIndex        =   129
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label ValorOri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   128
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label ValorOri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   127
         Top             =   960
         Width           =   3195
      End
      Begin VB.Label ValorOri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   126
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label ValorOri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   125
         Top             =   1440
         Width           =   3195
      End
      Begin VB.Label ValorOri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   124
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Label ValorOri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   123
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label ValorOri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   122
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label ValorOri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   121
         Top             =   2400
         Width           =   3195
      End
      Begin VB.Label ValorOri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   120
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label ValorOri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   6550
         TabIndex        =   119
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
         TabIndex        =   118
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   108
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
         TabIndex        =   107
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   106
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Std2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   105
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   104
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   103
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   102
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   101
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   100
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   99
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   98
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   3350
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Descri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   94
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Descri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   93
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   91
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   90
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   88
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   87
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   86
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   85
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
         TabIndex        =   84
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   83
         Top             =   720
         Width           =   3200
      End
      Begin VB.Label Std12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   82
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   81
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   80
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   79
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   78
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   77
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   76
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   75
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Std30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   69
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Std29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   68
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Std28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   67
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Std27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   66
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Std26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   65
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Std25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   64
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Std24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   63
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Std23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   62
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Std22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   61
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Std21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71650
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   480
         Width           =   3200
      End
      Begin VB.Label Descri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   58
         Top             =   2880
         Width           =   3200
      End
      Begin VB.Label Descri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   57
         Top             =   2640
         Width           =   3200
      End
      Begin VB.Label Descri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   56
         Top             =   2400
         Width           =   3200
      End
      Begin VB.Label Descri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   55
         Top             =   2160
         Width           =   3200
      End
      Begin VB.Label Descri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   1920
         Width           =   3200
      End
      Begin VB.Label Descri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   1680
         Width           =   3200
      End
      Begin VB.Label Descri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   52
         Top             =   1440
         Width           =   3200
      End
      Begin VB.Label Descri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   51
         Top             =   1200
         Width           =   3200
      End
      Begin VB.Label Descri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   50
         Top             =   960
         Width           =   3200
      End
      Begin VB.Label Descri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   720
         Width           =   3200
      End
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
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   5280
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
      TabIndex        =   13
      Top             =   4920
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
      TabIndex        =   12
      Top             =   4560
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgRevalidaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZEnsayo(30) As Integer
Dim ZEnsayoActual(30) As Integer
Dim ZValor(30) As String
Dim WRevalida As Integer

Private Sub Cancela_click()
    PrgRevalidaConsulta.Hide
    Unload Me
    PrgPruart.Show
End Sub

Private Sub Form_Load()

    Lote.Text = ZLoteRevalida
    Articulo.Text = ZArticuloRevalida
    DesArticulo.Caption = ZDesArticuloRevalida
    Revalida.Text = ZNroRevalida
    
    ZSql = ""
    ZSql = ZSql + "Select Revalida.Revalida, Revalida.Articulo, Revalida.Lote, Revalida.Fecha, Revalida.Vencimiento, Revalida.Resultado, Revalida.Observaciones, Revalida.Responsable,"
    ZSql = ZSql + " Revalida.Codigo1, Revalida.Codigo2, Revalida.Codigo3, Revalida.Codigo4, Revalida.Codigo5, Revalida.Codigo6, Revalida.Codigo7, Revalida.Codigo8, Revalida.Codigo9, Revalida.Codigo10, Revalida.Codigo11, Revalida.Codigo12, Revalida.Codigo13, Revalida.Codigo14, Revalida.Codigo15, Revalida.Codigo16, Revalida.Codigo17, Revalida.Codigo18, Revalida.Codigo19, Revalida.Codigo20, Revalida.Codigo21, Revalida.Codigo22, Revalida.Codigo23, Revalida.Codigo24, Revalida.Codigo25, Revalida.Codigo26, Revalida.Codigo27, Revalida.Codigo28, Revalida.Codigo29, Revalida.Codigo30"
    ZSql = ZSql + " FROM Revalida"
    ZSql = ZSql + " Where Revalida.Revalida = " + "'" + Revalida.Text + "'"
    ZSql = ZSql + " and Revalida.Articulo = " + "'" + Articulo.Text + "'"
    ZSql = ZSql + " and Revalida.Lote = " + "'" + Lote.Text + "'"
    spRevalida = ZSql
    Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
    If rstRevalida.RecordCount > 0 Then
    
        Fecha.Text = rstRevalida!Fecha
        Vencimiento.Text = rstRevalida!Vencimiento
        Resultado.Text = rstRevalida!Resultado
        Observaciones.Text = rstRevalida!Observaciones
        Responsable.Text = rstRevalida!Responsable
        
        ZEnsayoActual(1) = rstRevalida!Codigo1
        ZEnsayoActual(2) = rstRevalida!Codigo2
        ZEnsayoActual(3) = rstRevalida!Codigo3
        ZEnsayoActual(4) = rstRevalida!Codigo4
        ZEnsayoActual(5) = rstRevalida!Codigo5
        ZEnsayoActual(6) = rstRevalida!Codigo6
        ZEnsayoActual(7) = rstRevalida!Codigo7
        ZEnsayoActual(8) = rstRevalida!Codigo8
        ZEnsayoActual(9) = rstRevalida!Codigo9
        ZEnsayoActual(10) = rstRevalida!Codigo10
        ZEnsayoActual(11) = IIf(IsNull(rstRevalida!Codigo11), "0", rstRevalida!Codigo11)
        ZEnsayoActual(12) = IIf(IsNull(rstRevalida!Codigo12), "0", rstRevalida!Codigo12)
        ZEnsayoActual(13) = IIf(IsNull(rstRevalida!Codigo13), "0", rstRevalida!Codigo13)
        ZEnsayoActual(14) = IIf(IsNull(rstRevalida!Codigo14), "0", rstRevalida!Codigo14)
        ZEnsayoActual(15) = IIf(IsNull(rstRevalida!Codigo15), "0", rstRevalida!Codigo15)
        ZEnsayoActual(16) = IIf(IsNull(rstRevalida!Codigo16), "0", rstRevalida!Codigo16)
        ZEnsayoActual(17) = IIf(IsNull(rstRevalida!Codigo17), "0", rstRevalida!Codigo17)
        ZEnsayoActual(18) = IIf(IsNull(rstRevalida!Codigo18), "0", rstRevalida!Codigo18)
        ZEnsayoActual(19) = IIf(IsNull(rstRevalida!Codigo19), "0", rstRevalida!Codigo19)
        ZEnsayoActual(20) = IIf(IsNull(rstRevalida!Codigo20), "0", rstRevalida!Codigo20)
        ZEnsayoActual(21) = IIf(IsNull(rstRevalida!Codigo21), "0", rstRevalida!Codigo21)
        ZEnsayoActual(22) = IIf(IsNull(rstRevalida!Codigo22), "0", rstRevalida!Codigo22)
        ZEnsayoActual(23) = IIf(IsNull(rstRevalida!Codigo23), "0", rstRevalida!Codigo23)
        ZEnsayoActual(24) = IIf(IsNull(rstRevalida!Codigo24), "0", rstRevalida!Codigo24)
        ZEnsayoActual(25) = IIf(IsNull(rstRevalida!Codigo25), "0", rstRevalida!Codigo25)
        ZEnsayoActual(26) = IIf(IsNull(rstRevalida!Codigo26), "0", rstRevalida!Codigo26)
        ZEnsayoActual(27) = IIf(IsNull(rstRevalida!Codigo27), "0", rstRevalida!Codigo27)
        ZEnsayoActual(28) = IIf(IsNull(rstRevalida!Codigo28), "0", rstRevalida!Codigo28)
        ZEnsayoActual(29) = IIf(IsNull(rstRevalida!Codigo29), "0", rstRevalida!Codigo29)
        ZEnsayoActual(30) = IIf(IsNull(rstRevalida!Codigo30), "0", rstRevalida!Codigo30)
        
        
        
        rstRevalida.Close
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select Revalida.std1, Revalida.std2, Revalida.std3, Revalida.std4, Revalida.std5, Revalida.std6, Revalida.std7, Revalida.std8, Revalida.std9, Revalida.std10, Revalida.std11, Revalida.std12, Revalida.std13, Revalida.std14, Revalida.std15"
    ZSql = ZSql + " FROM Revalida"
    ZSql = ZSql + " Where Revalida.Revalida = " + "'" + Revalida.Text + "'"
    ZSql = ZSql + " and Revalida.Articulo = " + "'" + Articulo.Text + "'"
    ZSql = ZSql + " and Revalida.Lote = " + "'" + Lote.Text + "'"
    spRevalida = ZSql
    Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
    If rstRevalida.RecordCount > 0 Then
    
        Std1.Caption = rstRevalida!Std1
        Std2.Caption = rstRevalida!Std2
        Std3.Caption = rstRevalida!Std3
        Std4.Caption = rstRevalida!Std4
        Std5.Caption = rstRevalida!Std5
        Std6.Caption = rstRevalida!Std6
        Std7.Caption = rstRevalida!Std7
        Std8.Caption = rstRevalida!Std8
        Std9.Caption = rstRevalida!Std9
        Std10.Caption = rstRevalida!Std10
        Std11.Caption = IIf(IsNull(rstRevalida!Std11), "", rstRevalida!Std11)
        Std12.Caption = IIf(IsNull(rstRevalida!Std12), "", rstRevalida!Std12)
        Std13.Caption = IIf(IsNull(rstRevalida!Std13), "", rstRevalida!Std13)
        Std14.Caption = IIf(IsNull(rstRevalida!Std14), "", rstRevalida!Std14)
        Std15.Caption = IIf(IsNull(rstRevalida!Std15), "", rstRevalida!Std15)
        
        rstRevalida.Close
    End If
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select Revalida.std16, Revalida.std17, Revalida.std18, Revalida.std19, Revalida.std20, Revalida.std21, Revalida.std22, Revalida.std23, Revalida.std24, Revalida.std25, Revalida.std26, Revalida.std27, Revalida.std28, Revalida.std29, Revalida.std30"
    ZSql = ZSql + " FROM Revalida"
    ZSql = ZSql + " Where Revalida.Revalida = " + "'" + Revalida.Text + "'"
    ZSql = ZSql + " and Revalida.Articulo = " + "'" + Articulo.Text + "'"
    ZSql = ZSql + " and Revalida.Lote = " + "'" + Lote.Text + "'"
    spRevalida = ZSql
    Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
    If rstRevalida.RecordCount > 0 Then
    
        Std16.Caption = IIf(IsNull(rstRevalida!Std16), "", rstRevalida!Std16)
        Std17.Caption = IIf(IsNull(rstRevalida!Std17), "", rstRevalida!Std17)
        Std18.Caption = IIf(IsNull(rstRevalida!Std18), "", rstRevalida!Std18)
        Std19.Caption = IIf(IsNull(rstRevalida!Std19), "", rstRevalida!Std19)
        Std20.Caption = IIf(IsNull(rstRevalida!Std20), "", rstRevalida!Std20)
        Std21.Caption = IIf(IsNull(rstRevalida!Std21), "", rstRevalida!Std21)
        Std22.Caption = IIf(IsNull(rstRevalida!Std22), "", rstRevalida!Std22)
        Std23.Caption = IIf(IsNull(rstRevalida!Std23), "", rstRevalida!Std23)
        Std24.Caption = IIf(IsNull(rstRevalida!Std24), "", rstRevalida!Std24)
        Std25.Caption = IIf(IsNull(rstRevalida!Std25), "", rstRevalida!Std25)
        Std26.Caption = IIf(IsNull(rstRevalida!Std26), "", rstRevalida!Std26)
        Std27.Caption = IIf(IsNull(rstRevalida!Std27), "", rstRevalida!Std27)
        Std28.Caption = IIf(IsNull(rstRevalida!Std28), "", rstRevalida!Std28)
        Std29.Caption = IIf(IsNull(rstRevalida!Std29), "", rstRevalida!Std29)
        Std30.Caption = IIf(IsNull(rstRevalida!Std30), "", rstRevalida!Std30)
        
        rstRevalida.Close
    End If
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select Revalida.valor1, Revalida.valor2, Revalida.valor3, Revalida.valor4, Revalida.valor5, Revalida.valor6, Revalida.valor7, Revalida.valor8, Revalida.valor9, Revalida.valor10, Revalida.valor11, Revalida.valor12, Revalida.valor13, Revalida.valor14, Revalida.valor15, Revalida.valor16, Revalida.valor17, Revalida.valor18, Revalida.valor19, Revalida.valor20, Revalida.valor21, Revalida.valor22, Revalida.valor23, Revalida.valor24, Revalida.valor25, Revalida.valor26, Revalida.valor27, Revalida.valor28, Revalida.valor29, Revalida.Valor30"
    ZSql = ZSql + " FROM Revalida"
    ZSql = ZSql + " Where Revalida.Revalida = " + "'" + Revalida.Text + "'"
    ZSql = ZSql + " and Revalida.Articulo = " + "'" + Articulo.Text + "'"
    ZSql = ZSql + " and Revalida.Lote = " + "'" + Lote.Text + "'"
    spRevalida = ZSql
    Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
    If rstRevalida.RecordCount > 0 Then
    
        Valor1.Text = rstRevalida!Valor1
        valor2.Text = rstRevalida!valor2
        Valor3.Text = rstRevalida!Valor3
        valor4.Text = rstRevalida!valor4
        valor5.Text = rstRevalida!valor5
        valor6.Text = rstRevalida!valor6
        valor7.Text = rstRevalida!valor7
        valor8.Text = rstRevalida!valor8
        valor9.Text = rstRevalida!valor9
        valor10.Text = rstRevalida!valor10
        Valor11.Text = IIf(IsNull(rstRevalida!Valor11), "", rstRevalida!Valor11)
        Valor12.Text = IIf(IsNull(rstRevalida!Valor12), "", rstRevalida!Valor12)
        Valor13.Text = IIf(IsNull(rstRevalida!Valor13), "", rstRevalida!Valor13)
        Valor14.Text = IIf(IsNull(rstRevalida!Valor14), "", rstRevalida!Valor14)
        Valor15.Text = IIf(IsNull(rstRevalida!Valor15), "", rstRevalida!Valor15)
        Valor16.Text = IIf(IsNull(rstRevalida!Valor16), "", rstRevalida!Valor16)
        Valor17.Text = IIf(IsNull(rstRevalida!Valor17), "", rstRevalida!Valor17)
        Valor18.Text = IIf(IsNull(rstRevalida!Valor18), "", rstRevalida!Valor18)
        Valor19.Text = IIf(IsNull(rstRevalida!Valor19), "", rstRevalida!Valor19)
        Valor20.Text = IIf(IsNull(rstRevalida!Valor20), "", rstRevalida!Valor20)
        Valor21.Text = IIf(IsNull(rstRevalida!Valor21), "", rstRevalida!Valor21)
        Valor22.Text = IIf(IsNull(rstRevalida!Valor22), "", rstRevalida!Valor22)
        Valor23.Text = IIf(IsNull(rstRevalida!Valor23), "", rstRevalida!Valor23)
        Valor24.Text = IIf(IsNull(rstRevalida!Valor24), "", rstRevalida!Valor24)
        Valor25.Text = IIf(IsNull(rstRevalida!Valor25), "", rstRevalida!Valor25)
        Valor26.Text = IIf(IsNull(rstRevalida!Valor26), "", rstRevalida!Valor26)
        Valor27.Text = IIf(IsNull(rstRevalida!Valor27), "", rstRevalida!Valor27)
        Valor28.Text = IIf(IsNull(rstRevalida!Valor28), "", rstRevalida!Valor28)
        Valor29.Text = IIf(IsNull(rstRevalida!Valor29), "", rstRevalida!Valor29)
        Valor30.Text = IIf(IsNull(rstRevalida!Valor30), "", rstRevalida!Valor30)
        
        rstRevalida.Close
    End If
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Prueart"
    ZSql = ZSql + " Where Lote = " + "'" + Lote.Text + "'"
    spPrueart = ZSql
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueart.RecordCount > 0 Then
        ZValor(1) = rstPrueart!Valor1
        ZValor(2) = rstPrueart!valor2
        ZValor(3) = rstPrueart!Valor3
        ZValor(4) = rstPrueart!valor4
        ZValor(5) = rstPrueart!valor5
        ZValor(6) = rstPrueart!valor6
        ZValor(7) = rstPrueart!valor7
        ZValor(8) = rstPrueart!valor8
        ZValor(9) = rstPrueart!valor9
        ZValor(10) = rstPrueart!valor10
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
        rstEspecificacionesUnifica.Close
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
        descri2.Caption = Auxi1 + " - " + rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        descri2.Caption = ""
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
        descri2.Caption = ""
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
        
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo21), "0", rstEspecificacionesUnificaVersionII!Ensayo21)
            ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo22), "0", rstEspecificacionesUnificaVersionII!Ensayo22)
            ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo23), "0", rstEspecificacionesUnificaVersionII!Ensayo23)
            ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo24), "0", rstEspecificacionesUnificaVersionII!Ensayo24)
            ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo25), "0", rstEspecificacionesUnificaVersionII!Ensayo25)
            ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo26), "0", rstEspecificacionesUnificaVersionII!Ensayo26)
            ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo27), "0", rstEspecificacionesUnificaVersionII!Ensayo27)
            ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo28), "0", rstEspecificacionesUnificaVersionII!Ensayo28)
            ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo29), "0", rstEspecificacionesUnificaVersionII!Ensayo29)
            ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo30), "0", rstEspecificacionesUnificaVersionII!Ensayo30)
                            
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

