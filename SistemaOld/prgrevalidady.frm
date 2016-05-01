VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRevalidady 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Revalida de Fecha de Vencimiento de DY"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11880
   Begin VB.TextBox RevalidaConsulta 
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
      Left            =   10440
      MaxLength       =   6
      TabIndex        =   69
      Text            =   " "
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3240
      TabIndex        =   65
      Top             =   2520
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   67
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   66
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
         TabIndex        =   68
         Top             =   240
         Width           =   2895
      End
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
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   61
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Valor1 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "prgrevalidady.frx":0000
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Valor2 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "prgrevalidady.frx":0002
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Valor3 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "prgrevalidady.frx":0004
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Valor4 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "prgrevalidady.frx":0006
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Valor5 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "prgrevalidady.frx":0008
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox Valor6 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "prgrevalidady.frx":000A
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox Valor7 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "prgrevalidady.frx":000C
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox Valor8 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "prgrevalidady.frx":000E
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox Valor9 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "prgrevalidady.frx":0010
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Valor10 
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
      Left            =   8520
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "prgrevalidady.frx":0012
      Top             =   5640
      Width           =   2895
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
      Left            =   3720
      TabIndex        =   26
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
      Left            =   5760
      TabIndex        =   25
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
      TabIndex        =   13
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
      TabIndex        =   14
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
      TabIndex        =   12
      Text            =   " "
      Top             =   6240
      Width           =   5295
   End
   Begin VB.TextBox Lote 
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
      MaxLength       =   30
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1455
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3720
      TabIndex        =   15
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
      TabIndex        =   16
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
      TabIndex        =   63
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
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consulta Revalida"
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
      Left            =   8400
      TabIndex        =   70
      Top             =   6240
      Width           =   1935
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
      TabIndex        =   64
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
      Left            =   5280
      TabIndex        =   62
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblensayo 
      BackColor       =   &H00FFFF00&
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
      Left            =   0
      TabIndex        =   60
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblDescri 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      Left            =   960
      TabIndex        =   59
      Top             =   960
      Width           =   4275
   End
   Begin VB.Label lblresultado 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
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
      Left            =   5400
      TabIndex        =   58
      Top             =   960
      Width           =   3000
   End
   Begin VB.Label Descri1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   57
      Top             =   1320
      Width           =   4400
   End
   Begin VB.Label descri2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   56
      Top             =   1800
      Width           =   4400
   End
   Begin VB.Label Descri3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   55
      Top             =   2280
      Width           =   4400
   End
   Begin VB.Label Descri4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   54
      Top             =   2760
      Width           =   4400
   End
   Begin VB.Label Descri5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   53
      Top             =   3240
      Width           =   4400
   End
   Begin VB.Label Descri6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   52
      Top             =   3720
      Width           =   4400
   End
   Begin VB.Label Descri7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   51
      Top             =   4200
      Width           =   4400
   End
   Begin VB.Label Descri8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   50
      Top             =   4680
      Width           =   4400
   End
   Begin VB.Label Descri9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   49
      Top             =   5160
      Width           =   4400
   End
   Begin VB.Label Descri10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   840
      TabIndex        =   48
      Top             =   5640
      Width           =   4400
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Obtenido"
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
      Left            =   8520
      TabIndex        =   47
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Std1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   46
      Top             =   1320
      Width           =   3000
   End
   Begin VB.Label Std2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   45
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Label Std3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   44
      Top             =   2280
      Width           =   3000
   End
   Begin VB.Label Std4 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   43
      Top             =   2760
      Width           =   3000
   End
   Begin VB.Label Std5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   42
      Top             =   3240
      Width           =   3000
   End
   Begin VB.Label Std6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   41
      Top             =   3720
      Width           =   3000
   End
   Begin VB.Label Std7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   40
      Top             =   4200
      Width           =   3000
   End
   Begin VB.Label Std8 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   39
      Top             =   4680
      Width           =   3000
   End
   Begin VB.Label Std9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   38
      Top             =   5160
      Width           =   3000
   End
   Begin VB.Label Std10 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   495
      Left            =   5400
      TabIndex        =   37
      Top             =   5640
      Width           =   3000
   End
   Begin VB.Label Ensayo1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   36
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Ensayo2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   35
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Ensayo3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Ensayo4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   33
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Ensayo5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Ensayo6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   31
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Ensayo7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Ensayo8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   29
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Ensayo9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Ensayo10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   120
      TabIndex        =   27
      Top             =   5640
      Width           =   615
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      Left            =   2640
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgRevalidady"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WGraba As String
Dim ZEnsayo(10) As Integer
Dim ZEnsayoActual(10) As Integer
Dim ZValor(10) As String
Dim WRevalida As Integer
Dim WProceso As Integer

Dim XMes As String
Dim XAno As String

Private Sub Cancela_click()
    PrgRevalidady.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()
    Lote.Text = ""
    Fecha.Text = "  /  /    "
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
End Sub

Private Sub RevalidaConsulta_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Revalida"
        ZSql = ZSql + " Where Lote = " + "'" + Lote.Text + "'"
        ZSql = ZSql + " and Revalida = " + "'" + RevalidaConsulta.Text + "'"
        spRevalida = ZSql
        Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
        If rstRevalida.RecordCount > 0 Then
        
            Fecha.Text = rstRevalida!Fecha
            Revalida.Text = rstRevalida!Revalida
            VtoAnterior.Text = "  /  /    "
            Vencimiento.Text = rstRevalida!Vencimiento
            
            Resultado.Text = rstRevalida!Resultado
            Observaciones.Text = rstRevalida!Observaciones
            Responsable.Text = rstRevalida!Responsable
            
            Ensayo1.Caption = rstRevalida!Codigo1
            Ensayo2.Caption = rstRevalida!Codigo2
            Ensayo3.Caption = rstRevalida!Codigo3
            Ensayo4.Caption = rstRevalida!Codigo4
            Ensayo5.Caption = rstRevalida!Codigo5
            Ensayo6.Caption = rstRevalida!Codigo6
            Ensayo7.Caption = rstRevalida!Codigo7
            Ensayo8.Caption = rstRevalida!Codigo8
            Ensayo9.Caption = rstRevalida!Codigo9
            Ensayo10.Caption = rstRevalida!Codigo10
            
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
            
            rstRevalida.Close
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
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri1.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri1.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            descri2.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            descri2.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri3.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri3.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri4.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri4.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri5.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri5.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri6.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri6.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri7.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri7.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri8.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri8.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri9.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri9.Caption = ""
        End If
            
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri10.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
                Else
            Descri10.Caption = ""
        End If
        
        Call Conecta_Empresa
    
        m$ = "Pulse ACEPTAR para finalizar la consulta"
        A% = MsgBox(m$, 0, "Consulta de Revalidas")
        
        Lote.Text = ""
        Fecha.Text = "  /  /    "
        Revalida.Text = ""
        VtoAnterior.Text = "  /  /    "
        Articulo.Text = "  -   -   "
        DesArticulo.Caption = ""
        Vencimiento.Text = "  /  /    "
    
        Ensayo1.Caption = ""
        Ensayo2.Caption = ""
        Ensayo3.Caption = ""
        Ensayo4.Caption = ""
        Ensayo5.Caption = ""
        Ensayo6.Caption = ""
        Ensayo7.Caption = ""
        Ensayo8.Caption = ""
        Ensayo9.Caption = ""
        Ensayo10.Caption = ""
        
        Descri1.Caption = ""
        descri2.Caption = ""
        Descri3.Caption = ""
        Descri4.Caption = ""
        Descri5.Caption = ""
        Descri6.Caption = ""
        Descri7.Caption = ""
        Descri8.Caption = ""
        Descri9.Caption = ""
        Descri10.Caption = ""
        
        Std1.Caption = ""
        Std2.Caption = ""
        Std3.Caption = ""
        Std4.Caption = ""
        Std5.Caption = ""
        Std6.Caption = ""
        Std7.Caption = ""
        Std8.Caption = ""
        Std9.Caption = ""
        Std10.Caption = ""
        
        Valor1.Text = ""
        valor2.Text = ""
        Valor3.Text = ""
        valor4.Text = ""
        valor5.Text = ""
        valor6.Text = ""
        valor7.Text = ""
        valor8.Text = ""
        valor9.Text = ""
        valor10.Text = ""
    
        Responsable.Text = ""
        Observaciones.Text = ""
        Resultado.Text = ""
        
        RevalidaConsulta.Text = ""
        
        Lote.SetFocus
    
    
    End If
    
End Sub

Private Sub Graba_Click()
    
    If WGraba <> "S" Then
    
        ZProceso = 0
        Call Ingresa_clave

               Else

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
        ZSql = ZSql + "'" + Ensayo1.Caption + "',"
        ZSql = ZSql + "'" + Ensayo2.Caption + "',"
        ZSql = ZSql + "'" + Ensayo3.Caption + "',"
        ZSql = ZSql + "'" + Ensayo4.Caption + "',"
        ZSql = ZSql + "'" + Ensayo5.Caption + "',"
        ZSql = ZSql + "'" + Ensayo6.Caption + "',"
        ZSql = ZSql + "'" + Ensayo7.Caption + "',"
        ZSql = ZSql + "'" + Ensayo8.Caption + "',"
        ZSql = ZSql + "'" + Ensayo9.Caption + "',"
        ZSql = ZSql + "'" + Ensayo10.Caption + "',"
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
        ZSql = ZSql + "'" + Valor1.Text + "',"
        ZSql = ZSql + "'" + valor2.Text + "',"
        ZSql = ZSql + "'" + Valor3.Text + "',"
        ZSql = ZSql + "'" + valor4.Text + "',"
        ZSql = ZSql + "'" + valor5.Text + "',"
        ZSql = ZSql + "'" + valor6.Text + "',"
        ZSql = ZSql + "'" + valor7.Text + "',"
        ZSql = ZSql + "'" + valor8.Text + "',"
        ZSql = ZSql + "'" + valor9.Text + "',"
        ZSql = ZSql + "'" + valor10.Text + "',"
        ZSql = ZSql + "'" + Revalida.Text + "')"
            
        spRevalida = ZSql
        Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        ZOrdVencimiento = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Laudo SET "
        ZSql = ZSql + "FechaVencimiento = " + "'" + Vencimiento.Text + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "',"
        ZSql = ZSql + "Revalida = " + "'" + Revalida.Text + "'"
        ZSql = ZSql + " Where PartiOri = " + "'" + Lote.Text + "'"
        
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        Call Cancela_click
    
    End If

End Sub

Private Sub Lote_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If Trim(Lote.Text) <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + Lote.Text + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
            
                Articulo.Text = rstLaudo!Articulo
                ZZFecha = rstLaudo!Fecha
                Revalida.Text = IIf(IsNull(rstLaudo!Revalida), "", rstLaudo!Revalida)
                Revalida.Text = Str$(Val(Revalida.Text) + 1)
                VtoAnterior.Text = IIf(IsNull(rstLaudo!fechavencimiento), "  /  /    ", rstLaudo!fechavencimiento)
                Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                rstLaudo.Close
                
                If VtoAnterior.Text = "  /  /    " Then
                
                    WVida = 0
                
                    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WVida = IIf(IsNull(rstArticulo!Meses), "0", rstArticulo!Meses)
                        rstArticulo.Close
                    End If
                        
                    WMes = Val(Mid$(ZZFecha, 4, 2))
                    WAno = Val(Right$(ZZFecha, 4))
                
                    For Ciclo = 1 To WVida
                        WMes = WMes + 1
                        If WMes > 12 Then
                            WAno = WAno + 1
                            WMes = 1
                        End If
                    Next Ciclo
                    If WVida <> 0 Then
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        VtoAnterior.Text = "01/" + XMes + "/" + XAno
                    End If
                
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
                    Ensayo1.Caption = rstEspecificacionesUnifica!Ensayo1
                    Ensayo2.Caption = rstEspecificacionesUnifica!Ensayo2
                    Ensayo3.Caption = rstEspecificacionesUnifica!Ensayo3
                    Ensayo4.Caption = rstEspecificacionesUnifica!Ensayo4
                    Ensayo5.Caption = rstEspecificacionesUnifica!Ensayo5
                    Ensayo6.Caption = rstEspecificacionesUnifica!Ensayo6
                    Ensayo7.Caption = rstEspecificacionesUnifica!Ensayo7
                    Ensayo8.Caption = rstEspecificacionesUnifica!Ensayo8
                    Ensayo9.Caption = rstEspecificacionesUnifica!Ensayo9
                    Ensayo10.Caption = rstEspecificacionesUnifica!Ensayo10
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
                    Std1.Caption = rstEspecificacionesUnifica!Valor1
                    Std2.Caption = rstEspecificacionesUnifica!valor2
                    Std3.Caption = rstEspecificacionesUnifica!Valor3
                    Std4.Caption = rstEspecificacionesUnifica!valor4
                    Std5.Caption = rstEspecificacionesUnifica!valor5
                    Std6.Caption = rstEspecificacionesUnifica!valor6
                    Std7.Caption = rstEspecificacionesUnifica!valor7
                    Std8.Caption = rstEspecificacionesUnifica!valor8
                    Std9.Caption = rstEspecificacionesUnifica!valor9
                    Std10.Caption = rstEspecificacionesUnifica!valor10
                    rstEspecificacionesUnifica.Close
                End If
                
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri1.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri1.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    descri2.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    descri2.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri3.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri3.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri4.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri4.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri5.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri5.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri6.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri6.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri7.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri7.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri8.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri8.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri9.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri9.Caption = ""
                End If
                    
                spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Caption + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    Descri10.Caption = rstEnsayo!Descripcion
                    rstEnsayo.Close
                        Else
                    Descri10.Caption = ""
                End If
                
                Call Conecta_Empresa
            
                Vencimiento.SetFocus
            
            End If
        
        End If
        
    End If
    
End Sub

Private Sub ValorOri2_Click()
End Sub

Private Sub ValorOri5_Click()

End Sub

Private Sub ValorOri8_Click()
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
        valor2.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor1.Text = ""
    End If
End Sub

Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor2.SelStart = 0
        Valor3.SetFocus
    End If
    If KeyAscii = 27 Then
        valor2.Text = ""
    End If
End Sub

Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Valor3.SelStart = 0
        valor4.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor3.Text = ""
    End If
End Sub

Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor4.SelStart = 0
        valor5.SetFocus
    End If
    If KeyAscii = 27 Then
        valor4.Text = ""
    End If
End Sub

Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor5.SelStart = 0
        valor6.SetFocus
    End If
    If KeyAscii = 27 Then
        valor5.Text = ""
    End If
End Sub

Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor6.SelStart = 0
        valor7.SetFocus
    End If
    If KeyAscii = 27 Then
        valor6.Text = ""
    End If
End Sub

Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor7.SelStart = 0
        valor8.SetFocus
    End If
    If KeyAscii = 27 Then
        valor7.Text = ""
    End If
End Sub

Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor8.SelStart = 0
        valor9.SetFocus
    End If
    If KeyAscii = 27 Then
        valor8.Text = ""
    End If
End Sub

Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor9.SelStart = 0
        valor10.SetFocus
    End If
    If KeyAscii = 27 Then
        valor9.Text = ""
    End If
End Sub

Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        valor10.SelStart = 0
        Resultado.SetFocus
    End If
    If KeyAscii = 27 Then
        valor10.Text = ""
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
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Composicion de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub

