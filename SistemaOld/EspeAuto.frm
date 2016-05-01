VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspeAuto 
   Caption         =   "Consulta de Versiones de Especificaciones de Producto Terminado"
   ClientHeight    =   8160
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   11685
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11685
   Begin VB.TextBox Version 
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
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   52
      Text            =   " "
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox FechaFinal 
      Enabled         =   0   'False
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
      Left            =   7800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   51
      Text            =   " "
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox FechaInicio 
      Enabled         =   0   'False
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
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   48
      Text            =   " "
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Valor1010 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   47
      Text            =   " "
      Top             =   6360
      Width           =   5655
   End
   Begin VB.TextBox Valor99 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   46
      Text            =   " "
      Top             =   5760
      Width           =   5655
   End
   Begin VB.TextBox Valor88 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   45
      Text            =   " "
      Top             =   5160
      Width           =   5655
   End
   Begin VB.TextBox Valor77 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   44
      Text            =   " "
      Top             =   4560
      Width           =   5655
   End
   Begin VB.TextBox Valor66 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   43
      Text            =   " "
      Top             =   3960
      Width           =   5655
   End
   Begin VB.TextBox Valor55 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   42
      Text            =   " "
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox Valor44 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   41
      Text            =   " "
      Top             =   2760
      Width           =   5655
   End
   Begin VB.TextBox Valor33 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      Top             =   2160
      Width           =   5655
   End
   Begin VB.TextBox Valor22 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   39
      Text            =   " "
      Top             =   1560
      Width           =   5655
   End
   Begin VB.TextBox Valor11 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   38
      Text            =   " "
      Top             =   960
      Width           =   5655
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
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
   Begin Crystal.CrystalReport Lista 
      Left            =   9840
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEspefUnifica.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   10560
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox valor10 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   35
      Text            =   " "
      Top             =   6120
      Width           =   5655
   End
   Begin VB.TextBox valor9 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   34
      Text            =   " "
      Top             =   5520
      Width           =   5655
   End
   Begin VB.TextBox valor8 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   4920
      Width           =   5655
   End
   Begin VB.TextBox valor7 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   32
      Text            =   " "
      Top             =   4320
      Width           =   5655
   End
   Begin VB.TextBox valor6 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   3720
      Width           =   5655
   End
   Begin VB.TextBox valor5 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   30
      Text            =   " "
      Top             =   3120
      Width           =   5655
   End
   Begin VB.TextBox valor4 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   2520
      Width           =   5655
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
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   28
      Text            =   " "
      Top             =   1920
      Width           =   5655
   End
   Begin VB.TextBox valor2 
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
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   27
      Text            =   " "
      Top             =   1250
      Width           =   5655
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
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   26
      Text            =   " "
      Top             =   720
      Width           =   5655
   End
   Begin VB.TextBox Ensayo10 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   25
      Text            =   " "
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox Ensayo9 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   24
      Text            =   " "
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Ensayo8 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   23
      Text            =   " "
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Ensayo7 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   22
      Text            =   " "
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Ensayo6 
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
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   21
      Text            =   " "
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Ensayo5 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   20
      Text            =   " "
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Ensayo4 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   19
      Text            =   " "
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Ensayo3 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   18
      Text            =   " "
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Ensayo2 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   17
      Text            =   " "
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Ensayo1 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   16
      Text            =   " "
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9960
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Version"
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
      Index           =   0
      Left            =   3360
      TabIndex        =   50
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
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
      Index           =   2
      Left            =   5640
      TabIndex        =   49
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
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
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descri10 
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
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   6120
      Width           =   4980
   End
   Begin VB.Label Descri9 
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
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   5520
      Width           =   4980
   End
   Begin VB.Label Descri8 
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
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   4920
      Width           =   4980
   End
   Begin VB.Label Descri7 
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
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   4320
      Width           =   4980
   End
   Begin VB.Label Descri6 
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
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   3720
      Width           =   4980
   End
   Begin VB.Label Descri5 
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
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   3120
      Width           =   4980
   End
   Begin VB.Label Descri4 
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
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2520
      Width           =   4980
   End
   Begin VB.Label Descri3 
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
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   4980
   End
   Begin VB.Label descri2 
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
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1320
      Width           =   4980
   End
   Begin VB.Label Descri1 
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
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   720
      Width           =   4980
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
      Left            =   5880
      TabIndex        =   5
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblDescri 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label lblensayo 
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
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgEspeAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim EspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim XParam As String
Dim ZFecha As String
Dim CargaEmpresa(12, 2) As String

Private Sub cmdClose_Click()
    PrgEspeAuto.Hide
    Unload Me
    PrgHoja.Show
End Sub

Private Sub Form_Load()
        
    Producto.Text = ZTerminado
    Version.Text = ZVersion
    
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
    
    Producto.Text = UCase(Producto.Text)
    Sql1 = "Select *"
    Sql2 = " FROM Terminado"
    Sql3 = " Where Terminado.Codigo = " + "'" + Producto.Text + "'"
    spTerminado = Sql1 + Sql2 + Sql3
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        XVersion = IIf(IsNull(rstTerminado!VersionI), "0", rstTerminado!VersionI)
        rstTerminado.Close
    End If


    If Val(XVersion) <> Val(Version.Text) Then
    
    
        Sql1 = "Select *"
        Sql2 = " FROM EspecifUnificaVersion"
        Sql3 = " Where EspecifUnificaVersion.Producto = " + "'" + Producto.Text + "'"
        Sql4 = " and EspecifUnificaVersion.Version = " + "'" + Version.Text + "'"
        spEspecifUnificaVersion = Sql1 + Sql2 + Sql3 + Sql4
        Set EspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
        If EspecifUnificaVersion.RecordCount > 0 Then
    
            Ensayo1.Text = EspecifUnificaVersion!Ensayo1
            Ensayo2.Text = EspecifUnificaVersion!Ensayo2
            Ensayo3.Text = EspecifUnificaVersion!Ensayo3
            Ensayo4.Text = EspecifUnificaVersion!Ensayo4
            Ensayo5.Text = EspecifUnificaVersion!Ensayo5
            Ensayo6.Text = EspecifUnificaVersion!Ensayo6
            Ensayo7.Text = EspecifUnificaVersion!Ensayo7
            Ensayo8.Text = EspecifUnificaVersion!Ensayo8
            Ensayo9.Text = EspecifUnificaVersion!Ensayo9
            Ensayo10.Text = EspecifUnificaVersion!Ensayo10
            
            Valor1.Text = EspecifUnificaVersion!Valor1
            valor2.Text = EspecifUnificaVersion!valor2
            Valor3.Text = EspecifUnificaVersion!Valor3
            valor4.Text = EspecifUnificaVersion!valor4
            valor5.Text = EspecifUnificaVersion!valor5
            valor6.Text = EspecifUnificaVersion!valor6
            valor7.Text = EspecifUnificaVersion!valor7
            valor8.Text = EspecifUnificaVersion!valor8
            valor9.Text = EspecifUnificaVersion!valor9
            valor10.Text = EspecifUnificaVersion!valor10
            Valor11.Text = IIf(IsNull(EspecifUnificaVersion!Valor11), "", EspecifUnificaVersion!Valor11)
            Valor22.Text = IIf(IsNull(EspecifUnificaVersion!Valor22), "", EspecifUnificaVersion!Valor22)
            Valor33.Text = IIf(IsNull(EspecifUnificaVersion!Valor33), "", EspecifUnificaVersion!Valor33)
            Valor44.Text = IIf(IsNull(EspecifUnificaVersion!Valor44), "", EspecifUnificaVersion!Valor44)
            Valor55.Text = IIf(IsNull(EspecifUnificaVersion!Valor55), "", EspecifUnificaVersion!Valor55)
            Valor66.Text = IIf(IsNull(EspecifUnificaVersion!Valor66), "", EspecifUnificaVersion!Valor66)
            Valor77.Text = IIf(IsNull(EspecifUnificaVersion!Valor77), "", EspecifUnificaVersion!Valor77)
            Valor88.Text = IIf(IsNull(EspecifUnificaVersion!Valor88), "", EspecifUnificaVersion!Valor88)
            Valor99.Text = IIf(IsNull(EspecifUnificaVersion!Valor99), "", EspecifUnificaVersion!Valor99)
            Valor1010.Text = IIf(IsNull(EspecifUnificaVersion!Valor1010), "", EspecifUnificaVersion!Valor1010)
            
            FechaInicio.Text = EspecifUnificaVersion!FechaInicio
            FechaFinal.Text = EspecifUnificaVersion!FechaFinal
        
            EspecifUnificaVersion.Close
        
        End If
        
            Else
            
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM EspecifUnifica"
        ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
        spEspecifUnifica = ZSql
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
        
            Ensayo1.Text = rstEspecifUnifica!Ensayo1
            Ensayo2.Text = rstEspecifUnifica!Ensayo2
            Ensayo3.Text = rstEspecifUnifica!Ensayo3
            Ensayo4.Text = rstEspecifUnifica!Ensayo4
            Ensayo5.Text = rstEspecifUnifica!Ensayo5
            Ensayo6.Text = rstEspecifUnifica!Ensayo6
            Ensayo7.Text = rstEspecifUnifica!Ensayo7
            Ensayo8.Text = rstEspecifUnifica!Ensayo8
            Ensayo9.Text = rstEspecifUnifica!Ensayo9
            Ensayo10.Text = rstEspecifUnifica!Ensayo10
            
            Valor1.Text = rstEspecifUnifica!Valor1
            valor2.Text = rstEspecifUnifica!valor2
            Valor3.Text = rstEspecifUnifica!Valor3
            valor4.Text = rstEspecifUnifica!valor4
            valor5.Text = rstEspecifUnifica!valor5
            valor6.Text = rstEspecifUnifica!valor6
            valor7.Text = rstEspecifUnifica!valor7
            valor8.Text = rstEspecifUnifica!valor8
            valor9.Text = rstEspecifUnifica!valor9
            valor10.Text = rstEspecifUnifica!valor10
            Valor11.Text = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
            Valor22.Text = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
            Valor33.Text = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
            Valor44.Text = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
            Valor55.Text = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
            Valor66.Text = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
            Valor77.Text = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
            Valor88.Text = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
            Valor99.Text = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
            Valor1010.Text = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
            
            FechaInicio.Text = rstEspecifUnifica!Fecha
            FechaFinal.Text = rstEspecifUnifica!Fecha
            
            rstEspecifUnifica.Close
        
        End If
        
    End If
            
    
    
    
    
    
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri2.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If
    
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
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub





