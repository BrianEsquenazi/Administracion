VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11625
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   9
      Text            =   " "
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   6
      Text            =   " "
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   285
      Left            =   7560
      TabIndex        =   4
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Trabajo"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
