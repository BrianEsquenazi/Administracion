VERSION 5.00
Begin VB.Form prevar 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Texto 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "prevar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call ciclo
End Sub

Private Sub Command2_Click()
        Texto.Caption = "aaaaaa"
End Sub

Private Sub ciclo()
        Texto.Caption = Time
End Sub


