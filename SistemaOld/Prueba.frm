VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form prgPrueba 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid dada 
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3201
      _Version        =   327680
      Rows            =   4000
      Cols            =   7
   End
End
Attribute VB_Name = "prgPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    dada.Clear
    
    For da = 1 To 3000
        For da1 = 1 To 6
            dada.Col = da1
            dada.Row = da
            dada.Text = Str$(da) + Str$(da1)
        Next da1
    Next da
End Sub
