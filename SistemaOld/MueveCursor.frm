VERSION 5.00
Begin VB.Form MueveCursor 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ValorY 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox ValorX 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.Shape dada 
      Height          =   1215
      Left            =   960
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "MueveCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    a1 = dada.Top
    a2 = dada.Left
    a3 = dada.Height + dada.Top
    a4 = dada.Width + dada.Left
    If X >= a2 And X <= a4 Then
        If Y >= a1 And Y <= a3 Then
            Rem mueve el dibujo
            Rem esta parte tenes que trabajarla un poco mas
            Rem para que se mueva por toda la pantalla
            Rem sin que salga
            dada.Top = dada.Top + 2000
            dada.Left = dada.Left + 2000
        End If
    End If
End Sub

