VERSION 5.00
Begin VB.Form PrgForm1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2520
      Picture         =   "PrgForm1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   1320
   End
End
Attribute VB_Name = "PrgForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DeltaX, DeltaY As Integer   ' Declara variables.
Private Sub Timer1_Timer()
    Picture1.Move Picture1.Left + DeltaX, Picture1.Top + DeltaY
    If Picture1.Left < ScaleLeft Then DeltaX = 100
    If Picture1.Left + Picture1.Width > ScaleWidth + ScaleLeft Then
        DeltaX = -100
    End If
    If Picture1.Top < ScaleTop Then DeltaY = 100
    If Picture1.Top + Picture1.Height > ScaleHeight + ScaleTop Then
        DeltaY = -100
    End If
End Sub

Private Sub Form_Load()

Timer1.Interval = 1000  ' Establece el intervalo.
    DeltaX = 100    ' Inicializa variables.
    DeltaY = 100
End Sub
