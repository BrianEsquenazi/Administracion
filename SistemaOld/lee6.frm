VERSION 5.00
Begin VB.Form Prglee6 
   Caption         =   "Generacion de traspaso de datos"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      Caption         =   "Control de Grabacion"
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Prglee6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLinea As String

Private Sub Acepta_Click()
 
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prglee6.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
            With rstProveedor
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        .Edit
                        !provincia = 1
                        !Iva = 3
                        .Update
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With






    Call Cancela_click


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
    
Error:
Stop
     coderr = Err
     
     Resume Next
     
End Sub




