VERSION 5.00
Begin VB.Form PrgLimpter 
   Caption         =   "Proceso de Limpieza de Stock de Productos Terminados"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   " "
      Height          =   1935
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   4815
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   1560
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "PrgLimpter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    With rstTerminado
        .Index = "Codigo"
        .MoveFirst
        Do
            If .EOF = False Then
                .Edit
                !Inicial = 0
                !Entradas = 0
                !Salidas = 0
                !Proceso = 0
                .Update
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    Call Cancela_click

End Sub

Private Sub Cancela_click()

    With rstTerminado
        .Close
    End With
    
    DbsVentas.Close
    
    PrgLimpter.Hide
    Unload Me
    Menu.Show
End Sub



Private Sub Form_Activate()
    OPEN_FILE_TERMINADO
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgLimpter.Caption = "Proceso de Limpieza de Stock de Productos Terminados :  " + !Nombre
        End If
    End With

End Sub
