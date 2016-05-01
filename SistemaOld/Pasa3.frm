VERSION 5.00
Begin VB.Form PrgPasa3 
   Caption         =   "Traspaso de Precios de Clientes"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "PrgPasa3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    With rstPrecios
        .Close
    End With
    DbsVentas.Close
    PrgPasa3.Hide
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Open "A:" + WEmpresa + "prec.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WCliente = Mid$(Linea, 1, 6)
        WTerminado = Mid$(Linea, 8, 12)
        WPrecio = Val(Mid$(Linea, 52, 10))
        WDescripcion = Mid$(Linea, 21, 30)
        
        With rstPrecios
        
            .Index = "Clave"
            .Seek "=", WCliente + WTerminado
            If .NoMatch Then
                .AddNew
                !Cliente = WCliente
                !Terminado = WTerminado
                !Precio = WPrecio
                !Descripcion = WDescripcion
                !Clave = !Cliente + !Terminado
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.

End Sub


