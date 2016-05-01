VERSION 5.00
Begin VB.Form PrgPasa2 
   Caption         =   "Traspaso de Productos Terminados"
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
Attribute VB_Name = "PrgPasa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelar_Click()
    With rstTerminado
        .Close
    End With
    DbsVentas.Close
    PrgPasa2.Hide
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Open "A:" + WEmpresa + "ter.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WCodigo = Mid$(Linea, 1, 12)
        WDescripcion = Mid$(Linea, 13, 30)
        WLinea = Val(Mid$(Linea, 44, 4))
        WUnidad = Mid$(Linea, 49, 5)
        WInicial = Val(Mid$(Linea, 60, 10))
        WEntradas = Val(Mid$(Linea, 70, 10))
        WSalidas = Val(Mid$(Linea, 80, 10))
        WMinimo = Val(Mid$(Linea, 90, 10))
        WDeposito = ""
        WPedido = Val(Mid$(Linea, 114, 10))
        WEnvase = Val(Mid$(Linea, 153, 3))
        WEnvase1 = Val(Mid$(Linea, 126, 3))
        WEnvase2 = Val(Mid$(Linea, 130, 3))
        WEnvase3 = Val(Mid$(Linea, 134, 3))
        WEnvase4 = Val(Mid$(Linea, 138, 3))
        WEnvase5 = Val(Mid$(Linea, 142, 3))
        WEnvase6 = Val(Mid$(Linea, 146, 3))
        WProceso = Val(Mid$(Linea, 151, 10))
        
        With rstTerminado
        
            .Index = "Codigo"
            .Seek "=", Codigo
            If .NoMatch Then
                .AddNew
                !Codigo = WCodigo
                !Descripcion = WDescripcion
                !Linea = WLinea
                !Unidad = WUnidad
                !Inicial = WInicial
                !Entradas = WEntradas
                !Salidas = WSalidas
                !Minimo = WMinimo
                !Deposito = ""
                !Pedido = WPedido
                Rem !Envase = WEnvase
                !Envase1 = WEnvase1
                !Envase2 = WEnvase2
                !Envase3 = WEnvase3
                !Envase4 = WEnvase4
                !Envase5 = WEnvase5
                !Envase6 = WEnvase6
                !Proceso = WProceso
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.

End Sub


