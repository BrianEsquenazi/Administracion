VERSION 5.00
Begin VB.Form Prgrepro 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cash Flow"
   ClientHeight    =   6330
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   6135
   Begin VB.CommandButton aceptar 
      Caption         =   "aceptar"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Prgrepro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WSaldo As Double
Private Wvencimiento As String
Private WCliente As String

Private Sub aceptar_Click()

    'PRUEBAS DE MATERIAS PRIMAS
        
    coderr = 0
    
    Open "c:\prueba\ventas\" + WEmpresa + "xx.txt" For Input As #1
    
    Do
                Line Input #1, WLinea

                WCliente = Mid$(WLinea, 1, 6)
                WProducto = Mid$(WLinea, 8, 2) + "-" + Mid$(WLinea, 11, 5) + "-" + Mid$(WLinea, 17, 3)
                WFecha = Mid$(WLinea, 21, 2) + "/" + Mid$(WLinea, 23, 2) + "/19" + Mid$(WLinea, 25, 2)
                Wfactura = Mid$(WLinea, 28, 6)
                Wprecio = Val(Mid$(WLinea, 34, 11))
                Wcantidad = Val(Mid$(WLinea, 45, 10))
                
                With rstPrecios
                        .Index = "Clave"
                        .Seek "=", WCliente + WProducto
                        If .NoMatch Then
                            .AddNew
                            !Cliente = WCliente
                            !Terminado = WProducto
                            !Fecha1 = "  /  /    "
                            !Factura1 = 0
                            !Precio1 = 0
                            !Cantidad1 = 0
                            !Fecha2 = "  /  /    "
                            !Factura2 = 0
                            !Precio2 = 0
                            !Cantidad2 = 0
                            !Fecha3 = "  /  /    "
                            !Factura3 = 0
                            !Precio3 = 0
                            !Cantidad3 = 0
                            !Fecha4 = "  /  /    "
                            !Factura4 = 0
                            !Precio4 = 0
                            !Cantidad4 = 0
                            !Fecha5 = WFecha
                            !Factura5 = Wfactura
                            !Precio5 = Wprecio
                            !Cantidad5 = Wcantidad
                            !Clave = !Cliente + !Terminado
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                                If !Cantidad5 = 0 Then
                                    !Fecha5 = WFecha
                                    !Factura5 = Wfactura
                                    !Precio5 = Wprecio
                                    !Cantidad5 = Wcantidad
                                        Else
                                    If !Cantidad4 = 0 Then
                                        !Fecha4 = WFecha
                                        !Factura4 = Wfactura
                                        !Precio4 = Wprecio
                                        !Cantidad4 = Wcantidad
                                            Else
                                        If !Cantidad3 = 0 Then
                                            !Fecha3 = WFecha
                                            !Factura3 = Wfactura
                                            !Precio3 = Wprecio
                                            !Cantidad3 = Wcantidad
                                                Else
                                            If !Cantidad2 = 0 Then
                                                !Fecha2 = WFecha
                                                !Factura2 = Wfactura
                                                !Precio2 = Wprecio
                                                !Cantidad2 = Wcantidad
                                                    Else
                                                If !Cantidad1 = 0 Then
                                                    !Fecha1 = WFecha
                                                    !Factura1 = Wfactura
                                                    !Precio1 = Wprecio
                                                    !Cantidad1 = Wcantidad
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1

        
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    With rstPrecios
        .Close
    End With
    DbsAdminis.Close
    Prgrepro.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DADA()
Rem terminado DADA

    Open "c:\prueba\ventas\" + WEmpresa + "ter.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WCodigo = Mid$(Linea, 1, 12)
        WLinea = Val(Mid$(Linea, 13, 4))
        WUnidad = Mid$(Linea, 18, 5)
        WInicial = Val(Mid$(Linea, 29, 10))
        WEntradas = Val(Mid$(Linea, 39, 11))
        WSalidas = Val(Mid$(Linea, 50, 11))
        WMinimo = Val(Mid$(Linea, 61, 11))
        WDeposito = ""
        WProceso = Val(Mid$(Linea, 72, 11))
        WPedido = Val(Mid$(Linea, 83, 11))
        WEnvase1 = Val(Mid$(Linea, 95, 3))
        WEnvase2 = Val(Mid$(Linea, 99, 3))
        WEnvase3 = Val(Mid$(Linea, 103, 3))
        WEnvase4 = Val(Mid$(Linea, 107, 3))
        WEnvase5 = Val(Mid$(Linea, 111, 3))
        WEnvase6 = Val(Mid$(Linea, 115, 3))
        WEnvase = Val(Mid$(Linea, 119, 3))
        WDescripcion = Mid$(Linea, 135, 30)
        
        With rstTerminado
        
            .Index = "Codigo"
            .Seek "=", WCodigo
            If .NoMatch Then
                .AddNew
                !Codigo = WCodigo
                !Descripcion = WDescripcion
                !Linea = WLinea
                !Unidad = WUnidad
                !Inicial = Val(WInicial)
                !Entradas = Val(WEntradas)
                !Salidas = Val(WSalidas)
                !MInimo = Val(WMinimo)
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
                    Else
                .Edit
                !Codigo = WCodigo
                !Descripcion = WDescripcion
                !Linea = WLinea
                !Unidad = WUnidad
                !Inicial = Val(WInicial)
                Rem !Entradas = Val(WEntradas)
                Rem !Salidas = Val(WSalidas)
                !MInimo = Val(WMinimo)
                !Deposito = ""
                !Pedido = WPedido
                Rem !Envase = WEnvase
                !Envase1 = WEnvase1
                !Envase2 = WEnvase2
                !Envase3 = WEnvase3
                !Envase4 = WEnvase4
                !Envase5 = WEnvase5
                !Envase6 = WEnvase6
                Rem !Proceso = WProceso
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.
    
    
        With rstCtaCte
    
        .Index = "Cliente"
        .MoveFirst
        If .NoMatch = False Then
            Do
                .Edit
                
                If Right$(!Vencimiento, 4) = "1900" Then
                    !Vencimiento = Left$(!Vencimiento, 6) + "2000"
                    !OrdVencimiento = "2000" + Right$(!OrdVencimiento, 4)
                End If
                If Right$(!Vencimiento1, 4) = "1900" Then
                    !Vencimiento1 = Left$(!Vencimiento1, 6) + "2000"
                    !OrdVencimiento1 = "2000" + Right$(!OrdVencimiento1, 4)
                End If
                .Update
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With



End Sub
