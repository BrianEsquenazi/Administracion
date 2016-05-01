VERSION 5.00
Begin VB.Form Prglee4 
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
Attribute VB_Name = "Prglee4"
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
    Prglee4.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
Stop
    'envases
        
    coderr = 0
    
    Open "c:\prueba\ventas\" + WEmpresa + "env.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                WEnvases = Mid$(WLinea, 5, 3)
                WDescripcion = Mid$(WLinea, 8, 30)
                WAbreviatura = Mid$(WLinea, 38, 10)
                WKilos = Mid$(WLinea, 48, 10)
                
                With rstEnvases
                        .Index = "envases"
                        .Seek "=", Val(WEnvases)
                        If .NoMatch Then
                            .AddNew
                            !Envases = WEnvases
                            !Descripcion = WDescripcion
                            !Abreviatura = WAbreviatura
                            !Kilos = WKilos
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Envases = WEnvases
                            !Descripcion = WDescripcion
                            !Abreviatura = WAbreviatura
                            !Kilos = WKilos
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1

    coderr = 0
    
    Open "c:\prueba\ventas\" + WEmpresa + "LIN.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                WCodigo = Mid$(WLinea, 5, 4)
                WDescripcion = Mid$(WLinea, 9, 30)
                
                With rstLineas
                        .Index = "linea"
                        .Seek "=", Val(WCodigo)
                        If .NoMatch Then
                            .AddNew
                            !Linea = WCodigo
                            !Nombre = WDescripcion
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Linea = WCodigo
                            !Nombre = WDescripcion
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1


    coderr = 0
    
    Open "c:\prueba\ventas\" + WEmpresa + "rub.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                WCodigo = Mid$(WLinea, 5, 4)
                WDescripcion = Mid$(WLinea, 9, 30)
                
                With rstRubros
                        .Index = "rubro"
                        .Seek "=", Val(WCodigo)
                        If .NoMatch Then
                            .AddNew
                            !Rubro = WCodigo
                            !Nombre = WDescripcion
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Rubro = WCodigo
                            !Nombre = WDescripcion
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1


    coderr = 0
    
    Open "c:\prueba\ventas\" + WEmpresa + "vend.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                WCodigo = Mid$(WLinea, 5, 4)
                WDescripcion = Mid$(WLinea, 9, 30)
                
                With rstVendedores
                        .Index = "vendedor"
                        .Seek "=", Val(WCodigo)
                        If .NoMatch Then
                            .AddNew
                            !Vendedor = WCodigo
                            !Nombre = WDescripcion
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Vendedor = WCodigo
                            !Nombre = WDescripcion
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1

 Rem clientes


    Open "c:\prueba\ventas\" + WEmpresa + "clie.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        Cliente = Mid$(Linea, 1, 6)
        Razon = Mid$(Linea, 8, 40)
        Direccion = Mid$(Linea, 49, 40)
        Localidad = Mid$(Linea, 90, 40)
        Postal = Mid$(Linea, 131, 4)
        Telefono = Mid$(Linea, 138, 15) + Mid$(Linea, 255, 15)
        Contacto = ""
        Observaciones = Mid$(Linea, 270, 50)
        Cuit = Mid$(Linea, 156, 15)
        Vendedor = Val(Mid$(Linea, 182, 4))
        email = ""
        fax = ""
        Rubro = Val(Mid$(Linea, 177, 4))
        Horario = ""
        Pago1 = Val(Mid$(Linea, 172, 4))
        pago2 = Val(Mid$(Linea, 232, 4))
        Limite = Val(Mid$(Linea, 237, 9))
        MInimo = Val(Mid$(Linea, 246, 9))
        DirEntrega = Mid$(Linea, 191, 40)
        
        provincia = "1"
        Select Case Mid$(Linea, 136, 1)
            Case "C"
                provincia = "0"
            Case "B"
                provincia = "1"
            Case "K"
                provincia = "2"
            Case "X"
                provincia = "3"
            Case "W"
                provincia = "4"
            Case "H"
                provincia = "5"
            Case "U"
                provincia = "6"
            Case "E"
                provincia = "7"
            Case "P"
                provincia = "8"
            Case "Y"
                provincia = "9"
            Case "L"
                provincia = "10"
            Case "F"
                provincia = "11"
            Case "M"
                provincia = "12"
            Case "N"
                provincia = "13"
            Case "Q"
                provincia = "14"
            Case "R"
                provincia = "15"
            Case "A"
                provincia = "16"
            Case "J"
                provincia = "17"
            Case "D"
                provincia = "18"
            Case "Z"
                provincia = "19"
            Case "S"
                provincia = "20"
            Case "G"
                provincia = "21"
            Case "T"
                provincia = "22"
            Case "V"
                provincia = "23"
            Case Else
                provincia = "24"
        End Select
        
        Select Case Val(Mid$(Linea, 154, 1))
            Case 0
                Iva = "3"
            Case 1
                Iva = "1"
            Case 2
                Iva = "2"
            Case 3
                Iva = "4"
            Case 4
                Iva = "6"
            Case Else
                Iva = "5"
        End Select
        
        With rstClientes
        
            .Index = "Cliente"
            .Seek "=", Cliente
            If .NoMatch Then
                .AddNew
                !Cliente = Cliente
                !Razon = Razon
                !Direccion = Direccion
                !Localidad = Localidad
                !Postal = Postal
                !Telefono = Left$(Telefono, 20)
                !Contacto = Contacto
                !Observaciones = Observaciones
                !Cuit = Cuit
                !Vendedor = Vendedor
                !email = email
                !fax = fax
                !Rubro = Rubro
                !Horario = Horario
                !Pago1 = Pago1
                !pago2 = pago2
                !Limite = Limite
                !MInimo = MInimo
                !DirEntrega = DirEntrega
                !provincia = provincia
                !Iva = Iva
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.

Stop

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
            .Seek "=", Codigo
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
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.

Rem precios
Stop

    Open "c:\prueba\ventas\" + WEmpresa + "prec.txt" For Input As #1
    
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

Stop


Rem matreia prima


    Open "c:\prueba\ventas\" + WEmpresa + "art.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WProducto = Mid$(Linea, 1, 10)
        WDescripcion = Mid$(Linea, 11, 40)
        WCosto1 = Val(Mid$(Linea, 51, 11))
        WCosto2 = Val(Mid$(Linea, 62, 11))
        WInicial = Val(Mid$(Linea, 73, 11))
        WEntradas = Val(Mid$(Linea, 84, 11))
        WSalidas = Val(Mid$(Linea, 95, 11))
        WMinimo = Val(Mid$(Linea, 106, 11))
        WUnidad = Mid$(Linea, 117, 10)
        WDeposito = Mid$(Linea, 130, 10)
        WLaboratorio = Val(Mid$(Linea, 157, 11))
        WPedido = Val(Mid$(Linea, 168, 11))
        WEnvase = Mid$(Linea, 179, 4)
        WRs = Mid$(Linea, 184, 1)
        WFlete = Val(Mid$(Linea, 183, 13))
        WMoneda = Mid$(Linea, 196, 3)
        
        With rstArticulo
        
            .Index = "Codigo"
            .Seek "=", Codigo
            If .NoMatch Then
                .AddNew
                !Codigo = WProducto
                !Descripcion = WDescripcion
                !Costo1 = WCosto1
                !Costo2 = WCosto2
                !Inicial = WInicial
                !Entradas = WEntradas
                !Salidas = WSalidas
                !MInimo = WMinimo
                !Unidad = WUnidad
                !Deposito = WDeposito
                !Laboratorio = WLaboratorio
                !Pedido = WPedido
                !Envase = Val(WEnvase)
                !Rs = WRs
                !Flete = WFlete
                !Moneda = WMoneda
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.





Rem FOR,MULA


    Open "c:\prueba\ventas\" + WEmpresa + "com.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WTerminado = Mid$(Linea, 1, 12)
        WRenglon = Mid$(Linea, 13, 2)
        If Mid$(Linea, 15, 1) = "T" Then
            WTipo = "T"
                Else
            WTipo = "M"
        End If
        If WTipo = "M" Then
             WArticulo1 = Mid$(Linea, 16, 10)
             WArticulo2 = "  -     -   "
                Else
            WArticulo2 = Mid$(Linea, 26, 12)
            WArticulo1 = "  -   -  "
        End If
        WCantidad = Val(Mid$(Linea, 38, 11)) / 100
        WClave = WTerminado + WRenglon
        
        With rstComposicion
        
            .Index = "clave"
            .Seek "=", WClave
            If .NoMatch Then
                .AddNew
                !Terminado = WTerminado
                !Renglon = WRenglon
                !Tipo = WTipo
                !Articulo1 = WArticulo1
                !Articulo2 = WArticulo2
                !Cantidad = WCantidad
                !Clave = WClave
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.







Rem estadistica


    Open "c:\prueba\ventas\" + WEmpresa + "est.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WTipo = Mid$(Linea, 1, 2)
        WNUmero = Mid$(Linea, 3, 6)
        WRenglon = Mid$(Linea, 9, 2)
        WArticulo = Mid$(Linea, 11, 12)
        WCantidad = Val(Mid$(Linea, 23, 9))
        WPrecio = Val(Mid$(Linea, 32, 9))
        WPrecioUs = WPrecio
        WImporte = WPrecio * WCantidad
        WImporteUs = WPrecio * WCantidad
        WCliente = Mid$(Linea, 41, 6)
        WParidad = 1
        WVendedor = Mid$(Linea, 47, 4)
        WRubro = Mid$(Linea, 51, 4)
        WLinea = Mid$(Linea, 55, 4)
        WCosto1 = Val(Mid$(Linea, 59, 9))
        WCosto2 = Val(Mid$(Linea, 59, 9))
        WCoeficiente = 0
        WPedido = Mid$(Linea, 68, 6)
        WFecha = Mid$(Linea, 78, 2) + "/" + Mid$(Linea, 76, 2) + "/19" + Mid$(Linea, 74, 2)
        WImporte1 = 0
        WImporte2 = 0
        WImporte3 = 0
        WImporte4 = 0
        WOrdfecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WWArticulo = Left$(WArticulo, 8)
        WRemito = 0
        WClave = WTipo + WNUmero + WRenglon
        
        With rstEstadistica
        
            .Index = "clave"
            .Seek "=", WClave
            If .NoMatch Then
                .AddNew
                !Tipo = WTipo
                !Numero = WNUmero
                !Renglon = WRenglon
                !Articulo = WArticulo
                !Cantidad = WCantidad
                !Precio = WPrecio
                !PrecioUs = WPrecioUs
                !Importe = WImporte
                !ImporteUs = WImporteUs
                !Cliente = WCliente
                !Paridad = WParidad
                !Vendedor = WVendedor
                !Rubro = WRubro
                !Linea = WLinea
                !Costo1 = WCosto1
                !Costo2 = WCosto2
                !Coeficiente = WCoeficiente
                !Pedido = Val(WPedido)
                !Fecha = WFecha
                !Importe1 = 0
                !Importe2 = 0
                !Importe3 = 0
                !Importe4 = 0
                !OrdFecha = WOrdfecha
                !WArticulo = WWArticulo
                !Remito = WRemito
                !Clave = WClave
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.





Rem pedido


    Open "c:\prueba\ventas\" + WEmpresa + "ped.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WPedido = Mid$(Linea, 1, 6)
        WRenglon = Mid$(Linea, 7, 2)
        WFecha = Mid$(Linea, 9, 2) + "/" + Mid$(Linea, 11, 2) + "/19" + Mid$(Linea, 13, 2)
        WCliente = Mid$(Linea, 15, 6)
        WFecEntrega = Mid$(Linea, 21, 2) + "/" + Mid$(Linea, 23, 2) + "/19" + Mid$(Linea, 25, 2)
        WHora = Mid$(Linea, 27, 5)
        WTerminado = Mid$(Linea, 36, 12)
        WCantidad = Val(Mid$(Linea, 48, 9))
        WPrecio = Val(Mid$(Linea, 57, 9))
        WWPrecio = Mid$(Linea, 57, 9)
WFacturado = Val(Mid$(Linea, 66, 9))
        WLinea = Mid$(Linea, 75, 4)
        WObservaciones = Mid$(Linea, 83, 40)
        WEnvase1 = 0
        WCanti1 = 0
        WEnvase2 = 0
        WCanti2 = 0
        WEnvase3 = 0
        WCanti3 = 0
        WEnvase4 = 0
        WCanti4 = 0
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WImporte = 0
        WClave = WPedido + WRenglon
        
        With rstPedido
        
            .Index = "clave"
            .Seek "=", WClave
            If .NoMatch Then
                .AddNew
                !Pedido = WPedido
                !Renglon = WRenglon
                !Cliente = WCliente
                !Fecha = WFecha
                !FecEntrega = WFecEntrega
                !Hora = WHora
                !Observaciones = WObservaciones
                !Terminado = WTerminado
                !Cantidad = WCantidad
                !Envase1 = 0
                !Canti1 = 0
                !Envase2 = 0
                !Canti2 = 0
                !Envase3 = 0
                !Canti3 = 0
                !Envase4 = 0
                !Canti4 = 0
                !fechaord = WFechaord
                !Precio = WPrecio
                !Linea = WLinea
                !Facturado = WFacturado
                !Importe = WImporte
                !Clave = WClave
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.


Stop

Rem cuenta corriente


    Open "c:\prueba\ventas\" + WEmpresa + "cta.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WTipo = Mid$(Linea, 1, 2)
        WNUmero = Mid$(Linea, 3, 6)
        WRenglon = Mid$(Linea, 9, 2)
        WCliente = Mid$(Linea, 17, 6)
        WFecha = Mid$(Linea, 15, 2) + "/" + Mid$(Linea, 13, 2) + "/19" + Mid$(Linea, 11, 2)
        WNeto = Val(Mid$(Linea, 23, 9))
        WIva1 = Val(Mid$(Linea, 32, 9))
        WIva2 = Val(Mid$(Linea, 41, 9))
        WTotal = Val(Mid$(Linea, 59, 9))
        WTotalUs = Val(Mid$(Linea, 68, 9))
        WSaldo = Val(Mid$(Linea, 77, 9))
        WSaldoUs = Val(Mid$(Linea, 86, 9))
        If Val(Mid$(Linea, 95, 2)) <> 0 Then
            WVencimiento = Mid$(Linea, 95, 2) + "/" + Mid$(Linea, 97, 2) + "/19" + Mid$(Linea, 99, 2)
                Else
            WVencimiento = "  /  /    "
        End If
        WPedido = Mid$(Linea, 101, 6)
        WVendedor = Mid$(Linea, 107, 4)
        WRubro = Mid$(Linea, 111, 4)
        WRemito = Mid$(Linea, 121, 8)
        WOrden = Mid$(Linea, 129, 10)
        If Val(Mid$(Linea, 115, 2)) <> 0 Then
            WVencimiento1 = Mid$(Linea, 115, 2) + "/" + Mid$(Linea, 117, 2) + "/19" + Mid$(Linea, 119, 2)
                Else
            WVencimiento1 = "  /  /    "
        End If
        WEstado = "0"
        WOrdfecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WOrdvencimiento = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
        WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
        Select Case Val(WTipo)
            Case 1
                WImpre = "FC"
            Case 2
                WImpre = "Dv"
                WNeto = WNeto * -1
                WIva1 = WIva1 * -1
                WIva2 = WIva2 * -1
                WTotal = WTotal * -1
                WTotalUs = WTotalUs * -1
                WSaldo = WSaldo * -1
                WSaldoUs = WSaldoUs * -1
            Case 3
                WImpre = "Fc"
            Case 4
                WImpre = "Nd"
            Case 5
                WImpre = "Nc"
                WNeto = WNeto * -1
                WIva1 = WIva1 * -1
                WIva2 = WIva2 * -1
                WTotal = WTotal * -1
                WTotalUs = WTotalUs * -1
                WSaldo = WSaldo * -1
                WSaldoUs = WSaldoUs * -1
                
            Case 6
                WTipo = 7
                WImpre = "An"
                WNeto = WNeto * -1
                WIva1 = WIva1 * -1
                WIva2 = WIva2 * -1
                WTotal = WTotal * -1
                WTotalUs = WTotalUs * -1
                WSaldo = WSaldo * -1
                WSaldoUs = WSaldoUs * -1
                
            Case 7
                WTipo = 6
                WImpre = "Rc"
                WNeto = WNeto * -1
                WIva1 = WIva1 * -1
                WIva2 = WIva2 * -1
                WTotal = WTotal * -1
                WTotalUs = WTotalUs * -1
                WSaldo = WSaldo * -1
                WSaldoUs = WSaldoUs * -1
                
            Case 10
                WImpre = "Fr"
            Case 50
                WImpre = "CD"
            Case 51
                WImpre = "Rc"
                WNeto = WNeto * -1
                WIva1 = WIva1 * -1
                WIva2 = WIva2 * -1
                WTotal = WTotal * -1
                WTotalUs = WTotalUs * -1
                WSaldo = WSaldo * -1
                WSaldoUs = WSaldoUs * -1
                
            Case 52
                WImpre = "op"
                WNeto = WNeto * -1
                WIva1 = WIva1 * -1
                WIva2 = WIva2 * -1
                WTotal = WTotal * -1
                WTotalUs = WTotalUs * -1
                WSaldo = WSaldo * -1
                WSaldoUs = WSaldoUs * -1
                
            Case 60
                WImpre = "Fr"
            Case Else
            Stop
                WImpre = ""
        End Select
                
        WParidad = 1
        WProvincia = ""
        WComprobante = ""
        WAceptada = ""
        WCosto = 0
        WImporte1 = 0
        WImporte2 = 0
        WImporte3 = 0
        WImporte4 = 0
        WImporte5 = 0
        WImporte6 = 0
        WImporte7 = 0
        WClave = WTipo + WNUmero + WRenglon
        
        
        With rstCtaCte
        
            .Index = "clave"
            .Seek "=", WClave
            If .NoMatch Then
                .AddNew
                !Tipo = WTipo
                !Numero = WNUmero
                !Renglon = WRenglon
                !Cliente = WCliente
                !Fecha = WFecha
                !Estado = WEstado
                !Vencimiento = WVencimiento
                !Vencimiento1 = WVencimiento1
                !Total = WTotal
                !TotalUs = WTotalUs
                !Saldo = WSaldo
                !SaldoUs = WSaldoUs
                !OrdFecha = WOrdfecha
                !OrdVencimiento = WOrdvencimiento
                !OrdVencimiento1 = WOrdVencimiento1
                !Impre = WImpre
                !Neto = WNeto
                !Iva1 = WIva1
                !Iva2 = WIva2
                !Pedido = WPedido
                !Remito = WRemito
                !Orden = WOrden
                !Paridad = WParidad
                !provincia = WProvincia
                !Vendedor = WVendedor
                !Rubro = WRubro
                !Comprobante = ""
                !Aceptada = ""
                !Costo = 0
                !Importe1 = 0
                !Importe2 = 0
                !Importe3 = 0
                !Importe4 = 0
                !Importe5 = 0
                !Importe6 = 0
                !Importe7 = 0
                !Clave = WClave
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.



Stop



Rem condicion de pago


    Open "c:\prueba\ventas\" + WEmpresa + "pag.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WPago = Val(Mid$(Linea, 1, 4))
        WNombre = Mid$(Linea, 5, 30)
        WDias = Val(Mid$(Linea, 35, 3))
        WPlazo = Val(Mid$(Linea, 38, 9))
        WTasa = Val(Mid$(Linea, 47, 9))
        WDescuento = Val(Mid$(Linea, 56, 9))
        
        With rstPago
        
            .Index = "pago"
            .Seek "=", WPago
            If .NoMatch Then
                .AddNew
                !Pago = WPago
                !Nombre = WNombre
                !Dias = WDias
                !Plazo = WPlazo
                !Tasa = WTasa
                !Descuento = WDescuento
                
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.


Stop



Rem movimientos de envases dada


    Open "c:\prueba\ventas\" + WEmpresa + "pag.txt" For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, Linea
        
        WPago = Val(Mid$(Linea, 1, 4))
        WNombre = Mid$(Linea, 5, 30)
        WDias = Val(Mid$(Linea, 35, 3))
        WPlazo = Val(Mid$(Linea, 38, 9))
        WTasa = Val(Mid$(Linea, 47, 9))
        WDescuento = Val(Mid$(Linea, 56, 9))
        
        With rstPago
        
            .Index = "pago"
            .Seek "=", WPago
            If .NoMatch Then
                .AddNew
                !Pago = WPago
                !Nombre = WNombre
                !Dias = WDias
                !Plazo = WPlazo
                !Tasa = WTasa
                !Descuento = WDescuento
                
                .Update
            End If
        End With
        
    Loop
    Close #1    ' Cierra el archivo.






    Call Cancela_click


    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
    
Error:
Stop
     coderr = Err
     
     Resume Next
     
End Sub




