VERSION 5.00
Begin VB.Form PrgPaso2 
   Caption         =   "Recepcion de datos de ventas"
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
Attribute VB_Name = "PrgPaso2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgRecep.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
   Rem cuenta corriente
        
    If Val(WEmpresa) = 2 Then
        
    coderr = 0
    With rstWCtaCte4
            .Index = "Clave"
            .MoveFirst
            Do
                
                WTipo = !Tipo
                WNumero = !Numero
                Wrenglon = !Renglon
                WCliente = !Cliente
                WFecha = !Fecha
                WEstado = !Estado
                Wvencimiento = !Vencimiento
                WVencimiento1 = !Vencimiento1
                WTotal = !Total
                WTotalUs = !TotalUs
                WSaldo = !Saldo
                WSaldoUs = !SaldoUS
                WOrdFecha = !OrdFecha
                WOrdVencimiento = !OrdVencimiento
                WOrdVencimiento1 = !OrdVencimiento1
                WImpre = !Impre
                WNeto = !Neto
                WIva1 = !Iva1
                WIva2 = !Iva2
                WPedido = !Pedido
                WRemito = !Remito
                WOrden = !Orden
                WParidad = !Paridad
                WProvincia = !Provincia
                WVendedor = !Vendedor
                WRubro = !Rubro
                WComprobante = !Comprobante
                WAceptada = !Aceptada
                WCosto = !Costo
                WImporte1 = !Importe1
                WImporte2 = !Importe2
                WImporte3 = !Importe3
                WImporte4 = !Importe4
                WImporte5 = !Importe5
                WImporte6 = !Importe6
                WImporte7 = !Importe7
                WClave = !Clave
                    
                With rstCtaCte
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !SaldoUS = WSaldoUs
                            !OrdFecha = WOrdFecha
                            !OrdVencimiento = WOrdVencimiento
                            !OrdVencimiento1 = WOrdVencimiento1
                            !Impre = WImpre
                            !Neto = WNeto
                            !Iva1 = WIva1
                            !Iva2 = WIva2
                            !Pedido = WPedido
                            !Remito = WRemito
                            !Orden = WOrden
                            !Paridad = WParidad
                            !Provincia = WProvincia
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Comprobante = WComprobante
                            !Aceptada = WAceptada
                            !Costo = WCosto
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !Importe5 = WImporte5
                            !Importe6 = WImporte6
                            !Importe7 = WImporte7
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !SaldoUS = WSaldoUs
                            !OrdFecha = WOrdFecha
                            !OrdVencimiento = WOrdVencimiento
                            !OrdVencimiento1 = WOrdVencimiento1
                            !Impre = WImpre
                            !Neto = WNeto
                            !Iva1 = WIva1
                            !Iva2 = WIva2
                            !Pedido = WPedido
                            !Remito = WRemito
                            !Orden = WOrden
                            !Paridad = WParidad
                            !Provincia = WProvincia
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Comprobante = WComprobante
                            !Aceptada = WAceptada
                            !Costo = WCosto
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !Importe5 = WImporte5
                            !Importe6 = WImporte6
                            !Importe7 = WImporte7
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
        Else
        
    coderr = 0
    With rstWCtaCte2
            .Index = "Clave"
            .MoveFirst
            Do
                
                WTipo = !Tipo
                WNumero = !Numero
                Wrenglon = !Renglon
                WCliente = !Cliente
                WFecha = !Fecha
                WEstado = !Estado
                Wvencimiento = !Vencimiento
                WVencimiento1 = !Vencimiento1
                WTotal = !Total
                WTotalUs = !TotalUs
                WSaldo = !Saldo
                WSaldoUs = !SaldoUS
                WOrdFecha = !OrdFecha
                WOrdVencimiento = !OrdVencimiento
                WOrdVencimiento1 = !OrdVencimiento1
                WImpre = !Impre
                WNeto = !Neto
                WIva1 = !Iva1
                WIva2 = !Iva2
                WPedido = !Pedido
                WRemito = !Remito
                WOrden = !Orden
                WParidad = !Paridad
                WProvincia = !Provincia
                WVendedor = !Vendedor
                WRubro = !Rubro
                WComprobante = !Comprobante
                WAceptada = !Aceptada
                WCosto = !Costo
                WImporte1 = !Importe1
                WImporte2 = !Importe2
                WImporte3 = !Importe3
                WImporte4 = !Importe4
                WImporte5 = !Importe5
                WImporte6 = !Importe6
                WImporte7 = !Importe7
                WClave = !Clave
                    
                With rstCtaCte
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !SaldoUS = WSaldoUs
                            !OrdFecha = WOrdFecha
                            !OrdVencimiento = WOrdVencimiento
                            !OrdVencimiento1 = WOrdVencimiento1
                            !Impre = WImpre
                            !Neto = WNeto
                            !Iva1 = WIva1
                            !Iva2 = WIva2
                            !Pedido = WPedido
                            !Remito = WRemito
                            !Orden = WOrden
                            !Paridad = WParidad
                            !Provincia = WProvincia
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Comprobante = WComprobante
                            !Aceptada = WAceptada
                            !Costo = WCosto
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !Importe5 = WImporte5
                            !Importe6 = WImporte6
                            !Importe7 = WImporte7
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = Wrenglon
                            !Cliente = WCliente
                            !Fecha = WFecha
                            !Estado = WEstado
                            !Vencimiento = Wvencimiento
                            !Vencimiento1 = WVencimiento1
                            !Total = WTotal
                            !TotalUs = WTotalUs
                            !Saldo = WSaldo
                            !SaldoUS = WSaldoUs
                            !OrdFecha = WOrdFecha
                            !OrdVencimiento = WOrdVencimiento
                            !OrdVencimiento1 = WOrdVencimiento1
                            !Impre = WImpre
                            !Neto = WNeto
                            !Iva1 = WIva1
                            !Iva2 = WIva2
                            !Pedido = WPedido
                            !Remito = WRemito
                            !Orden = WOrden
                            !Paridad = WParidad
                            !Provincia = WProvincia
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Comprobante = WComprobante
                            !Aceptada = WAceptada
                            !Costo = WCosto
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !Importe5 = WImporte5
                            !Importe6 = WImporte6
                            !Importe7 = WImporte7
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
            
    End If
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub




