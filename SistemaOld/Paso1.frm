VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPaso1 
   Caption         =   "Traspaso de informacion de ventas"
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "PrgPaso1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    WFectraspaso = Mid$(Fecha.Text, 4, 2) + "-" + Left$(Fecha.Text, 2) + "-" + Right$(Fecha.Text, 4)
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prgtraspa.Hide
    Unload Me
    Menu.SetFocus
End Sub


Sub Form_Load()
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    'cuenta corriente de clientes
    
    coderr = 0
    With rstWCtaCte
        .Index = "Clave"
        .MoveFirst
        If coderr = 0 Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Rem cuenta corriente
        
    coderr = 0
    With rstCtaCte
            .Index = "Clave"
            .MoveFirst
            Do
                If !Wdate = WFectraspaso Then
                
                    WTipo = !Tipo
                    WNumero = !Numero
                    WRenglon = !Renglon
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
                    
                    With rstWCtaCte
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
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
                            !Renglon = WRenglon
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
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub




