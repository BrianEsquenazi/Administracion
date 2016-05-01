VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAscii 
   AutoRedraw      =   -1  'True
   Caption         =   "Generacion de Archivo Ascii"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2535
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   4815
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "PrgAscii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Producto As String
Private Costo As Double
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim XParam As String
Dim Recibo(1000, 5) As String
Dim Facturas(1000, 5) As String
Dim Vector(100, 20) As String
Dim Cobro(100, 10) As String
Dim WNroFactura As Double
Dim WNroRecibo As Double
Dim WFactura As String
Dim WImporte As Double
Dim Porce As Double
Dim WImpo As Double
Dim WRenglon As String
Dim XRecibo As String

Private Sub Acepta_Click()

    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Rem Open "esta.txt" For Output As #1
    Rem procesa las ventas
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spEstadistica = "ListaEstadisticaFecha" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then

    With rstEstadistica
            .MoveFirst
            Do
                
                If WDesde <= rstEstadistica!OrdFecha And rstEstadistica!OrdFecha <= WHasta Then
                
                    WTipo = rstEstadistica!Tipo
                    WNumero = rstEstadistica!Numero
                    WRenglon = rstEstadistica!Renglon
                    WArticulo = rstEstadistica!Articulo
                    WCantidad = rstEstadistica!Cantidad
                    WPrecio = rstEstadistica!Precio
                    WPrecioUs = rstEstadistica!PrecioUs
                    WImporte = rstEstadistica!Importe
                    WimporteUs = rstEstadistica!ImporteUs
                    WCliente = rstEstadistica!Cliente
                    WParidad = rstEstadistica!Paridad
                    WVendedor = rstEstadistica!Vendedor
                    WRubro = rstEstadistica!Rubro
                    WLinea = rstEstadistica!Linea
                    WCosto1 = rstEstadistica!Costo1
                    WCosto2 = rstEstadistica!Costo2
                    WCoeficiente = rstEstadistica!Coeficiente
                    WPedido = rstEstadistica!Pedido
                    WFecha = rstEstadistica!Fecha
                    WImporte1 = rstEstadistica!Importe1
                    WImporte2 = rstEstadistica!Importe2
                    WImporte3 = rstEstadistica!Importe3
                    WImporte4 = rstEstadistica!Importe4
                    WOrdFecha = rstEstadistica!OrdFecha
                    WWArticulo = rstEstadistica!WArticulo
                    WRemito = rstEstadistica!Remito
                    WClave = rstEstadistica!Clave
                    
                    With rstEsta1
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !PrecioUs = WPrecioUs
                            !Importe = WImporte
                            !ImporteUs = WimporteUs
                            !Cliente = WCliente
                            !Paridad = WParidad
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Linea = WLinea
                            !Costo1 = WCosto1
                            !Costo2 = WCosto2
                            !Coeficiente = WCoeficiente
                            !Pedido = WPedido
                            !Fecha = WFecha
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !OrdFecha = WOrdFecha
                            !WArticulo = WWArticulo
                            !Remito = WRemito
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Tipo = WTipo
                            !Numero = WNumero
                            !Renglon = WRenglon
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Precio = WPrecio
                            !PrecioUs = WPrecioUs
                            !Importe = WImporte
                            !ImporteUs = WimporteUs
                            !Cliente = WCliente
                            !Paridad = WParidad
                            !Vendedor = WVendedor
                            !Rubro = WRubro
                            !Linea = WLinea
                            !Costo1 = WCosto1
                            !Costo2 = WCosto2
                            !Coeficiente = WCoeficiente
                            !Pedido = WPedido
                            !Fecha = WFecha
                            !Importe1 = WImporte1
                            !Importe2 = WImporte2
                            !Importe3 = WImporte3
                            !Importe4 = WImporte4
                            !OrdFecha = WOrdFecha
                            !WArticulo = WWArticulo
                            !Remito = WRemito
                            !Clave = WClave
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                
                    Rem Impre1 = Pusing("###,###,###,###.##", Str$(rstEstadistica!Precio * 100))
                    Rem Impre2 = Pusing("###,###,###,###.##", Str$(Abs(rstEstadistica!Cantidad) * 100))
                
                    Rem Print #1, Tab(1); rstEstadistica!Tipo;
                    Rem Print #1, Tab(4); rstEstadistica!Numero;
                    Rem Print #1, Tab(15); rstEstadistica!Renglon;
                    Rem Print #1, Tab(20); rstEstadistica!Articulo;
                    Rem Print #1, Tab(35); Impre2;
                    Rem Print #1, Tab(55); Impre1;
                    Rem Print #1, Tab(65); rstEstadistica!Paridad;
                    Rem Print #1, Tab(75); rstEstadistica!Vendedor;
                    Rem Print #1, Tab(85); rstEstadistica!Rubro;
                    Rem Print #1, Tab(95); rstEstadistica!Linea;
                    Rem Print #1, Tab(105); rstEstadistica!Fecha;
                    Rem Print #1, Tab(120); rstEstadistica!Cliente
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    
    rstEstadistica.Close
    
    End If
    
    Rem procesa las notas de debito varias
    
    SumaFac = 0
    SumaRec = 0
    
    Erase Facturas
    Erase Recibo
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spCtacte = "ListaCtacteFecha" + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then

    With rstCtacte
            .MoveFirst
            Do
                
                If WDesde <= rstCtacte!OrdFecha And rstCtacte!OrdFecha <= WHasta Then
                    WNroFactura = IIf(IsNull(rstCtacte!NroFactura), "0", rstCtacte!NroFactura)
                    WNroRecibo = IIf(IsNull(rstCtacte!NroRecibo), "0", rstCtacte!NroRecibo)
                    If WNroFactura <> 0 Or WNroRecibo <> 0 Then
                        If WNroFactura <> 0 Then
                            SumaFac = SumaFac + 1
                            Facturas(SumaFac, 1) = Str$(rstCtacte!NroFactura)
                            Facturas(SumaFac, 2) = Str$(rstCtacte!Neto)
                            Facturas(SumaFac, 3) = Str$(rstCtacte!Numero)
                            Facturas(SumaFac, 4) = rstCtacte!Fecha
                            Facturas(SumaFac, 5) = ""
                                Else
                            SumaRec = SumaRec + 1
                            Recibo(SumaRec, 1) = Str$(rstCtacte!NroRecibo)
                            Recibo(SumaRec, 2) = Str$(rstCtacte!Neto)
                            Recibo(SumaRec, 3) = Str$(rstCtacte!Numero)
                            Recibo(SumaRec, 4) = rstCtacte!Fecha
                            Recibo(SumaRec, 5) = ""
                        End If
                    End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    
    rstCtacte.Close
    
    End If
    
    
    For Ciclo = 1 To SumaRec
    
        XRecibo = Recibo(Ciclo, 1)
        XImporte = Val(Recibo(Ciclo, 2))
        XNumero = Val(Recibo(Ciclo, 3))
        XFecha = Recibo(Ciclo, 4)
        
        Call Ceros(XRecibo, 6)
        
        Erase Cobro
        SumaCob = 0
        TotalRec = 0
        
        XParam = "'" + XRecibo + "'"
    
        spRecibo = "ConsultaRecibos " + XParam
        Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibo.RecordCount > 0 Then
    
            With rstRecibo
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        If Val(rstRecibo!tipo1) = 1 Then
    
                            SumaCob = SumaCob + 1
            
                            Cobro(SumaCob, 1) = rstRecibo!Numero1
                            Cobro(SumaCob, 2) = Str$(rstRecibo!Importe1)
                            Cobro(SumaCob, 3) = Str$(SumaCob)
                            TotalRec = TotalRec + rstRecibo!Importe1
                        
                        End If
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstRecibo.Close
        End If
        
        For dada = 1 To SumaCob
            
            If TotalRec <> 0 Then
                Porce = Val(Cobro(dada, 2)) / TotalRec * 100
                Call Redondeo(Porce)
                WImpo = XImporte * ((Porce / 100))
                Call Redondeo(WImpo)
            End If
            
            SumaFac = SumaFac + 1
            
            Facturas(SumaFac, 1) = Str$(Cobro(dada, 1))
            Facturas(SumaFac, 2) = Str$(WImpo)
            Facturas(SumaFac, 3) = XNumero
            Facturas(SumaFac, 4) = XFecha
            Facturas(SumaFac, 5) = Str$(dada)
            
        Next dada
        
    Next Ciclo
            
    Rem Procesa las notas de debitos
    
    For Ciclo = 1 To SumaFac
    
        XFactura = Facturas(Ciclo, 1)
        XImporte = Val(Facturas(Ciclo, 2))
        XNumero = Val(Facturas(Ciclo, 3))
        XFecha = Facturas(Ciclo, 4)
        Xlugar = Facturas(Ciclo, 5)
        
        Suma = 0
        Total = 0
        Erase Vector
        
        XParam = "'" + "01" + "','" _
                + XFactura + "'"
    
        spEstadistica = "ConsultaEstadistica1 " + XParam
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
    
            With rstEstadistica
                .MoveFirst
                Do
                    If .EOF = False Then
    
                        Suma = Suma + 1
            
                        Vector(Suma, 1) = rstEstadistica!Articulo
                        Vector(Suma, 2) = Str$(rstEstadistica!Cantidad)
                        Vector(Suma, 3) = Str$(rstEstadistica!Precio)
                        Vector(Suma, 4) = Str$(rstEstadistica!Cantidad * rstEstadistica!Precio)
                        Total = Total + rstEstadistica!Cantidad * rstEstadistica!Precio
                        Vector(Suma, 6) = rstEstadistica!Cliente
                        Vector(Suma, 7) = rstEstadistica!Paridad
                        Vector(Suma, 8) = rstEstadistica!Vendedor
                        Vector(Suma, 9) = rstEstadistica!Rubro
                        Vector(Suma, 10) = rstEstadistica!Linea
                        Vector(Suma, 11) = rstEstadistica!Pedido
                        Vector(Suma, 12) = XFecha
                        Vector(Suma, 13) = rstEstadistica!WArticulo
                        Vector(Suma, 14) = rstEstadistica!Remito
                        Vector(Suma, 15) = XNumero
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEstadistica.Close
        End If
        
        For dada = 1 To Suma
            
            If Total <> 0 Then
                Porce = Val(Vector(dada, 4)) / Total * 100
                Call Redondeo(Porce)
                WImpo = XImporte * ((Porce / 100))
                Call Redondeo(WImpo)
            End If
            
            If WImpo > 0 Then
                WTipo = "01"
                    Else
                WTipo = "02"
            End If
            WNumero = Vector(dada, 15)
            WRenglon = Str$(dada)
            If Val(Xlugar) <> 0 Then
                WRenglon = Str$(dada + ((Val(Xlugar) - 1) * 10))
            End If
            WArticulo = Vector(dada, 1)
            WCantidad = "1"
            WPrecio = Abs(WImpo)
            WPrecioUs = Abs(WImpo)
            WImporte = WImpo
            WimporteUs = WImpo
            WCliente = Vector(dada, 6)
            WParidad = "1"
            WVendedor = Vector(dada, 8)
            WRubro = Vector(dada, 9)
            WLinea = Vector(dada, 10)
            WCosto1 = "0"
            WCosto2 = "0"
            WCoeficiente = "0"
            WPedido = Vector(dada, 11)
            WFecha = Vector(dada, 12)
            WImporte1 = "0"
            WImporte2 = "0"
            WImporte3 = "0"
            WImporte4 = "0"
            WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            WWArticulo = Vector(dada, 13)
            WRemito = Vector(dada, 14)
            Call Ceros(WNumero, 8)
            Call Ceros(WRenglon, 2)
            WClave = WTipo + WNumero + WRenglon
                    
            With rstEsta2
                .Index = "Clave"
                .Seek "=", WClave
                If .NoMatch Then
                    .AddNew
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Renglon = WRenglon
                    !Articulo = WArticulo
                    !Cantidad = WCantidad
                    !Precio = WPrecio
                    !PrecioUs = WPrecioUs
                    !Importe = WImporte
                    !ImporteUs = WimporteUs
                    !Cliente = WCliente
                    !Paridad = WParidad
                    !Vendedor = WVendedor
                    !Rubro = WRubro
                    !Linea = WLinea
                    !Costo1 = WCosto1
                    !Costo2 = WCosto2
                    !Coeficiente = WCoeficiente
                    !Pedido = WPedido
                    !Fecha = WFecha
                    !Importe1 = WImporte1
                    !Importe2 = WImporte2
                    !Importe3 = WImporte3
                    !Importe4 = WImporte4
                    !OrdFecha = WOrdFecha
                    !WArticulo = WWArticulo
                    !Remito = WRemito
                    !Clave = WClave
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Renglon = WRenglon
                    !Articulo = WArticulo
                    !Cantidad = WCantidad
                    !Precio = WPrecio
                    !PrecioUs = WPrecioUs
                    !Importe = WImporte
                    !ImporteUs = WimporteUs
                    !Cliente = WCliente
                    !Paridad = WParidad
                    !Vendedor = WVendedor
                    !Rubro = WRubro
                    !Linea = WLinea
                    !Costo1 = WCosto1
                    !Costo2 = WCosto2
                    !Coeficiente = WCoeficiente
                    !Pedido = WPedido
                    !Fecha = WFecha
                    !Importe1 = WImporte1
                    !Importe2 = WImporte2
                    !Importe3 = WImporte3
                    !Importe4 = WImporte4
                    !OrdFecha = WOrdFecha
                    !WArticulo = WWArticulo
                    !Remito = WRemito
                    !Clave = WClave
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
            
        Next dada
        
    Next Ciclo
    
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    PrgAscii.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
End Sub

Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            DesdeFec.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub



