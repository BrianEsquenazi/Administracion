VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgSalvaPrecios 
   AutoRedraw      =   -1  'True
   Caption         =   "Salva Precios x Cliente"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2775
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEstaVen.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "salvaprecios.frx":0000
      Left            =   840
      List            =   "salvaprecios.frx":0007
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgSalvaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Costo As Double
Private Producto As String
Private Auxiliar(100, 7) As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim XParam As String
Private Vector(10000, 2) As String
Dim Posi As Integer

Private Sub Acepta_Click()

    On Error GoTo WError
    
    With rstEsta
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    XParam = "'" + "00000000" + "','" _
                 + "99999999" + "'"
    spEstadistica = "ListaEstadisticaFecha" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
            
                If Val(rstEstadistica!Tipo) = 1 Then
                
                    WTipo = rstEstadistica!Tipo
                    WNumero = rstEstadistica!numero
                    WRenglon = rstEstadistica!Renglon
                    WArticulo = rstEstadistica!Articulo
                    WCantidad = rstEstadistica!Cantidad
                    WPrecio = rstEstadistica!Precio
                    WPrecioUs = rstEstadistica!PrecioUs
                    WImporte = rstEstadistica!Importe
                    WimporteUs = rstEstadistica!ImporteUs
                    WCliente = rstEstadistica!Cliente
                    WParidad = rstEstadistica!Paridad
                    wvendedor = rstEstadistica!Vendedor
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
                
                    With rstEsta
        
                        .Index = "Clave"
                        .AddNew
                        !Tipo = WTipo
                        !numero = WNumero
                        !Renglon = WRenglon
                        !Articulo = WArticulo
                        !Cantidad = WCantidad
                        !Precio = WPrecio
                        !PrecioUs = WPrecioUs
                        !Importe = WImporte
                        !ImporteUs = WimporteUs
                        !Cliente = WCliente
                        !Paridad = WParidad
                        !Vendedor = wvendedor
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
                        
                    End With
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        rstEstadistica.Close
    End If

    With rstEsta
        .Index = "ordfecha"
        .MoveLast
        If .NoMatch = False Then
            Do
                .Edit
                
                WCliente = rstEsta!Cliente
                WArticulo = rstEsta!Articulo
                WCantidad = rstEsta!Cantidad
                WPrecio = rstEsta!Precio
                WFecha = rstEsta!Fecha
                WFactura = rstEsta!numero
                
                ClavePrecio = WCliente + WArticulo
            
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecio + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
            
                    WFecha1 = IIf(IsNull(rstPrecios!Fecha1), "", rstPrecios!Fecha1)
                    WFactura1 = IIf(IsNull(rstPrecios!Factura1), "", rstPrecios!Factura1)
                    WPrecio1 = Str$(rstPrecios!Precio1)
                    WCantidad1 = Str$(rstPrecios!Cantidad1)
            
                    WFecha2 = IIf(IsNull(rstPrecios!Fecha2), "", rstPrecios!Fecha2)
                    WFactura2 = IIf(IsNull(rstPrecios!Factura2), "", rstPrecios!Factura2)
                    WPrecio2 = Str$(rstPrecios!Precio2)
                    WCantidad2 = Str$(rstPrecios!Cantidad2)
            
                    WFecha3 = IIf(IsNull(rstPrecios!Fecha3), "", rstPrecios!Fecha3)
                    WFactura3 = IIf(IsNull(rstPrecios!Factura3), "", rstPrecios!Factura3)
                    WPrecio3 = Str$(rstPrecios!Precio3)
                    WCantidad3 = Str$(rstPrecios!Cantidad3)
            
                    WFecha4 = IIf(IsNull(rstPrecios!Fecha4), "", rstPrecios!Fecha4)
                    WFactura4 = IIf(IsNull(rstPrecios!Factura4), "", rstPrecios!Factura4)
                    WPrecio4 = Str$(rstPrecios!Precio4)
                    WCantidad4 = Str$(rstPrecios!Cantidad4)
            
                    WFecha5 = IIf(IsNull(rstPrecios!Fecha5), "", rstPrecios!Fecha5)
                    WFactura5 = IIf(IsNull(rstPrecios!Factura5), "", rstPrecios!Factura5)
                    WPrecio5 = Str$(rstPrecios!Precio5)
                    WCantidad5 = Str$(rstPrecios!Cantidad5)

                    If WFactura <> Val(WFactura1) And WFactura <> Val(WFactura2) And WFactura <> Val(WFactura3) And WFactura <> Val(WFactura4) And WFactura <> Val(WFactura5) Then
                    Rem Stop
                        Graba = "N"
                    
                        If Val(WCantidad5) = O And Graba = "N" Then
                            WFecha5 = WFecha
                            WFactura5 = Str$(WFactura)
                            WPrecio5 = Str$(WPrecio)
                            WCantidad5 = Str$(WCantidad)
                            Graba = "S"
                        End If
                    
                        If Val(WCantidad4) = O And Graba = "N" Then
                            WFecha4 = WFecha
                            WFactura4 = Str$(WFactura)
                            WPrecio4 = Str$(WPrecio)
                            WCantidad4 = Str$(WCantidad)
                            Graba = "S"
                        End If
                    
                        If Val(WCantidad3) = O And Graba = "N" Then
                            WFecha3 = WFecha
                            WFactura3 = Str$(WFactura)
                            WPrecio3 = Str$(WPrecio)
                            WCantidad3 = Str$(WCantidad)
                            Graba = "S"
                        End If
                    
                        If Val(WCantidad2) = O And Graba = "N" Then
                            WFecha2 = WFecha
                            WFactura2 = Str$(WFactura)
                            WPrecio2 = Str$(WPrecio)
                            WCantidad2 = Str$(WCantidad)
                            Graba = "S"
                        End If
                    
                        If Val(WCantidad1) = O And Graba = "N" Then
                            WFecha1 = WFecha
                            WFactura1 = Str$(WFactura)
                            WPrecio1 = Str$(WPrecio)
                            WCantidad1 = Str$(WCantidad)
                            Graba = "S"
                        End If
                    
                        If Graba = "S" Then
                    
                            WDate = Date$
                            rstPrecios.Close
            
                            XParam = "'" + ClavePrecio + "','" _
                                + WFecha1 + "','" _
                                + WFactura1 + "','" _
                                + WPrecio1 + "','" _
                                + WCantidad1 + "','" _
                                + WFecha2 + "','" _
                                + WFactura2 + "','" _
                                + WPrecio2 + "','" _
                                + WCantidad2 + "','" _
                                + WFecha3 + "','" _
                                + WFactura3 + "','" _
                                + WPrecio3 + "','" _
                                + WCantidad3 + "','" _
                                + WFecha4 + "','" _
                                + WFactura4 + "','" _
                                + WPrecio4 + "','" _
                                + WCantidad4 + "','" _
                                + WFecha5 + "','" _
                                + WFactura5 + "','" _
                                + WPrecio5 + "','" _
                                + WCantidad5 + "','" _
                                + WDate + "'"
                                           
                            spPrecios = "ModificaPreciosFactura " + XParam
                            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    
                    End If
                    
                End If
                
                .MovePrevious
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEsta
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    PrgSalvaPrecios.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Esta
End Sub
