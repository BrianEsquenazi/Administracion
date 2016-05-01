VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgActuaFactuexpo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturacion de Pedidos"
   ClientHeight    =   8310
   ClientLeft      =   1125
   ClientTop       =   420
   ClientWidth     =   10020
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   10020
   Visible         =   0   'False
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   8280
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Cliente 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   7
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Numero 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   570
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreRemito.rpt"
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6855
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12091
      _Version        =   327680
      Rows            =   1000
      Cols            =   7
   End
   Begin VB.Label Label11 
      Caption         =   "Pedido"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Factura"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgActuaFactuexpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WPlazo1 As Integer
Private WPlazo2 As Integer
Private WDias1 As Integer
Private WDias2 As Integer
Private WFecha As String
Private Wvencimiento As String
Private WVencimiento1 As String
Private WPago1 As Integer
Private WPago2 As Integer
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WImporte As Double
Private WCodIva As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private Precio As String
Private dada As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WImpoIb As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private parcial As String
Private WSeguro As Double
Private WFlete As Double
Private WGastos As Double
Private WTexto1 As String
Private WTexto2 As String
Private Auxiliar(100, 50) As String

Dim ZZControlLote(100, 60) As String
Dim ControlLote(12, 2) As String
Dim ControlEnvase(12, 2) As String

Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstImpreRemito As Recordset
Dim spImpreRemio As String

Dim XParam As String
Dim WLote(12, 2) As String
Dim WImpresion(100, 10) As String
Dim XEnvase(100, 6) As String
Dim XCanti As String
Private WTipoPedido As String

Dim VectorCosto(100, 3) As String
Dim ZZZProducto As String
Dim ZZZCosto As Double

Dim ZZClave As String
Dim ZZNumero As String
Dim ZZRenglon As String
Dim ZZFecha As String
Dim ZZNombre As String
Dim ZZDireccion As String
Dim ZZLocalidad As String
Dim ZZPedido As String
Dim ZZCliente As String
Dim ZZOrden As String
Dim ZZDescripcion As String
Dim ZZCantidad As String
Dim ZZRemito As String

Dim ZZVector(100, 10) As String
Dim ZZImpre(100, 10) As String
Dim ZZCampo1 As String
Dim ZZCampo2 As String
Dim ZLote6 As Double
Dim ZLote7 As Double
Dim ZLote8 As Double
Dim ZLote9 As Double
Dim ZLote10 As Double
Dim ZLote11 As Double
Dim ZLote12 As Double

Dim ZZComprobante As Integer
Dim ZZCuit As String
Dim ZZPais As String
Dim ZZCuitII As String
Dim ZZRazon As String
Dim ZZDomicilio As String

Dim ZZGrabaFactura As String


Private Sub cmdClose_Click()
    PrgActuaFactuexpo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub wvector1_DBLCLICK()
        
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        Select Case rstPedido!TipoPedido
            Case 1
                ZZPasaTipoPedido = "CO"
            Case 3
                ZZPasaTipoPedido = "BI"
            Case 4
                ZZPasaTipoPedido = "FA"
            Case 5
                ZZPasaTipoPedido = "PG"
            Case Else
                ZZPasaTipoPedido = "PT"
        End Select
        rstPedido.Close
    End If
    
    WVector1.Col = 6
    ZZPasaClave = WVector1.Text
    WVector1.Col = 1
    ZZPasaTerminado = WVector1.Text
    WVector1.Col = 5
    ZZPasaCantidad = Val(WVector1.Text)
    ZSuma = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.Clave = " + "'" + ZZPasaClave + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        XLote = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
        ZSuma = ZSuma + XLote
        XLote = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
        ZSuma = ZSuma + XLote
        XLote = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
        ZSuma = ZSuma + XLote
        XLote = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
        ZSuma = ZSuma + XLote
        XLote = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
        ZSuma = ZSuma + XLote
                
        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
        
        If Len(Trim(WLoteAdicional)) = 98 Then
            XLote = Val(Mid$(WLoteAdicional, 9, 6))
            ZSuma = ZSuma + XLote
            XLote = Val(Mid$(WLoteAdicional, 23, 6))
            ZSuma = ZSuma + XLote
            XLote = Val(Mid$(WLoteAdicional, 37, 6))
            ZSuma = ZSuma + XLote
            XLote = Val(Mid$(WLoteAdicional, 51, 6))
            ZSuma = ZSuma + XLote
            XLote = Val(Mid$(WLoteAdicional, 65, 6))
            ZSuma = ZSuma + XLote
            XLote = Val(Mid$(WLoteAdicional, 79, 6))
            ZSuma = ZSuma + XLote
            XLote = Val(Mid$(WLoteAdicional, 93, 6))
            ZSuma = ZSuma + XLote
        End If
        
        rstEstadistica.Close
        
    End If
    
    If ZSuma <> 0 Then
        Exit Sub
    End If
    
    PrgModactuaFactuExpo.Show
    
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Call Limpia_Vector
    
    Renglon = 0
    Numero.SetFocus

End Sub


Private Sub Form_Load()

    Call Limpia_Vector

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Rem Numero.SetFocus
     
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Proceso1_Click()

    WNeto = 0
    
    Call Limpia_Vector
    
    Renglon = 0
    Erase Auxiliar
    
    XParam = "'" + "01" + "','" _
                + Numero.Text + "'"
    
    spEstadistica = "ConsultaEstadistica1 " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Renglon = Renglon + 1
                    
                    WVector1.Row = Renglon
            
                    WVector1.Col = 1
                    WVector1.Text = rstEstadistica!Articulo
                    Auxi1 = rstEstadistica!Articulo
                
                    dada = Str$(rstEstadistica!Cantidad)
                    WVector1.Col = 3
                    WVector1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!PrecioUs)
                    WVector1.Col = 4
                    WVector1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Cantidad)
                    WVector1.Col = 5
                    WVector1.Text = Pusing("###,###.##", dada)
                    
                    WVector1.Col = 6
                    WVector1.Text = rstEstadistica!Clave
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!PrecioUs)
                    End If
                    
                    Auxiliar(Renglon, 1) = Auxi1
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    XRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To XRenglon
    
        Auxi1 = Auxiliar(Da, 1)
        
        If Left$(Auxi1, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Renglon = Renglon + 1
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
            Case Else
                ClavePrecios = Cliente.Text + Auxi1
        
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Renglon = Renglon + 1
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
        End Select
    Next Da
    
    WVector1.Col = 1
    WVector1.Row = 1
    WVector1.TopRow = 1

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "01" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
        
            Pedido.Text = rstCtacte!Pedido
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!pago2
                WVendedor = rstCliente!Vendedor
                WProv = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
            End If
            
            Call Proceso1_Click
            
                    Else
                    
            Rem .Index = "Numero"
            Rem .Seek "=", Val(Numero.Text)
            Rem If .NoMatch = False Then
            Rem     m$ = "Comprobante ya existente"
            Rem     A% = MsgBox(m$, 0, "Ingreso de Facturas")
            Rem     Numero.SetFocus
            Rem        Else
            Rem     WNumero = Numero.Text
            Rem    Rem Call Limpia_Click
            Rem    Numero.Text = WNumero
            Rem    Pedido.SetFocus
            Rem End If
            WNumero = Numero.Text
            Rem Call Limpia_Click
            Numero.Text = WNumero
            Pedido.SetFocus
                
        End If
    End If
End Sub



Private Sub Limpia_Vector()

    WVector1.Clear
    
    WVector1.ColWidth(0) = 150
    WVector1.ColWidth(1) = 1200
    WVector1.ColWidth(2) = 2000
    WVector1.ColWidth(3) = 1000
    WVector1.ColWidth(4) = 100
    WVector1.ColWidth(5) = 1000
    WVector1.ColWidth(6) = 10
    
    WVector1.Row = 0
    
    WVector1.Col = 1
    WVector1.Text = "Producto"
    
    WVector1.Col = 2
    WVector1.Text = "Descripcion"
    
    WVector1.Col = 3
    WVector1.Text = "Cantidad S/Pedido"
    
    WVector1.Col = 4
    WVector1.Text = ""
    
    WVector1.Col = 5
    WVector1.Text = "Parcial"
    
    WVector1.Col = 6
    WVector1.Text = ""
    
End Sub
