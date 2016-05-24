VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrgActualizaPedidoFarma 
   AutoRedraw      =   -1  'True
   Caption         =   "Seleccion de Pedidos en Condiciones de Emitir factura"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   Begin VB.CommandButton cmdclose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   6600
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10821
      _Version        =   327680
      Rows            =   4000
      Cols            =   8
   End
End
Attribute VB_Name = "PrgActualizaPedidoFarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstMuestra As Recordset
Dim spMuestra As String
Dim XParam As String
Dim TotalPedidos As Integer
Dim WGraba As String
Dim ZVector(100, 4) As String
Dim CargaPedido(100, 10) As String
Dim ZDirEntrega(10) As String
Private CargaEmpresa(10, 2) As String

Dim ZCodigo As String
Dim ZTerminado As String
Dim ZArticulo As String
Dim ZEnsayo As String
Dim ZNombre As String
Dim ZFecha As String
Dim ZFechaOrd As String
Dim ZCantidad As String
Dim ZCliente As String
Dim ZRazon As String
Dim ZDescriCliente As String
Dim ZVendedor As String
Dim ZDesVendedor As String
Dim ZObservaciones As String
Dim ZAutoriza As String
Dim ZImpresion As String
Dim ZPedido As String
Dim ZLugarDirEntrega As String
Dim ZDescriDirEntrega As String

Private Sub cmdClose_Click()
    PrgActualizaPedidoFarma.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 1000
    Muestra.ColWidth(2) = 1400
    Muestra.ColWidth(3) = 1000
    Muestra.ColWidth(4) = 3000
    Muestra.ColWidth(5) = 1400
    Muestra.ColWidth(6) = 1700
    Muestra.ColWidth(7) = 1400
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Importe"
    
    Call Proceso_Click
    
End Sub

Private Sub Proceso_Click()

    WSalida = "N"
        
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Importe"
    
    Renglon = 0
    WSaldo = 0
    
    Pasa = 0
    Pedido = ""
    Fecha = "  /  /    "
    Cliente = ""
    Razon = ""
    FEntrega = "  /  /    "
    Tipo = 0
    Importe = 0
    Estado = ""
    
    Rem
    Rem Pedidos de Venta
    Rem
    
    XEmpresa = Wempresa
    
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where MarcaFactura = " + "'" + "9" + "'"
    ZSql = ZSql + " Order by Clave"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    If Pasa = 0 Then
                        corte = rstPedido!Pedido
                        Fecha = rstPedido!Fecha
                        Cliente = rstPedido!Cliente
                        FEntrega = rstPedido!FecEntrega
                        Tipo = rstPedido!Tipoped
                        Importe = 0
                        Estado = rstPedido!Autorizo
                        Impresa = rstPedido!Impresion
                        Pasa = 1
                    End If
                
                    If corte <> rstPedido!Pedido Then
                
                        Renglon = Renglon + 1
        
                        Muestra.Row = Renglon
                    
                        Muestra.Col = 1
                        Muestra.Text = Pusing("######", Str$(corte))
                    
                        Muestra.Col = 2
                        Muestra.Text = Fecha
            
                        Muestra.Col = 3
                        Muestra.Text = Cliente
                    
                        Muestra.Col = 4
                        Muestra.Text = ""
                    
                        Muestra.Col = 5
                        Muestra.Text = FEntrega
                    
                        Select Case Tipo
                            Case 0
                                Muestra.Col = 6
                                Muestra.Text = "Normal"
                            Case 1
                                Muestra.Col = 6
                                Muestra.Text = "A Fecha"
                            Case 2
                                Muestra.Col = 6
                                Muestra.Text = "Fecha LImite"
                            Case 3
                                Muestra.Col = 6
                                Muestra.Text = "Urgente"
                            Case 4
                                Muestra.Col = 6
                                Muestra.Text = "Retira Cliente"
                            Case 5
                                Muestra.Col = 6
                                Muestra.Text = "MUESTRA"
                            Case Else
                                Muestra.Col = 6
                                Muestra.Text = ""
                        End Select
                    
                        Muestra.Col = 7
                        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                    
                        corte = rstPedido!Pedido
                        Fecha = rstPedido!Fecha
                        Cliente = rstPedido!Cliente
                        FEntrega = rstPedido!FecEntrega
                        Tipo = rstPedido!Tipoped
                        Importe = 0
                        Estado = rstPedido!Autorizo
                        Impresa = rstPedido!Impresion
                        Pasa = 1
                
                    End If
                
                    Importe = Importe + ((rstPedido!Cantidad - rstPedido!Facturado) * rstPedido!Precio)
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
    
        rstPedido.Close
    
    End If
    
    If Pasa <> 0 Then
                    
        Renglon = Renglon + 1
            
        Muestra.Row = Renglon
                        
        Muestra.Col = 1
        Muestra.Text = Pusing("######", Str$(corte))
                        
        Muestra.Col = 2
        Muestra.Text = Fecha
                
        Muestra.Col = 3
        Muestra.Text = Cliente
                        
        Muestra.Col = 4
        Muestra.Text = ""
                        
        Muestra.Col = 5
        Muestra.Text = FEntrega
                        
        Muestra.Col = 6
        Muestra.Text = Str$(Tipo)
        
        Select Case Tipo
            Case 0
                Muestra.Col = 6
                Muestra.Text = "Normal"
            Case 1
                Muestra.Col = 6
                Muestra.Text = "A Fecha"
            Case 2
                Muestra.Col = 6
                Muestra.Text = "Fecha LImite"
            Case 3
                Muestra.Col = 6
                Muestra.Text = "Urgente"
            Case 4
                Muestra.Col = 6
                Muestra.Text = "Retira Cliente"
            Case 5
                Muestra.Col = 6
                Muestra.Text = "MUESTRA"
            Case Else
                Muestra.Col = 6
                Muestra.Text = ""
        End Select
                                
        Muestra.Col = 7
        Muestra.Text = Pusing("###,###,###.##", Str$(Importe))
                        
    End If
    
    For dada = 1 To Renglon
    
        Muestra.Row = dada
                        
        Muestra.Col = 3
        WCliente = Muestra.Text
    
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra.Col = 4
            Muestra.Text = rstCliente!Razon
            rstCliente.Close
        End If
        
    Next dada
    
    Call Conecta_Empresa
    
    TotalPedidos = Renglon
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.TopRow = 1
    
    Rem Muestra.SetFocus

End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Cliente"
    
    Muestra.Col = 4
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 5
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 6
    Muestra.Text = "Tipo"
    
    Muestra.Col = 7
    Muestra.Text = "Importe"
    
End Sub

Private Sub Muestra_DblClick()

    ZZProcesoFactura = 99
    Muestra.Col = 1
    PrgMuestraPedido.Pedido.Text = Muestra.Text
    PrgActualizaPedidoFarma.Hide
    Unload Me
    PrgMuestraPedido.Show

    Call cmdClose_Click
    
End Sub
