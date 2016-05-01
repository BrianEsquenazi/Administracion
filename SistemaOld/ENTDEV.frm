VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEntdev 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Devoluvion de Mercaderia"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   330
   ClientWidth     =   11805
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11805
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Reimpresion"
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8880
      MaxLength       =   6
      TabIndex        =   31
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Cliente 
      Height          =   285
      Left            =   7080
      MaxLength       =   6
      TabIndex        =   27
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   480
      Width           =   5655
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11040
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   2520
      TabIndex        =   16
      Top             =   6480
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   4680
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Codigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   12
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   1320
      TabIndex        =   9
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   2520
      TabIndex        =   7
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   11415
      Begin VB.TextBox WLaboratorio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WLote 
         Height          =   285
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   18
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Verificado"
         Height          =   255
         Left            =   10080
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         Height          =   255
         Left            =   8880
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   7440
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto Terminado"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3825
      Left            =   120
      OleObjectBlob   =   "ENTDEV.frx":0000
      TabIndex        =   3
      Top             =   840
      Width           =   11565
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2205
      ItemData        =   "ENTDEV.frx":09EA
      Left            =   3840
      List            =   "ENTDEV.frx":09F1
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   500
      Left            =   1320
      TabIndex        =   0
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Nro.Solicitud :"
      Height          =   255
      Left            =   7560
      TabIndex        =   32
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8160
      TabIndex        =   28
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label9 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Movimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgEntdev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 5 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Terminado As String
Private Auxiliar(100, 10) As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstGuia As Recordset
Dim spGuia As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedidoDevol As Recordset
Dim spPedidoDevol As String
Dim XParam As String

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    WTerminado.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WLote.Text = ""
    WLaboratorio.Text = ""
    WLinea.Text = ""
    
    WTerminado.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    PrgEntdev.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
  
    
    
  Listado.SQLQuery = "SELECT  Entdev.Codigo, Entdev.Fecha, Entdev.Terminado, Entdev.Cantidad, Entdev.Observaciones, Entdev.Lote, Entdev.Cliente, Entdev.Laboratorio, Entdev.ImpreTerminado, Entdev.Pedido.cliente.cliente,cliente.razon " _
                + "From " _
                + DSQ + ".dbo.entdev entdev, " _
                + DSQ + ".dbo.cliente cliente " _
                + "Where " _
                + "entdev.Cliente= cliente.cliente  AND " _
                + "entdev.Codigo >= '" + Codigo.Text + "' AND " _
                + "entdev.Codigo <= '" + Codigo.Text + "'"
    
               
 
 
 Listado.ReportFileName = "entdev2.rpt"

    Listado.Destination = 0
    Listado.Action = 1
    Rem by nan
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Productos Terminados"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    Rem OPEN_FILE_Entdev
    Rem OPEN_FILE_TERMINADO
End Sub

Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            XEmpresa = Wempresa
            If Val(Wempresa) <> 8 Then
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
        
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                IngresaItem = rstCliente!Cliente + "     " + rstCliente!Razon
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstCliente!Cliente
                                WIndice.AddItem IngresaItem
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                End With
                rstCliente.Close
            End If
            
            Call Conecta_Empresa
    
        Case 1
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Left$(rstTerminado!Codigo, 2) = "NK" Then
                                IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstTerminado!Codigo
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WTerminado.Text = DBGrid1.Text
            Else
        WTerminado.Text = "  -     -   "
        WLinea.Text = ""
    End If

    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    WCantidad.Text = DBGrid1.Text
    
    DBGrid1.Col = 3
    WLote.Text = DBGrid1.Text
    
    DBGrid1.Col = 4
    WLaboratorio.Text = DBGrid1.Text
    
    WTerminado.SetFocus

End Sub

Private Sub Graba_Click()

    If Val(Pedido.Text) = 0 Then
    
        m$ = "No se ha informado el numero de solicitud de pedido de devolucion correspondiente"
        G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
        Exit Sub
        
            Else
            
        XEmpresa = Wempresa
        If Val(Wempresa) <> 8 Then
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoDevol"
        ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + Pedido.Text + "'"
        ZSql = ZSql + " and PedidoDevol.Cliente = " + "'" + Cliente.Text + "'"
        spPedidoDevol = ZSql
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoDevol.RecordCount > 0 Then
            rstPedidoDevol.Close
                Else
            Call Conecta_Empresa
            m$ = "Hay datos de la solicitud de devolucion que son incorrectos"
            G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
            Exit Sub
        End If
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoDevol"
        ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + Pedido.Text + "'"
        spPedidoDevol = ZSql
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoDevol.RecordCount > 0 Then
            WBloqueo = IIf(IsNull(rstPedidoDevol!Bloqueo), "", rstPedidoDevol!Bloqueo)
            If WBloqueo = "S" Or WBloqueo = "P" Then
                m$ = "Los productos deben remitirse a Laboratorio para su analisis respectivo"
                G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
                    Else
                T$ = "Entrada de devolucion de mercaderia"
                m$ = "Los productos se encuentran liberados y no deben ser enviados a laboratorio. " + Chr$(13) + "Confirma que los productos se encuentran en condiciones optimas"
                Respuesta% = MsgBox(m$, 32 + 4, T$)
                If Respuesta% = 7 Then
                    T$ = "Entrada de devolucion de mercaderia"
                    m$ = "Confirma el envio de los productos a Laboratorio para su analisis"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 7 Then
                        WBloqueo = "P"
                    End If
                End If
            End If
            rstPedidoDevol.Close
        End If
        
        Call Conecta_Empresa
    
    End If

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = 0
    Erase Auxiliar
    
    DBGrid1.Refresh
                
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Terminado = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Lote = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Laboratorio = DBGrid1.Text
                    
            If Cantidad <> "" Then
            
                Rem
                Rem verifica si existe el codigo
                Rem
        
                If Left$(Terminado, 2) = "DK" Or Left$(Terminado, 2) = "NS" Then
                
                    WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                    
                        rstArticulo.Close
                        
                            Else
                            
                        If Left$(Terminado, 2) = "DK" Then
                            WArti = "DY-" + Right$(Terminado, 7)
                            ZCodigo = "DK-" + Right$(Terminado, 7)
                                Else
                            If Left$(Terminado, 2) = "NQ" Then
                                WArti = "DQ-" + Right$(Terminado, 7)
                                ZCodigo = "NQ-" + Right$(Terminado, 7)
                                    Else
                                WArti = "DS-" + Right$(Terminado, 7)
                                ZCodigo = "NS-" + Right$(Terminado, 7)
                            End If
                        End If
                               
                        If Left$(Terminado, 2) = "DK" Then
                        
                            WArti = "DY-" + Right$(Terminado, 7)
                            
                            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                            
                                ZCodigo = "DK-" + Right$(rstArticulo!Codigo, 7)
                                ZDescripcion = rstArticulo!Descripcion
                                ZUnidad = rstArticulo!Unidad
                                ZDeposito = rstArticulo!Deposito
                                ZInicial = ""
                                ZEntradas = ""
                                ZSalidas = ""
                                ZMinimo = ""
                                ZMinimo1 = ""
                                ZLaboratorio = ""
                                ZPedido = ""
                                ZVenta = ""
                                ZEnvase = Str$(rstArticulo!Envase)
                                ZCosto1 = Str$(rstArticulo!Costo1)
                                ZCosto2 = Str$(rstArticulo!Costo2)
                                ZCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", Str$(rstArticulo!Costo3))
                                ZWCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", Str$(rstArticulo!WCosto1))
                                ZWCosto2 = IIf(IsNull(rstArticulo!WCosto2), "0", Str$(rstArticulo!WCosto2))
                                ZWCosto3 = IIf(IsNull(rstArticulo!WCosto3), "0", Str$(rstArticulo!WCosto3))
                                ZCodSedronar = IIf(IsNull(rstArticulo!CodSedronar), "", rstArticulo!CodSedronar)
                                ZRs = rstArticulo!Rs
                                ZFlete = Str$(rstArticulo!Flete)
                                ZMoneda = rstArticulo!Moneda
                                ZControla = IIf(IsNull(rstArticulo!Controla), "0", Str$(rstArticulo!Controla))
                                ZDensidad = IIf(IsNull(rstArticulo!Densidad), "", rstArticulo!Densidad)
                                If rstArticulo!Proveedor <> "" Then
                                    ZProveedor = rstArticulo!Proveedor
                                        Else
                                    ZProveedor = ""
                                End If
                                ZFecha = ""
                                ZOrden = ""
                                ZDife = ""
                                ZDate = Date$
                                    
                                rstArticulo.Close
                                
                                XParam = "'" + ZCodigo + "','" _
                                         + ZDescripcion + "','" _
                                         + ZCosto1 + "','" _
                                         + ZCosto2 + "','" _
                                         + ZInicial + "','" _
                                         + ZEntradas + "','" _
                                         + ZSalidas + "','" _
                                         + ZMinimo + "','" _
                                         + ZLaboratorio + "','" _
                                         + ZUnidad + "','" _
                                         + ZPedido + "','" _
                                         + ZDeposito + "','" _
                                         + ZEnvase + "','" _
                                         + ZRs + "','" _
                                         + ZFecha + "','" _
                                         + ZOrden + "','" _
                                         + ZDife + "','" _
                                         + ZProveedor + "','" _
                                         + ZDate + "','" + ZFlete + "','" _
                                         + ZMoneda + "','" + ZControla + "','" _
                                         + ZDensidad + "','" + ZCosto3 + "','" _
                                         + ZWCosto1 + "','" + ZWCosto2 + "','" _
                                         + ZWCosto3 + "','" _
                                         + ZVenta + "'"
                                Set rstArticulo = db.OpenRecordset("AltaArticuloII " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                               
                            End If
                            
                        End If
                        
                                
                                
                                
                                
                        If Left$(Terminado, 2) = "NS" Then
                        
                            WArti = "DS-" + Right$(Terminado, 7)
                            
                            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                            
                                ZCodigo = "NS-" + Right$(rstArticulo!Codigo, 7)
                                ZDescripcion = rstArticulo!Descripcion
                                ZUnidad = rstArticulo!Unidad
                                ZDeposito = rstArticulo!Deposito
                                ZInicial = ""
                                ZEntradas = ""
                                ZSalidas = ""
                                ZMinimo = ""
                                ZMinimo1 = ""
                                ZLaboratorio = ""
                                ZPedido = ""
                                ZVenta = ""
                                ZEnvase = Str$(rstArticulo!Envase)
                                ZCosto1 = Str$(rstArticulo!Costo1)
                                ZCosto2 = Str$(rstArticulo!Costo2)
                                ZCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", Str$(rstArticulo!Costo3))
                                ZWCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", Str$(rstArticulo!WCosto1))
                                ZWCosto2 = IIf(IsNull(rstArticulo!WCosto2), "0", Str$(rstArticulo!WCosto2))
                                ZWCosto3 = IIf(IsNull(rstArticulo!WCosto3), "0", Str$(rstArticulo!WCosto3))
                                ZRs = rstArticulo!Rs
                                ZFlete = Str$(rstArticulo!Flete)
                                ZMoneda = rstArticulo!Moneda
                                ZControla = IIf(IsNull(rstArticulo!Controla), "0", Str$(rstArticulo!Controla))
                                ZDensidad = IIf(IsNull(rstArticulo!Densidad), "", rstArticulo!Densidad)
                                If rstArticulo!Proveedor <> "" Then
                                    ZProveedor = rstArticulo!Proveedor
                                        Else
                                    ZProveedor = ""
                                End If
                                ZFecha = ""
                                ZOrden = ""
                                ZDife = ""
                                ZDate = Date$
                                    
                                rstArticulo.Close
                                
                                XParam = "'" + ZCodigo + "','" _
                                         + ZDescripcion + "','" _
                                         + ZCosto1 + "','" _
                                         + ZCosto2 + "','" _
                                         + ZInicial + "','" _
                                         + ZEntradas + "','" _
                                         + ZSalidas + "','" _
                                         + ZMinimo + "','" _
                                         + ZLaboratorio + "','" _
                                         + ZUnidad + "','" _
                                         + ZPedido + "','" _
                                         + ZDeposito + "','" _
                                         + ZEnvase + "','" _
                                         + ZRs + "','" _
                                         + ZFecha + "','" _
                                         + ZOrden + "','" _
                                         + ZDife + "','" _
                                         + ZProveedor + "','" _
                                         + ZDate + "','" + ZFlete + "','" _
                                         + ZMoneda + "','" + ZControla + "','" _
                                         + ZDensidad + "','" + ZCosto3 + "','" _
                                         + ZWCosto1 + "','" + ZWCosto2 + "','" _
                                         + ZWCosto3 + "','" _
                                         + ZVenta + "'"
                                Set rstArticulo = db.OpenRecordset("AltaArticuloII " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                               
                            End If
                            
                        End If
                        
                        
                        
                        
                        If Left$(Terminado, 2) = "NQ" Then
                        
                            WArti = "DQ-" + Right$(Terminado, 7)
                            
                            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                            
                                ZCodigo = "NQ-" + Right$(rstArticulo!Codigo, 7)
                                ZDescripcion = rstArticulo!Descripcion
                                ZUnidad = rstArticulo!Unidad
                                ZDeposito = rstArticulo!Deposito
                                ZInicial = ""
                                ZEntradas = ""
                                ZSalidas = ""
                                ZMinimo = ""
                                ZMinimo1 = ""
                                ZLaboratorio = ""
                                ZPedido = ""
                                ZVenta = ""
                                ZEnvase = Str$(rstArticulo!Envase)
                                ZCosto1 = Str$(rstArticulo!Costo1)
                                ZCosto2 = Str$(rstArticulo!Costo2)
                                ZCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", Str$(rstArticulo!Costo3))
                                ZWCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", Str$(rstArticulo!WCosto1))
                                ZWCosto2 = IIf(IsNull(rstArticulo!WCosto2), "0", Str$(rstArticulo!WCosto2))
                                ZWCosto3 = IIf(IsNull(rstArticulo!WCosto3), "0", Str$(rstArticulo!WCosto3))
                                ZRs = rstArticulo!Rs
                                ZFlete = Str$(rstArticulo!Flete)
                                ZMoneda = rstArticulo!Moneda
                                ZControla = IIf(IsNull(rstArticulo!Controla), "0", Str$(rstArticulo!Controla))
                                ZDensidad = IIf(IsNull(rstArticulo!Densidad), "", rstArticulo!Densidad)
                                If rstArticulo!Proveedor <> "" Then
                                    ZProveedor = rstArticulo!Proveedor
                                        Else
                                    ZProveedor = ""
                                End If
                                ZFecha = ""
                                ZOrden = ""
                                ZDife = ""
                                ZDate = Date$
                                    
                                rstArticulo.Close
                                
                                XParam = "'" + ZCodigo + "','" _
                                         + ZDescripcion + "','" _
                                         + ZCosto1 + "','" _
                                         + ZCosto2 + "','" _
                                         + ZInicial + "','" _
                                         + ZEntradas + "','" _
                                         + ZSalidas + "','" _
                                         + ZMinimo + "','" _
                                         + ZLaboratorio + "','" _
                                         + ZUnidad + "','" _
                                         + ZPedido + "','" _
                                         + ZDeposito + "','" _
                                         + ZEnvase + "','" _
                                         + ZRs + "','" _
                                         + ZFecha + "','" _
                                         + ZOrden + "','" _
                                         + ZDife + "','" _
                                         + ZProveedor + "','" _
                                         + ZDate + "','" + ZFlete + "','" _
                                         + ZMoneda + "','" + ZControla + "','" _
                                         + ZDensidad + "','" + ZCosto3 + "','" _
                                         + ZWCosto1 + "','" + ZWCosto2 + "','" _
                                         + ZWCosto3 + "','" _
                                         + ZVenta + "'"
                                Set rstArticulo = db.OpenRecordset("AltaArticuloII " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                               
                            End If
                            
                        End If
                        
                        
                        
                        
                    End If
                End If
                    
                Anterior = "N"
                WClaveEntDev = ""
                
                If Left$(Terminado, 2) = "DK" Or Left$(Terminado, 2) = "NS" Then
                
                    Sql1 = "Select *"
                    Sql2 = " FROM EntDev"
                    Sql3 = " Where EntDev.Terminado = " + "'" + Terminado + "'"
                    Sql4 = " and EntDev.PartiOri = " + "'" + Lote + "'"
                    Sql5 = " Order by Clave"
                    spEntdev = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEntdev.RecordCount > 0 Then
                        WClaveEntDev = rstEntdev!Clave
                        Anterior = "S"
                        rstEntdev.Close
                    End If
                    
                        Else
                        
                    Sql1 = "Select *"
                    Sql2 = " FROM EntDev"
                    Sql3 = " Where EntDev.Terminado = " + "'" + Terminado + "'"
                    Sql4 = " and EntDev.Lote = " + "'" + Lote + "'"
                    spEntdev = Sql1 + Sql2 + Sql3 + Sql4
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEntdev.RecordCount > 0 Then
                        WClaveEntDev = rstEntdev!Clave
                        Anterior = "S"
                        rstEntdev.Close
                    End If
                    
                End If
                
                PartiOri = ""
                
                If Left$(Terminado, 2) = "DK" Or Left$(Terminado, 2) = "NS" Then
                
                    PartiOri = Lote
                    Lote = ""
                    WEntra = "N"
                    If Left$(Terminado, 2) = "DK" Then
                        WArti = "DY-" + Right$(Terminado, 7)
                            Else
                        If Left$(Terminado, 2) = "NS" Then
                            WArti = "DS-" + Right$(Terminado, 7)
                                Else
                            WArti = "DQ-" + Right$(Terminado, 7)
                        End If
                    End If
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + PartiOri + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            Lote = Str$(rstLaudo!Laudo)
                            WEntra = "S"
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + PartiOri + "'"
                        ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                Lote = Str$(rstMovguia!Lote)
                                WEntra = "S"
                                rstMovguia.Close
                            End With
                        End If
                    End If
                    
                End If
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Codigo.Text)
                Call Ceros(Auxi1, 6)
                
                WCodigo = Codigo.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WTerminado = Terminado
                WCantidad = Cantidad
                WObservaciones = Observaciones.Text
                WPedido = Pedido.Text
                WClave = Auxi1 + Auxi
                WMarca = ""
                WLote = Lote
                WPartiOri = PartiOri
                WCliente = Cliente.Text
                WLaboratorio = Laboratorio
                If Anterior = "N" Then
                    WSaldo = Str$(Val(WCantidad) - Val(WLaboratorio))
                        Else
                    WSaldo = Str$(Val(WCantidad))
                End If
                
                Auxiliar(Renglon, 1) = WTerminado
                Auxiliar(Renglon, 2) = WCantidad
                Auxiliar(Renglon, 3) = WLote
                Auxiliar(Renglon, 4) = WPartiOri
                
                WImpresion = "N"
                WImpreTerminado = ""
                WTrabajo = ""
                WImpreLaboI = "N"
                WImpreLaboII = "N"
                WImpreProdI = "N"
                WImpreProdII = "N"
                WImpreProdIII = "N"
                WImpreProdIV = "N"
                
                XTipoPro = ""
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DK" Or Left$(WTerminado, 2) = "DS" Or Left$(WTerminado, 2) = "NS" Or Left$(WTerminado, 2) = "DQ" Or Left$(WTerminado, 2) = "NQ" Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 12999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    XTipoPro = "PT"
                                End If
                            End If
                        End If
                    End If
                End If
                
                ZLinea = 0
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZLinea = rstTerminado!Linea
                    rstTerminado.Close
                End If
                
                Select Case ZLinea
                    Case 8
                        XTipoPro = "PG"
                    Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                        XTipoPro = "FA"
                    Case Else
                End Select
                
                WTipopro = XTipoPro
                
                If Left$(Terminado, 2) = "DK" Or Left$(Terminado, 2) = "NS" Then
                
                    If Left$(Terminado, 2) = "NS" Then
                        WArti = "DS-" + Right$(Terminado, 7)
                            Else
                        If Left$(Terminado, 2) = "DK" Then
                             WArti = "DY-" + Right$(Terminado, 7)
                                Else
                            WArti = "DQ-" + Right$(Terminado, 7)
                         End If
                    End If
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WImpreTerminado = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                            
                        Else
                        
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WImpreTerminado = rstTerminado!Descripcion
                        rstTerminado.Close
                    End If
                    
                End If
                
                
                
                
                
                WEstado = "PT"
                If WBloqueo = "S" Or WBloqueo = "P" Then
                    WEstado = "NK"
                End If
                
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO EntDev ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "FechaOrd ,"
                ZSql = ZSql + "Observaciones,"
                ZSql = ZSql + "Pedido,"
                ZSql = ZSql + "Marca,"
                ZSql = ZSql + "Lote,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Saldo,"
                ZSql = ZSql + "Laboratorio,"
                ZSql = ZSql + "PartiOri,"
                ZSql = ZSql + "Impresion,"
                ZSql = ZSql + "ImpreTerminado,"
                ZSql = ZSql + "Trabajo,"
                ZSql = ZSql + "ImpreLaboI,"
                ZSql = ZSql + "ImpreLaboII,"
                ZSql = ZSql + "ImpreProdI,"
                ZSql = ZSql + "ImpreProdII,"
                ZSql = ZSql + "ImpreProdIII,"
                ZSql = ZSql + "ImpreProdIV,"
                ZSql = ZSql + "TipoPro,"
                ZSql = ZSql + "Estado )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WCodigo + "',"
                ZSql = ZSql + "'" + WRenglon + "',"
                ZSql = ZSql + "'" + WFecha + "',"
                ZSql = ZSql + "'" + WTerminado + "',"
                ZSql = ZSql + "'" + WCantidad + "',"
                ZSql = ZSql + "'" + WFechaord + "',"
                ZSql = ZSql + "'" + WObservaciones + "',"
                ZSql = ZSql + "'" + WPedido + "',"
                ZSql = ZSql + "'" + WMarca + "',"
                ZSql = ZSql + "'" + WLote + "',"
                ZSql = ZSql + "'" + WCliente + "',"
                ZSql = ZSql + "'" + WSaldo + "',"
                ZSql = ZSql + "'" + WLaboratorio + "',"
                ZSql = ZSql + "'" + WPartiOri + "',"
                ZSql = ZSql + "'" + WImpresion + "',"
                ZSql = ZSql + "'" + WImpreTerminado + "',"
                ZSql = ZSql + "'" + WTrabajo + "',"
                ZSql = ZSql + "'" + WImpreLaboI + "',"
                ZSql = ZSql + "'" + WImpreLaboII + "',"
                ZSql = ZSql + "'" + WImpreProdI + "',"
                ZSql = ZSql + "'" + WImpreProdII + "',"
                ZSql = ZSql + "'" + WImpreProdIII + "',"
                ZSql = ZSql + "'" + WImpreProdIV + "',"
                ZSql = ZSql + "'" + WTipopro + "',"
                ZSql = ZSql + "'" + WEstado + "')"
        
                spEntdev = ZSql
                Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                
                If Anterior = "S" Then
                    Sql1 = "UPDATE EntDev SET "
                    Sql2 = "Saldo = Saldo + " + "'" + WCantidad + "'"
                    Sql3 = " Where Clave = " + "'" + WClaveEntDev + "'"
                    spEntdev = Sql1 + Sql2 + Sql3
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                If WBloqueo = "S" Or WBloqueo = "P" Then
                
                    XEmpresa = Wempresa
                    If Val(Wempresa) <> 8 Then
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End If
                
                    Sql1 = "UPDATE PedidoDevol SET "
                    Sql2 = "ImpreLaboI =  " + "'" + "N" + "',"
                    Sql3 = "ImpreLaboII =  " + "'" + "N" + "',"
                    Sql4 = "NroDev =  " + "'" + Codigo.Text + "'"
                    Sql5 = " Where Pedido = " + "'" + Pedido.Text + "'"
                    spPedidoDevol = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Call Conecta_Empresa
                    
                        Else
                        
                    XEmpresa = Wempresa
                    If Val(Wempresa) <> 8 Then
                        Wempresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End If
                
                    Sql1 = "UPDATE PedidoDevol SET "
                    Sql2 = "NroDev =  " + "'" + Codigo.Text + "'"
                    Sql3 = " Where Pedido = " + "'" + Pedido.Text + "'"
                    spPedidoDevol = Sql1 + Sql2 + Sql3
                    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Call Conecta_Empresa
                    
                End If
                
            End If
                
        Next iRow
            
    Next A
    
    For Da = 1 To Renglon
    
        Terminado = Auxiliar(Da, 1)
        Cantidad = Auxiliar(Da, 2)
        Lote = Auxiliar(Da, 3)
        PartiOri = Auxiliar(Da, 4)
        
        If Left$(Terminado, 2) = "DY" Or Left$(Terminado, 2) = "DK" Or Left$(Terminado, 2) = "DS" Or Left$(Terminado, 2) = "NS" Or Left$(Terminado, 2) = "DQ" Or Left$(Terminado, 2) = "NQ" Then
        
            WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
            
            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = rstArticulo!Codigo
                WEntradas = Str$(rstArticulo!Entradas + Val(Cantidad))
                WSalidas = Str$(rstArticulo!Salidas)
                rstArticulo.Close
                XParam = "'" + WCodigo + "','" _
                        + WEntradas + "','" _
                        + WSalidas + "','" _
                        + WDate + "'"
                spArticulo = "ModificaArticuloMovimientos " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            Rem WEntra = "N"
            Rem If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Then
            Rem     WArti = "DY-" + Right$(WTerminado.Text, 7)
            Rem         Else
            Rem     WArti = "DW-" + Right$(WTerminado.Text, 7)
            Rem End If
            Rem
            Rem Sql1 = "Select *"
            Rem Sql2 = " FROM Laudo"
            Rem Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
            Rem Sql4 = " and Laudo.Lote = " + "'" + WLote.Text + "'"
            Rem Sql5 = " Order by Laudo.Laudo"
            Rem spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstLaudo.RecordCount > 0 Then
            Rem     rstLaudo.Close
            Rem     WEntra = "S"
            Rem     WMarcaEstado = "N"
            Rem     Sql1 = "UPDATE Laudo SET "
            Rem     Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
            Rem     Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
            Rem     Sql4 = " and Laudo.Lote = " + "'" + WLote.Text + "'"
            Rem     spLaudo = Sql1 + Sql2 + Sql3 + Sql4
            Rem     Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            Rem End If
            Rem
            Rem If WEntra = "N" Then
            Rem     Sql1 = "Select *"
            Rem     Sql2 = " FROM Guia"
            Rem     Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
            Rem     Sql4 = " and Guia.PartiOri = " + "'" + PartiOri + "'"
            Rem     Sql5 = " Order by Guia.Saldo desc"
            Rem     spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Rem     Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstMovguia.RecordCount > 0 Then
            Rem         rstMovguia.Close
            Rem         WMarcaEstado = "N"
            Rem         Sql1 = "UPDATE Guia SET "
            Rem         Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
            Rem         Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
            Rem         Sql4 = " and Guia.PartiOri = " + "'" + PartiOri + "'"
            Rem         spMovguia = Sql1 + Sql2 + Sql3 + Sql4
            Rem         Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            Rem     End If
            Rem End If
            
                    Else
                
            spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = Terminado
                WEntradas = Str$(rstTerminado!Entradas + Val(Cantidad))
                WSalidas = Str$(rstTerminado!Salidas)
                rstTerminado.Close
                XParam = "'" + WCodigo + "','" _
                        + WEntradas + "','" _
                        + WSalidas + "','" _
                        + WDate + "'"
                spTerminado = "ModificaTerminadoMovimientos " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            Rem WEntra = "N"
            Rem WTer = "PT" + Mid$(WTerminado.Text, 3, 10)
            Rem
            Rem XParam = "'" + WLote.Text + "','" _
            rem              + WTer + "'"
            Rem spHoja = "ListaHojaProducto " + XParam
            Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstHoja.RecordCount > 0 Then
            Rem     rstHoja.Close
            Rem     WEntra = "S"
            Rem     WMarcaEstado = "N"
            Rem     Sql1 = "UPDATE Hoja SET "
            Rem     Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
            Rem     Sql3 = " Where Hoja.Producto = " + "'" + WTer + "'"
            Rem     Sql4 = " and Hoja.Hoja = " + "'" + WLote.Text + "'"
            Rem     spHoja = Sql1 + Sql2 + Sql3 + Sql4
            Rem     Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            Rem End If
            Rem
            Rem If WEntra = "N" Then
            Rem     XParam = "'" + WTer + "','" _
            rem                  + WLote.Text + "'"
            Rem     spMovguia = "ListaMovguiaLote1 " + XParam
            Rem     Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstMovguia.RecordCount > 0 Then
            Rem         rstMovguia.Close
            Rem         WMarcaEstado = "N"
            Rem         Sql1 = "UPDATE Guia SET "
            Rem         Sql2 = "Estado  = " + "'" + WMarcaEstado + "'"
            Rem         Sql3 = " Where Guia.Terminado = " + "'" + WTer + "'"
            Rem         Sql4 = " and Guia.Lote = " + "'" + WLote.Text + "'"
            Rem         spMovguia = Sql1 + Sql2 + Sql3 + Sql4
            Rem         Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            Rem     End If
            Rem End If
            
        End If
        
        If WBloqueo = "N" Then
        
            Rem graba la liberafcion automatica
            
            Select Case Left$(Terminado, 2)
                Case "NK"
                    ZTerminado = "PT" + Mid$(Terminado, 3, 10)
                Case "DK"
                    ZTerminado = "DY" + Mid$(Terminado, 3, 10)
                Case "NS"
                    ZTerminado = "DS" + Mid$(Terminado, 3, 10)
                Case "NQ"
                    ZTerminado = "DQ" + Mid$(Terminado, 3, 10)
                Case Else
                    ZTerminado = "PT" + Mid$(Terminado, 3, 10)
            End Select
                            
            Rem ZTerminado = "PT" + Mid$(Terminado, 3, 20)
            ZCantidad = Cantidad
            ZLote = Lote
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            
            ZSql = ""
            ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
            ZSql = ZSql + " FROM LiberaTerminado"
            spLiberaTerminado = ZSql
            Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstLiberaTerminado.RecordCount > 0 Then
                rstLiberaTerminado.MoveLast
                WCodigoMayor = IIf(IsNull(rstLiberaTerminado!CodigoMayor), "0", rstLiberaTerminado!CodigoMayor)
                Lote = Str$(WCodigoMayor)
                rstLiberaTerminado.Close
                    Else
                Lote = "0"
            End If
            
            WCodigo = Str$(Val(Lote) + 1)
            WProducto = ZTerminado
            WFecha = ZFecha
            WFechaord = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
            WPartida = ZLote
            WPartiOri = ""
            WValor1 = ""
            Wvalor2 = ""
            WValor3 = ""
            Wvalor4 = ""
            Wvalor5 = ""
            Wvalor6 = ""
            Wvalor7 = ""
            Wvalor8 = ""
            Wvalor9 = ""
            Wvalor10 = ""
            WEnsayo = ""
            WAspecto = ""
            WObservaciones = "Liberacion Automatica"
            WConfecciono = ""
            WMarca = "N"
            WCliente = Cliente.Text
            WObserva = ""
            Select Case Left$(Terminado, 2)
                Case "DK", "NS", "NQ"
                    WObserva = "Partida : " + PartiOri
                Case Else
                    WObserva = ""
            End Select
            WCantidad = ZCantidad
            WFacturado = "0"
            WOrigen = "L"
            WTipo = Left$(ZTerminado, 2)
            WImpreProdI = "N"
            WImpreProdII = "N"
            WImpreProdIII = "N"
            WImpreVentas = "N"
            WTipopro = ""
            
            XTipoPro = ""
            XCodigo = Val(Mid$(ZTerminado, 4, 5))
            If Left$(ZTerminado, 2) = "DY" Or Left$(ZTerminado, 2) = "DS" Or Left$(ZTerminado, 2) = "DQ" Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 0 And XCodigo <= 999 Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 11000 And XCodigo <= 12999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 25000 And XCodigo <= 25999 Then
                            XTipoPro = "FA"
                                Else
                            If XCodigo >= 2300 And XCodigo <= 2399 Then
                                XTipoPro = "BI"
                                    Else
                                XTipoPro = "PT"
                            End If
                        End If
                    End If
                End If
            End If
            
            ZLinea = 0
            spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZLinea = rstTerminado!Linea
                rstTerminado.Close
            End If
            
            Select Case ZLinea
                Case 8
                    XTipoPro = "PG"
                Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                    XTipoPro = "FA"
                Case Else
            End Select
            
            WTipopro = XTipoPro
            
            Rem Select Case WTipopro
            Rem     Case "CO", "PG"
            Rem         WImpreProdI = "S"
            Rem     Case "BI", "PT"
            Rem         WImpreProdII = "S"
            Rem     Case "FA"
            Rem         WImpreProdIII = "S"
            Rem     Case Else
            Rem End Select
            
            ZSql = ""
            ZSql = ZSql & "INSERT INTO LiberaTerminado ("
            ZSql = ZSql & "Codigo, "
            ZSql = ZSql & "Producto, "
            ZSql = ZSql & "Fecha, "
            ZSql = ZSql & "OrdFecha, "
            ZSql = ZSql & "Partida, "
            ZSql = ZSql & "PartiOri, "
            ZSql = ZSql & "PedidoDevol, "
            ZSql = ZSql & "Valor1, "
            ZSql = ZSql & "Valor2, "
            ZSql = ZSql & "Valor3, "
            ZSql = ZSql & "Valor4, "
            ZSql = ZSql & "Valor5, "
            ZSql = ZSql & "Valor6, "
            ZSql = ZSql & "Valor7, "
            ZSql = ZSql & "Valor8, "
            ZSql = ZSql & "Valor9, "
            ZSql = ZSql & "Valor10, "
            ZSql = ZSql & "Ensayo, "
            ZSql = ZSql & "Aspecto, "
            ZSql = ZSql & "Observaciones, "
            ZSql = ZSql & "Confecciono, "
            ZSql = ZSql & "Marca, "
            ZSql = ZSql & "Cliente, "
            ZSql = ZSql & "Cantidad, "
            ZSql = ZSql & "Facturado, "
            ZSql = ZSql & "Observa, "
            ZSql = ZSql & "Origen, "
            ZSql = ZSql & "Tipo, "
            ZSql = ZSql & "ImpreProdI, "
            ZSql = ZSql & "ImpreProdII, "
            ZSql = ZSql & "ImpreProdIII, "
            ZSql = ZSql & "ImpreVentas, "
            ZSql = ZSql & "TipoPro) "
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + WCodigo + "',"
            ZSql = ZSql & "'" + WProducto + "',"
            ZSql = ZSql & "'" + WFecha + "',"
            ZSql = ZSql & "'" + WOrdFecha + "',"
            ZSql = ZSql & "'" + WPartida + "',"
            ZSql = ZSql & "'" + WPartiOri + "',"
            ZSql = ZSql & "'" + Codigo.Text + "',"
            ZSql = ZSql & "'" + WValor1 + "',"
            ZSql = ZSql & "'" + Wvalor2 + "',"
            ZSql = ZSql & "'" + WValor3 + "',"
            ZSql = ZSql & "'" + Wvalor4 + "',"
            ZSql = ZSql & "'" + Wvalor5 + "',"
            ZSql = ZSql & "'" + Wvalor6 + "',"
            ZSql = ZSql & "'" + Wvalor7 + "',"
            ZSql = ZSql & "'" + Wvalor8 + "',"
            ZSql = ZSql & "'" + Wvalor9 + "',"
            ZSql = ZSql & "'" + Wvalor10 + "',"
            ZSql = ZSql & "'" + WEnsayo + "',"
            ZSql = ZSql & "'" + WAspecto + "',"
            ZSql = ZSql & "'" + WObservaciones + "',"
            ZSql = ZSql & "'" + WConfecciono + "',"
            ZSql = ZSql & "'" + WMarca + "',"
            ZSql = ZSql & "'" + WCliente + "',"
            ZSql = ZSql & "'" + WCantidad + "',"
            ZSql = ZSql & "'" + WFacturado + "',"
            ZSql = ZSql & "'" + WObserva + "',"
            ZSql = ZSql & "'" + WOrigen + "',"
            ZSql = ZSql & "'" + WTipo + "',"
            ZSql = ZSql & "'" + WImpreProdI + "',"
            ZSql = ZSql & "'" + WImpreProdII + "',"
            ZSql = ZSql & "'" + WImpreProdIII + "',"
            ZSql = ZSql & "'" + WImpreVentas + "',"
            ZSql = ZSql & "'" + WTipopro + "')"
            
            spLiberaTerminado = ZSql
            Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)


            XEmpresa = Wempresa
            If Val(Wempresa) <> 8 Then
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            Sql1 = "UPDATE PedidoDevol SET "
            Sql2 = "Facturado =  Cantidad"
            Sql3 = " Where Pedido = " + "'" + Pedido.Text + "'"
            Sql4 = " and Terminado = " + "'" + WProducto + "'"
            spPedidoDevol = Sql1 + Sql2 + Sql3 + Sql4
            Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
            
            Call Conecta_Empresa
            
        End If
        
    Next Da
        
  Rem by nan
  
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
  
     Listado.SQLQuery = "SELECT  Entdev.Codigo, Entdev.Fecha, Entdev.Terminado, Entdev.Cantidad, Entdev.Observaciones, Entdev.Lote, Entdev.Cliente, Entdev.Laboratorio, Entdev.ImpreTerminado, Entdev.Pedido.cliente.cliente,cliente.razon " _
                + "From " _
                + DSQ + ".dbo.entdev entdev, " _
                + DSQ + ".dbo.cliente cliente " _
                + "Where " _
                + "entdev.Cliente= cliente.cliente  AND " _
                + "entdev.Codigo >= '" + Codigo.Text + "' AND " _
                + "entdev.Codigo <= '" + Codigo.Text + "'"
    
               
 
 
      Listado.ReportFileName = "entdev2.rpt"
    
  
    Listado.Destination = 1
    Listado.Action = 1
    Rem by nan
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
        
        
End Sub


Private Sub Ingresa_Click()

    WLinea.Text = ""
    WTerminado.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WLote.Text = ""
    WLaboratorio.Text = ""
    
    WTerminado.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WTerminado.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WLote.Text = ""
    WLaboratorio.Text = ""

    Codigo.Text = ""
    Observaciones.Text = ""
    Pedido.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = ""
    DesCliente.Caption = ""
    
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 4
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Codigo.Text = "1"
    
    spEntdev = "ListaEntdevNumero"
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
        With rstEntdev
            .MoveLast
            Codigo.Text = rstEntdev!Codigo + 1
        End With
        rstEntdev.Close
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Codigo.SetFocus

End Sub

Private Sub WTerminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WTerminado.Text = UCase(WTerminado.Text)
    
        If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Or Left$(WTerminado.Text, 2) = "DS" Or Left$(WTerminado.Text, 2) = "NS" Or Left$(WTerminado.Text, 2) = "DQ" Or Left$(WTerminado.Text, 2) = "NQ" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Then
                    WArti = "DY-" + Right$(WTerminado.Text, 7)
                        Else
                    If Left$(WTerminado.Text, 2) = "DQ" Or Left$(WTerminado.Text, 2) = "NQ" Then
                        WArti = "DQ-" + Right$(WTerminado.Text, 7)
                            Else
                        WArti = "DS-" + Right$(WTerminado.Text, 7)
                    End If
                End If
                    
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                    WCantidad.SetFocus
                        Else
                    WTerminado.SetFocus
                End If
                
            Case Else
                spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescripcion.Caption = rstTerminado!Descripcion
                    rstTerminado.Close
                    WCantidad.SetFocus
                        Else
                    WTerminado.SetFocus
                End If
                
        End Select
        
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WLote.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Or Left$(WTerminado.Text, 2) = "DS" Or Left$(WTerminado.Text, 2) = "NS" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WEntra = "N"
                If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Or Left$(WTerminado.Text, 2) = "DS" Or Left$(WTerminado.Text, 2) = "NS" Then
                    If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Then
                        WArti = "DY-" + Right$(WTerminado.Text, 7)
                            Else
                        WArti = "DS-" + Right$(WTerminado.Text, 7)
                    End If
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote.Text + "'"
                ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WEntra = "S"
                    rstLaudo.Close
                End If
                        
                If WEntra = "N" Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote.Text + "'"
                    ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                If WEntra = "S" Then
                    WLaboratorio.SetFocus
                        Else
                    m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
                    WLote.SetFocus
                End If
    
            Case Else
                WEntra = "N"
                WTer = "PT" + Mid$(WTerminado.Text, 3, 10)
                
                XParam = "'" + WLote.Text + "','" _
                             + WTer + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTer + "','" _
                            + WLote.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                If WEntra = "S" Then
                    WLaboratorio.SetFocus
                        Else
                    m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
                    WLote.SetFocus
                End If
                
        End Select
        
    End If
    
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub WLaboratorio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLaboratorio.Text = Pusing("###,###.##", WLaboratorio.Text)
        
        XEmpresa = Wempresa
        If Val(Wempresa) <> 8 Then
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        If Left$(WTerminado.Text, 2) = "NK" Or Left$(WTerminado.Text, 2) = "DK" Or Left$(WTerminado.Text, 2) = "NS" Then
            Select Case Left$(WTerminado.Text, 2)
                Case "DK"
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM PedidoDevol"
                    ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + Pedido.Text + "'"
                    ZSql = ZSql + " and PedidoDevol.Cliente = " + "'" + Cliente.Text + "'"
                    ZSql = ZSql + " and PedidoDevol.Terminado = " + "'" + "DY" + Mid$(WTerminado.Text, 3, 20) + "'"
                
                Case "NS"
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM PedidoDevol"
                    ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + Pedido.Text + "'"
                    ZSql = ZSql + " and PedidoDevol.Cliente = " + "'" + Cliente.Text + "'"
                    ZSql = ZSql + " and PedidoDevol.Terminado = " + "'" + "DS" + Mid$(WTerminado.Text, 3, 20) + "'"

                Case Else
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM PedidoDevol"
                    ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + Pedido.Text + "'"
                    ZSql = ZSql + " and PedidoDevol.Cliente = " + "'" + Cliente.Text + "'"
                    ZSql = ZSql + " and PedidoDevol.Terminado = " + "'" + "PT" + Mid$(WTerminado.Text, 3, 20) + "'"
                    
            End Select
            
            spPedidoDevol = ZSql
            Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedidoDevol.RecordCount > 0 Then
                ZPedido = IIf(IsNull(rstPedidoDevol!nroDev), "0", rstPedidoDevol!nroDev)
                rstPedidoDevol.Close
                If ZPedido <> 0 Then
                    m$ = "Pedido ya actualizado"
                    G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
                        Else
                    WTerminado.SetFocus
                    Call Alta_Vector
                    Call Ingresa_Click
                End If
                    Else
                m$ = "Los valores no coinciden con la Solicitud de devolucion"
                G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
            End If
            
                Else
                
            m$ = "El producto a informar debe ser NK o DK o NW O DS"
            G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
            
        End If
        
        Call Conecta_Empresa
        
        WTerminado.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            XEmpresa = Wempresa
            If Val(Wempresa) <> 8 Then
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
        
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = Claveven$
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
            End If
            
            Call Conecta_Empresa
    
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WTerminado.Text = Claveven$
                WDescripcion.Caption = rstTerminado!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = rstTerminado!Codigo
                DBGrid1.Col = 1
                DBGrid1.Text = rstTerminado!Descripcion
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                rstTerminado.Close
                    
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 4, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 4
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Prod.Terminado"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3620
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Verificado"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Codigo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = ""
    DesCliente.Caption = ""
    
    Codigo.Text = "1"
    
    spEntdev = "ListaEntdevNumero"
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
        With rstEntdev
            .MoveLast
            Codigo.Text = rstEntdev!Codigo + 1
        End With
        rstEntdev.Close
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgEntdev.Caption = "Entrada de Devolucion de Producto :  " + !Nombre
        End If
    End With
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 4
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Erase Auxiliar
    Renglon = 0
    
    spEntdev = "ListaEntdev " + "'" + Codigo.Text + "'"
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
        With rstEntdev
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstEntdev!Terminado
                    Auxi1 = rstEntdev!Terminado
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", rstEntdev!Cantidad)
                
                    If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DK" Then
                        DBGrid1.Col = 3
                        DBGrid1.Text = IIf(IsNull(rstEntdev!PartiOri), "", rstEntdev!PartiOri)
                            Else
                        DBGrid1.Col = 3
                        DBGrid1.Text = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                    End If
                    
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", rstEntdev!Laboratorio)
                    
                    Observaciones.Text = rstEntdev!Observaciones
                    Pedido.Text = IIf(IsNull(rstEntdev!Pedido), "", rstEntdev!Pedido)
                    Cliente.Text = rstEntdev!Cliente
                    
                    Auxiliar(Renglon, 1) = Auxi1
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEntdev.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0

    For Da = 1 To WRenglon
    
        Auxi1 = Auxiliar(Da, 1)
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DK" Then
            WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                DBGrid1.Col = 1
                DBGrid1.Text = rstArticulo!Descripcion
                rstArticulo.Close
            End If
                    Else
            spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DBGrid1.Col = 1
                DBGrid1.Text = rstTerminado!Descripcion
                rstTerminado.Close
            End If
        End If
        
    Next Da
    
    XEmpresa = Wempresa
    If Val(Wempresa) <> 8 Then
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
            Else
        DesCliente.Caption = ""
    End If
    
    Call Conecta_Empresa

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    WTerminado.SetFocus

End Sub

Private Sub Alta_Vector()

    If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Or Left$(WTerminado.Text, 2) = "DS" Or Left$(WTerminado.Text, 2) = "NS" Then
        WTipopro = "M"
            Else
        WTipopro = "T"
    End If
    
    Select Case WTipopro
        Case "M"
            WEntra = "N"
            If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Or Left$(WTerminado.Text, 2) = "DS" Or Left$(WTerminado.Text, 2) = "NS" Then
                If Left$(WTerminado.Text, 2) = "DY" Or Left$(WTerminado.Text, 2) = "DK" Then
                    WArti = "DY-" + Right$(WTerminado.Text, 7)
                        Else
                    WArti = "DS-" + Right$(WTerminado.Text, 7)
                End If
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
            ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote.Text + "'"
            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WEntra = "S"
                rstLaudo.Close
            End If
                    
            If WEntra = "N" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Guia"
                ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote.Text + "'"
                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                spMovguia = ZSql
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WEntra = "S"
                    rstMovguia.Close
                End If
            End If
            
            If WEntra = "S" Then
                Rem WLaboratorio.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
                Exit Sub
            End If

        Case Else
            WEntra = "N"
            WTer = "PT" + Mid$(WTerminado.Text, 3, 10)
            
            XParam = "'" + WLote.Text + "','" _
                         + WTer + "'"
            spHoja = "ListaHojaProducto " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                WEntra = "S"
                rstHoja.Close
            End If
            
            If WEntra = "N" Then
                XParam = "'" + WTer + "','" _
                        + WLote.Text + "'"
                spMovguia = "ListaMovguiaLote1 " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WEntra = "S"
                    rstMovguia.Close
                End If
            End If
            
            If WEntra = "S" Then
                Rem WLaboratorio.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
                Exit Sub
            End If
            
    End Select

    If Val(WLinea.Text) = 0 Then

        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        WAnterior = DBGrid1.Row
                    
        DBGrid1.Col = 0
        DBGrid1.Text = WTerminado.Text
            
        DBGrid1.Col = 1
        DBGrid1.Text = WDescripcion.Caption
                
        DBGrid1.Col = 2
        DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
        DBGrid1.Col = 3
        DBGrid1.Text = WLote.Text
            
        DBGrid1.Col = 4
        DBGrid1.Text = Pusing("###,###.##", WLaboratorio.Text)
            
        Rem DBGrid1.Row = Renglon
        DBGrid1.Col = 0
            
            Else
                
        DBGrid1.Row = Val(WLinea.Text) - 1
            
        WAnterior = DBGrid1.Row
            
        DBGrid1.Col = 0
        DBGrid1.Text = WTerminado.Text
            
        DBGrid1.Col = 1
        DBGrid1.Text = WDescripcion.Caption
                
        DBGrid1.Col = 2
        DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
        DBGrid1.Col = 3
        DBGrid1.Text = WLote.Text
            
        DBGrid1.Col = 4
        DBGrid1.Text = Pusing("###,###.##", WLaboratorio.Text)
            
        Rem DBGrid1.Row = Renglon
        DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEntdev = "ListaEntdev " + "'" + Codigo.Text + "'"
        Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        If rstEntdev.RecordCount > 0 Then
            Graba.Enabled = False
            Fecha.Text = rstEntdev!Fecha
            rstEntdev.Close
            Call Proceso_Click
                Else
            Graba.Enabled = True
            WCodigo = Codigo.Text
            Call Limpia_Click
            Codigo.Text = WCodigo
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
        
            XEmpresa = Wempresa
            If Val(Wempresa) <> 8 Then
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
        
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
                Pedido.SetFocus
                    Else
                Cliente.SetFocus
            End If
            
            Call Conecta_Empresa
            
        End If
    End If
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = Wempresa
        If Val(Wempresa) <> 8 Then
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoDevol"
        ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + Pedido.Text + "'"
        ZSql = ZSql + " and PedidoDevol.Cliente = " + "'" + Cliente.Text + "'"
        spPedidoDevol = ZSql
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoDevol.RecordCount > 0 Then
            rstPedidoDevol.Close
            WTerminado.SetFocus
                Else
            m$ = "Hay datos de la solicitud de devolucion que son incorrectos"
            G% = MsgBox(m$, 0, "Entrada de devolucion de mercaderia")
            Pedido.SetFocus
        End If
        
        Call Conecta_Empresa
        
    End If
End Sub

