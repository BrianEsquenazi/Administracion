VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVerificaPedido 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Pedidos Ingresados"
   ClientHeight    =   8385
   ClientLeft      =   105
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   8385
   ScaleWidth      =   11850
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Actualiza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3480
      TabIndex        =   2
      Top             =   7560
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12938
      _Version        =   327680
      Rows            =   4000
      Cols            =   8
      BackColor       =   16777088
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6120
      TabIndex        =   0
      Top             =   7560
      Width           =   1335
   End
End
Attribute VB_Name = "PrgVerificaPedido"
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
Dim XParam As String
Dim WGraba As String

Dim ZImpreConcepto(1000) As String
Dim ZZHasta As Integer

Private Sub cmdClose_Click()
    PrgVerificaPedido.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    ZImpreConcepto(0) = ""
    ZImpreConcepto(1) = "Error del Sistema"
    ZImpreConcepto(2) = "Varios"
    ZImpreConcepto(3) = "Problemas Vehiculos"
    ZImpreConcepto(4) = "Problemas Logistica"
    ZImpreConcepto(5) = "Problemas Recepcion Cliente"
    ZImpreConcepto(6) = "Varios"
    ZImpreConcepto(7) = "Corte de Luz"
    ZImpreConcepto(8) = "Pedido por el Cliente"
    ZImpreConcepto(9) = "Falta de Pago"
    ZImpreConcepto(10) = "Confirmacion Pedido Parcial"
    ZImpreConcepto(11) = "Envase"
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 450
    Muestra.ColWidth(1) = 800
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 800
    Muestra.ColWidth(4) = 2000
    Muestra.ColWidth(5) = 1200
    Muestra.ColWidth(6) = 1200
    Muestra.ColWidth(7) = 3700
    
    Muestra.ColAlignment(7) = 0
    
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
    Muestra.Text = "Aviso"
    
    Muestra.Row = 0
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        WTitulo(Ciclo).Text = Muestra.Text
        WTitulo(Ciclo).Left = Muestra.CellLeft + Muestra.Left
        WTitulo(Ciclo).Top = Muestra.CellTop + Muestra.Top
        WTitulo(Ciclo).Width = Muestra.CellWidth
        WTitulo(Ciclo).Height = Muestra.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    WPosi1 = 1
    WPosi2 = 1
    
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Else
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If

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
    Muestra.Text = "Aviso"
    
    Renglon = 0
    
    Sql1 = "Select DISTINCT Pedido.Pedido, Pedido.Fecha, Pedido.Cliente, Pedido.TipoPed, Pedido.FecEntrega, Cliente.Razon as [WRazon]"
    Sql2 = " FROM Pedido, CLiente"
    Sql3 = " Where Pedido.Cantidad - Pedido.Facturado > 0"
    Sql4 = " and Pedido.TipoPedido <> 1"
    Sql5 = " and Pedido.Autorizo <> " + "'" + "N" + "'"
    Sql6 = " and Pedido.Cliente = Cliente.Cliente"
    Sql7 = " Order by Pedido.Pedido"
    
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Muestra.TextMatrix(Renglon, 0) = Pusing("###", Str$(Renglon))
                    Muestra.TextMatrix(Renglon, 1) = Pusing("######", Str$(rstPedido!Pedido))
                    Muestra.TextMatrix(Renglon, 2) = rstPedido!Fecha
                    Muestra.TextMatrix(Renglon, 3) = rstPedido!Cliente
                    Muestra.TextMatrix(Renglon, 4) = rstPedido!WRazon
                    Muestra.TextMatrix(Renglon, 5) = rstPedido!FecEntrega
                    Select Case rstPedido!Tipoped
                        Case 0
                            Muestra.TextMatrix(Renglon, 6) = "Normal"
                        Case 1
                            Muestra.TextMatrix(Renglon, 6) = "A Fecha"
                        Case 2
                            Muestra.TextMatrix(Renglon, 6) = "Fecha Limite"
                        Case 3
                            Muestra.TextMatrix(Renglon, 6) = "Urgente"
                        Case 4
                            Muestra.TextMatrix(Renglon, 6) = "Retira Cliente"
                        Case 5
                            Muestra.TextMatrix(Renglon, 6) = "Muestra"
                        Case Else
                            Muestra.TextMatrix(Renglon, 6) = ""
                    End Select
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    ZZHasta = Renglon
    Call Carga_Atrasos
    
    Rem For dada = 1 To Renglon
    Rem     WCliente = Muestra.TextMatrix(dada, 3)
    Rem     spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Rem     Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    Rem     If rstCliente.RecordCount > 0 Then
    Rem         Muestra.TextMatrix(dada, 4) = rstCliente!Razon
    Rem         rstCliente.Close
    Rem     End If
    Rem Next dada
    
    Call Conecta_Empresa
    
    WTop = Renglon - 25
    If WTop < 1 Then
        WTop = 1
    End If
    
    WPosi1 = WTop
    WPosi2 = 1
    
    Muestra.TopRow = WPosi1
    Muestra.Row = WPosi2

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
    Muestra.Text = "Aviso"
    
End Sub

Private Sub Muestra_DblClick()

    WPosi1 = Muestra.TopRow
    WPosi2 = Muestra.Row
    WOrigenPosi = 0

    Muestra.Col = 1
    WXPed = Muestra.Text
    
    PrgVerificaPedidoConsulta.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    If WOrigenPosi <> 1 Then
        Call Proceso_Click
    End If
    WOrigenPosi = 0
    Muestra.TopRow = WPosi1
    Muestra.Row = WPosi2
End Sub

Private Sub WTitulo_DblClick(Index As Integer)
    
    XEmpresa = WEmpresa
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Else
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If

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
    Muestra.Text = "Aviso"
    
    Renglon = 0
    
    Sql1 = "Select DISTINCT Pedido.Pedido, Pedido.Fecha, Pedido.Cliente, Pedido.TipoPed, Pedido.FecEntrega,Pedido.OrdFecEntrega,Pedido.Fechaord, Cliente.Razon as [WRazon]"
    Sql2 = " FROM Pedido, CLiente"
    Sql3 = " Where Pedido.Cantidad - Pedido.Facturado > 0"
    Sql4 = " and Pedido.TipoPedido <> 1"
    Sql5 = " and Pedido.Autorizo <> " + "'" + "N" + "'"
    Sql6 = " and Pedido.Cliente = Cliente.Cliente"
    Select Case Index
        Case 1
            Sql7 = " Order by Pedido.Pedido"
        Case 2
            Sql7 = " Order by Pedido.FechaOrd, Pedido.Pedido"
        Case 3
            Sql7 = " Order by Pedido.Cliente, Pedido.Pedido"
        Case 4
            Sql7 = " Order by Pedido.CLiente, Pedido.Pedido"
        Case 5
            Sql7 = " Order by Pedido.OrdFecEntrega, Pedido.Pedido"
        Case Else
            Sql7 = " Order by Pedido.Pedido"
    End Select
    
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Muestra.TextMatrix(Renglon, 0) = Pusing("###", Str$(Renglon))
                    Muestra.TextMatrix(Renglon, 1) = Pusing("######", Str$(rstPedido!Pedido))
                    Muestra.TextMatrix(Renglon, 2) = rstPedido!Fecha
                    Muestra.TextMatrix(Renglon, 3) = rstPedido!Cliente
                    Muestra.TextMatrix(Renglon, 4) = rstPedido!WRazon
                    Muestra.TextMatrix(Renglon, 5) = rstPedido!FecEntrega
                    Select Case rstPedido!Tipoped
                        Case 0
                            Muestra.TextMatrix(Renglon, 6) = "Normal"
                        Case 1
                            Muestra.TextMatrix(Renglon, 6) = "A Fecha"
                        Case 2
                            Muestra.TextMatrix(Renglon, 6) = "Fecha Limite"
                        Case 3
                            Muestra.TextMatrix(Renglon, 6) = "Urgente"
                        Case 4
                            Muestra.TextMatrix(Renglon, 6) = "Retira Cliente"
                        Case 5
                            Muestra.TextMatrix(Renglon, 6) = "Muestra"
                        Case Else
                            Muestra.TextMatrix(Renglon, 6) = ""
                    End Select
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    ZZHasta = Renglon
    Call Carga_Atrasos
    
    Rem For dada = 1 To Renglon
    Rem     WCliente = Muestra.TextMatrix(dada, 3)
    Rem     spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Rem     Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    Rem     If rstCliente.RecordCount > 0 Then
    Rem         Muestra.TextMatrix(dada, 4) = rstCliente!Razon
    Rem         rstCliente.Close
    Rem     End If
    Rem Next dada
    
    
    
    Call Conecta_Empresa
    
    WTop = Renglon - 25
    If WTop < 1 Then
        WTop = 1
    End If
    
    WPosi1 = WTop
    WPosi2 = 1
    
    Muestra.TopRow = WPosi1
    Muestra.Row = WPosi2
    
End Sub


Private Sub Carga_Atrasos()

    For ZZCiclo = 1 To ZZHasta
    
        ZZPedido = Muestra.TextMatrix(ZZCiclo, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Atraso"
        ZSql = ZSql + " Where Atraso.Pedido = " + "'" + ZZPedido + "'"
        spAtraso = ZSql
        Set rstatraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        If rstatraso.RecordCount > 0 Then
            ZZImpre = Mid$(rstatraso!Fecha, 1, 5) + " " + Trim(rstatraso!problema) + " " + Trim(rstatraso!Articulo) + " " + ZImpreConcepto(rstatraso!Concepto)
            Muestra.TextMatrix(ZZCiclo, 7) = ZZImpre
            rstatraso.Close
        End If
    
    Next ZZCiclo
    

End Sub

