VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAutorizaDesarrollo 
   AutoRedraw      =   -1  'True
   Caption         =   "Autorizacion de Desarrollos "
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11033
      _Version        =   327680
      Rows            =   4000
      Cols            =   6
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "autorizadesarrollo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "autorizadesarrollo.frx":030A
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "PrgAutorizaDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstOrdenPedidoTRabajo As Recordset
Dim spOrdenPedidoTRabajo As String
Dim XParam As String
Dim TotalPedidos As Integer
Dim WGraba As String
Dim ZVector(100, 4) As String

Private Sub cmdClose_Click()
    PrgAutorizaDesarrollo.Hide
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
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 2200
    Muestra.ColWidth(4) = 2200
    Muestra.ColWidth(5) = 4000
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 4
    Muestra.Text = "Vendedor"
    
    Muestra.Col = 5
    Muestra.Text = "Desarrollo"
    
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
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 4
    Muestra.Text = "Vendedor"
    
    Muestra.Col = 5
    Muestra.Text = "Desarrollo"
    
    
    Renglon = 0
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM PedidoOrdenTRabajo"
    Sql3 = " Where PedidoOrdenTRabajo.Respuesta = 0"
    Sql4 = " Order by Pedido"
    spPedidoOrdenTrabajo = Sql1 + Sql2 + Sql3 + Sql4
    Set rstPedidoordentrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoordentrabajo.RecordCount > 0 Then
    
        With rstPedidoordentrabajo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Muestra.TextMatrix(Renglon, 1) = Pusing("######", Str$(rstPedidoordentrabajo!Pedido))
                    Muestra.TextMatrix(Renglon, 2) = rstPedidoordentrabajo!Fecha
                    Muestra.TextMatrix(Renglon, 5) = rstPedidoordentrabajo!Observaciones
                    
                    ZVector(Renglon, 1) = rstPedidoordentrabajo!Cliente
                    ZVector(Renglon, 2) = rstPedidoordentrabajo!Vendedor
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedidoordentrabajo.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        spCliente = "ConsultaCliente " + "'" + ZVector(Ciclo, 1) + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra.TextMatrix(Ciclo, 3) = rstCliente!Razon
            rstCliente.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Vendedor"
        ZSql = ZSql + " Where Vendedor.Vendedor = " + "'" + ZVector(Ciclo, 2) + "'"
        spVendedor = ZSql
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            Muestra.TextMatrix(Ciclo, 4) = rstVendedor!Nombre
            rstVendedor.Close
        End If
    
    Next Ciclo
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    
    Muestra.TopRow = 1
    
End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Pedido"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 4
    Muestra.Text = "Vendedor"
    
    Muestra.Col = 5
    Muestra.Text = "Desarrollo"
    
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 1
    WXPed = Muestra.Text
    PrgAutorizaDesarrollo.Hide
    Unload Me
    PrgpedidoOrdenTrabajoAutoriza.Show
    
End Sub

