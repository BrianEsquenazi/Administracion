VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaDesarrollo 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Estado de Desarrollo de Pedidos"
   ClientHeight    =   7320
   ClientLeft      =   90
   ClientTop       =   690
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11850
   Begin VB.ComboBox Estado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   10
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Vendedor 
      Alignment       =   1  'Right Justify
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8705
      _Version        =   327680
      Rows            =   4000
      Cols            =   7
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin MSMask.MaskEdBox HastaFecha 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DesdeFecha 
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label DesVendedor 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde Fecha"
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
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "ConsultaDesarrollo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaDesarrollo.frx":030A
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "PrgConsultaDesarrollo"
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
Dim ZVector(100, 8) As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdClose_Click()
    PrgConsultaDesarrollo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Estado_click()
    Call Proceso_Click
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 800
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 2000
    Muestra.ColWidth(4) = 1500
    Muestra.ColWidth(5) = 3500
    Muestra.ColWidth(6) = 2000
    
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
    
    Muestra.Col = 6
    Muestra.Text = "Estado"
    
    Estado.Clear
    
    Estado.AddItem "Todos"
    Estado.AddItem "Pendientes de Aprobar"
    Estado.AddItem "Aprobados"
    Estado.AddItem "Rechazados"
    Estado.AddItem "Aprobados Laboratorio"
    Estado.AddItem "Aprobados Laboratorio Pendientes"
    Estado.AddItem "Aprobados Laboratorio Finalizado"
    Estado.AddItem "Aprobados Desarrollo"
    
    Estado.ListIndex = 0
    
    Rem Call Proceso_Click
    
End Sub


Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

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
    
    Muestra.Col = 6
    Muestra.Text = "Estado"
    
    Renglon = 0
    
    If DesdeFecha.Text <> "  /  /    " And HastaFecha.Text <> "  /  /    " Then
        WAno = Right$(DesdeFecha.Text, 4)
        WMes = Mid$(DesdeFecha.Text, 4, 2)
        WDia = Left$(DesdeFecha.Text, 2)
        WDesde = WAno + WMes + WDia
        WAno = Right$(HastaFecha.Text, 4)
        WMes = Mid$(HastaFecha.Text, 4, 2)
        WDia = Left$(HastaFecha.Text, 2)
        WHasta = WAno + WMes + WDia
            Else
        WDesde = "00000000"
        WHasta = "99999999"
    End If
    
    Select Case Estado.ListIndex
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " Order by Pedido"

        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Respuesta = 0"
            ZSql = ZSql + " Order by Pedido"
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Respuesta = 1"
            ZSql = ZSql + " Order by Pedido"
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Respuesta = 2"
            ZSql = ZSql + " Order by Pedido"
            
        Case 4
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Respuesta = 1"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Destino = 2"
            ZSql = ZSql + " Order by Pedido"
            
        Case 5
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Respuesta = 1"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Destino = 2"
            ZSql = ZSql + " and PedidoOrdenTRabajo.EstadoLabora <> 'N'"
            ZSql = ZSql + " Order by Pedido"
            
        Case 6
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Respuesta = 1"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Destino = 2"
            ZSql = ZSql + " and PedidoOrdenTRabajo.EstadoLabora = 'N'"
            ZSql = ZSql + " Order by Pedido"
            
        Case 7
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM PedidoOrdenTRabajo"
            ZSql = ZSql + " Where PedidoOrdenTRabajo.OrdFecha >= " + "'" + WDesde + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.OrdFecha <= " + "'" + WHasta + "'"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Respuesta = 1"
            ZSql = ZSql + " and PedidoOrdenTRabajo.Destino = 1"
            ZSql = ZSql + " Order by Pedido"

        Case Else
    End Select
            
    spPedidoOrdenTrabajo = ZSql
    Set rstPedidoordentrabajo = db.OpenRecordset(spPedidoOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoordentrabajo.RecordCount > 0 Then
    
        With rstPedidoordentrabajo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Val(Vendedor.Text) = 0 Or Val(Vendedor.Text) = !Vendedor Then
                
                        Renglon = Renglon + 1
            
                        Muestra.TextMatrix(Renglon, 1) = Pusing("######", Str$(rstPedidoordentrabajo!Pedido))
                        Muestra.TextMatrix(Renglon, 2) = rstPedidoordentrabajo!Fecha
                        Muestra.TextMatrix(Renglon, 5) = rstPedidoordentrabajo!Observaciones
                        
                        ZEstado = ""
                        
                        Select Case rstPedidoordentrabajo!Respuesta
                            Case 0
                                ZEstado = "Pendiente de Aprobar"
                                
                            Case 1
                                ZEstado = "Aprobada"
                                If rstPedidoordentrabajo!Destino = 1 Then
                                    ZEstado = "Desarrollo"
                                End If
                                If rstPedidoordentrabajo!Destino = 2 Then
                                    If rstPedidoordentrabajo!estadolabora = "N" Then
                                        ZEstado = "Laboratorio Finalizado"
                                            Else
                                        ZEstado = "Laboratorio Pendiente"
                                    End If
                                End If
                                
                            Case 2
                                ZEstado = ZEstado + "Rechazada"
                                
                            Case Else
                        End Select
                        
                        Muestra.TextMatrix(Renglon, 6) = ZEstado
                    
                        ZVector(Renglon, 1) = rstPedidoordentrabajo!Cliente
                        ZVector(Renglon, 2) = rstPedidoordentrabajo!Vendedor
                        
                    End If
                        
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
    
    Call Conecta_Empresa
    
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
    
    Muestra.Col = 6
    Muestra.Text = "Estado"
    
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 1
    WXPed = Muestra.Text
    PrgConsultaDesarrollo.Hide
    Rem Unload Me
    PrgpedidoOrdenTrabajoConsulta.Show
    
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Vendedor.Text) <> 0 Then
        
            XEmpresa = WEmpresa
    
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            spVendedor = "ConsultaVendedor " + "'" + Vendedor.Text + "'"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                Vendedor.Text = rstVendedor!Vendedor
                DesVendedor.Caption = rstVendedor!Nombre
                rstVendedor.Close
                Call Conecta_Empresa
                Call Proceso_Click
                    Else
                Call Conecta_Empresa
                Vendedor.SetFocus
            End If
                Else
            Vendedor.Text = ""
            DesVendedor.Caption = ""
            Call Proceso_Click
        End If
    End If
    If KeyAscii = 27 Then
        Vendedor.Text = ""
        DesVendedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
                Else
            DesdeFecha.SetFocus
        End If
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Call Proceso_Click
            DesdeFecha.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
End Sub



