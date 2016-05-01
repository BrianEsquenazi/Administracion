VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAvisoRechaProduV 
   AutoRedraw      =   -1  'True
   Caption         =   "Aviso de Entrada de Productos para Devolucion"
   ClientHeight    =   4605
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4605
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3975
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Aviso de Productos a Bloquear por   Devolucion"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   5535
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PedpenII.rpt"
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
End
Attribute VB_Name = "PrgAvisoRechaProduV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPedidoDevol As Recordset
Dim spPedidoDevol As String

Dim XParam As String
Dim WVector(1000) As String
Dim LeeAviso(100, 3) As String

Dim LugarVector As Integer

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase WVector
    LugarVector = 0
    
    ZMarca = "N"
    ZBloqueo = "S"
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PedidoDevol"
    ZSql = ZSql + " Where PedidoDevol.ImpreProdV = " + "'" + ZMarca + "'"
    ZSql = ZSql + " and PedidoDevol.Bloqueo = " + "'" + ZBloqueo + "'"
    ZSql = ZSql + " and PedidoDevol.Renglon = 1"
    ZSql = ZSql + " Order by Clave"
    spPedidoDevol = ZSql
    Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedidoDevol.RecordCount > 0 Then
        With rstPedidoDevol
        .MoveFirst
            Do
                If .EOF = False Then
                    If rstPedidoDevol!TipoProII = "BI" Or rstPedidoDevol!TipoProII = "PT" Then
                        LugarVector = LugarVector + 1
                        WVector(LugarVector) = rstPedidoDevol!Pedido
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedidoDevol.Close
    End If
    
    If LugarVector > 0 Then
        PrgAvisoRechaProduV.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgAvisoRechaProduV.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    Listado.WindowTitle = "Impresion de Entrada de Devolucion de Mercaderia"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    For ZCiclo = 1 To LugarVector
    
        WCodigo = WVector(ZCiclo)
        
        Erase LeeAviso
        Renglon = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PedidoDevol"
        ZSql = ZSql + " Where PedidoDevol.Pedido = " + "'" + WCodigo + "'"
        ZSql = ZSql + " Order by Clave"
        spPedidoDevol = ZSql
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedidoDevol.RecordCount > 0 Then
            With rstPedidoDevol
                .MoveFirst
                Do
                    If .EOF = False Then
                        Renglon = rstPedidoDevol!Renglon
                        LeeAviso(Renglon, 1) = rstPedidoDevol!Clave
                        LeeAviso(Renglon, 2) = rstPedidoDevol!Terminado
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedidoDevol.Close
        End If
    
        For Da = 1 To Renglon
        
            WClave = LeeAviso(Da, 1)
            Auxi1 = LeeAviso(Da, 2)
            WImpreTerminado = ""
            
            If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DK" Or Left$(Auxi1, 2) = "DW" Or Left$(Auxi1, 2) = "NW" Then
            
                If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DK" Then
                    WArti = "DY-" + Right$(Auxi1, 7)
                        Else
                    WArti = "DW-" + Right$(Auxi1, 7)
                End If
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WImpreTerminado = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                        Else
                Auxi1 = "PT-" + Mid$(LeeAviso(Da, 2), 4, 9)
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WImpreTerminado = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
            End If
            
            Sql1 = "UPDATE PedidoDevol SET "
            Sql2 = "ImpreTerminado  = " + "'" + WImpreTerminado + "'"
            Sql3 = " Where Clave = " + "'" + WClave + "'"
            spPedidoDevol = Sql1 + Sql2 + Sql3
            Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Da
        
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.ReportFileName = "AvisoRechaProdu.rpt"
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.GroupSelectionFormula = "{PedidoDevol.Pedido} in " + WCodigo + " to " + WCodigo
        Listado.SQLQuery = "SELECT PedidoDevol.Pedido, PedidoDevol.Cliente, PedidoDevol.Fecha, PedidoDevol.Observaciones, PedidoDevol.Terminado, PedidoDevol.Cantidad, PedidoDevol.ImpreTerminado, PedidoDevol.Partida, " _
                        + "Cliente.Razon " _
                        + "From " _
                        + DSQ + ".dbo.PedidoDevol PedidoDevol, " _
                        + DSQ + ".dbo.Cliente Cliente " _
                        + "Where " _
                        + "PedidoDevol.Cliente = Cliente.Cliente AND " _
                        + "PedidoDevol.Pedido >= " + WCodigo + " AND " _
                        + "PedidoDevol.Pedido <= " + WCodigo
    
        Listado.Connect = Connect()
        Listado.Action = 1
            
        WMarcaImpresion = "S"
        Sql1 = "UPDATE PedidoDevol SET "
        Sql2 = "ImpreProdV  = " + "'" + WMarcaImpresion + "'"
        Sql3 = " Where Pedido = " + "'" + WCodigo + "'"
        spPedidoDevol = Sql1 + Sql2 + Sql3
        Set rstPedidoDevol = db.OpenRecordset(spPedidoDevol, dbOpenSnapshot, dbSQLPassThrough)
        
    Next ZCiclo
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgAvisoRechaProduV.Hide
    Unload Me
    Close
    End
End Sub




