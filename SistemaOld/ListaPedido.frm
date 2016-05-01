VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaPedido 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Pedidos"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Pedidos a Retirar por los clientes vencidos"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedPen.rpt"
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
Attribute VB_Name = "PrgListaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPedido As Recordset
Dim spPedido As String

Private Sub Acepta_Click()

    Rem WEmpresa = "0009"
    Rem txtOdbc = "Empresa09"
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    WDesdeFec = "01/01/2005"
    WHastaFec = ""

    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia

    spPedido = "ModificaPedpen0"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spPedido = "ModificaPedpen " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem XParam = "'" + WDesde + "','" _
    Rem              + WHasta + "'"
    Rem spPedido = "ModificaPedidonousar " + XParam
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    Rem With rstPedido
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem             .Edit
    Rem             !Importe = 0
    Rem             If !Pedido >= Desde.Text And !Pedido <= Hasta.Text Then
    Rem                 If !FechaOrd >= WDesde And !FechaOrd <= WHasta Then
    Rem                     !Importe = !Cantidad - !Facturado
    Rem                 End If
    Rem             End If
    Rem             .Update
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem            End If
    Rem         Loop
    Rem End With
    
    Erase WVector
    LugarVector = 0
    
    spPedido = "ListaPedidoPend"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    EntraVector = "S"
                    For Ciclo = 1 To LugarVector
                        If WVector(Ciclo, 1) = rstPedido!Terminado Then
                            EntraVector = "N"
                            Exit For
                        End If
                    Next Ciclo
                    If EntraVector = "S" Then
                        LugarVector = LugarVector + 1
                        WVector(LugarVector, 1) = rstPedido!Terminado
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    For Ciclo = 1 To LugarVector
        WProducto = WVector(Ciclo, 1)
        WTipopro = Left$(WProducto, 2)
        Select Case WTipopro
            Case "DY", "DW"
                WArticulo = Left$(WProducto, 3) + Right$(WProducto, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                XParam = "'" + WArticulo + "','" _
                        + WDescripcion + "'"
                spPedido = "ModificaPedidoArticulo " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Case Else
                WTerminado = WProducto
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescripcion = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                
                XParam = "'" + WTerminado + "','" _
                        + WDescripcion + "'"
                spPedido = "ModificaPedidoTerminado " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        End Select
    Next Ciclo

    
    If Lugar > 0 Then
        PrgImpreped.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgImpreped.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Pedidos Pendientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.FechaOrd, Pedido.Facturado, Pedido.Importe, Pedido.Autorizo, Pedido.Tipoped, Pedido.Descripcion, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.Importe > 0 AND " _
                    + "Pedido.Autorizo <> 'N'"
    
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()

    Listado.Action = 1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgImpreped.Hide
    Unload Me
    End
End Sub


Private Sub Form_Load()
    Call Acepta_Click
End Sub

