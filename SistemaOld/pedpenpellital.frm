VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgPedPenPellital 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Pedidos Pendientes de Fazon de Pellital"
   ClientHeight    =   3165
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedpenPellital.rpt"
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
      Left            =   6480
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgPedPenPellital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Uno As String
Private Dos As String
Private Tres As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim WVector(1000, 5) As String
Dim LugarVector As Integer
Dim WTipopro As String

Private Sub Acepta_Click()

    XEmpresa = WEmpresa
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    WDesde = "00000000"
    WHasta = "99999999"

    spPedido = "ModificaPedpen0"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spPedido = "ModificaPedpen " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase WVector
    LugarVector = 0
    
    spPedido = "ListaPedidoPend"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstPedido!Terminado = "P00005" Then
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
        If Left$(WProducto, 2) <> "PT" And Left$(WProducto, 2) <> "PE" Then
        
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
            
                Else
            
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
        
        End If
    Next Ciclo
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Pedido SET "
    ZSql = ZSql + " Descripcion = NombreComercial"
    ZSql = ZSql + " Where Terminado >= " + "'" + "ML-00000-000" + "'"
    ZSql = ZSql + " and Terminado <= " + "'" + "ML-99999-999" + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = ""
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Pedidos Pendientes de Fazon de Pellital"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Uno = "{Pedido.FechaOrd} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Rem Dos = " and {Pedido.Pedido} in " + Desde.Text + " to " + Hasta.Text
    Rem Tres = " and {Pedido.Importe} > 0 "
    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres
   
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
                    + "Pedido.Cliente = 'P00005' AND " _
                    + "Pedido.Importe > 0 AND " _
                    + "Pedido.Autorizo <> 'N'"
    
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()

    Listado.Action = 1
    
    Call Conecta_Empresa
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    PrgPedPenPellital.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Sub Form_Load()
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

