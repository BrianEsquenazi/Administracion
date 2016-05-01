VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedPen 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Pedidos Pendientes"
   ClientHeight    =   4305
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8220
   LinkTopic       =   "Form2"
   ScaleHeight     =   4305
   ScaleWidth      =   8220
   Begin VB.TextBox vendedor 
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
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   14
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox hcliente 
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
      Left            =   3840
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox dcliente 
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
      Left            =   2520
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   5775
      Begin VB.OptionButton disco 
         Caption         =   "Disco"
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
         Left            =   3480
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Left            =   1800
         TabIndex        =   6
         Top             =   2880
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
         Left            =   240
         TabIndex        =   5
         Top             =   2880
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
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   975
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
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1680
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   3120
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Caption         =   "Producto"
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente "
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedpen.rpt"
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
      Left            =   6960
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgPedPen"
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
Dim ZZVector(10000, 10) As String


Private Sub Acepta_Click()
      Desde.Text = UCase(Desde.Text)
      Hasta.Text = UCase(Hasta.Text)
      
    Tipol = 0
    
    If dcliente = "" Or hcliente = "" Then
        Tipol = 1
        If vendedor <> "" Then
            Tipol = 3
        End If
    End If
    
    If dcliente <> "" And hcliente <> "" Then
        Tipol = 2
    End If
    
    If (Left$(Desde.Text, 2) = "PT" Or Left$(Hasta.Text, 2) = "DY") And dcliente = "" Then
        Tipol = 4
    End If
    
    If (Left$(Desde.Text, 2) = "PT" Or Left$(Hasta.Text, 2) = "DY") And dcliente <> "" Then
        Tipol = 5
    End If
    
    
    
    
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    Erase ZZVector
    Renglon = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Pedido"
    Sql3 = " Where Pedido.Cantidad > Pedido.Facturado"
    Sql4 = " and (Pedido.TipoPed = 5 or Pedido.TipoPed = 6)"
    Sql5 = " Order by Pedido.fechaord"
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    
                    WVector(Renglon, 1) = rstPedido!Pedido
                    WVector(Renglon, 2) = rstPedido!Terminado
                    WVector(Renglon, 3) = rstPedido!Cantidad
                    WVector(Renglon, 4) = rstPedido!Clave
                    WVector(Renglon, 5) = rstPedido!Fecha
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    
    For Ciclo = 1 To Renglon
    
    
        ZZPedido = WVector(Ciclo, 1)
        ZZTerminado = WVector(Ciclo, 2)
        ZZCantidad = WVector(Ciclo, 3)
        ZZClavePedido = WVector(Ciclo, 4)
        ZZFecha = WVector(Ciclo, 5)
    
        Select Case Left$(ZZTerminado, 2)
            Case "PT", "PE", "YQ", "YF", "YP", "YH"
                ZTerminado = ZZTerminado
                ZArticulo = ""
                Proceso = 1
            Case Else
                ZTerminado = ""
                ZArticulo = Left$(ZZTerminado, 3) + Right$(ZZTerminado, 7)
                Proceso = 2
        End Select
        
        ZZRemito = 0
        If Proceso = 1 Then
    
            Sql1 = "Select *"
            Sql2 = " FROM Muestra"
            Sql3 = " Where Muestra.Pedido = " + "'" + ZZPedido + "'"
            Sql4 = " and Muestra.Producto = " + "'" + ZTerminado + "'"
            spMuestra = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                ZZRemito = IIf(IsNull(rstMuestra!Remito), "", rstMuestra!Remito)
                rstMuestra.Close
            End If
            
                Else
        
            Sql1 = "Select *"
            Sql2 = " FROM Muestra"
            Sql3 = " Where Muestra.Pedido = " + "'" + ZZPedido + "'"
            Sql4 = " and Muestra.Articulo = " + "'" + ZArticulo + "'"
            spMuestra = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMuestra = db.OpenRecordset(spMuestra, dbOpenSnapshot, dbSQLPassThrough)
            If rstMuestra.RecordCount > 0 Then
                ZZRemito = IIf(IsNull(rstMuestra!Remito), "", rstMuestra!Remito)
                rstMuestra.Close
            End If
    
        End If
        
        If Val(ZZRemito) <> 0 Then
            ZSql = ""
            ZSql = ZSql & "UPDATE Pedido SET "
            ZSql = ZSql & "Facturado = Cantidad"
            ZSql = ZSql & " Where Clave = " + "'" + ZZClavePedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
    Next Ciclo





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
        If Left$(WProducto, 2) <> "PT" Then
        
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
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Pedidos Pendientes"
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
        If disco.Value = True Then
            Listado.Destination = 0
            Listado.ReportFileName = "wpedpenddisco.rpt"
                Else
            Listado.Destination = 0
                Listado.ReportFileName = "Wpedpen.rpt"
        End If
    End If
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    
    If Tipol = 1 Then
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
    End If

    If Tipol = 2 Then
        Rem by nan agregado busca por cliente
        Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.FechaOrd, Pedido.Facturado, Pedido.Importe, Pedido.Autorizo, Pedido.Tipoped, Pedido.Descripcion, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.cliente >= '" + dcliente.Text + "' AND " _
                    + "Pedido.cliente <= '" + hcliente.Text + "' AND " _
                    + "Pedido.Importe > 0 AND " _
                    + "Pedido.Autorizo <> 'N'"
  
        Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
    
    End If
    Rem *******BY NAN POR VENDEDOR*************
    
    
    If Tipol = 3 Then
          Listado.ReportFileName = "wpedpen2.rpt"
        Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.FechaOrd, Pedido.Facturado, Pedido.Importe, Pedido.Autorizo, Pedido.Tipoped, Pedido.Descripcion, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "cliente.vendedor = '" + vendedor.Text + "' AND " _
                    + "Pedido.Importe > 0 AND " _
                    + "Pedido.Autorizo <> 'N'"
        Rem + "Pedido.cliente >= '" + dcliente.Text + "' AND " _
        rem                + "Pedido.cliente <= '" + hcliente.Text + "' AND " _

        Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
    
        Rem BY NAN FIN
    End If
     
       If Tipol = 4 Then
             Listado.ReportFileName = "wpedpen2.rpt"
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.FechaOrd, Pedido.Facturado, Pedido.Importe, Pedido.Autorizo, Pedido.Tipoped, Pedido.Descripcion, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.Terminado >= '" + Desde.Text + "' AND " _
                    + "Pedido.Terminado <= '" + Hasta.Text + "' AND " _
                    + "Pedido.Importe > 0 AND " _
                    + "Pedido.Autorizo <> 'N'"
            
            
    
            
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
    
    
      End If
    If Tipol = 5 Then
             Listado.ReportFileName = "wpedpen2.rpt"
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.FechaOrd, Pedido.Facturado, Pedido.Importe, Pedido.Autorizo, Pedido.Tipoped, Pedido.Descripcion, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.cliente >= '" + dcliente.Text + "' AND " _
                    + "Pedido.cliente <= '" + hcliente.Text + "' AND " _
                    + "Pedido.Terminado >= '" + Desde.Text + "' AND " _
                    + "Pedido.Terminado <= '" + Hasta.Text + "' AND " _
                    + "Pedido.Importe > 0 AND " _
                    + "Pedido.Autorizo <> 'N'"
            
            
    
            
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
    
    
      End If
    
    
    
    
    
    
    
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    DesdeFec.SetFocus
    PrgPedPen.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub dcliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   dcliente.Text = UCase(dcliente)
   hcliente.Text = dcliente
   hcliente.SetFocus
End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub DesdeFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
End Sub


Private Sub Hasta_Keypress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        dcliente.SetFocus
    End If
End Sub

Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            DesdeFec.SetFocus
                Else
            HastaFec.SetFocus
        End If
    Desde.SetFocus
    End If

End Sub

Sub Form_Load()
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Panta.Value = True
    Impresora.Value = False
    Frame2.Visible = True
End Sub

Private Sub hcliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
hcliente.Text = UCase(hcliente)
vendedor.SetFocus
End If

End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then



End If
End Sub
