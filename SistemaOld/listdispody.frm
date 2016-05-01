VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListDispoDy 
   Caption         =   "Listado de Disponibilidad  de Stock de DY"
   ClientHeight    =   6600
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6600
   ScaleWidth      =   8145
   Begin VB.CommandButton AvisoError 
      Caption         =   "No se puede emitir el informe. Sistema sin Conexion con las otras plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2640
      Picture         =   "listdispody.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   0
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   6015
      Begin VB.ComboBox Tipo 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1920
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
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
         Left            =   2040
         TabIndex        =   9
         Top             =   2040
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
         Left            =   600
         TabIndex        =   8
         Top             =   2040
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
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   600
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
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Listado"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
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
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
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
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "listdispody.frx":0742
      Left            =   0
      List            =   "listdispody.frx":0749
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListDispoDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vector(10000, 6) As String
Dim WFechaCierre As String
Dim WOrdFechaCierre As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private XLote(100, 7) As String

Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim CargaEmpresa(12, 2) As String

Private Sub Acepta_Click()

    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0006"
    CargaEmpresa(2, 2) = "Empresa06"
    CargaEmpresa(3, 1) = "0007"
    CargaEmpresa(3, 2) = "Empresa07"
    CargaEmpresa(4, 1) = "0008"
    CargaEmpresa(4, 2) = "Empresa08"
                    
    For Cicla = 1 To 4
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    If WSalidaError = "N" Then Exit Sub

    On Error GoTo WError
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    With rstStockDy
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    XEmpresa = WEmpresa
    
    
    
    Rem
    Rem calcula los pedidos pendientess
    Rem
    
    spArticulo = "ModificaArticuloVenta0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    spPedido = "ModificaPedpen0"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spPedido = "ModificaPedpenDy " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spPedido = "ListaPedidoPendDesdeHasta " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    EntraVector = "S"
                    For Ciclo = 1 To Renglon
                        If Vector(Ciclo, 1) = rstPedido!Terminado Then
                            Vector(Ciclo, 2) = Str$(Val(Vector(Ciclo, 2)) + rstPedido!Importe)
                            EntraVector = "N"
                            Exit For
                        End If
                    Next Ciclo
                    If EntraVector = "S" Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = rstPedido!Terminado
                        Vector(Renglon, 2) = Str$(rstPedido!Importe)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    For Ciclo = 1 To Renglon
        WProducto = Vector(Ciclo, 1)
        WImporte = Vector(Ciclo, 2)
        WArticulo = Left$(WProducto, 3) + Right$(WProducto, 7)
        XParam = "'" + WArticulo + "','" _
                     + WImporte + "','" _
                     + WDate + "'"
        spArticulo = "ModificaArticuloVenta " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
    
    
    
    
    
    Rem
    Rem recalcula saldo de orden de compra
    Rem
    
    For CicloEmpresa = 1 To 4
    
        Select Case CicloEmpresa
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        Erase Vector
        Renglon = 0
        
    
        XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
        spArticulo = "ModificaArticuloPedido0DesdeHasta" + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            
            
            
        ZSql = ""
        ZSql = ZSql + "Select Orden.Clave, Orden.Orden, Orden.Articulo, Orden.Cantidad, Orden.FechaOrd, Orden.Recibida"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Articulo >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Orden.Articulo <= " + "'" + Hasta.Text + "'"
        ZSql = ZSql + " Order by Orden.Orden"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        If !FechaOrd > "20020101" Then
                            ZSaldo = rstOrden!Cantidad - rstOrden!Recibida
                            If (rstOrden!Orden < 900000) Or (rstOrden!Orden > 900000 And ZSaldo <> 0) Then
                                Renglon = Renglon + 1
                                Vector(Renglon, 1) = rstOrden!Clave
                                Vector(Renglon, 2) = rstOrden!Orden
                                Vector(Renglon, 3) = rstOrden!Articulo
                                Vector(Renglon, 4) = rstOrden!Cantidad
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
    
    
        For Ciclo = 1 To Renglon
    
            WClave = Vector(Ciclo, 1)
            WOrden = Vector(Ciclo, 2)
            WArticulo = Vector(Ciclo, 3)
            WCantidad = Val(Vector(Ciclo, 4))
            WResta = 0
            
            XParam = "'" + WOrden + "','" _
                         + WArticulo + "'"
            spInforme = "ListaInformeOrdenArticulo" + XParam
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                With rstInforme
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                            WResta = WResta + rstInforme!Resta
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                        Loop
                    End If
                End With
                rstInforme.Close
            End If
            
            WDife = WCantidad - WResta
            If WDife > 0 Then
                XParam = "'" + WArticulo + "','" _
                             + Str$(WDife) + "'"
                spArticulo = "ModificaArticuloPedido " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
        Next Ciclo
    
    
    Next CicloEmpresa
    
    


    Rem
    Rem recalcula saldo del stock
    Rem

    
    For CicloEmpresa = 1 To 4
            
        Select Case CicloEmpresa
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        
        Erase Vector
        Renglon = 0
        
    
        XParam = "'" + Desde.Text + "','" _
                     + Hasta.Text + "'"
        spArticulo = "ListaArticuloDesdeHasta " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        With rstArticulo
                .MoveFirst
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstArticulo!Codigo
                    Vector(Renglon, 2) = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
                    Vector(Renglon, 3) = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
        End With
        rstArticulo.Close
    
        For Da = 1 To Renglon
    
            WEntradas = 0
            WSalidas = 0
            WArticulo = Vector(Da, 1)
            XCodigo = Vector(Da, 1)
            WFechaCierre = Vector(Da, 2)
            WOrdFechaCierre = Vector(Da, 3)
            XDate = Date$
        
            Call calcula_datos
        
            XEntradas = Str$(WEntradas)
            XSalidas = Str$(WSalidas)
        
            XParam = "'" + XCodigo + "','" _
                         + XEntradas + "','" _
                         + XSalidas + "','" _
                         + XDate + "'"
                                           
            spArticulo = "ModificaArticuloMovimientos " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Next Da
    
    Next CicloEmpresa
    
    
    
    
    
    Rem
    Rem reproceso de calculo de datos para el listado
    Rem
    
    For CicloEmpresa = 1 To 4
            
        Select Case CicloEmpresa
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        XParam = "'" + Desde.Text + "','" _
                     + Hasta.Text + "'"
    
        spArticulo = "ListaArticuloDesdeHasta" + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            
            With rstArticulo
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        WArticulo = rstArticulo!Codigo
                        WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                        WDescripcion = rstArticulo!Descripcion
                        WLaboratorio = rstArticulo!Laboratorio
                        WOrden = rstArticulo!Pedido
                        WPedido = rstArticulo!Venta
                    
                        With rstStockDy
                            .Index = "Codigo"
                            .Seek "=", WArticulo
                            If .NoMatch = True Then
                                .AddNew
                                !Codigo = WArticulo
                                !Descripcion = WDescripcion
                                !Stock1 = 0
                                !Stock2 = 0
                                !Stock3 = 0
                                !Stock4 = 0
                                !Stock5 = 0
                                !Stock1 = !Stock1 + WStock
                                !Titulo1 = "Surfactan S.A."
                                !Titulo2 = ""
                                !Laboratorio = WLaboratorio
                                !Orden = WOrden
                                !Pedido = WPedido
                                !Familia = Mid$(WArticulo, 4, 3)
                                .Update
                                    Else
                                .Edit
                                !Stock1 = !Stock1 + WStock
                                !Laboratorio = !Laboratorio + WLaboratorio
                                !Orden = !Orden + WOrden
                                !Pedido = !Pedido + WPedido
                                .Update
                            End If
                        End With
                    
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            
            rstArticulo.Close
        End If
        
    Next CicloEmpresa
    
    
    Call Conecta_Empresa
        
    
    If Tipo.ListIndex = 0 Then
    
        With rstStockDy
            .Index = "Codigo"
            .Seek ">=", ""
            If .NoMatch = False Then
                Do
                    Suma = !Stock1 + !Stock2 + !Stock3 + !Stock4 + !Stock5
                    Orden = !Orden
                    Laboratorio = !Laboratorio
                    If Suma = 0 And Orden = 0 And Laboratorio = 0 Then
                        .Delete
                    End If
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        
    End If

    Listado.WindowTitle = "Listado de Disponibilidad de Stock de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "WListDispoDy.rpt"
    
    Listado.Action = 1
    
    Exit Sub

WError:
    Resume Next
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
    
End Sub

Private Sub AvisoError_Click()
    AvisoError.Visible = False
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstStockDy
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgListDispoDy.Hide
    Unload Me
    Menu.Show
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
    OPEN_FILE_StockDy
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListDispoDy.Caption = "Listado de Disponibilidad  de Stock de Materias Primas :  " + !Nombre
        End If
    End With
    
    Tipo.Clear
    
    Tipo.AddItem "C/Stock"
    Tipo.AddItem "Todos los Articulos"
    
    Tipo.ListIndex = 0
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstArticulo
            
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                    Desde.Text = rstArticulo!Codigo
                    Hasta.Text = rstArticulo!Codigo
                End If
            End With
            Desde.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(rstArticulo!Descripcion) - WEspacios
                For Aaa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next Aaa
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    rstArticulo.Close
    
    End If
    
    End If

End Sub

Private Sub calcula_datos()


    Rem PROCESA LOS LAUDOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoRepro" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                If rstLaudo!Marca = "X" And WSaldo = 0 Then
                        Else
                    If rstLaudo!Articulo = WArticulo Then
                        WEntradas = WEntradas + rstLaudo!Liberada
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spHoja = "ListaHojaRepro" + XParam
    Rem spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                        Else
                    If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                        XX = rstHoja!Clave
                        WSalidas = WSalidas + rstHoja!Cantidad
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                If rstHoja!Articulo > WArticulo Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstHoja.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovvar = "ListaMovvarRepro1" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        With rstMovvar
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovvar!Marca = "X" Then
                        Else
                    If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                        If rstMovvar!Movi = "E" Then
                            WEntradas = WEntradas + rstMovvar!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovvar!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstMovvar.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaRepro1" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                        Else
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                        If rstMovguia!Movi = "E" Then
                            WEntradas = WEntradas + rstMovguia!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovguia!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spMovlab = "ListaMovlabRepro1" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
        With rstMovlab
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovlab!Marca = "X" Then
                        Else
                    If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                        If rstMovlab!Movi = "E" Then
                            WEntradas = WEntradas + rstMovlab!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovlab!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstMovlab.Close
    End If
    
    
    
    Rem PROCESA LAS VENTAS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spEstadistica = "ListaEstadisticaReproDy " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstEstadistica!Marca = "X" Then
                
                        Else
                        
                    WTipo = rstEstadistica!Tipo
                    
                    XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                    XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                    XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                    XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                    XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                    XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                    XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                        
                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                     
                    If Len(Trim(WLoteAdicional)) = 98 Then
                         XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                         XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                         XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                         XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                         XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                         XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                         XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                         XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                         XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                         XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                         XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                         XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                         XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                         XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                             Else
                         XLote(6, 1) = "0"
                         XLote(6, 2) = "0"
                         XLote(7, 1) = "0"
                         XLote(7, 2) = "0"
                         XLote(8, 1) = "0"
                         XLote(8, 2) = "0"
                         XLote(9, 1) = "0"
                         XLote(9, 2) = "0"
                         XLote(10, 1) = "0"
                         XLote(10, 2) = "0"
                         XLote(11, 1) = "0"
                         XLote(11, 2) = "0"
                         XLote(12, 1) = "0"
                         XLote(12, 2) = "0"
                    End If
                        
                    For Da = 1 To 12
                        WLote = XLote(Da, 1)
                        WCantidad = XLote(Da, 2)
                        If Val(WCantidad) <> 0 Then
                            If WTipo = 2 Then
                                WEntradas = WEntradas + Abs(Val(WCantidad))
                                    Else
                                WSalidas = WSalidas + WCantidad
                            End If
                        End If
                    Next Da
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstEstadistica.Close
    End If
    
End Sub



