VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEstaVen 
   AutoRedraw      =   -1  'True
   Caption         =   "1.-Listado de Estadisitica de Ventas por Vendedor, Rubro y Linea"
   ClientHeight    =   5805
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5805
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox TipoCosto 
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
         Left            =   1680
         TabIndex        =   18
         Top             =   2280
         Width           =   2055
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   1680
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
         TabIndex        =   15
         Top             =   1200
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
      Begin VB.TextBox Hasta 
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
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Desde 
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
         MaxLength       =   4
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   975
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
         Left            =   2520
         TabIndex        =   11
         Top             =   2760
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
         TabIndex        =   10
         Top             =   2760
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
         Left            =   3480
         TabIndex        =   9
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
         Left            =   3480
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Costo"
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
         TabIndex        =   17
         Top             =   2280
         Width           =   1575
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
         TabIndex        =   14
         Top             =   1680
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
         TabIndex        =   13
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Vendedor"
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
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Vendedor"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEstaVen.rpt"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "EstaVen.frx":0000
      Left            =   840
      List            =   "EstaVen.frx":0007
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgEstaVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Costo As Double
Private Producto As String
Private Auxiliar(100, 7) As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstLinea As Recordset
Dim spLinea As String
Dim rstRubro As Recordset
Dim spRubro As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Private Vecosto(5000, 2) As String
Dim Posi As Integer
Dim WDescuento As Double

Dim ZDescriLinea(10000) As String
Dim ZDescriRubro(10000) As String
Dim ZDescriVendedor(10000) As String

Private Sub Acepta_Click()

    On Error GoTo WError

    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    Whasta = WAno + WMes + WDia
    
    Select Case TipoCosto.ListIndex
        Case 0
            WTitulo = "del " + DesdeFec.Text + " al " + HastaFec.Text + " (Costo actual)"
        Case 1
            WTitulo = "del " + DesdeFec.Text + " al " + HastaFec.Text + " (Costo F.Fact.)"
    End Select
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            Nomempresa = !Nombre
        End If
    End With
    
    With rstEsta
        .Index = "Clave"
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
    
    Erase ZDescriLinea
    Erase ZDescriRubro
    
    spLinea = "ListaLinea"
    Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstLinea.RecordCount > 0 Then
        With rstLinea
            .MoveFirst
            Do
                If .EOF = False Then
                    ZDescriLinea(rstLinea!Linea) = rstLinea!Nombre
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstLinea.Close
    End If
    
    
    spRubro = "ListaRubro"
    Set rstRubro = db.OpenRecordset(spRubro, dbOpenSnapshot, dbSQLPassThrough)
    If rstRubro.RecordCount > 0 Then
        With rstRubro
            .MoveFirst
            Do
                If .EOF = False Then
                    ZDescriRubro(rstRubro!Rubro) = rstRubro!Nombre
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRubro.Close
    End If
    
    
    
    
    spVendedor = "ListaVendedor"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstVendedor.RecordCount > 0 Then
        With rstVendedor
            .MoveFirst
            Do
                If .EOF = False Then
                    ZDescriVendedor(rstVendedor!Vendedor) = rstVendedor!Nombre
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstVendedor.Close
    End If
    
    
    
    Sql1 = "Select Estadistica.Tipo, Estadistica.Numero, Estadistica.Renglon, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Precio, Estadistica.PrecioUs, Estadistica.Importe, Estadistica.ImporteUs, Estadistica.Cliente, Estadistica.Linea, Estadistica.Costo1, Estadistica.Costo2, Estadistica.Coeficiente, Estadistica.Pedido, Estadistica.Fecha, Estadistica.OrdFecha, Estadistica.Articulo, Estadistica.Remito, Estadistica.Clave, Estadistica.WArticulo, Estadistica.Paridad, Estadistica.Importe1, Estadistica.Importe2, Estadistica.Importe3, Estadistica.Importe4, Cliente.Razon as [WDesCliente], Cliente.Vendedor as [WCodigoVendedor], CLiente.Rubro as [WCodigoRubro]"
    Sql2 = " FROM Estadistica, Cliente"
    Sql3 = " Where Estadistica.Cliente = Cliente.Cliente"
    Sql5 = " and Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    Sql6 = " and Estadistica.OrdFecha <= " + "'" + Whasta + "'"
    Sql7 = " and Cliente.Vendedor >= " + "'" + Desde.Text + "'"
    Sql8 = " and Cliente.Vendedor <= " + "'" + Hasta.Text + "'"
    Sql9 = " Order by Clave"
    
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
            
                WTipo = rstEstadistica!Tipo
                WNumero = rstEstadistica!numero
                WRenglon = rstEstadistica!Renglon
                WArticulo = rstEstadistica!Articulo
                WCantidad = rstEstadistica!Cantidad
                WPrecio = rstEstadistica!Precio
                WPrecioUs = rstEstadistica!PrecioUs
                WImporte = rstEstadistica!Importe
                WimporteUs = rstEstadistica!ImporteUs
                WCliente = rstEstadistica!Cliente
                WParidad = rstEstadistica!Paridad
                wvendedor = rstEstadistica!WCodigoVendedor
                WRubro = rstEstadistica!WCodigoRubro
                WLinea = rstEstadistica!Linea
                WCosto1 = rstEstadistica!Costo1
                WCosto2 = rstEstadistica!Costo2
                WCoeficiente = rstEstadistica!Coeficiente
                WPedido = rstEstadistica!Pedido
                WFecha = rstEstadistica!Fecha
                WImporte1 = rstEstadistica!Importe1
                WImporte2 = rstEstadistica!Importe2
                WImporte3 = rstEstadistica!Importe3
                WImporte4 = rstEstadistica!Importe4
                WOrdFecha = rstEstadistica!OrdFecha
                WWArticulo = rstEstadistica!WArticulo
                WRemito = rstEstadistica!Remito
                WClave = rstEstadistica!Clave
                Rem WDescriVendedor = rstEstadistica!WDesVendedor
                WDescriVendedor = ZDescriVendedor(wvendedor)
                WDescriLinea = ZDescriLinea(rstEstadistica!Linea)
                WDescriRubro = ZDescriRubro(WRubro)
                WDescriCliente = rstEstadistica!WDesCliente
                
                With rstEsta
                
                    .Index = "Clave"
                    .AddNew
                    !Tipo = WTipo
                    !numero = WNumero
                    !Renglon = WRenglon
                    !Articulo = WArticulo
                    !Cantidad = WCantidad
                    !Precio = WPrecio
                    !PrecioUs = WPrecioUs
                    !Importe = WImporte
                    !ImporteUs = WimporteUs
                    !Cliente = WCliente
                    !Paridad = WParidad
                    !Vendedor = wvendedor
                    !Rubro = WRubro
                    !Linea = WLinea
                    !Costo1 = WCosto1
                    !Costo2 = WCosto2
                    !Coeficiente = WCoeficiente
                    !Pedido = WPedido
                    !Fecha = WFecha
                    !Importe1 = WImporte1
                    !Importe2 = WImporte2
                    !Importe3 = WImporte3
                    !Importe4 = WImporte4
                    !OrdFecha = WOrdFecha
                    !WArticulo = WWArticulo
                    !Remito = WRemito
                    !Clave = WClave
                    !Nomempresa = Nomempresa
                    !Varios = WTitulo
                    !DescriVendedor = Trim(WDescriVendedor)
                    !DescriLinea = Trim(WDescriLinea)
                    !DescriRubro = Trim(WDescriRubro)
                    !DescriCliente = Trim(WDescriCliente)
                    .Update
                        
                End With
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
    End If
    
    With rstEsta
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                !WCantidad = 0
                !WImporte = 0
                !WimporteUs = 0
                !Costo2 = 0
                
                Producto = !Articulo
                Costo = 0
                If !Costo1 <> 0 And TipoCosto.ListIndex = 1 Then
                    Costo = !Costo1
                        Else
                    Call Calcula_Costo(Producto, Costo)
                End If
                        
                If !Tipo = 2 Then
                    !WCantidad = Abs(!Cantidad) * -1
                    !WImporte = Abs(!Importe) * -1
                    !WimporteUs = Abs(!ImporteUs) * -1
                    !Costo2 = Costo * Abs(!Cantidad) * -1
                        Else
                    !WCantidad = !Cantidad
                    !WImporte = !Importe
                    !WimporteUs = !ImporteUs
                    !Costo2 = Costo * !Cantidad
                End If
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Select Case TipoCosto.ListIndex
        Case 0
            WTitulo = "del " + DesdeFec.Text + " al " + HastaFec.Text + " (Costo actual)"
        Case 1
            WTitulo = "del " + DesdeFec.Text + " al " + HastaFec.Text + " (Costo F.Fact.)"
    End Select
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "1.-Listado de Estadistica de Ventas por Vendedor, Rubro y Linea"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    Dos = " and {Estadistica.Vendedor} in " + Desde.Text + " to " + Hasta.Text
    Listado.GroupSelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.ReportFileName = "WEstavenii.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEsta
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgEstaVen.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Command1_Click()
Stop
    ZSql = ""
    ZSql = ZSql + "DELETE Estadistica"
    ZSql = ZSql + " Where Clave = 'NULL'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    Stop
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Esta
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFec.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
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
    If KeyAscii = 27 Then
        DesdeFec.Text = "  /  /    "
    End If
End Sub

Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    TipoCosto.Clear
    
    TipoCosto.AddItem "Actual"
    TipoCosto.AddItem "Fecha Facturacion"
    
    TipoCosto.ListIndex = 0

    Desde.Text = "0"
    Hasta.Text = "9999"
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    If Left$(Producto, 2) = "PT" Or Left$(Producto, 2) = "PE" Or Left$(Producto, 2) = "DW" Or Left$(Producto, 2) = "NK" Or Left$(Producto, 2) = "RE" Then
    
    If Left$(Producto, 2) = "NK" Or Left$(Producto, 2) = "RE" Then
        Producto = "PT" + Mid$(Producto, 3, 10)
    End If
    
    If Left$(Producto, 2) = "NW" Then
        Producto = "DW" + Mid$(Producto, 3, 10)
    End If

    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    For da = 1 To Posi
        If Producto = Vecosto(da, 1) Then
            Costo = Val(Vecosto(da, 2))
            Exit Sub
        End If
    Next da
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
        
            Entra = "S"
    
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        If Left$(Articulo1, 2) = "DW" Then
                            Tipo = "T"
                            Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        End If
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
            If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
                Renglon = Renglon + 1
                Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
                Auxiliar(Renglon, 2) = 1
                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
    
    If Renglon > 0 Then
                    
        For da = 1 To Renglon
            Articulo = Auxiliar(da, 1)
            Cantidad = Auxiliar(da, 2)
            XVector = Auxiliar(da, 3)
            
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
                Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
                rstArticulo.Close
            End If
        Next da
        
            Else
            
        XArti = Left$(Producto, 3) + Right$(Producto, 7)
        spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Costo = rstArticulo!Costo2
            rstArticulo.Close
        End If
    
    End If
            
    
    Posi = Posi + 1
    Vecosto(Posi, 1) = Producto
    Vecosto(Posi, 2) = Str$(Costo)
    
        Else
        
    XArti = Left$(Producto, 3) + Right$(Producto, 7)
    spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Costo = rstArticulo!Costo2
        rstArticulo.Close
    End If

    End If
    
End Sub

