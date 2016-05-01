VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAnalisisDesarrollo 
   Caption         =   "Analisis de Desarrollo"
   ClientHeight    =   3225
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3225
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   6015
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
         Left            =   2280
         TabIndex        =   7
         Top             =   1680
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
         Left            =   840
         TabIndex        =   6
         Top             =   1680
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
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   5
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
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   2400
         TabIndex        =   1
         Top             =   960
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   2400
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgAnalisisDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstOrdenTrabajo As Recordset
Dim spOrdenTrabajo As String
Dim rstAnalisisDesarrollo As Recordset
Dim spAnalisisDesarrollo As String
Dim Empe(10, 10) As String
Dim Vector(1000, 10) As String
Dim VectorII(1000, 10) As String

Private Sub Acepta_Click()

    WDesdeFecha = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHastaFecha = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            ZEmpresa = !Nombre
        End If
    End With
    
    ZTitulo = "Desde el " + DesdeFecha.Text + " hasta el " + HastaFecha.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM OrdenTrabajo"
    ZSql = ZSql + " Order by Orden"
    spOrdenTrabajo = ZSql
    Set rstOrdenTrabajo = db.OpenRecordset(spOrdenTrabajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrdenTrabajo.RecordCount > 0 Then
        With rstOrdenTrabajo
            .MoveFirst
            Do
                If .EOF = False Then
                    WFecha = Right$(rstOrdenTrabajo!Fecha, 4) + Mid$(rstOrdenTrabajo!Fecha, 4, 2) + Left$(rstOrdenTrabajo!Fecha, 2)
                    If WFecha >= WDesdeFecha And WFecha <= WHastaFecha Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = rstOrdenTrabajo!Orden
                        Vector(Lugar, 2) = rstOrdenTrabajo!Fecha
                        Vector(Lugar, 3) = rstOrdenTrabajo!Observaciones
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrdenTrabajo.Close
    End If
    
    aa = WEmpresa
    
    
    ZSql = "DELETE AnalisisDesarrollo"
    spAnalisisDesarrollo = ZSql
    Set rstAnalisisDesarrollo = db.OpenRecordset(spAnalisisDesarrollo, dbOpenSnapshot, dbSQLPassThrough)
    
    For Ciclo = 1 To Lugar
    
        ZOrden = Vector(Ciclo, 1) + "-100"
        ZFecha = Vector(Ciclo, 2)
        ZObservaciones = Vector(Ciclo, 3)
        
        LugarII = 0
        Erase VectorII
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Terminado = " + "'" + ZOrden + "'"
        ZSql = ZSql + " and Hoja.Renglon = 1"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        If Left$(rstHoja!Producto, 2) = "PT" Then
                            
                            LugarII = LugarII + 1
                            
                            VectorII(LugarII, 1) = ZOrden
                            VectorII(LugarII, 2) = ZFecha
                            VectorII(LugarII, 3) = ZObservaciones
                            VectorII(LugarII, 4) = Str$(rstHoja!Hoja)
                            VectorII(LugarII, 5) = rstHoja!Producto
                            VectorII(LugarII, 6) = Str$(rstHoja!Cantidad)

                        End If
                                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
        End If
        
        If LugarII = 0 Then
            LugarII = LugarII + 1
            VectorII(LugarII, 1) = ZOrden
            VectorII(LugarII, 2) = ZFecha
            VectorII(LugarII, 3) = ZObservaciones
            VectorII(LugarII, 4) = ""
            VectorII(LugarII, 5) = ""
            VectorII(LugarII, 6) = ""
        End If
        
        ZImpo1 = 0
        ZImpo2 = 0
        
        For CicloII = 1 To LugarII
    
            ZOrden = VectorII(CicloII, 1)
            ZFecha = VectorII(CicloII, 2)
            ZObservaciones = VectorII(CicloII, 3)
            ZHoja = VectorII(CicloII, 4)
            ZProducto = VectorII(CicloII, 5)
            ZCantidad = VectorII(CicloII, 6)
        
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
                Case Else
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            WVenta = 0
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + ZProducto + "'"
            ZSql = ZSql + " and Estadistica.Tipo = 1"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            If rstEstadistica!Lote1 = Val(ZHoja) Then
                                WVenta = WVenta + rstEstadistica!canti1
                            End If
                            If rstEstadistica!Lote2 = Val(ZHoja) Then
                                WVenta = WVenta + rstEstadistica!canti2
                            End If
                            If rstEstadistica!Lote3 = Val(ZHoja) Then
                                WVenta = WVenta + rstEstadistica!canti3
                            End If
                            If rstEstadistica!Lote4 = Val(ZHoja) Then
                                WVenta = WVenta + rstEstadistica!canti4
                            End If
                            If rstEstadistica!Lote5 = Val(ZHoja) Then
                                WVenta = WVenta + rstEstadistica!Canti5
                            End If
                                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEstadistica.Close
            End If
        
            Call Conecta_Empresa
            
            If WVenta > Val(ZCantidad) Then
                WVenta = Val(ZCantidad)
            End If
        
            ZVenta = Str$(WVenta)
            ZImpo1 = "0"
            ZImpo2 = "0"
            If CicloII = 1 Then
                ZImpo3 = "1"
                    Else
                ZImpo3 = "0"
            End If
            ZRenglon = Str$(CicloII)
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO AnalisisDesarrollo ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Produccion ,"
            ZSql = ZSql + "Venta ,"
            ZSql = ZSql + "Hoja ,"
            ZSql = ZSql + "Empresa ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Impo1 ,"
            ZSql = ZSql + "Impo2 ,"
            ZSql = ZSql + "Impo3 ,"
            ZSql = ZSql + "Articulo )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZOrden + "',"
            ZSql = ZSql + "'" + ZObservaciones + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZRenglon + "',"
            ZSql = ZSql + "'" + ZCantidad + "',"
            ZSql = ZSql + "'" + ZVenta + "',"
            ZSql = ZSql + "'" + ZHoja + "',"
            ZSql = ZSql + "'" + ZEmpresa + "',"
            ZSql = ZSql + "'" + ZTitulo + "',"
            ZSql = ZSql + "'" + ZImpo1 + "',"
            ZSql = ZSql + "'" + ZImpo2 + "',"
            ZSql = ZSql + "'" + ZImpo3 + "',"
            ZSql = ZSql + "'" + ZProducto + "')"
                
            spAnalisisDesarrollo = ZSql
            Set rstAnalisisDesarrollo = db.OpenRecordset(spAnalisisDesarrollo, dbOpenSnapshot, dbSQLPassThrough)
            
            If Val(ZCantidad) <> 0 Then
                ZImpo1 = 1
            End If
            If Val(ZVenta) <> 0 Then
                ZImpo2 = 1
            End If
        
        Next CicloII
        
        If ZImpo1 = 1 Then
            ZSql = ""
            ZSql = ZSql + "UPDATE AnalisisDesarrollo SET "
            ZSql = ZSql + "Impo1 = 1"
            ZSql = ZSql + " Where Codigo = " + "'" + ZOrden + "'"
            ZSql = ZSql + " and Renglon = 1"
            spAnalisisDesarrollo = ZSql
            Set rstAnalisisDesarrollo = db.OpenRecordset(spAnalisisDesarrollo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        If ZImpo2 = 1 Then
            ZSql = ""
            ZSql = ZSql + "UPDATE AnalisisDesarrollo SET "
            ZSql = ZSql + "Impo2 = 1"
            ZSql = ZSql + " Where Codigo = " + "'" + ZOrden + "'"
            ZSql = ZSql + " and Renglon = 1"
            spAnalisisDesarrollo = ZSql
            Set rstAnalisisDesarrollo = db.OpenRecordset(spAnalisisDesarrollo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    
    Next Ciclo
    
    Listado.WindowTitle = "Analisis de Desarrollo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT AnalisisDesarrollo.Codigo, AnalisisDesarrollo.Descripcion, AnalisisDesarrollo.Fecha, AnalisisDesarrollo.Produccion, AnalisisDesarrollo.Venta, AnalisisDesarrollo.Hoja, AnalisisDesarrollo.Articulo, AnalisisDesarrollo.Empresa, AnalisisDesarrollo.Titulo, AnalisisDesarrollo.Impo1, AnalisisDesarrollo.Impo2, AnalisisDesarrollo.Impo3 " _
            + "From " _
            + DSQ + ".dbo.AnalisisDesarrollo AnalisisDesarrollo " _
            + "Where " _
            + "AnalisisDesarrollo.Codigo >= ' ' AND " _
            + "AnalisisDesarrollo.Codigo <= 'ZZZZZZZZZZ'"
    
    Listado.Connect = Connect()
    
    Rem Listado.GroupSelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "WAnalisisDesarrollo.rpt"
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    PrgAnalisisDesarrollo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFecha.SetFocus
    End If
End Sub

Sub Form_Load()
    
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub


