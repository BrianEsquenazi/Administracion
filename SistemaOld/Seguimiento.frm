VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgSeguimiento 
   Caption         =   "Seguimiento de Ordenes de Compra de Importacion"
   ClientHeight    =   6150
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   8145
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
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   7215
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
         Left            =   2040
         TabIndex        =   15
         Top             =   1200
         Width           =   2775
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
         Height          =   540
         Left            =   5280
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox HastaProv 
         Height          =   285
         Left            =   2880
         MaxLength       =   11
         TabIndex        =   11
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox DesdeProv 
         Height          =   285
         Left            =   2880
         MaxLength       =   11
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1335
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
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   1920
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
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   1920
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
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   480
         Width           =   1335
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
         Height          =   495
         Left            =   5280
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Listado"
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
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Proveedor"
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
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Proveedor"
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WOrdPenPrv.rpt"
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
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "Seguimiento.frx":0000
      Left            =   240
      List            =   "Seguimiento.frx":0007
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6960
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstSeguimiento As Recordset
Dim spSeguimiento As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim XParam As String
Dim XIndice As Integer

Dim Vector(5000, 15) As String
Dim VectorII(5000, 15) As String
Dim VectorIII(5000, 15) As String
Dim Lugar As Integer
Dim LugarII As Integer
Dim LugarIII As Integer
Dim ZZLugar As Integer

Dim WProveedor As String
Dim WArticulo As String
Dim WOrden As String
Dim WFecha As String
Dim WFecha2 As String
Dim WCantidad As String
Dim WPrecio As String
Dim WCarpeta As String
Dim WFechaLlegada As String
Dim WPagoDespacho As String
Dim WTexto1 As String
Dim WTexto2 As String
Dim WRenglon As String
Dim Empe(12, 10) As String


Private Sub Acepta_Click()

    ZSql = "DELETE Seguimiento"
    spSeguimiento = ZSql
    Set rstSeguimiento = db.OpenRecordset(spSeguimiento, dbOpenSnapshot, dbSQLPassThrough)
    

    Erase Vector
    Erase VectorII
    
    Lugar = 0
    LugarII = 0
    
    
    XEmpresa = WEmpresa
    
    If Val(WEmpresa) = 1 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        Hasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        Hasta = 4
    End If
    
    For CiclaEmpresa = 1 To Hasta
    
        WEmpresa = Empe(CiclaEmpresa, 1)
        txtOdbc = Empe(CiclaEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Cantidad > Orden.Recibida"
        ZSql = ZSql + " and Orden.Tipo = 1"
        ZSql = ZSql + " and Orden.Proveedor >= " + "'" + DesdeProv.Text + "'"
        ZSql = ZSql + " and Orden.Proveedor <= " + "'" + HastaProv.Text + "'"
        ZSql = ZSql + " Order by Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        ZEntregaI = IIf(IsNull(rstOrden!EntregaI), "0", rstOrden!EntregaI)
                        ZEntregaII = IIf(IsNull(rstOrden!EntregaII), "0", rstOrden!EntregaII)
                
                        ZEntra = "N"
                
                        If Tipo.ListIndex = 0 Then
                            ZEntra = "S"
                                Else
                            If Tipo.ListIndex = 1 And ZEntregaI = 0 Then
                                ZEntra = "S"
                                    Else
                                If Tipo.ListIndex = 2 And ZEntregaII = 0 Then
                                    ZEntra = "S"
                                End If
                            End If
                        End If
                        
                        If ZEntra = "S" Then
                
                            Lugar = Lugar + 1
                   
                            Vector(Lugar, 1) = rstOrden!Proveedor
                            Vector(Lugar, 2) = rstOrden!Articulo
                            Vector(Lugar, 3) = Str$(rstOrden!Orden)
                            Vector(Lugar, 4) = rstOrden!Fecha
                            Vector(Lugar, 5) = rstOrden!fecha2
                            Vector(Lugar, 6) = Str$(rstOrden!Cantidad)
                            Vector(Lugar, 7) = Str$(rstOrden!Precio)
                            Vector(Lugar, 8) = Str$(rstOrden!Carpeta)
                            Vector(Lugar, 9) = IIf(IsNull(rstOrden!FechaLlegada), "  /  /    ", rstOrden!FechaLlegada)
                            Vector(Lugar, 10) = IIf(IsNull(rstOrden!PagoDespacho), "0", rstOrden!PagoDespacho)
                            Vector(Lugar, 11) = IIf(IsNull(rstOrden!EntregaI), "0", rstOrden!EntregaI)
                            Vector(Lugar, 12) = IIf(IsNull(rstOrden!EntregaII), "0", rstOrden!EntregaII)
                        
                        End If
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
    
        End If
        
    Next CiclaEmpresa
    
    Call Conecta_Empresa
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1
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
        
    Pasa = 0
    LugarIII = 0
    Erase VectorIII
    
    For Ciclo = 1 To Lugar
    
        WCarpeta = Vector(Ciclo, 8)
        
        Rem If Val(WCarpeta) = 2741 Then Stop
           
        If Pasa = 0 Then
        
            Corte = WCarpeta
            Pasa = 1
            Erase VectorII
            LugarII = 0
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM ObservaOrden"
            ZSql = ZSql + " Where ObservaOrden.Carpeta = " + "'" + WCarpeta + "'"
            ZSql = ZSql + " Order by ObservaOrden.Clave"
            spObservaOrden = ZSql
            Set rstObservaOrden = db.OpenRecordset(spObservaOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstObservaOrden.RecordCount > 0 Then
                With rstObservaOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            
                            LugarII = LugarII + 1
                            
                            VectorII(LugarII, 1) = rstObservaOrden!texto1
                            VectorII(LugarII, 2) = rstObservaOrden!texto2
                            VectorII(LugarII, 3) = rstObservaOrden!Renglon
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstObservaOrden.Close
            End If
            ZZLugar = LugarII
            LugarII = 0
            
        End If
        
        
        If Val(Corte) <> Val(WCarpeta) Then
        
            If ZZLugar > LugarII Then
                For ZZZCiclo = LugarII + 1 To ZZLugar
                    LugarIII = LugarIII + 1
                    VectorIII(LugarIII, 1) = WProveedor
                    VectorIII(LugarIII, 2) = ""
                    VectorIII(LugarIII, 3) = WOrden
                    VectorIII(LugarIII, 4) = WFecha
                    VectorIII(LugarIII, 5) = WFecha2
                    VectorIII(LugarIII, 6) = ""
                    VectorIII(LugarIII, 7) = ""
                    VectorIII(LugarIII, 8) = Corte
                    VectorIII(LugarIII, 9) = WFechaLlegada
                    VectorIII(LugarIII, 10) = WPagoDespacho
                    VectorIII(LugarIII, 11) = VectorII(ZZZCiclo, 3)
                    VectorIII(LugarIII, 12) = VectorII(ZZZCiclo, 1)
                    VectorIII(LugarIII, 13) = VectorII(ZZZCiclo, 2)
                    VectorIII(LugarIII, 14) = WEntregaI
                    VectorIII(LugarIII, 15) = WEntregaII
                Next ZZZCiclo
            End If
        
            Corte = WCarpeta
            Erase VectorII
            LugarII = 0
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM ObservaOrden"
            ZSql = ZSql + " Where ObservaOrden.Carpeta = " + "'" + WCarpeta + "'"
            ZSql = ZSql + " Order by ObservaOrden.Clave"
            spObservaOrden = ZSql
            Set rstObservaOrden = db.OpenRecordset(spObservaOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstObservaOrden.RecordCount > 0 Then
                With rstObservaOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            
                            LugarII = LugarII + 1
                            
                            VectorII(LugarII, 1) = rstObservaOrden!texto1
                            VectorII(LugarII, 2) = rstObservaOrden!texto2
                            VectorII(LugarII, 3) = rstObservaOrden!Renglon
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstObservaOrden.Close
            End If
            ZZLugar = LugarII
            LugarII = 0
            
        End If
        
        WProveedor = Vector(Ciclo, 1)
        WArticulo = Vector(Ciclo, 2)
        WOrden = Vector(Ciclo, 3)
        WFecha = Vector(Ciclo, 4)
        WFecha2 = Vector(Ciclo, 5)
        WCantidad = Vector(Ciclo, 6)
        WPrecio = Vector(Ciclo, 7)
        WCarpeta = Vector(Ciclo, 8)
        WFechaLlegada = Vector(Ciclo, 9)
        WPagoDespacho = Vector(Ciclo, 10)
        WEntregaI = Vector(Ciclo, 11)
        WEntregaII = Vector(Ciclo, 12)
        
        
        LugarIII = LugarIII + 1
        LugarII = LugarII + 1
        
        VectorIII(LugarIII, 1) = WProveedor
        VectorIII(LugarIII, 2) = WArticulo
        VectorIII(LugarIII, 3) = WOrden
        VectorIII(LugarIII, 4) = WFecha
        VectorIII(LugarIII, 5) = WFecha2
        VectorIII(LugarIII, 6) = WCantidad
        VectorIII(LugarIII, 7) = WPrecio
        VectorIII(LugarIII, 8) = WCarpeta
        VectorIII(LugarIII, 9) = WFechaLlegada
        VectorIII(LugarIII, 10) = WPagoDespacho
        VectorIII(LugarIII, 11) = VectorII(LugarII, 3)
        VectorIII(LugarIII, 12) = VectorII(LugarII, 1)
        VectorIII(LugarIII, 13) = VectorII(LugarII, 2)
        VectorIII(LugarIII, 14) = WEntregaI
        VectorIII(LugarIII, 15) = WEntregaII
        
    Next Ciclo
    
    If ZZLugar > LugarII Then
        For ZZZCiclo = LugarII + 1 To ZZLugar
            LugarIII = LugarIII + 1
            VectorIII(LugarIII, 1) = WProveedor
            VectorIII(LugarIII, 2) = ""
            VectorIII(LugarIII, 3) = WOrden
            VectorIII(LugarIII, 4) = WFecha
            VectorIII(LugarIII, 5) = WFecha2
            VectorIII(LugarIII, 6) = ""
            VectorIII(LugarIII, 7) = ""
            VectorIII(LugarIII, 8) = WCarpeta
            VectorIII(LugarIII, 9) = WFechaLlegada
            VectorIII(LugarIII, 10) = WPagoDespacho
            VectorIII(LugarIII, 11) = VectorII(ZZZCiclo, 3)
            VectorIII(LugarIII, 12) = VectorII(ZZZCiclo, 1)
            VectorIII(LugarIII, 13) = VectorII(ZZZCiclo, 2)
            VectorIII(LugarIII, 14) = WEntregaI
            VectorIII(LugarIII, 15) = WEntregaII
        Next ZZZCiclo
    End If
    
    Call Conecta_Empresa
    
    For Ciclo = 1 To LugarIII
    
        WProveedor = VectorIII(Ciclo, 1)
        WArticulo = VectorIII(Ciclo, 2)
        WOrden = VectorIII(Ciclo, 3)
        WFecha = VectorIII(Ciclo, 4)
        WFecha2 = VectorIII(Ciclo, 5)
        WCantidad = VectorIII(Ciclo, 6)
        WPrecio = VectorIII(Ciclo, 7)
        WCarpeta = VectorIII(Ciclo, 8)
        WFechaLlegada = VectorIII(Ciclo, 9)
        WPagoDespacho = VectorIII(Ciclo, 10)
        WRenglon = Str$(Ciclo)
        WTexto1 = VectorIII(Ciclo, 12)
        WTexto2 = VectorIII(Ciclo, 13)
        WEntregaI = VectorIII(Ciclo, 14)
        WEntregaII = VectorIII(Ciclo, 15)
        WDesArticulo = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDesArticulo = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Seguimiento ("
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "DesArticulo ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Fecha2 ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Texto1 ,"
        ZSql = ZSql + "Texto2 ,"
        ZSql = ZSql + "EntregaI ,"
        ZSql = ZSql + "EntregaII ,"
        ZSql = ZSql + "FechaLLegada ,"
        ZSql = ZSql + "PagoDespacho )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WProveedor + "',"
        ZSql = ZSql + "'" + WArticulo + "',"
        ZSql = ZSql + "'" + WDesArticulo + "',"
        ZSql = ZSql + "'" + WOrden + "',"
        ZSql = ZSql + "'" + WRenglon + "',"
        ZSql = ZSql + "'" + WFecha + "',"
        ZSql = ZSql + "'" + WFecha2 + "',"
        ZSql = ZSql + "'" + WCarpeta + "',"
        ZSql = ZSql + "'" + WCantidad + "',"
        ZSql = ZSql + "'" + WPrecio + "',"
        ZSql = ZSql + "'" + WTexto1 + "',"
        ZSql = ZSql + "'" + WTexto2 + "',"
        ZSql = ZSql + "'" + WEntregaI + "',"
        ZSql = ZSql + "'" + WEntregaII + "',"
        ZSql = ZSql + "'" + WFechaLlegada + "',"
        ZSql = ZSql + "'" + WPagoDespacho + "')"
       
        spSeguimiento = ZSql
        Set rstSeguimiento = db.OpenRecordset(spSeguimiento, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
        

    Listado.WindowTitle = "Seguimiento de Ordenes de Compra de Importaciones"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Seguimiento.Proveedor} in " + Chr$(34) + "0" + Chr$(34) + " to " + Chr$(34) + "99999999999" + Chr$(34)
    Listado.GroupSelectionFormula = Uno
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Seguimiento.Proveedor, Seguimiento.Articulo, Seguimiento.Orden, Seguimiento.Renglon, Seguimiento.Fecha, Seguimiento.Fecha2, Seguimiento.Carpeta, Seguimiento.Cantidad, Seguimiento.Precio, Seguimiento.Texto1, Seguimiento.Texto2, Seguimiento.FechaLLegada, Seguimiento.PagoDespacho, Seguimiento.EntregaI, Seguimiento.EntregaII, Seguimiento.DesArticulo, " _
            + "Proveedor.Nombre " _
            + "From " _
            + DSQ + ".dbo.Seguimiento Seguimiento, " _
            + DSQ + ".dbo.Proveedor Proveedor " _
            + "Where " _
            + "Seguimiento.Proveedor = Proveedor.Proveedor AND " _
            + "Seguimiento.Proveedor >= '0' AND " _
            + "Seguimiento.Proveedor <= '99999999999'"
    
    Listado.ReportFileName = "WSeguimiento.rpt"
    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    
    PrgSeguimiento.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaProv.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub HastaProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeProv.SetFocus
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Pend. Certificado"
    Tipo.AddItem "Pend. Docum Embarque"
    
    Tipo.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgSeguimiento.Caption = "Seguimiento de Ordenes de Compra de Importacion :  " + !Nombre
        End If
    End With
    DesdeProv.Text = ""
    HastaProv.Text = "99999999999"
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
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            spProveedor = "ListaProveedoresOrd"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = RstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = RstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.SetFocus

End Sub



Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrd"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(RstProveedor!Nombre) - WEspacios
                
                For aa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        Auxi = Str$(RstProveedor!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next aa
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    RstProveedor.Close
    
    End If

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WProveedor = WIndice.List(Indice)
            DesdeProv.Text = WProveedor
            HastaProv.Text = WProveedor
            
            Ayuda.Visible = False
            Pantalla.Visible = False
        Case Else
    End Select
    
End Sub


