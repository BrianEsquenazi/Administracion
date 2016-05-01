VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompos 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Composicion de Productos Terminados"
   ClientHeight    =   5505
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5505
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1935
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox TipoListado 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   2895
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Listado "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wcompos.rpt"
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
      ItemData        =   "compos.frx":0000
      Left            =   120
      List            =   "compos.frx":0007
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCompos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Cantidad As Double
Private Producto As String
Private Costo As Double
Private Costo1 As Double
Private Costo2 As Double
Private WCosto1 As String
Private WCosto2 As String
Private Auxiliar(100, 7) As String
Private XVector(20000, 6) As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim XParam As String

Dim ZCosto1 As Double
Dim WWCosto1 As Double
Dim ZZOrdenI As Double
Dim ZZOrdenII As Double
Dim ZZOrdenIII As Double
Dim ZZPtaOrdenI As Double
Dim ZZPtaOrdenII As Double
Dim ZZPtaOrdenIII As Double
    
Dim ZZFechaOrdenI As String
Dim ZZFechaOrdenII As String
Dim ZZFechaOrdenIII As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Erase XVector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                + Hasta.Text + "'"
                                         
    Set rstComposicion = db.OpenRecordset("ListaComposicionDesdeHasta " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicion.RecordCount > 0 Then

    With rstComposicion
    
        .MoveFirst
        If .NoMatch = False Then
        
            Do
            
                Renglon = Renglon + 1
                    
                XVector(Renglon, 1) = rstComposicion!Tipo
                XVector(Renglon, 2) = rstComposicion!Articulo1
                XVector(Renglon, 3) = rstComposicion!Articulo2
                XVector(Renglon, 4) = rstComposicion!Cantidad
                XVector(Renglon, 5) = rstComposicion!Clave
                XVector(Renglon, 6) = rstComposicion!Terminado
                    
                .MoveNext
                    
                If .EOF = True Then
                    Exit Do
                End If
                        
            Loop
            
        End If
            
    End With
    rstComposicion.Close
    
    End If
    
    For Da = 1 To Renglon
    
        Tipo = XVector(Da, 1)
        Articulo1 = XVector(Da, 2)
        Articulo2 = XVector(Da, 3)
        Cantidad = XVector(Da, 4)
        Clave = XVector(Da, 5)
        Terminado = XVector(Da, 6)
        
        DescriTerminado = ""
        DescriArticulo1 = ""
        DescriArticulo2 = ""
        
        Rem If Left$(Articulo1, 2) = "DW" Then
        Rem     Tipo = "T"
        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
        Rem End If
        
        Select Case Tipo
            Case "T"
                Producto = Articulo2
                Call Calcula_Costo(Producto, Costo)
                spTerminado = "ConsultaTerminado " + "'" + Articulo2 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    DescriArticulo2 = Left$(rstTerminado!Descripcion, 30)
                    rstTerminado.Close
                End If
                WDescriTipo = ""
                
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Articulo1 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                        DescriArticulo1 = Left$(rstArticulo!Descripcion, 30)
                        Select Case TipoListado.ListIndex
                            Case 0
                                Costo = rstArticulo!Costo2
                                ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                                If ZTipoCosto = 1 Then
                                    WDescriTipo = "Estimado"
                                        Else
                                    WDescriTipo = ""
                                End If
                                rstArticulo.Close
                            Case 1
                                Costo = rstArticulo!Costo1
        
                                Costo1 = rstArticulo!Costo1
                                WWCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                                ZCosto1 = IIf(IsNull(rstArticulo!ZCosto1), "0", rstArticulo!ZCosto1)
                                ZZOrdenI = IIf(IsNull(rstArticulo!OrdenI), "0", rstArticulo!OrdenI)
                                ZZOrdenII = IIf(IsNull(rstArticulo!OrdenII), "0", rstArticulo!OrdenII)
                                ZZOrdenIII = IIf(IsNull(rstArticulo!OrdenIII), "0", rstArticulo!OrdenIII)
                                ZZPtaOrdenI = IIf(IsNull(rstArticulo!PtaOrdenI), "0", rstArticulo!PtaOrdenI)
                                ZZPtaOrdenII = IIf(IsNull(rstArticulo!PtaOrdenII), "0", rstArticulo!PtaOrdenII)
                                ZZPtaOrdenIII = IIf(IsNull(rstArticulo!PtaOrdenIII), "0", rstArticulo!PtaOrdenIII)
                                    
                                ZZFechaOrdenI = ""
                                ZZFechaOrdenII = ""
                                ZZFechaOrdenIII = ""
                                
                                ZZMoneda = ""
                                
                                rstArticulo.Close
                                
                                XEmpresa = WEmpresa
        
                                If ZZPtaOrdenI <> 0 And ZZOrdenI <> 0 Then
                                
                                    Select Case ZZPtaOrdenI
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
                                        Case 11
                                            WEmpresa = "0011"
                                            txtOdbc = "Empresa11"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                    End Select
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Orden"
                                    ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenI) + "'"
                                    spOrden = ZSql
                                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstOrden.RecordCount > 0 Then
                                        ZZFechaOrdenI = rstOrden!Fecha
                                        Select Case rstOrden!Moneda
                                            Case 0
                                                ZZMoneda = "U$S"
                                            Case 1
                                                ZZMoneda = "$"
                                            Case Else
                                                ZZMoneda = "Eur"
                                        End Select
                                        rstOrden.Close
                                    End If
                                    
                                    Call Conecta_Empresa
                                    
                                End If
                                
                                
                                If ZZPtaOrdenII <> 0 And ZZOrdenII <> 0 Then
                                    
                                    Select Case ZZPtaOrdenII
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
                                        Case 11
                                            WEmpresa = "0011"
                                            txtOdbc = "Empresa11"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                    End Select
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Orden"
                                    ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenII) + "'"
                                    spOrden = ZSql
                                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstOrden.RecordCount > 0 Then
                                        ZZFechaOrdenII = rstOrden!Fecha
                                        rstOrden.Close
                                    End If
                                    
                                    Call Conecta_Empresa
                                    
                                End If
                                
                                If ZZPtaOrdenIII <> 0 And ZZOrdenIII <> 0 Then
                                    
                                    Select Case ZZPtaOrdenIII
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
                                        Case 11
                                            WEmpresa = "0011"
                                            txtOdbc = "Empresa11"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                    End Select
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Orden"
                                    ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenIII) + "'"
                                    spOrden = ZSql
                                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstOrden.RecordCount > 0 Then
                                        ZZFechaOrdenIII = rstOrden!Fecha
                                        rstOrden.Close
                                    End If
                                    
                                    Call Conecta_Empresa
                                    
                                End If
                                
                                Rem DADA
                                Rem spCambio = "ConsultaCambio " + "'" + ZZFecha + "'"
                                Rem Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                                Rem If rstCambio.RecordCount > 0 Then
                                Rem     ZZParidad = rstCambio!Cambio
                                Rem     rstCambio.Close
                                Rem End If
                                
                                Rem If ZZParidad <> 0 Then
                                Rem     ZZCostoPesos = WPrecio * ZZParidad
                                Rem End If
                                        
                                If ZZFechaOrdenI <> "" Then
                                    WFechaOrdI = Right$(ZZFechaOrdenI, 4) + Mid$(ZZFechaOrdenI, 4, 2) + Left$(ZZFechaOrdenI, 2)
                                        Else
                                    WFechaOrdI = ""
                                End If
                                If ZZFechaOrdenII <> "" Then
                                    WFechaOrdII = Right$(ZZFechaOrdenII, 4) + Mid$(ZZFechaOrdenII, 4, 2) + Left$(ZZFechaOrdenII, 2)
                                        Else
                                    WFechaOrdII = ""
                                End If
                                If ZZFechaOrdenIII <> "" Then
                                    WFechaOrdIII = Right$(ZZFechaOrdenIII, 4) + Mid$(ZZFechaOrdenIII, 4, 2) + Left$(ZZFechaOrdenIII, 2)
                                        Else
                                    WFechaOrdIII = ""
                                End If
                                
                                If WFechaOrdI <> "" And WFechaOrdI > WFechaOrdII And WFechaOrdI > WFechaOrdIII Then
                                    Costo = Costo1
                                End If
                                
                                If WFechaOrdII <> "" And WFechaOrdII > WFechaOrdI And WFechaOrdII > WFechaOrdIII Then
                                
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                     spCambios = "ConsultaCambio  " + "'" + ZZFechaOrdenII + "'"
                                     Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                                     If rstCambios.RecordCount > 0 Then
                                         ZZZZParidad = rstCambios!Cambio
                                         rstCambios.Close
                                         If ZZZZParidad <> 0 Then
                                             ZZCosto1Dol = WWCosto1 / ZZZZParidad
                                         End If
                                     End If
                                    Call Conecta_Empresa
                                
                                    Costo = ZZCosto1Dol
                                End If
                                
                                If WFechaOrdIII <> "" And WFechaOrdIII > WFechaOrdI And WFechaOrdIII > WFechaOrdII Then
                                    Costo = ZCosto1
                                End If
                                WDescriTipo = ""
                                                        
                                Rem WEmpresa = "0001"
                                Rem txtOdbc = "Empresa01"
                                Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                Rem spCambios = "ConsultaCambio  " + "'" + FechaOrdenII.Text + "'"
                                Rem Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                                Rem If rstCambios.RecordCount > 0 Then
                                Rem     ZZParidad = rstCambios!Cambio
                                Rem     rstCambios.Close
                                Rem     If ZZParidad <> 0 Then
                                Rem         ZZCosto1Dol = Val(WCosto1.Text) / ZZParidad
                                Rem         WCosto1Dol.Text = Str$(ZZCosto1Dol)
                                Rem         WCosto1Dol.Text = Pusing("###,###.##", WCosto1Dol.Text)
                                Rem     End If
                                Rem End If
                                
                            Case 2
                                Costo = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                                WDescriTipo = ""
                                rstArticulo.Close
                            Case 3
                                Costo = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
                                WDescriTipo = "Reposicion"
                                If Costo = 0 Then
                                    Costo = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
                                    ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                                    If ZTipoCosto = 1 Then
                                        WDescriTipo = "Estimado"
                                            Else
                                        WDescriTipo = ""
                                    End If
                                End If
                                rstArticulo.Close
                            Case 4
                                Costo = IIf(IsNull(rstArticulo!UltimoFob), "0", rstArticulo!UltimoFob)
                                WDescriTipo = ""
                                If Costo = 0 Then
                                    Costo = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
                                    ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                                    If ZTipoCosto = 1 Then
                                        WDescriTipo = "Estimado"
                                            Else
                                        WDescriTipo = ""
                                    End If
                                End If
                                rstArticulo.Close
                            Case Else
                                Costo = 0
                                rstArticulo.Close
                        End Select
                End If
                
            Case Else
        End Select
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DescriTerminado = Left$(rstTerminado!Descripcion, 30)
            rstTerminado.Close
        End If
        
        Cantidad = XVector(Da, 4)
        Costo1 = Costo
        Call Redondeo(Costo1)
        WCosto1 = Costo1
        Costo2 = Costo * Cantidad
        Call Redondeo(Costo2)
        WCosto2 = Costo2
        WCosto1 = Pusing("###,###.##", WCosto1)
        WCosto2 = Pusing("###,###.##", WCosto2)
            
        XParam = "'" + Clave + "','" _
                    + WCosto1 + "','" _
                    + WCosto2 + "','" _
                    + DescriTerminado + "','" _
                    + DescriArticulo1 + "','" _
                    + DescriArticulo2 + "'"
                                           
        spComposicion = "ModificaComposicionCosto " + XParam
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Composicion SET "
        ZSql = ZSql + " DescriTipo = " + "'" + WDescriTipo + "'"
        ZSql = ZSql + " Where Clave = " + "'" + Clave + "'"
        spComposicion = ZSql
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Da

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            Select Case TipoListado.ListIndex
                Case 0
                    !varios = "(Costo Standard o Estimado)"
                Case 1
                    !varios = "(Costo Ultima Compra)"
                Case 2
                    !varios = "(Costo Promedio)"
                Case 3
                    !varios = "(Costo Reposicion)"
                Case 4
                    !varios = "(Costo Standard Ultima Compra)"
                Case Else
                    !varios = ""
            End Select
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Composicion de Productos Terminados"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Composicion.terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    Listado.SQLQuery = "SELECT Composicion.Clave , Composicion.Terminado, Composicion.Tipo, Composicion.Articulo1, Composicion.Articulo2, Composicion.Cantidad, Composicion.Costo1, Composicion.Costo2, Composicion.DescriTerminado, Composicion.DescriArticulo1, Composicion.DescriArticulo2, Composicion.DescriTipo " + _
                        "From " + DSQ + ".dbo.Composicion Composicion " + _
                        "Where Composicion.Terminado >= '" + Desde.Text + "' AND Composicion.Terminado <= '" + Hasta.Text + "'"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgCompos.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
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
            PrgCompos.Caption = "Listado de composicion de Productos Terminados :  " + !Nombre
        End If
    End With
    
    TipoListado.Clear
    
    TipoListado.AddItem "Costo Standard y Estimado"
    TipoListado.AddItem "Costo Ultima Compra"
    TipoListado.AddItem "Costo Promedio"
    TipoListado.AddItem "Costo Reposicion"
    TipoListado.AddItem "Costo Standard Ultima Compra"
    
    TipoListado.ListIndex = 0

    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
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
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstTerminado!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTerminado.Close
            
            End If
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub


Private Sub pantalla_Click()

    Pantalla.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                    Desde.Text = rstTerminado!Codigo
                    Hasta.Text = rstTerminado!Codigo
                    rstTerminado.Close
            End If
            Desde.SetFocus
    End Select
    
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
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
                        
                        Rem If Left$(Articulo1, 2) = "DW" Then
                        Rem     Tipo = "T"
                        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        Rem End If
                        
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
            
            Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
            Rem     Renglon = Renglon + 1
            Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
            Rem     Auxiliar(Renglon, 2) = 1
            Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Auxiliar(Da, 2)
        WVector = Auxiliar(Da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            Rem Select Case TipoListado.ListIndex
            Rem     Case 0
            Rem         WCosto = (Cantidad * rstArticulo!Costo2 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(WVector))
            Rem     Case 1
            Rem         WCosto = (Cantidad * rstArticulo!Costo1 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * rstArticulo!Costo1 * Val(WVector))
            Rem     Case 2
            Rem         Costo3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
            Rem         WCosto = (Cantidad * Costo3 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * Costo3 * Val(WVector))
            Rem     Case 3
            Rem         Costo4 = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
            Rem         If Costo4 = 0 Then
            Rem             Costo4 = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
            Rem         End If
            Rem         WCosto = (Cantidad * Costo4 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * Costo4 * Val(WVector))
            Rem     Case 4
            Rem         Costo4 = IIf(IsNull(rstArticulo!UltimoFob), "0", rstArticulo!UltimoFob)
            Rem         If Costo4 = 0 Then
            Rem             Costo4 = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
            Rem         End If
            Rem         WCosto = (Cantidad * Costo4 * Val(WVector))
            Rem         Costo = Costo + (Cantidad * Costo4 * Val(WVector))
            Rem     Case Else
            Rem         WCosto = 0
            Rem         Costo = 0
            Rem End Select
            
            Select Case TipoListado.ListIndex
                Case 0
                    ZZCosto = rstArticulo!Costo2
                    ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                    If ZTipoCosto = 1 Then
                        WDescriTipo = "Estimado"
                            Else
                        WDescriTipo = ""
                    End If
                    rstArticulo.Close
                Case 1
                    ZZCosto = rstArticulo!Costo1

                    Costo1 = rstArticulo!Costo1
                    WWCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                    ZCosto1 = IIf(IsNull(rstArticulo!ZCosto1), "0", rstArticulo!ZCosto1)
                    ZZOrdenI = IIf(IsNull(rstArticulo!OrdenI), "", rstArticulo!OrdenI)
                    ZZOrdenII = IIf(IsNull(rstArticulo!OrdenII), "", rstArticulo!OrdenII)
                    ZZOrdenIII = IIf(IsNull(rstArticulo!OrdenIII), "", rstArticulo!OrdenIII)
                    ZZPtaOrdenI = IIf(IsNull(rstArticulo!PtaOrdenI), "0", rstArticulo!PtaOrdenI)
                    ZZPtaOrdenII = IIf(IsNull(rstArticulo!PtaOrdenII), "0", rstArticulo!PtaOrdenII)
                    ZZPtaOrdenIII = IIf(IsNull(rstArticulo!PtaOrdenIII), "0", rstArticulo!PtaOrdenIII)
                        
                    ZZFechaOrdenI = ""
                    ZZFechaOrdenII = ""
                    ZZFechaOrdenIII = ""
                    
                    ZZMoneda = ""
                    
                    rstArticulo.Close
                    
                    XEmpresa = WEmpresa

                    If ZZPtaOrdenI <> 0 And ZZOrdenI <> 0 Then
                    
                        Select Case ZZPtaOrdenI
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
                            Case 11
                                WEmpresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenI) + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZFechaOrdenI = rstOrden!Fecha
                            Select Case rstOrden!Moneda
                                Case 0
                                    ZZMoneda = "U$S"
                                Case 1
                                    ZZMoneda = "$"
                                Case Else
                                    ZZMoneda = "Eur"
                            End Select
                            rstOrden.Close
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                    
                    
                    If ZZPtaOrdenII <> 0 And ZZOrdenII <> 0 Then
                        
                        Select Case ZZPtaOrdenII
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
                            Case 11
                                WEmpresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenII) + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZFechaOrdenII = rstOrden!Fecha
                            rstOrden.Close
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                    
                    If ZZPtaOrdenIII <> 0 And ZZOrdenIII <> 0 Then
                        
                        Select Case ZZPtaOrdenIII
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
                            Case 11
                                WEmpresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenIII) + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZFechaOrdenIII = rstOrden!Fecha
                            rstOrden.Close
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                    
                    Rem DADA
                    Rem spCambio = "ConsultaCambio " + "'" + ZZFecha + "'"
                    Rem Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstCambio.RecordCount > 0 Then
                    Rem     ZZParidad = rstCambio!Cambio
                    Rem     rstCambio.Close
                    Rem End If
                    
                    Rem If ZZParidad <> 0 Then
                    Rem     ZZCostoPesos = WPrecio * ZZParidad
                    Rem End If
                            
                    If ZZFechaOrdenI <> "" Then
                        WFechaOrdI = Right$(ZZFechaOrdenI, 4) + Mid$(ZZFechaOrdenI, 4, 2) + Left$(ZZFechaOrdenI, 2)
                            Else
                        WFechaOrdI = ""
                    End If
                    If ZZFechaOrdenII <> "" Then
                        WFechaOrdII = Right$(ZZFechaOrdenII, 4) + Mid$(ZZFechaOrdenII, 4, 2) + Left$(ZZFechaOrdenII, 2)
                            Else
                        WFechaOrdII = ""
                    End If
                    If ZZFechaOrdenIII <> "" Then
                        WFechaOrdIII = Right$(ZZFechaOrdenIII, 4) + Mid$(ZZFechaOrdenIII, 4, 2) + Left$(ZZFechaOrdenIII, 2)
                            Else
                        WFechaOrdIII = ""
                    End If
                    
                    If WFechaOrdI <> "" And WFechaOrdI > WFechaOrdII And WFechaOrdI > WFechaOrdIII Then
                        ZZCosto = Costo1
                    End If
                    
                    If WFechaOrdII <> "" And WFechaOrdII > WFechaOrdI And WFechaOrdII > WFechaOrdIII Then
                    
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                         spCambios = "ConsultaCambio  " + "'" + ZZFechaOrdenII + "'"
                         Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                         If rstCambios.RecordCount > 0 Then
                             ZZZZParidad = rstCambios!Cambio
                             rstCambios.Close
                             If ZZZZParidad <> 0 Then
                                 ZZCosto1Dol = WWCosto1 / ZZZZParidad
                             End If
                         End If
                        Call Conecta_Empresa
                    
                        ZZCosto = ZZCosto1Dol
                    End If
                    
                    If WFechaOrdIII <> "" And WFechaOrdIII > WFechaOrdI And WFechaOrdIII > WFechaOrdII Then
                        ZZCosto = ZCosto1
                    End If
                    WDescriTipo = ""
                                            
                    Rem WEmpresa = "0001"
                    Rem txtOdbc = "Empresa01"
                    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    Rem spCambios = "ConsultaCambio  " + "'" + FechaOrdenII.Text + "'"
                    Rem Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstCambios.RecordCount > 0 Then
                    Rem     ZZParidad = rstCambios!Cambio
                    Rem     rstCambios.Close
                    Rem     If ZZParidad <> 0 Then
                    Rem         ZZCosto1Dol = Val(WCosto1.Text) / ZZParidad
                    Rem         WCosto1Dol.Text = Str$(ZZCosto1Dol)
                    Rem         WCosto1Dol.Text = Pusing("###,###.##", WCosto1Dol.Text)
                    Rem     End If
                    Rem End If
                    
                Case 2
                    ZZCosto = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                    WDescriTipo = ""
                    rstArticulo.Close
                Case 3
                    ZZCosto = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
                    WDescriTipo = "Reposicion"
                    If ZZCosto = 0 Then
                        ZZCosto = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
                        ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                        If ZTipoCosto = 1 Then
                            WDescriTipo = "Estimado"
                                Else
                            WDescriTipo = ""
                        End If
                    End If
                    rstArticulo.Close
                Case 4
                    ZZCosto = IIf(IsNull(rstArticulo!UltimoFob), "0", rstArticulo!UltimoFob)
                    WDescriTipo = ""
                    If ZZCosto = 0 Then
                        ZZCosto = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
                        ZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
                        If ZTipoCosto = 1 Then
                            WDescriTipo = "Estimado"
                                Else
                            WDescriTipo = ""
                        End If
                    End If
                    rstArticulo.Close
                Case Else
                    ZZCosto = 0
                    rstArticulo.Close
            End Select
            
            WCosto = (Cantidad * ZZCosto * Val(WVector))
            Costo = Costo + WCosto
            
        End If
    Next Da
    
End Sub

