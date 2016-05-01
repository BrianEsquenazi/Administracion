VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCostoOrden 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Calculo de Costo de Importacion "
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.CommandButton AvisoError 
      Caption         =   "No se puede trabajar con este modulo. El sistema se encuentra sin conexion con las demas plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2400
      Picture         =   "costoorden.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2055
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   4815
      Begin VB.TextBox Carpeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCarpeta.rpt"
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
      ItemData        =   "costoorden.frx":0742
      Left            =   840
      List            =   "costoorden.frx":0749
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCostoOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim Vector(100, 20) As String
Dim Gastos(100, 10) As String

Dim rstCarpeta As Recordset
Dim spCarpeta As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstCambios As Recordset
Dim spCambios As String

Dim WArancel As String
Dim WCarpeta As String
Dim WCosto As Double
Dim WLeyenda As Integer
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
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
    CargaEmpresa(10, 1) = "0010"
    CargaEmpresa(10, 2) = "Empresa10"
    CargaEmpresa(11, 1) = "0011"
    CargaEmpresa(11, 2) = "Empresa11"
                    
    For Cicla = 1 To 11
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    On Error GoTo 0
    If WSalidaError = "N" Then Exit Sub

    WCarpeta = Carpeta.Text
    Call Ceros(WCarpeta, 6)

    spCarpeta = "BorrarCarpeta"
    Set rstCarpeta = db.OpenRecordset(spCarpeta, dbOpenSnapshot, dbSQLPassThrough)

    WMarca = ""
    Renglon = 0
    Erase Gastos
    ImpoGastos = 0
    ImpoSeguro = 0
    ImpoFlete = 0
    
    spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovgas.RecordCount > 0 Then
    
        With rstMovgas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WMarca = IIf(IsNull(rstMovgas!Marca), "", rstMovgas!Marca)
                    If rstMovgas!Concepto <> 10 Then
                        WArancel = Str$(rstMovgas!Derechos)
                        WOrden = Str$(rstMovgas!Orden)
                        WEmpresaOtro = rstMovgas!Empresa
                        Select Case rstMovgas!Concepto
                            Case 2
                                ImpoSeguro = ImpoSeguro + rstMovgas!Importe
                            Case 4, 5
                                ImpoFlete = ImpoFlete + rstMovgas!Importe
                            Case Else
                                Renglon = Renglon + 1
                                Gastos(Renglon, 1) = Str$(rstMovgas!Concepto)
                                Gastos(Renglon, 2) = Str$(rstMovgas!Importe)
                                ImpoGastos = ImpoGastos + rstMovgas!Importe
                        End Select
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovgas.Close
                
    End If
    
    WGastos = Renglon
    ImpoArancel = 0
    WTotal = 0
    
    Renglon = 0
    Erase Vector
    
    EmpresaAnterior = WEmpresa
    Select Case WEmpresaOtro
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
    
    ZCoeParidad = 1
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Orden = " + "'" + Str$(WOrden) + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
               
        ZNroInforme = rstInforme!Informe
        ZFechaInforme = rstInforme!Fecha
        rstInforme.Close
        
        ZZZEmpresa = WEmpresa
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        spCambios = "ConsultaCambio  " + "'" + ZFechaInforme + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            ZParidad = rstCambios!Cambio
            ZParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
            If ZParidadII <> 0 And ZParidad <> 0 Then
                ZCoeParidad = ZParidadII / ZParidad
            End If
            rstCambios.Close
        End If
        
        Select Case ZZZEmpresa
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
        
    End If
  
    
    WTotalImpo = 0
    WTotalPeso = 0
    
    spOrden = "ListaOrden " + "'" + WOrden + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                    
                    WMoneda = rstOrden!Moneda
                    ZZPrecio = rstOrden!Precio
                    If WMoneda = 2 Then
                        ZZPrecio = ZZPrecio * ZCoeParidad
                    End If
                    
                    Vector(Renglon, 1) = Carpeta.Text
                    Vector(Renglon, 2) = rstOrden!Articulo
                    Vector(Renglon, 3) = Str$(rstOrden!Cantidad)
                    Vector(Renglon, 4) = Str$(ZZPrecio)
                    Vector(Renglon, 5) = Str$(rstOrden!Cantidad * ZZPrecio)
                  
                    If IsNull(rstOrden!Derechos) Then
                        Vector(Renglon, 6) = ""
                             Else
                        Vector(Renglon, 6) = Str$(rstOrden!Derechos)
                    End If
                    
                    Vector(Renglon, 7) = ""
                    Vector(Renglon, 8) = ""
                    Vector(Renglon, 9) = ""
                    Vector(Renglon, 10) = ""
                    Vector(Renglon, 11) = WCarpeta + "01"
                    
                    WTotalImpo = WTotalImpo + (rstOrden!Cantidad * ZZPrecio)
                    WTotalPeso = WTotalPeso + rstOrden!Cantidad
                    
                    WLeyenda = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Ciclo = 1 To Renglon
        WSeguro = 0
        WFlete = 0
        If ImpoSeguro <> 0 Then
            If WTotalImpo <> 0 Then
                WSeguro = ((Val(Vector(Ciclo, 5)) / WTotalImpo) * ImpoSeguro) / Val(Vector(Ciclo, 3))
            End If
        End If
        If ImpoFlete <> 0 Then
            If WTotalPeso <> 0 Then
                WFlete = ((Val(Vector(Ciclo, 3)) / WTotalPeso) * ImpoFlete) / Val(Vector(Ciclo, 3))
            End If
        End If
        aa = Vector(Ciclo, 2)
        WCosto = Val(Vector(Ciclo, 4)) + WSeguro + WFlete
        Call Redondeo(WCosto)
        Vector(Ciclo, 4) = Str$(WCosto)
    Next Ciclo
    
    WTotal = 0
    
    For Ciclo = 1 To Renglon
        WArticulo = Vector(Ciclo, 2)
        
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Vector(Ciclo, 5) = Str$(Val(Vector(Ciclo, 3)) * Val(Vector(Ciclo, 4)))
            XArancel = Val(Vector(Ciclo, 5)) * Val(Vector(Ciclo, 6)) / 100
            
            Vector(Ciclo, 7) = Str$(Val(Vector(Ciclo, 5)) + XArancel)
            WTotal = WTotal + Val(Vector(Ciclo, 5))
            ImpoArancel = ImpoArancel + XArancel
            rstArticulo.Close
        End If
    Next Ciclo
    
    For Ciclo = 1 To Renglon
        If WTotal <> 0 Then
            Vector(Ciclo, 8) = Str$((Val(Vector(Ciclo, 5)) / WTotal) * ImpoGastos)
        End If
        If Val(Vector(Ciclo, 3)) <> 0 Then
            Vector(Ciclo, 9) = Str$((Val(Vector(Ciclo, 7)) + Val(Vector(Ciclo, 8))) / Val(Vector(Ciclo, 3)))
        End If
    Next Ciclo
    
    If WTotal <> 0 Then
        WCoeficiente = Str$((ImpoGastos + ImpoArancel) / (WTotal / 100))
        WCoeficiente = Str$(1 + (Val(WCoeficiente) / 100))
            Else
        WCoeficiente = ""
    End If
    
    Select Case Val(EmpresaAnterior)
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
    
    XLeyenda = Str$(WLeyenda)
    
    For Ciclo = 1 To Renglon
        XParam = "'" + Vector(Ciclo, 1) + "','" _
                    + Vector(Ciclo, 2) + "','" _
                    + Vector(Ciclo, 3) + "','" _
                    + Vector(Ciclo, 4) + "','" _
                    + Vector(Ciclo, 5) + "','" _
                    + Vector(Ciclo, 6) + "','" _
                    + Vector(Ciclo, 7) + "','" _
                    + Vector(Ciclo, 8) + "','" _
                    + Vector(Ciclo, 9) + "','" _
                    + WCoeficiente + "','" _
                    + Vector(Ciclo, 11) + "','" _
                    + XLeyenda + "'"
        
        spCarpeta = "AltaCarpeta " + XParam
        Set rstCarpeta = db.OpenRecordset(spCarpeta, dbOpenSnapshot, dbSQLPassThrough)
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Calculo de Costo de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Carpeta.Carpeta} in " + Carpeta.Text + " to " + Carpeta.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Carpeta.Carpeta, Carpeta.Articulo, Carpeta.Cantidad, Carpeta.CostoFlete, Carpeta.Importe, Carpeta.Arancel, Carpeta.Costo, Carpeta.Gastos, Carpeta.Precio, Carpeta.Coeficiente, " _
                    + "Articulo.Descripcion, " _
                    + "Movgas.Fecha, Movgas.Orden, Movgas.Proveedor, Movgas.Origen, Movgas.Moneda, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Carpeta Carpeta, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Movgas Movgas, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Carpeta.Articulo = Articulo.Codigo AND " _
                    + "Carpeta.Clave = Movgas.Clave AND " _
                    + "Movgas.Proveedor = Proveedor.Proveedor AND " _
                    + "Carpeta.Carpeta >= 0 AND Carpeta.Carpeta <= 999999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()

    If WMarca <> "X" Then
        Listado.ReportFileName = "WcarpetaFalta.rpt"
            Else
        Listado.ReportFileName = "Wcarpeta.rpt"
    End If

    Listado.Action = 1
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Carpeta.SetFocus
    PrgCostoOrden.Hide
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

Private Sub Carpeta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Carpeta.SetFocus
    End If
End Sub

Sub Form_Load()
    Carpeta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub







