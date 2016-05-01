VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgActuaOrden 
   AutoRedraw      =   -1  'True
   Caption         =   "Actualizacion de Costos de Importacion"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Pendiente 
      Caption         =   "Carpetas Pendientes de Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5760
      TabIndex        =   10
      Top             =   480
      Width           =   1695
   End
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
      Left            =   2880
      Picture         =   "ActuaOrden.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
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
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
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
      ItemData        =   "ActuaOrden.frx":0742
      Left            =   840
      List            =   "ActuaOrden.frx":0749
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
Attribute VB_Name = "PrgActuaOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim XParamII As String
Dim Vector(100, 20) As String
Dim Gastos(100, 10) As String
Dim ZZOrden(10000, 3) As String
Dim ZZOrdenII(1000, 7) As String

Dim WOrden As String
Dim WEmpresaOtro As String
Dim WFechaOrden As String

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
Dim XCosto1 As Double
Dim XCosto2 As Double
Dim XCosto3 As Double
Dim ZCostoCompara As Double
Dim WLeyenda As Integer
Dim CargaEmpresa(12, 2) As String
Dim ZZTipoImpo As Integer
Dim ZZUltimo(100) As String
Dim ZZArticulo(100) As String
Dim ZZZCosto2 As String
Dim ZZZTipoCosto As String

Dim WWEmpresa As String
Dim WWOrden As String
Dim WWArticulo As String
Dim WWCarpeta As String
Dim WWCosto As String


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

    spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovgas.RecordCount > 0 Then
        WMarca = IIf(IsNull(rstMovgas!Marca), "", rstMovgas!Marca)
        Rem WMarca = ""
        rstMovgas.Close
        If WMarca = "X" Then
            m$ = "La Carpeta ya fue actualizada"
            a% = MsgBox(m$, 0, "Calculo de Costo de Importacion")
            Carpeta.SetFocus
            Exit Sub
        End If
            Else
        m$ = "No Existe la carpeta solicitada"
        a% = MsgBox(m$, 0, "Calculo de Costo de Importacion")
        Carpeta.SetFocus
        Exit Sub
    End If

    WCarpeta = Carpeta.Text
    Call Ceros(WCarpeta, 6)
    
    ZTipoImpo = 0
    ZEmpresa = 0
    ZOrden = 0
    
    spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovgas.RecordCount > 0 Then
        ZEmpresa = IIf(IsNull(rstMovgas!Empresa), "", rstMovgas!Empresa)
        ZOrden = IIf(IsNull(rstMovgas!Orden), "", rstMovgas!Orden)
        rstMovgas.Close
    End If
    
    EmpresaAnterior = WEmpresa
    
    Select Case ZEmpresa
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
    
    ZZLugar = 0
        
    spOrden = "ListaOrden " + "'" + Str$(ZOrden) + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZTipoImpo = IIf(IsNull(rstOrden!TipoImpo), "0", rstOrden!TipoImpo)
                    ZZFechaOrden = rstOrden!Fecha
                    
                    ZZLugar = ZZLugar + 1
                    ZZArticulo(ZZLugar) = rstOrden!Articulo
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    ZCoeParidad = 1
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Orden = " + "'" + Str$(ZOrden) + "'"
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
    
    
    
    If ZTipoImpo = 3 Then
        XParam = "'" + Carpeta.Text + "','" _
                     + "X" + "'"
        spMovgas = "ActualizaMovgasMarca " + XParam
        Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
        Call Cancela_click
        Exit Sub
    End If
    
    
    
    
    ZZGrabaII = "N"
        
    For ZZCiclo = 1 To ZZLugar
    
        ZZArti = ZZArticulo(ZZCiclo)
    
        XEmpresa = WEmpresa
        Erase ZZUltimo
        
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
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Orden"
                ZSql = ZSql + " Where Articulo = " + "'" + ZZArti + "'"
                ZSql = ZSql + " and Saldo <> 0"
                ZSql = ZSql + " Order by Orden.FechaOrd"
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    With rstOrden
                        .MoveLast
                        ZZFechaOrdenII = rstOrden!Fecha
                    End With
                    rstOrden.Close
                        Else
                    ZZFechaOrdenII = "00/00/0000"
                End If
                
                ZZUltimo(Cicla) = ZZFechaOrdenII
            
            End If
        Next Cicla
    
        Call Conecta_Empresa
        
        WCompara = Right$(ZZFechaOrden, 4) + Mid$(ZZFechaOrden, 4, 2) + Left$(ZZFechaOrden, 2)
        WCompara1 = Right$(ZZUltimo(1), 4) + Mid$(ZZUltimo(1), 4, 2) + Left$(ZZUltimo(1), 2)
        WCompara2 = Right$(ZZUltimo(2), 4) + Mid$(ZZUltimo(2), 4, 2) + Left$(ZZUltimo(2), 2)
        WCompara3 = Right$(ZZUltimo(3), 4) + Mid$(ZZUltimo(3), 4, 2) + Left$(ZZUltimo(3), 2)
        WCompara4 = Right$(ZZUltimo(4), 4) + Mid$(ZZUltimo(4), 4, 2) + Left$(ZZUltimo(4), 2)
        WCompara5 = Right$(ZZUltimo(5), 4) + Mid$(ZZUltimo(5), 4, 2) + Left$(ZZUltimo(5), 2)
        WCompara6 = Right$(ZZUltimo(6), 4) + Mid$(ZZUltimo(6), 4, 2) + Left$(ZZUltimo(6), 2)
        WCompara7 = Right$(ZZUltimo(7), 4) + Mid$(ZZUltimo(7), 4, 2) + Left$(ZZUltimo(7), 2)
        WCompara8 = Right$(ZZUltimo(8), 4) + Mid$(ZZUltimo(8), 4, 2) + Left$(ZZUltimo(8), 2)
        WCompara9 = Right$(ZZUltimo(9), 4) + Mid$(ZZUltimo(9), 4, 2) + Left$(ZZUltimo(9), 2)
        
        
        If WCompara1 < WCompara Then
            If WCompara2 < WCompara Then
                If WCompara3 < WCompara Then
                    If WCompara4 < WCompara Then
                        If WCompara5 < WCompara Then
                            If WCompara6 < WCompara Then
                                If WCompara7 < WCompara Then
                                    If WCompara8 < WCompara Then
                                        If WCompara9 < WCompara Then
                                            ZZGrabaII = "S"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    Next ZZCiclo
    
    
    
    

    spCarpeta = "BorrarCarpeta"
    Set rstCarpeta = db.OpenRecordset(spCarpeta, dbOpenSnapshot, dbSQLPassThrough)

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
                    WFechaOrden = rstOrden!Fecha
                    ZZPrecio = rstOrden!Precio
                    If WMoneda = 2 Then
                        ZZPrecio = ZZPrecio * ZCoeParidad
                    End If
                    
                    Vector(Renglon, 1) = Carpeta.Text
                    Vector(Renglon, 2) = rstOrden!Articulo
                    Vector(Renglon, 3) = Str$(rstOrden!Cantidad)
                    Vector(Renglon, 4) = Str$(ZZPrecio)
                    Vector(Renglon, 5) = Str$(rstOrden!Cantidad * ZZPrecio)
                    Vector(Renglon, 6) = Str$(rstOrden!Derechos)
                    Vector(Renglon, 7) = ""
                    Vector(Renglon, 8) = ""
                    Vector(Renglon, 9) = ""
                    Vector(Renglon, 10) = ""
                    Vector(Renglon, 11) = WCarpeta + "01"
                    Vector(Renglon, 12) = Str$(ZZPrecio)
                    
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
            Rem Vector(Ciclo, 4) = Str$(rstArticulo!Flete)
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

    For Ciclo = 1 To Renglon
    
        WArticulo = Vector(Ciclo, 2)
        
        XCosto1 = Val(Vector(Ciclo, 9))
        XCosto2 = Val(Vector(Ciclo, 9)) * 1.03
        XCosto3 = Val(Vector(Ciclo, 9))
        Aaa = WArticulo
        CostoImpo = Val(Vector(Ciclo, 4))
        If CostoImpo <> 0 Then
            ZZCoeficiente = Str$(XCosto1 / CostoImpo)
                Else
            ZZCoeficiente = "0"
        End If
        
        Call Redondeo(XCosto1)
        Call Redondeo(XCosto2)
        Call Redondeo(XCosto3)
        
        WCosto1 = Str$(XCosto1)
        WCosto2 = Str$(XCosto2)
        WCosto3 = Str$(XCosto3)
        WFlete = Vector(Ciclo, 4)
        
        If WLeyenda > 0 Then
            XLeyenda = Str$(WLeyenda - 1)
                Else
            XLeyenda = "0"
        End If
        
        XParam = "'" + WArticulo + "','" _
                    + WCosto1 + "','" _
                    + WCosto2 + "','" _
                    + WCosto3 + "','" _
                    + WFlete + "'"
                    
        XParamII = "'" + WArticulo + "','" _
                       + XLeyenda + "'"
                       
                       
        ZLaudo = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " and Laudo.Orden = " + "'" + Str$(ZOrden) + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            ZLaudo = Str$(rstLaudo!Laudo)
            rstLaudo.Close
        End If
                       
                    
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZCostoCompara = rstArticulo!Costo2
            Call Redondeo(ZCostoCompara)
            ZZZCosto2 = Str$(ZCostoCompara)
            ZZZTipoCosto = IIf(IsNull(rstArticulo!TipoCosto), "0", rstArticulo!TipoCosto)
            rstArticulo.Close
        End If
            
        
        ZGraba = "N"
        If XCosto2 <> ZCostoCompara Then
            ZGraba = "S"
        End If
        
        
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
       
       
        Rem TOTO
        Rem TOTO
        Rem TOTO
       
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
            
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
            
        WEmpresa = "0003"
        txtOdbc = "Empresa03"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
            
        WEmpresa = "0004"
        txtOdbc = "Empresa04"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
            
        WEmpresa = "0005"
        txtOdbc = "Empresa05"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
            
        WEmpresa = "0006"
        txtOdbc = "Empresa06"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
            
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
            
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
            
        WEmpresa = "0009"
        txtOdbc = "Empresa09"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
            
        WEmpresa = "0010"
        txtOdbc = "Empresa10"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
            
        WEmpresa = "0011"
        txtOdbc = "Empresa11"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        spArticulo = "ModificaArticuloCostoImportacion " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        spArticulo = "ModificaArticuloLeyenda " + XParamII
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        If ZGraba = "S" And ZZGrabaII = "S" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo2Anterior = " + "'" + Str$(ZCostoCompara) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " OrdenI = " + "'" + WOrden + "',"
        ZSql = ZSql + " PtaOrdenI = " + "'" + WEmpresaOtro + "',"
        ZSql = ZSql + " Costo6 = " + "'" + "0" + "',"
        ZSql = ZSql + " TipoCosto = " + "'" + "0" + "',"
        ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
        ZSql = ZSql + " UltimoFob = " + "'" + WFlete + "',"
        ZSql = ZSql + " Factor = " + "'" + ZZCoeficiente + "',"
        ZSql = ZSql + " UltimoCosto = " + "'" + WCosto1 + "',"
        ZSql = ZSql + " UltimoTipo = " + "'" + XLeyenda + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem If ZZGrabaII = "N" Then
        If ZZGrabaII = "DADA" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " TipoCosto = " + "'" + ZZZTipoCosto + "',"
            ZSql = ZSql + " Costo2 = " + "'" + ZZZCosto2 + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
        WWEmpresa = ZEmpresa
        WWOrden = ZOrden
        WWArticulo = WArticulo
        WWCarpeta = Carpeta.Text
        WWCosto = WCosto1
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWOrden, 6)
        
        WWClave = WWEmpresa + WWOrden + WWArticulo
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CostoPartida ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Laudo ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWEmpresa + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWArticulo + "',"
        ZSql = ZSql + "'" + WWCarpeta + "',"
        ZSql = ZSql + "'" + WWLaudo + "',"
        ZSql = ZSql + "'" + WWCosto + "')"
        
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        
    Next Ciclo
    
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

    XParam = "'" + Carpeta.Text + "','" _
                 + "X" + "'"
        
    spMovgas = "ActualizaMovgasMarca " + XParam
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Cancela_click
    
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
    PrgActuaOrden.Hide
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
    Frame2.Visible = True
End Sub







Private Sub Pendiente_Click()

    Rem On Error GoTo WError
    
    Erase ZZOrdenII
    ZZLugarII = 0
    
    ZSql = ""
    ZSql = ZSql + "DELETE ListaOrdPen"
    spListaOrdPen = ZSql
    Set rstListaOrdPen = db.OpenRecordset(spListaOrdPen, dbOpenSnapshot, dbSQLPassThrough)
    
    
    XEmpresa = WEmpresa
    
    For CicloEmpresa = 1 To 7
    
        Select Case CicloEmpresa
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Orden SET "
        ZSql = ZSql + " MarcaActualiza = " + "'" + "" + "'"
        ZSql = ZSql + " Where MarcaActualiza IS NULL"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        Erase ZZOrden
        ZZLugar = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.MarcaActualiza = " + "'" + "" + "'"
        ZSql = ZSql + " and Orden.FechaOrd >= " + "'" + "20140101" + "'"
        ZSql = ZSql + " and Orden.Renglon = 1"
        ZSql = ZSql + " and Orden.Tipo = 1"
        ZSql = ZSql + " Order by Orden.Clave"
        
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
        
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZZLugar = ZZLugar + 1
                        
                        ZZOrden(ZZLugar, 1) = rstOrden!Orden
                        ZZOrden(ZZLugar, 2) = rstOrden!Carpeta
                            
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstOrden.Close
        End If
        

        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
        For Ciclo = 1 To ZZLugar
        
            ZZZZOrden = ZZOrden(Ciclo, 1)
            ZZZZCarpeta = ZZOrden(Ciclo, 2)
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM MovGas"
            ZSql = ZSql + " Where MovGas.Carpeta = " + "'" + ZZZZCarpeta + "'"
            
            spMovgas = ZSql
            Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovgas.RecordCount > 0 Then
                If Val(rstMovgas!Orden) <> Val(ZZZZOrden) Then
                    ZZZZMarca = ""
                        Else
                    ZZZZMarca = IIf(IsNull(rstMovgas!Marca), "", rstMovgas!Marca)
                End If
                
                ZZOrden(Ciclo, 3) = ZZZZMarca
                rstMovgas.Close
                
                    Else
                    
                ZZOrden(Ciclo, 3) = ""
            
            End If
            
        Next Ciclo
    
        Select Case CicloEmpresa
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        For Ciclo = 1 To ZZLugar
        
            ZZZZOrden = ZZOrden(Ciclo, 1)
            ZZZZCarpeta = ZZOrden(Ciclo, 2)
            ZZZZMarca = ZZOrden(Ciclo, 3)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Orden SET "
            ZSql = ZSql + " MarcaActualiza = " + "'" + ZZZZMarca + "'"
            ZSql = ZSql + " Where Orden = " + "'" + ZZZZOrden + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
            If Trim(ZZZZMarca) = "" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Orden"
                ZSql = ZSql + " Where Orden.Orden = " + "'" + ZZZZOrden + "'"
                spOrden = ZSql
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    ZZZZFecha = rstOrden!Fecha
                    ZZZZProveedor = rstOrden!Proveedor
                    ZZZZTipoImpo = rstOrden!TipoImpo
                    ZZZZRecibida = rstOrden!Recibida
                    
                    rstOrden.Close
                    
                    If ZZZZRecibida <> 0 Then
                        
                        ZZLugarII = ZZLugarII + 1
                        ZZOrdenII(ZZLugarII, 1) = ZZZZOrden
                        ZZOrdenII(ZZLugarII, 2) = ZZZZCarpeta
                        ZZOrdenII(ZZLugarII, 3) = ZZZZFecha
                        ZZOrdenII(ZZLugarII, 4) = ZZZZProveedor
                        ZZOrdenII(ZZLugarII, 5) = ZZZZTipoImpo
                        ZZOrdenII(ZZLugarII, 6) = WEmpresa
                        ZZOrdenII(ZZLugarII, 7) = ZZZZRecibida
                    
                    End If
                    
                End If
                
            End If
            
        Next Ciclo

    
    Next CicloEmpresa
    
    
    Call Conecta_Empresa
    
    For Ciclo = 1 To ZZLugarII
    
        ZZZZOrden = ZZOrdenII(Ciclo, 1)
        ZZZZCarpeta = ZZOrdenII(Ciclo, 2)
        ZZZZFecha = ZZOrdenII(Ciclo, 3)
        ZZZZProveedor = ZZOrdenII(Ciclo, 4)
        ZZZZTipoImpo = ZZOrdenII(Ciclo, 5)
        ZZZZEmpresa = ZZOrdenII(Ciclo, 6)
        ZZZZRecibida = ZZOrdenII(Ciclo, 7)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZZZProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            ZZZZDesProveedor = rstProveedor!Nombre
            rstProveedor.Close
        End If
        
        ZZZZDesTipoImpo = ""
        Select Case Val(ZZZZTipoImpo)
            Case 1
                ZZZZDesTipoImpo = "Maritimo"
            Case 2
                ZZZZDesTipoImpo = "Terrestre"
            Case 3
                ZZZZDesTipoImpo = "Aereo"
            Case Else
        End Select
        
        ZZZZDesEmpresa = ""
        Select Case Val(ZZZZEmpresa)
            Case 1
                ZZZZDesEmpresa = "SI"
            Case 2
                ZZZZDesEmpresa = "PI"
            Case 3
                ZZZZDesEmpresa = "SII"
            Case 4
                ZZZZDesEmpresa = "PII"
            Case 5
                ZZZZDesEmpresa = "SIII"
            Case 6
                ZZZZDesEmpresa = "SIV"
            Case 7
                ZZZZDesEmpresa = "SV"
            Case 8
                ZZZZDesEmpresa = "PIII"
            Case 9
                ZZZZDesEmpresa = "PV"
            Case 10
                ZZZZDesEmpresa = "SVI"
            Case 11
                ZZZZDesEmpresa = "SVII"
            Case Else
        End Select
        

        ZSql = ""
        ZSql = ZSql + "INSERT INTO ListaOrdPen ("
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "DesProveedor ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "DesEmpresa ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "TipoImpo ,"
        ZSql = ZSql + "DesTipoImpo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZZZOrden + "',"
        ZSql = ZSql + "'" + ZZZZCarpeta + "',"
        ZSql = ZSql + "'" + ZZZZProveedor + "',"
        ZSql = ZSql + "'" + ZZZZDesProveedor + "',"
        ZSql = ZSql + "'" + ZZZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZZZDesEmpresa + "',"
        ZSql = ZSql + "'" + ZZZZFecha + "',"
        ZSql = ZSql + "'" + ZZZZTipoImpo + "',"
        ZSql = ZSql + "'" + ZZZZDesTipoImpo + "')"

        spListaOrdPen = ZSql
        Set rstListaOrdPen = db.OpenRecordset(spListaOrdPen, dbOpenSnapshot, dbSQLPassThrough)
    

    Next Ciclo
    
    
    
    
    
    
    
    Rem by nan
    Listado.WindowTitle = "Listado de Carpetas Pendientes de Actualizar"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{ListaOrdpen.Orden} in 0 to 999999"
    
    Listado.GroupSelectionFormula = Uno
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
   
    Listado.SQLQuery = "SELECT ListaOrdpen.Orden, ListaOrdpen.Carpeta, ListaOrdpen.Proveedor, ListaOrdpen.DesProveedor, ListaOrdpen.Fecha, ListaOrdpen.TipoImpo, ListaOrdpen.DesTipoImpo, ListaOrdpen.Empresa, ListaOrdpen.DesEmpresa " _
            + "From " _
            + DSQ + ".dbo.ListaOrdpen ListaOrdpen " _
            + "Where " _
            + "ListaOrdpen.Orden >= 0 AND " _
            + "ListaOrdpen.Orden <= 999999"
    
    
    If WEmpresa = "0001" Then
    Listado.ReportFileName = "ListaOrdpen.rpt"
        Else
    
     Listado.ReportFileName = "ListaOrdpenpelli.rpt"
    End If
    
    Listado.Destination = 0
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next

End Sub





Private Sub PendienteAnterior_Click()

    Rem On Error GoTo WError
    
    
    
    XEmpresa = WEmpresa
    
    For CicloEmpresa = 1 To 7
    
        Select Case CicloEmpresa
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Orden SET "
        ZSql = ZSql + " MarcaActualiza = " + "'" + "" + "'"
        Rem ZSql = ZSql + " Where MarcaActualiza IS NULL"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        Erase ZZOrden
        ZZLugar = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.MarcaActualiza = " + "'" + "" + "'"
        ZSql = ZSql + " and Orden.FechaOrd >= " + "'" + "20140101" + "'"
        ZSql = ZSql + " and Orden.Renglon = 1"
        ZSql = ZSql + " and Orden.Tipo = 1"
        ZSql = ZSql + " Order by Orden.Clave"
        
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
        
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZZLugar = ZZLugar + 1
                        
                        ZZOrden(ZZLugar, 1) = rstOrden!Orden
                        ZZOrden(ZZLugar, 2) = rstOrden!Carpeta
                            
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstOrden.Close
        End If
        

        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
        For Ciclo = 1 To ZZLugar
        
            ZZZZOrden = ZZOrden(Ciclo, 1)
            ZZZZCarpeta = ZZOrden(Ciclo, 2)
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM MovGas"
            ZSql = ZSql + " Where MovGas.Carpeta = " + "'" + ZZZZCarpeta + "'"
            
            spMovgas = ZSql
            Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovgas.RecordCount > 0 Then
                If Val(rstMovgas!Orden) <> Val(ZZZZOrden) Then
                    ZZZZMarca = ""
                        Else
                    ZZZZMarca = IIf(IsNull(rstMovgas!Marca), "", rstMovgas!Marca)
                End If
                
                ZZOrden(Ciclo, 3) = ZZZZMarca
                rstMovgas.Close
                
                    Else
                    
                ZZOrden(Ciclo, 3) = ""
            
            End If
            
        Next Ciclo
    
        Select Case CicloEmpresa
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        For Ciclo = 1 To ZZLugar
        
            ZZZZOrden = ZZOrden(Ciclo, 1)
            ZZZZCarpeta = ZZOrden(Ciclo, 2)
            ZZZZMarca = ZZOrden(Ciclo, 3)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Orden SET "
            ZSql = ZSql + " MarcaActualiza = " + "'" + ZZZZMarca + "'"
            ZSql = ZSql + " Where Orden = " + "'" + ZZZZOrden + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Ciclo

    
    Next CicloEmpresa
    
    
    Call Conecta_Empresa

    
    
    
    
    
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE MovGas SET "
    ZSql = ZSql + " TipoImpo = " + "'" + "0" + "'"
    spMovgas = ZSql
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase ZZOrden
    ZZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Movgas"
    ZSql = ZSql + " Where Movgas.Marca <> " + "'" + "X" + "'"
    ZSql = ZSql + " and Movgas.Renglon = 1"
    ZSql = ZSql + " Order by Movgas.Clave"
    
    spMovgas = ZSql
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovgas.RecordCount > 0 Then
    
        With rstMovgas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZLugar = ZZLugar + 1
                    
                    ZZOrden(ZZLugar, 1) = rstMovgas!Empresa
                    ZZOrden(ZZLugar, 2) = rstMovgas!Orden
                    ZZOrden(ZZLugar, 3) = rstMovgas!Carpeta
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstMovgas.Close
    End If
    
    For ZZCiclo = 1 To ZZLugar
    
        XEmpresa = WEmpresa
        
        ZZEmpresa = ZZOrden(ZZCiclo, 1)
        ZZNroOrden = ZZOrden(ZZCiclo, 2)
        ZZCarpeta = ZZOrden(ZZCiclo, 3)
        Rem by nan
          
          
        Select Case Val(ZZEmpresa)
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
        
        ZZTipoImpo = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZZNroOrden + "'"
        ZSql = ZSql + " Order by Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            ZZTipoImpo = IIf(IsNull(rstOrden!TipoImpo), "0", rstOrden!TipoImpo)
            rstOrden.Close
        End If
        
        Call Conecta_Empresa
        
        If ZZTipoImpo = 3 Then
            ZSql = ""
            ZSql = ZSql + "UPDATE MovGas SET "
            ZSql = ZSql + " TipoImpo = " + "'" + "3" + "'"
            ZSql = ZSql + " Where Carpeta = " + "'" + ZZCarpeta + "'"
            spMovgas = ZSql
            Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next ZZCiclo
    Rem by nan
    Listado.WindowTitle = "Listado de Carpetas Pendientes de Actualizar"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{ListaOrdpen.Orden} in 0 to 999999"
    
    Listado.GroupSelectionFormula = Uno
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Movgas.Carpeta, Movgas.Renglon, Movgas.Fecha, Movgas.Orden, Movgas.Proveedor, Movgas.Origen, Movgas.Marca, Movgas.Empresa, Movgas.TipoImpo, " _
            + "Proveedor.Nombre " _
            + "From " _
            + DSQ + ".dbo.Movgas Movgas, " _
            + DSQ + ".dbo.Proveedor Proveedor " _
            + "Where " _
            + "Movgas.Proveedor = Proveedor.Proveedor AND " _
            + "Movgas.Renglon = 1 AND " _
            + "Movgas.Marca <> 'X' AND " _
            + "Movgas.TipoImpo = 0"
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Listado.ReportFileName = "ListaCarpetasPendiente.rpt"
        Case Else
            Listado.ReportFileName = "ListaCarpetaPendientePeli.rpt"
    End Select
            
    Listado.Destination = 0
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next

End Sub


