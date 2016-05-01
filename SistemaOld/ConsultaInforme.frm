VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgConsultaInforme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Remito"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   11835
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11835
   Visible         =   0   'False
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   12
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Impre 
      Caption         =   "Impresion"
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
      Left            =   5160
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   11
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   10
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CargaRemito 
      Height          =   285
      Left            =   7080
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   9
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   8
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   7
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   6
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   5
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreConsultaRemito.rpt"
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid WGrilla 
      Height          =   5655
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9975
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4200
      MouseIcon       =   "ConsultaInforme.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaInforme.frx":030A
      ToolTipText     =   "Salida"
      Top             =   6480
      Width           =   480
   End
   Begin VB.Label DesProveedor 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
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
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgConsultaInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZNroRemito(100) As String
Dim ZZRemito As String
Dim EmpresaTrabajo As String

Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Precio As Double
Private Condicion As String
Private Entra As String

Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstEnvases As Recordset
Dim spEnvases As String
Dim rstImpreConsultaRemito As Recordset
Dim spImpreConsultaRemito As String

Dim XParam As String

Private Sub cmdClose_Click()
    PrgConsultaInforme.Hide
    Unload Me
    Select Case ZZPasaProceso
        Case 0
            PrgCompras.Show
        Case Else
            PrgConsultaRemito.Show
    End Select
End Sub

Private Sub Form_Activate()
    CargaRemito.Text = ZZPasaRemito
End Sub

Private Sub Form_Load()
    
    Call Limpia_Grilla
    
    Proveedor.Text = ZZPasaProveedor
    DesProveedor.Caption = ""
    
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = RstProveedor!Nombre
        RstProveedor.Close
    End If
    
    CargaRemito.Text = Trim(ZZPasaRemito)
    ZLugarII = 0
    Erase ZNroRemito
    
    Do
        MyPos = InStr(CargaRemito.Text, ",")
        If MyPos = 0 Then
            ZLugarII = ZLugarII + 1
            ZNroRemito(ZLugarII) = CargaRemito.Text
            Exit Do
                Else
            ZLugarII = ZLugarII + 1
            ZNroRemito(ZLugarII) = Mid$(CargaRemito.Text, 1, MyPos - 1)
            CargaRemito.Text = Mid$(CargaRemito.Text, MyPos + 1, 100)
        End If
    Loop
    
    
    Call Busca_Empresa
    XEmpresa = WEmpresa
    
    Select Case Val(EmpresaTrabajo)
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
    
    For CicloII = 1 To ZLugarII
    
        ZZRemito = ZNroRemito(CicloII)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Informe"
        ZSql = ZSql + " Where Informe.Remito = " + "'" + ZZRemito + "'"
        ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
        spInforme = ZSql
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            With rstInforme
                .MoveFirst
                Do
                    If .EOF = False Then
                        
                        If rstInforme!Cantidad <> 0 Then
                        
                            WLugar = WLugar + 1
                        
                            WGrilla.TextMatrix(WLugar, 1) = ZZRemito
                            WGrilla.TextMatrix(WLugar, 2) = Str$(rstInforme!Orden)
                            WGrilla.TextMatrix(WLugar, 3) = rstInforme!Articulo
                            WGrilla.TextMatrix(WLugar, 9) = Str$(rstInforme!Informe)
                            
                            If UCase(Left$(rstInforme!Articulo, 2)) = "ZE" Then
                                WGrilla.TextMatrix(WLugar, 10) = Str$(rstInforme!Cantidad)
                            End If
                            
                        
                        End If
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstInforme.Close
        End If
        
    Next CicloII
    
    For Ciclo = 1 To WLugar
    
        ZZOrden = WGrilla.TextMatrix(Ciclo, 2)
        ZZArticulo = WGrilla.TextMatrix(Ciclo, 3)
        ZZInforme = WGrilla.TextMatrix(Ciclo, 9)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Codigo = " + "'" + ZZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WGrilla.TextMatrix(Ciclo, 4) = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZZOrden + "'"
        ZSql = ZSql + " and Orden.Articulo = " + "'" + ZZArticulo + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WGrilla.TextMatrix(Ciclo, 5) = Pusing("###,###.##", Str$(rstOrden!Cantidad))
            Select Case rstOrden!Moneda
                Case 0
                    WGrilla.TextMatrix(Ciclo, 6) = "U$S"
                Case 1
                    WGrilla.TextMatrix(Ciclo, 6) = "S"
                Case 2
                    WGrilla.TextMatrix(Ciclo, 6) = "Euro"
                Case Else
                    WGrilla.TextMatrix(Ciclo, 6) = ""
            End Select
            WGrilla.TextMatrix(Ciclo, 7) = Pusing("###,###.##", Str$(rstOrden!Precio))
            WGrilla.TextMatrix(Ciclo, 8) = Trim(rstOrden!Condicion)
            rstOrden.Close
        End If
        
        ZZLiberada = 0
        ZZRechazada = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Orden = " + "'" + ZZOrden + "'"
        ZSql = ZSql + " and Laudo.Informe = " + "'" + ZZInforme + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + ZZArticulo + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZZLiberada = ZZLiberada + rstLaudo!Liberada
                        ZZRechazada = ZZRechazada + rstLaudo!devuelta
                        WGrilla.TextMatrix(Ciclo, 12) = Left$(rstLaudo!Fecha, 5)
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        
            rstLaudo.Close
        End If
        
        If ZZLiberada > 0 Then
            WGrilla.TextMatrix(Ciclo, 10) = Pusing("###,###.##", Str$(ZZLiberada))
            WGrilla.TextMatrix(Ciclo, 11) = "Aprob."
                Else
            If ZZRechazada > 0 Then
                WGrilla.TextMatrix(Ciclo, 11) = "Rech."
            End If
        End If
    
    Next Ciclo
        
    Call Conecta_Empresa
    
End Sub

Private Sub Limpia_Grilla()

    WGrilla.Clear
    WGrilla.Font.Bold = True
    
    WGrilla.FixedCols = 1
    WGrilla.Cols = 13
    WGrilla.FixedRows = 1
    WGrilla.Rows = 101
    
    WGrilla.ColWidth(0) = 200
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        Select Case Ciclo
            Case 1
                WGrilla.Text = "Remito"
                WGrilla.ColWidth(Ciclo) = 800
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WGrilla.Text = "Orden"
                WGrilla.ColWidth(Ciclo) = 800
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WGrilla.Text = "Producto"
                WGrilla.ColWidth(Ciclo) = 1200
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WGrilla.Text = "Descripcion"
                WGrilla.ColWidth(Ciclo) = 1900
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WGrilla.Text = "Cant.Ped."
                WGrilla.ColWidth(Ciclo) = 1000
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WGrilla.Text = "M."
                WGrilla.ColWidth(Ciclo) = 400
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WGrilla.Text = "Precio"
                WGrilla.ColWidth(Ciclo) = 900
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 8
                WGrilla.Text = "C.Pago"
                WGrilla.ColWidth(Ciclo) = 1000
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 9
                WGrilla.Text = "Informe"
                WGrilla.ColWidth(Ciclo) = 800
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WGrilla.Text = "Cant.Ing."
                WGrilla.ColWidth(Ciclo) = 1000
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 11
                WGrilla.Text = "Est."
                WGrilla.ColWidth(Ciclo) = 600
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 12
                WGrilla.Text = "F.Apr."
                WGrilla.ColWidth(Ciclo) = 800
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        WTitulo(Ciclo).Text = WGrilla.Text
        WTitulo(Ciclo).Left = WGrilla.CellLeft + WGrilla.Left
        WTitulo(Ciclo).Top = WGrilla.CellTop + WGrilla.Top
        WTitulo(Ciclo).Width = WGrilla.CellWidth
        WTitulo(Ciclo).Height = WGrilla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WGrilla.Cols - 1
        WAncho = WAncho + WGrilla.ColWidth(Ciclo)
    Next Ciclo
    WGrilla.Width = WAncho

    ' Size the columns.
    Font.Name = WGrilla.Font.Name
    Font.Size = WGrilla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WGrilla.AllowUserResizing = flexResizeBoth
    
    WGrilla.Visible = True
    
    WGrilla.Col = 1
    WGrilla.Row = 1
    
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

Private Sub Impre_Click()

    ZSql = "DELETE ImpreConsultaRemito"
    spImpreConsultaRemito = ZSql
    Set rstImpreConsultaRemito = db.OpenRecordset(spImpreConsultaRemito, dbOpenSnapshot, dbSQLPassThrough)
        
    For iRow = 1 To 100

        ZZProveedor = Proveedor.Text
        ZZDesProveedor = DesProveedor.Caption
    
        ZZRemito = WGrilla.TextMatrix(iRow, 1)
        ZZOrden = WGrilla.TextMatrix(iRow, 2)
        ZZArticulo = WGrilla.TextMatrix(iRow, 3)
        ZZDesArticulo = WGrilla.TextMatrix(iRow, 4)
        ZZCantidad = WGrilla.TextMatrix(iRow, 5)
        ZZMoneda = WGrilla.TextMatrix(iRow, 6)
        ZZPrecio = WGrilla.TextMatrix(iRow, 7)
        ZZCondicion = WGrilla.TextMatrix(iRow, 8)
        ZZInforme = WGrilla.TextMatrix(iRow, 9)
        ZZCantidadII = WGrilla.TextMatrix(iRow, 10)
        ZZEstado = WGrilla.TextMatrix(iRow, 11)
        ZZFAprobacion = WGrilla.TextMatrix(iRow, 12)
        
        If Val(ZZRemito) <> 0 Or Val(ZZOrden) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreConsultaRemito ("
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "DesProveedor ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "DesArticulo ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Moneda ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "Condicion ,"
            ZSql = ZSql + "Informe ,"
            ZSql = ZSql + "CantidadII ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "FAprobacion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZProveedor + "',"
            ZSql = ZSql + "'" + ZZDesProveedor + "',"
            ZSql = ZSql + "'" + ZZRemito + "',"
            ZSql = ZSql + "'" + ZZOrden + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZDesArticulo + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZMoneda + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZCondicion + "',"
            ZSql = ZSql + "'" + ZZInforme + "',"
            ZSql = ZSql + "'" + ZZCantidadII + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + ZZFAprobacion + "')"
            
            spImpreConsultaRemito = ZSql
            Set rstImpreConsultaRemito = db.OpenRecordset(spImpreConsultaRemito, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
            
    Next iRow
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImpreConsultaRemito.Proveedor, ImpreConsultaRemito.DesProveedor, ImpreConsultaRemito.Remito, ImpreConsultaRemito.Orden, ImpreConsultaRemito.Renglon, ImpreConsultaRemito.Articulo, ImpreConsultaRemito.DesArticulo, ImpreConsultaRemito.Cantidad, ImpreConsultaRemito.Moneda, ImpreConsultaRemito.Precio, ImpreConsultaRemito.Condicion, ImpreConsultaRemito.Informe, ImpreConsultaRemito.CantidadII, ImpreConsultaRemito.Estado, ImpreConsultaRemito.FAprobacion " _
            + "From " _
            + DSQ + ".dbo.ImpreConsultaRemito ImpreConsultaRemito " _
            + "Where " _
            + "ImpreConsultaRemito.Orden >= 0 AND " _
            + "ImpreConsultaRemito.Orden <= 999999"
                            
    Listado.Connect = Connect()
    
    Listado.ReportFileName = "ImpreConsultaRemito.rpt"
    Listado.Destination = 1
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
    
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call cmdClose_Click
    End If
End Sub

Private Sub WGRilla_DblClick()

    ZZPasaOrden = WGrilla.TextMatrix(WGrilla.Row, 2)
    ZZPasaEmpresa = EmpresaTrabajo
    PrgConsultaOrden.Show

End Sub


Private Sub Busca_Empresa()

    EmpresaTrabajo = 0
    EmpresaAnterior = WEmpresa
    XEmpresa = WEmpresa
    
    If EmpresaAnterior = 1 Then

        For Va = 1 To 7
    
            Select Case Va
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
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Remito = " + "'" + ZNroRemito(1) + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                EmpresaTrabajo = WEmpresa
                rstInforme.Close
                Exit For
            End If
        
        Next Va
        
            Else
        
        For Va = 1 To 4
    
            Select Case Va
                Case 1
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Remito = " + "'" + ZNroRemito(1) + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                EmpresaTrabajo = WEmpresa
                rstInforme.Close
            End If
        
        Next Va
        
    End If
    
    Call Conecta_Empresa
    
End Sub

