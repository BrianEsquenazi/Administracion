VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgMiraPaso 
   AutoRedraw      =   -1  'True
   Caption         =   "Solicitudes de Pedido de Compra de Insumos Pendientes"
   ClientHeight    =   7320
   ClientLeft      =   135
   ClientTop       =   615
   ClientWidth     =   11580
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11580
   Begin VB.Frame PantaSolicitud 
      Height          =   1215
      Left            =   3600
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ComboBox TipoSolicitud 
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
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.CommandButton ImpreSolicitud 
      Caption         =   "Impre.Solic."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Cumplido 
      Caption         =   "Pedido Cumplido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Impresion 
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
      Height          =   300
      Left            =   6600
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Clave1 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "&Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Anula 
      Caption         =   "Anula Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   327680
      Rows            =   4000
      Cols            =   9
      BackColor       =   16777088
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaInsumo.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cierra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9840
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgMiraPaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim rstInsumo As Recordset
Dim spInsumo As String
Dim XParam As String
Dim WGraba As String
Private TotalSolicitud As Integer
Dim Auxiliar(10000, 15)
Dim ZTipoSolicitud As Integer

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgMiraInsumos.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anula_Click()
    Muestra.Col = 7
    Muestra.Text = "Anulado"
    Muestra.Col = 1
End Sub

Private Sub Cumplido_Click()
    Muestra.Col = 7
    Muestra.Text = "Ok"
    Muestra.Col = 1
End Sub

Private Sub Form_Load()

    XEmpresa = WEmpresa

    Call Limpia_Vector
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 900
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 1600
    Muestra.ColWidth(4) = 800
    Muestra.ColWidth(5) = 3500
    Muestra.ColWidth(6) = 1200
    Muestra.ColWidth(7) = 1200
    Muestra.ColWidth(8) = 600

    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Solicitud"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Solicitante"
    
    Muestra.Col = 4
    Muestra.Text = "Planta"
    
    Muestra.Col = 5
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 6
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 7
    Muestra.Text = "Estado"
    
    Muestra.Col = 8
    Muestra.Text = "Tipo"
    
    TipoSolicitud.Clear
    
    TipoSolicitud.AddItem "Insumos/Repuestos"
    TipoSolicitud.AddItem "Servicios"
    TipoSolicitud.AddItem "Sistemas"
    TipoSolicitud.AddItem "Total"
    
    ZTipoSolicitud = 3
    TipoSolicitud.ListIndex = ZTipoSolicitud
    
End Sub

Private Sub Graba_Click()

    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
            
        XEmpresa = WEmpresa
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        WGraba = ""

        For Ciclo = 1 To TotalSolicitud
        
            Muestra.Row = Ciclo
            Muestra.Col = 7
            
            If Muestra.Text = "Anulado" Then
                Muestra.Col = 1
                WSolicitud = Muestra.Text
                
                Sql1 = "UPDATE Insumos SET "
                Sql2 = " Estado = 6"
                Sql3 = " Where Solicitud = " + "'" + WSolicitud + "'"
                spInsumo = Sql1 + Sql2 + Sql3
                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            If Muestra.Text = "Ok" Then
                Muestra.Col = 1
                WSolicitud = Muestra.Text
                
                Sql1 = "UPDATE Insumos SET "
                Sql2 = " Estado = 7"
                Sql3 = " Where Solicitud = " + "'" + WSolicitud + "'"
                spInsumo = Sql1 + Sql2 + Sql3
                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            End If
    
        Next Ciclo
        
        Call Conecta_Empresa
    
        Call Proceso_Click
        Rem Call cmdClose_Click
        
    End If

End Sub

Private Sub Impresion_Click()

    Listado.ReportFileName = "ListaInsumo.rpt"
    
    Listado.WindowTitle = "Listado de Solicitudes de Insumos Pendiente de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Terminado.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.GroupSelectionFormula = ""
    
    Listado.SQLQuery = "SELECT Insumos.Solicitud, Insumos.Renglon, Insumos.Fecha, Insumos.Planta, Insumos.Solicitante, Insumos.Observaciones, Insumos.Entrega, Insumos.Estado " _
                    + "From " _
                    + DSQ + ".dbo.Insumos Insumos " _
                    + "Where " _
                    + "Insumos.Solicitud >= 0 AND " _
                    + "Insumos.Solicitud <= 999999 AND " _
                    + "Insumos.Renglon = 1 AND " _
                    + "(Insumos.Estado <> 7 AND " _
                    + "Insumos.Estado <> 6)"
    
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1

End Sub

Private Sub ImpreSolicitud_Click()

    XEmpresa = WEmpresa
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Listado.ReportFileName = "ImpreInsumos.rpt"
    
    Listado.WindowTitle = "Emision de Solicitud de Compras de Insumos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    For Ciclo = RowIni To Rowfin
    
        WSolicitud = Muestra.TextMatrix(Ciclo, 1)
    
        Listado.GroupSelectionFormula = "{Insumos.Solicitud} in " + WSolicitud + " to " + WSolicitud
        Listado.Destination = 1
        Rem Listado.Destination = 0
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        Listado.SQLQuery = "SELECT Insumos.Solicitud, Insumos.Fecha, Insumos.Planta, Insumos.Solicitante, Insumos.Observaciones, Insumos.Entrega, Insumos.Cantidad, Insumos.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Insumos Insumos " _
                    + "Where " _
                    + "Insumos.Solicitud >= " + WSolicitud + " AND Insumos.Solicitud <= " + WSolicitud
                            
        Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()
        Listado.Action = 1
    
    Next Ciclo
    
    Call Conecta_Empresa
    
End Sub

Private Sub TipoSolicitud_Click()
    ZTipoSolicitud = TipoSolicitud.ListIndex
    PantaSolicitud.Visible = False
    Call Proceso_Click
End Sub

Private Sub TipoSolicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call TipoSolicitud_Click
    End If
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Solicitud"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Solicitante"
    
    Muestra.Col = 4
    Muestra.Text = "Planta"
    
    Muestra.Col = 5
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 6
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 7
    Muestra.Text = "Estado"
    
    Muestra.Col = 8
    Muestra.Text = "Tipo"
    
    Renglon = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Insumos"
    Sql3 = " Where Renglon = 1"
    Sql4 = " and Estado <> 6 and Estado <> 7"
    Sql5 = " Order by Solicitud"
    spInsumo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumo.RecordCount > 0 Then
    
        With rstInsumo
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    If ZTipoSolicitud = 3 Or ZTipoSolicitud = rstInsumo!TipoSolicitud Then
                
                    Renglon = Renglon + 1
                    Muestra.Row = Renglon
                                                
                    Muestra.Col = 1
                    Muestra.Text = Pusing("######", Str$(rstInsumo!Solicitud))
                                            
                    Muestra.Col = 2
                    Muestra.Text = rstInsumo!Fecha
                
                    Muestra.Col = 3
                    Muestra.Text = rstInsumo!Solicitante
                        
                    Muestra.Col = 4
                    Muestra.Text = rstInsumo!Planta
                        
                    Muestra.Col = 5
                    Muestra.Text = rstInsumo!Observaciones
                    
                    Muestra.Col = 6
                    Muestra.Text = rstInsumo!Entrega
                    
                    WEstado = ""
                    Select Case rstInsumo!Estado
                        Case 1
                            WEstado = "En Proceso"
                        Case 2
                            WEstado = "Compra Realizada"
                        Case 3
                            WEstado = "Compra Parcial"
                        Case 4
                            WEstado = "En Espera de Fondos"
                        Case 5
                            WEstado = "En Espera de Autorizacion"
                        Case 6
                            WEstado = "Anulada"
                        Case 7
                            WEstado = "Pedido Cumplido"
                        Case Else
                    End Select
                    
                    Muestra.Col = 7
                    Muestra.Text = WEstado
                    
                    WTipoSolicitud = ""
                    Select Case rstInsumo!TipoSolicitud
                        Case 0
                            WTipoSolicitud = "  I "
                        Case 1
                            WTipoSolicitud = "Serv"
                        Case 2
                            WTipoSolicitud = "Sist"
                        Case Else
                    End Select
                    
                    Muestra.Col = 8
                    Muestra.Text = WTipoSolicitud
                    
                    End If
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        rstInsumo.Close
    End If
    
    Call Conecta_Empresa
    
    TotalSolicitud = Renglon
    
    Muestra.Font.Bold = True
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.Col = ZMuestraCol
    Muestra.Row = ZMuestraRow
    Muestra.TopRow = ZMuestraTopRow
    
End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Solicitud"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Solicitante"
    
    Muestra.Col = 4
    Muestra.Text = "Planta"
    
    Muestra.Col = 5
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 6
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 7
    Muestra.Text = "Estado"
    
    Muestra.Col = 8
    Muestra.Text = "Tipo"
    
End Sub

Private Sub Muestra_DblClick()

    ZMuestraCol = Muestra.Col
    ZMuestraRow = Muestra.Row
    ZMuestraTopRow = Muestra.TopRow

    If Muestra.Col = 8 Then
        PantaSolicitud.Visible = True
        Exit Sub
    End If

    Muestra.Col = 1
    WXSol = Muestra.Text
    PrgMiraInsumos.Hide
    Unload Me
    PrgInsumosII.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Sub Ingresa_clave()
    WClave.Text = ""
    Clave1.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    Clave1.Visible = False
End Sub


Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        WClave.Text = UCase(WClave.Text)
        If WClave.Text = "BAJA" Then
            WGraba = "S"
            Clave1.Visible = False
            Call Graba_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Solicitudes de Orden de Compra de Insumos")
            WClave.SetFocus
        End If
    End If

End Sub
