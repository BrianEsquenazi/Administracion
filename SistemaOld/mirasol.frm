VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgMirasol 
   AutoRedraw      =   -1  'True
   Caption         =   "Solicitudes de Pedido de Compra Pendientes"
   ClientHeight    =   7320
   ClientLeft      =   435
   ClientTop       =   780
   ClientWidth     =   10995
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   10995
   Begin VB.CommandButton Command3 
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   360
      Width           =   375
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
      Left            =   7920
      TabIndex        =   9
      Top             =   0
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
      Left            =   3000
      TabIndex        =   4
      Top             =   0
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
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   4000
      Cols            =   11
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wsoltot.rpt"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancela"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "PrgMIrasol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstSolic As Recordset
Dim spSolic As String
Dim rstSoltot As Recordset
Dim spSoltot As String
Dim XParam As String
Dim WGraba As String
Private TotalSolicitud As Integer
Dim Auxiliar(10000, 15)

Private Sub cmdClose_Click()

    WEmpe = Val(XEmpresa)
    WEmpresa = XEmpresa
    
    Select Case WEmpe
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
    
    With rstEmpresa
        .Close
    End With
    PrgMIrasol.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Anula_Click()

    Muestra.Col = 9
    Muestra.Text = "Anulado"
    Muestra.Col = 1

End Sub

Private Sub Command1_Click()
vercolumna = 2
Call Proceso_Click
End Sub

Private Sub Command2_Click()
vercolumna = 4
Call Proceso_Click
End Sub

Private Sub Command3_Click()
vercolumna = 8
Call Proceso_Click
End Sub

Private Sub Form_Load()
    vercolumna = 0
    XEmpresa = WEmpresa

    Call Limpia_Vector
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 800
    Muestra.ColWidth(2) = 1000
    Muestra.ColWidth(3) = 1300
    Muestra.ColWidth(4) = 800
    Muestra.ColWidth(5) = 1000
    Muestra.ColWidth(6) = 1600
    Muestra.ColWidth(7) = 800
    Muestra.ColWidth(8) = 1000
    Muestra.ColWidth(9) = 3000
    Muestra.ColWidth(10) = 200

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
    Muestra.Text = "Producto"
    
    Muestra.Col = 6
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 7
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 8
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 9
    Muestra.Text = "Observaciones"
    
    Muestra.Col = 10
    Muestra.Text = "Emp"
    
    Call Proceso_Click
    
End Sub

Private Sub Graba_Click()

    If WGraba <> "S" Then
        Call Ingresa_clave
            Else
            
        WGraba = ""

        For Ciclo = 1 To TotalSolicitud
        
    
            Muestra.Row = Ciclo
            Muestra.Col = 9
            If Muestra.Text = "Anulado" Then
            
                Muestra.Col = 10
                WEmpe = Val(Muestra.Text)
            
                Select Case WEmpe
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
            
                Muestra.Col = 1
                WSolicitud = Muestra.Text
                Muestra.Col = 5
                WArticulo = Muestra.Text
                WMarca = "X"
                XParam = "'" + WSolicitud + "','" _
                            + WArticulo + "','" _
                            + WMarca + "'"
                spSolic = "ModificaSolicitudMarca " + XParam
                Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
    
        Next Ciclo
    
        Call cmdClose_Click
        
    End If

End Sub

Private Sub Impresion_Click()

    Listado.WindowTitle = "Listado de Solicitudes de Ordenes de Compra Pendiente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Terminado.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Soltot.Solicitud, Soltot.Fecha, Soltot.Articulo, Soltot.Cantidad, Soltot.Entrega, Soltot.Planta, Soltot.Solicitante, Soltot.Obser, " _
                        + "Articulo.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.Soltot Soltot, " _
                        + DSQ + ".dbo.Articulo Articulo " _
                        + "Where " _
                        + "Soltot.Articulo = Articulo.Codigo AND " _
                        + "Soltot.Solicitud >= 0 AND Soltot.Solicitud <= 999999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1

End Sub

Private Sub Muestra_Click()
Rem Stop
Rem vercolumna = Muestra.ColSel
Rem Call Proceso_Click
End Sub

Private Sub Proceso_Click()
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    spSoltot = "BorrarSoltot "
    Set rstSoltot = db.OpenRecordset(spSoltot, dbOpenSnapshot, dbSQLPassThrough)

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
    Muestra.Text = "Producto"
    
    Muestra.Col = 6
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 7
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 8
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 9
    Muestra.Text = "Observaciones"
    
    Muestra.Col = 10
    Muestra.Text = "Emp"
    
    Erase Auxiliar
    WLugar = 0
    
    For Cicla = 1 To 11
    
        Select Case Cicla
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
        
        XParam = "'" + "X" + " '"
        spSolic = "ListaSolicitudPendiente " + XParam
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
            With rstSolic
    
                .MoveFirst
                If .NoMatch = False Then
                    Do
                
                        Corte = rstSolic!Solicitud
                        Fecha = rstSolic!Fecha
                        Solicitante = rstSolic!Solicitante
                        Planta = rstSolic!Planta
                        Observaciones = rstSolic!Observaciones
                        
                        WLugar = WLugar + 1
                        Auxiliar(WLugar, 1) = Pusing("######", Str$(rstSolic!Solicitud))
                        Auxiliar(WLugar, 2) = rstSolic!Fecha
                        Auxiliar(WLugar, 3) = rstSolic!Solicitante
                        Auxiliar(WLugar, 4) = rstSolic!Planta
                        Auxiliar(WLugar, 5) = rstSolic!Articulo
                        Auxiliar(WLugar, 6) = ""
                        Auxiliar(WLugar, 7) = rstSolic!Cantidad - rstSolic!Entregado
                        Auxiliar(WLugar, 8) = rstSolic!Entrega
                        Auxiliar(WLugar, 9) = rstSolic!Obser
                        Auxiliar(WLugar, 10) = WEmpresa
                        Auxiliar(WLugar, 11) = rstSolic!Clave
                    
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
        
            End With
            rstSolic.Close
        End If
        
    Next Cicla
    
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Cicla = 1 To WLugar
    
        ZSolicitud = Auxiliar(Cicla, 1)
        ZFecha = Auxiliar(Cicla, 2)
        ZSolicitante = Auxiliar(Cicla, 3)
        ZPlanta = Auxiliar(Cicla, 4)
        ZArticulo = Auxiliar(Cicla, 5)
        ZCantidad = Str$(Auxiliar(Cicla, 7))
        ZEntrega = Auxiliar(Cicla, 8)
        ZObser = Auxiliar(Cicla, 9)
        ZEmpresa = Auxiliar(Cicla, 10)
        ZClave = Auxiliar(Cicla, 11)
        ZBaja = ""
        ZRenglon = ""
        ZFechaOrd = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        ZObservaciones = ""
        ZOrdEntrega = Right$(ZEntrega, 4) + Mid$(ZEntrega, 4, 2) + Left$(ZEntrega, 2)
        ZDate = ""
        ZEntregado = ZEmpresa
        ZMarca = ""
        
        XParam = "'" + ZClave + "','" _
                + ZSolicitud + "','" _
                + ZRenglon + "','" _
                + ZFecha + "','" _
                + ZFechaOrd + "','" _
                + ZObservaciones + "','" _
                + ZArticulo + "','" _
                + ZCantidad + "','" _
                + ZEntrega + "','" _
                + ZOrdEntrega + "','" _
                + ZPlanta + "','" _
                + ZSolicitante + "','" _
                + ZDate + "','" _
                + ZMarca + "','" _
                + ZObser + "','" _
                + ZEntregado + "','" _
                + ZBaja + "'"
                         
        spSoltot = "AltaSoltot " + XParam
        Set rstSoltot = db.OpenRecordset(spSoltot, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Cicla
    
    
    Renglon = 0
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
  
  
Rem by nan
    Select Case vercolumna
        Case 2
            Rem     Stop
            Rem    Case Else
            Rem by nan
            spSoltotnan = "ListaSoltotnan"
            Set rstSoltotnan = db.OpenRecordset(spSoltotnan, dbOpenSnapshot, dbSQLPassThrough)
            If rstSoltotnan.RecordCount > 0 Then
                With rstSoltotnan
            
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                        
                            Corte = rstSoltotnan!Solicitud
                            Fecha = rstSoltotnan!Fecha
                            Solicitante = rstSoltotnan!Solicitante
                            Planta = rstSoltotnan!Planta
                            Observaciones = rstSoltotnan!Observaciones
                                
                            Renglon = Renglon + 1
                            Muestra.Row = Renglon
                                                        
                            Muestra.Col = 1
                            Muestra.Text = Pusing("######", Str$(Corte))
                                                    
                            Muestra.Col = 2
                            Muestra.Text = Fecha
                        
                            Muestra.Col = 3
                            Muestra.Text = Solicitante
                                
                            Muestra.Col = 4
                            Muestra.Text = Planta
                                
                            Muestra.Col = 5
                            Muestra.Text = rstSoltotnan!Articulo
                            
                            Muestra.Col = 6
                            Muestra.Text = ""
                            
                            Muestra.Col = 7
                            Muestra.Text = rstSoltotnan!Cantidad
                            
                            Muestra.Col = 8
                            Muestra.Text = rstSoltotnan!Entrega
                            
                            Muestra.Col = 9
                            Muestra.Text = rstSoltotnan!Obser
                                
                            Muestra.Col = 10
                            Muestra.Text = rstSoltotnan!Entregado
                            
                            .MoveNext
                        
                            If .EOF = True Then
                                Exit Do
                            End If
                        
                        Loop
                    End If
                
                End With
                rstSoltotnan.Close
            End If
  
        Case 4
            spSoltotnan1 = "ListaSoltotnan1"
            Set rstSoltotnan1 = db.OpenRecordset(spSoltotnan1, dbOpenSnapshot, dbSQLPassThrough)
            If rstSoltotnan1.RecordCount > 0 Then
                With rstSoltotnan1
          
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                      
                            Corte = rstSoltotnan1!Solicitud
                            Fecha = rstSoltotnan1!Fecha
                            Solicitante = rstSoltotnan1!Solicitante
                            Planta = rstSoltotnan1!Planta
                            Observaciones = rstSoltotnan1!Observaciones
                              
                            Renglon = Renglon + 1
                            Muestra.Row = Renglon
                                                        
                            Muestra.Col = 1
                            Muestra.Text = Pusing("######", Str$(Corte))
                                                    
                            Muestra.Col = 2
                            Muestra.Text = Fecha
                        
                            Muestra.Col = 3
                            Muestra.Text = Solicitante
                                
                            Muestra.Col = 4
                            Muestra.Text = Planta
                                
                            Muestra.Col = 5
                            Muestra.Text = rstSoltotnan1!Articulo
                            
                            Muestra.Col = 6
                            Muestra.Text = ""
                            
                            Muestra.Col = 7
                            Muestra.Text = rstSoltotnan1!Cantidad
                            
                            Muestra.Col = 8
                            Muestra.Text = rstSoltotnan1!Entrega
                            
                            Muestra.Col = 9
                            Muestra.Text = rstSoltotnan1!Obser
                                
                            Muestra.Col = 10
                            Muestra.Text = rstSoltotnan1!Entregado
                            
                            .MoveNext
                        
                            If .EOF = True Then
                                Exit Do
                            End If
                      
                        Loop
                    End If
              
                End With
                rstSoltotnan1.Close
            End If
            
        Rem by nan por fecha de entrega
        Case 8
            spSoltotnan2 = "ListaSoltotnan2"
            Set rstSoltotnan2 = db.OpenRecordset(spSoltotnan2, dbOpenSnapshot, dbSQLPassThrough)
            If rstSoltotnan2.RecordCount > 0 Then
                With rstSoltotnan2
           
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                       
                            Corte = rstSoltotnan2!Solicitud
                            Fecha = rstSoltotnan2!Fecha
                            Solicitante = rstSoltotnan2!Solicitante
                            Planta = rstSoltotnan2!Planta
                            Observaciones = rstSoltotnan2!Observaciones
                               
                            Renglon = Renglon + 1
                            Muestra.Row = Renglon
                                                        
                            Muestra.Col = 1
                            Muestra.Text = Pusing("######", Str$(Corte))
                                                    
                            Muestra.Col = 2
                            Muestra.Text = Fecha
                        
                            Muestra.Col = 3
                            Muestra.Text = Solicitante
                                
                            Muestra.Col = 4
                            Muestra.Text = Planta
                                
                            Muestra.Col = 5
                            Muestra.Text = rstSoltotnan2!Articulo
                            
                            Muestra.Col = 6
                            Muestra.Text = ""
                            
                            Muestra.Col = 7
                            Muestra.Text = rstSoltotnan2!Cantidad
                            
                            Muestra.Col = 8
                            Muestra.Text = rstSoltotnan2!Entrega
                            
                            Muestra.Col = 9
                            Muestra.Text = rstSoltotnan2!Obser
                                
                            Muestra.Col = 10
                            Muestra.Text = rstSoltotnan2!Entregado
                            
                            .MoveNext
                        
                            If .EOF = True Then
                                Exit Do
                            End If
                       
                        Loop
                    End If
                   
                End With
                rstSoltotnan2.Close
            End If
        
        Case Else
           
            spSoltot = "ListaSoltot "
            Set rstSoltot = db.OpenRecordset(spSoltot, dbOpenSnapshot, dbSQLPassThrough)
            If rstSoltot.RecordCount > 0 Then
               With rstSoltot
           
                   .MoveFirst
                   If .NoMatch = False Then
                       Do
                       
                           Corte = rstSoltot!Solicitud
                           Fecha = rstSoltot!Fecha
                           Solicitante = rstSoltot!Solicitante
                           Planta = rstSoltot!Planta
                           Observaciones = rstSoltot!Observaciones
                               
                           Renglon = Renglon + 1
                           Muestra.Row = Renglon
                                                       
                           Muestra.Col = 1
                           Muestra.Text = Pusing("######", Str$(Corte))
                                                   
                           Muestra.Col = 2
                           Muestra.Text = Fecha
                       
                           Muestra.Col = 3
                           Muestra.Text = Solicitante
                               
                           Muestra.Col = 4
                           Muestra.Text = Planta
                               
                           Muestra.Col = 5
                           Muestra.Text = rstSoltot!Articulo
                           
                           Muestra.Col = 6
                           Muestra.Text = ""
                           
                           Muestra.Col = 7
                           Muestra.Text = rstSoltot!Cantidad
                           
                           Muestra.Col = 8
                           Muestra.Text = rstSoltot!Entrega
                           
                           Muestra.Col = 9
                           Muestra.Text = rstSoltot!Obser
                               
                           Muestra.Col = 10
                           Muestra.Text = rstSoltot!Entregado
                           
                           .MoveNext
                       
                           If .EOF = True Then
                               Exit Do
                           End If
                       
                       Loop
                   End If
               
               End With
               rstSoltot.Close
           End If
    
    End Select
    Rem final del end case
    
    
    
    EmpAnt = 0
Rem by nan



    For Cicla = 1 To Renglon
    
        Muestra.Row = Cicla
        Muestra.Col = 5
        WArti = Muestra.Text
        
        Muestra.Col = 10
        WEmpe = Val(Muestra.Text)
    
        If EmpAnt <> WEmpe Then
        
            EmpAnt = WEmpe
        
            Select Case WEmpe
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
    
        spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Muestra.Col = 6
            Muestra.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    Next Cicla
    
    TotalSolicitud = Renglon
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.TopRow = 1
    
    Rem Muestra.SetFocus
    
    WEmpe = Val(XEmpresa)
    Select Case WEmpe
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
Rem End Select
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
    Muestra.Text = "Producto"
    
    Muestra.Col = 6
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 7
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 8
    Muestra.Text = "F.Entrega"
    
    Muestra.Col = 9
    Muestra.Text = "Observaciones"
    
    Muestra.Col = 10
    Muestra.Text = "Emp"
    
End Sub

Private Sub Muestra_DblClick()
        
    Muestra.Col = 10
    WEmpe = Val(Muestra.Text)
    
    Select Case WEmpe
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

    Muestra.Col = 1
    WXSol = Muestra.Text
    
    PrgSol.Show
    
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
            a% = MsgBox(m$, 0, "Solicitudes de Orden de Compra")
            WClave.SetFocus
        End If
    End If

End Sub

Private Sub Conecta()

    WEmpe = Val(XEmpresa)
    WEmpresa = XEmpresa
    
    Select Case WEmpe
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
    
End Sub




