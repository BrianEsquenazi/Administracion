VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCertifPend 
   AutoRedraw      =   -1  'True
   Caption         =   "Veriricacion de Certificados Pendientes de Recibir"
   ClientHeight    =   4605
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4605
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3975
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Mensaje 
         Alignment       =   2  'Center
         Caption         =   "Hay certificado pendientes de recibir"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   5055
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "CERTIFPEND.rpt"
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
End
Attribute VB_Name = "PrgCertifPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(10000, 5) As String
Dim ZLugar As Integer

Dim ZZOrden As String
Dim ZZfecha As String
Dim ZZProveedor As String
Dim ZZArticulo As String
Dim ZZPlanta As String

Dim rstCertifPend As Recordset
Dim spCertifPend As String
Dim rstOrden As Recordset
Dim spOrden As String

Dim XParam As String

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Acepta_Click()

    Erase ZVector
    ZLugar = 0
    
    For vA = 1 To 5

        Select Case vA
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
            Case Else
        End Select

        Sql1 = "UPDATE Orden SET "
        Sql2 = " EntregaI = " + "'" + "0" + "'"
        Sql3 = " Where EntregaI IS NULL"
        spOrden = Sql1 + Sql2 + Sql3
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Tipo = 1"
        ZSql = ZSql + " and Orden.EntregaI = 0"
        ZSql = ZSql + " and Orden.FechaOrd >= '20090101'"
        ZSql = ZSql + " Order by Orden.Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar, 1) = rstOrden!Orden
                        ZVector(ZLugar, 2) = rstOrden!Fecha
                        ZVector(ZLugar, 3) = rstOrden!Proveedor
                        ZVector(ZLugar, 4) = rstOrden!Articulo
                        ZVector(ZLugar, 5) = vA
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
        
    Next vA

    WEntra = "N"
    
    If ZLugar > 0 Then
        Mensaje.Visible = True
        WEntra = "S"
            Else
        Mensaje.Visible = False
    End If
    
    If WEntra = "S" Then
        PrgCertifPend.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgCertifPend.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    ZSql = ""
    ZSql = ZSql + "DELETE CertifPend"
    spCertifPend = ZSql
    Set rstCertifPend = db.OpenRecordset(spCertifPend, dbOpenSnapshot, dbSQLPassThrough)

    For Ciclo = 1 To ZLugar
    
        ZZOrden = ZVector(Ciclo, 1)
        
        ZZfecha = ZVector(Ciclo, 2)
        ZZProveedor = ZVector(Ciclo, 3)
        ZZArticulo = ZVector(Ciclo, 4)
        ZZPlanta = ZVector(Ciclo, 5)
        
        
        
        
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CertifPend ("
        ZSql = ZSql + "Planta,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "Articulo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZPlanta + "',"
        ZSql = ZSql + "'" + ZZOrden + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZProveedor + "',"
        ZSql = ZSql + "'" + ZZArticulo + "')"
        spCertifPend = ZSql
        Set rstCertifPend = db.OpenRecordset(spCertifPend, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo

    Listado.WindowTitle = "Certificados Pendientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.Destination = 1
    Rem Listado.Destination = 0

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = "SELECT CertifPend.Planta, CertifPend.Orden, CertifPend.Fecha, CertifPend.Proveedor, CertifPend.Articulo, " _
            + "Proveedor.Nombre, " _
            + "Articulo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.CertifPend CertifPend, " _
            + DSQ + ".dbo.Proveedor Proveedor, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "CertifPend.Proveedor = Proveedor.Proveedor AND " _
            + "CertifPend.Articulo = Articulo.Codigo AND " _
            + "CertifPend.Planta >= 0 AND " _
            + "CertifPend.Planta <= 999"

    Listado.Connect = Connect()

    Listado.Action = 1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgCertifPend.Hide
    Unload Me
    Close
    End
End Sub


