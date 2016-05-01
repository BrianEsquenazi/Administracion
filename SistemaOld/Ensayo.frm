VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgEnsayo 
   Caption         =   "Ingreso de Ensayos"
   ClientHeight    =   5985
   ClientLeft      =   1695
   ClientTop       =   750
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5985
   ScaleWidth      =   7725
   Begin VB.TextBox Unidad 
      Height          =   285
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   33
      Text            =   " "
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox WDescriII 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   840
      Width           =   6015
   End
   Begin VB.Frame XClave 
      BackColor       =   &H00C0C000&
      Height          =   1935
      Left            =   2640
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   30
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1575
      Left            =   1080
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   3735
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Pantalla 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   22
         Text            =   " "
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   21
         Text            =   " "
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.TextBox WDescri 
      Height          =   285
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   25
      Text            =   " "
      Top             =   480
      Width           =   6015
   End
   Begin VB.TextBox WEnsayo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin Crystal.CrystalReport EnsListado 
      Left            =   4080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEnsayos.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   "{Ensayos.Codigo} in 2 to 3"
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.ListBox EnsIndice 
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Enspantalla 
      Height          =   2400
      ItemData        =   "Ensayo.frx":0000
      Left            =   240
      List            =   "Ensayo.frx":0007
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton EnsList 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1320
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton EnsConsulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   3960
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
      Begin VB.CommandButton EnsAnterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton EnsSiguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton EnsUltimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton EnsPrimer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   300
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Unidad"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descricion Ingles"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEnsayo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim XParam As String
Dim EmpresaActual As String

Private WGraba As String
Private WGrabaII As String
Dim ZProceso As Integer

Private Sub Acepta_Click()
    
    Dim DbConnect$, DSN$, UID$, PWD$, DSQ$
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    EnsListado.WindowTitle = "Listado de Ensayos"
    EnsListado.WindowTop = 0
    EnsListado.WindowLeft = 0
    EnsListado.WindowWidth = Screen.Width
    EnsListado.WindowHeight = Screen.Height

    EnsListado.GroupSelectionFormula = "{Ensayos.Codigo} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        EnsListado.Destination = 1
            Else
        EnsListado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    EnsListado.SQLQuery = "SELECT Ensayos.Codigo, Ensayos.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Ensayos Ensayos " _
                    + "Where " _
                    + "Ensayos.Codigo >= " + Desde.Text + " AND " _
                    + "Ensayos.Codigo <= " + Hasta.Text
    
    EnsListado.DataFiles(1) = WEmpresa + "auxi.mdb"
    EnsListado.Connect = Connect()

    EnsListado.Action = 1
    Frame2.Visible = False
    
    Call Conecta_Empresa
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()

    If WGraba <> "S" Then
            
        ZProceso = 1
        Call Ingresa_clave

               Else

        If WEnsayo.Text <> "" Then
    
            WGraba = ""
            WGrabaII = ""
            
            XEmpresa = WEmpresa
        
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
    
            spEnsayo = "ConsultaEnsayos " + "'" + WEnsayo.Text + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                rstEnsayo.Close
                
                ZCodigo = WEnsayo.Text
                ZDescripcion = WDescri.Text
                ZDescripcionII = WDescriII.Text
                ZUnidad = Unidad.Text
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Ensayos SET "
                ZSql = ZSql + "Descripcion = " + "'" + ZDescripcion + "',"
                ZSql = ZSql + "DescripcionII = " + "'" + ZDescripcionII + "',"
                ZSql = ZSql + "Unidad = " + "'" + ZUnidad + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + ZCodigo + "'"
                 
                spEnsayos = ZSql
                Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
                    
                ZCodigo = WEnsayo.Text
                ZDescripcion = WDescri.Text
                ZDescripcionII = WDescriII.Text
                ZUnidad = Unidad.Text
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Ensayos ("
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Descripcion ,"
                ZSql = ZSql + "DescripcionII ,"
                ZSql = ZSql + "Unidad )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZCodigo + "',"
                ZSql = ZSql + "'" + ZDescripcion + "',"
                ZSql = ZSql + "'" + ZDescripcionII + "',"
                ZSql = ZSql + "'" + ZUnidad + "')"
                
                spEnsayos = ZSql
                Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        
        
            Call Conecta_Empresa
        
            Call CmdLimpiar_Click
            WEnsayo.SetFocus
            
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If WGraba <> "S" Then
    
        ZProceso = 2
        Call Ingresa_clave

               Else

        If WEnsayo.Text <> "" Then
    
            WGraba = ""
            WGrabaII = ""
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then

                XEmpresa = WEmpresa
        
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
    
                spEnsayo = "ConsultaEnsayos " + "'" + WEnsayo.Text + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    rstEnsayo.Close
                    spEnsayo = "BorrarEnsayos " + "'" + WEnsayo.Text + "'"
                    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenDynaset, dbSQLPassThrough)
                End If
        
                Call Conecta_Empresa
                
            End If
            
            Call CmdLimpiar_Click
            WEnsayo.SetFocus
            
        End If
        
    End If
End Sub

Private Sub CmdLimpiar_Click()
    WGraba = ""
    WGrabaII = ""
    WEnsayo.Text = ""
    WDescri.Text = ""
    WDescriII.Text = ""
    Unidad.Text = ""
    WEnsayo.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgEnsayo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
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
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEnsayo.Caption = "Ingreso de Ensayos :  " + !Nombre
        End If
    End With
    WEnsayo.Text = ""
    WDescri.Text = ""
    WDescriII.Text = ""
    Unidad.Text = ""
    WGraba = ""
    WGrabaII = ""
    EmpresaActual = WEmpresa
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub EnsAnterior_Click()

    XEmpresa = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    spEnsayo = "AnteriorEnsayos " + "'" + WEnsayo.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    
    If rstEnsayo.RecordCount > 0 Then
        With rstEnsayo
            .MoveLast
            WEnsayo.Text = rstEnsayo!Codigo
            WDescri.Text = Trim(rstEnsayo!Descripcion)
            WDescriII.Text = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
            WDescriII.Text = Trim(WDescriII.Text)
            Unidad.Text = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            Unidad.Text = Trim(Unidad.Text)
        End With
        rstEnsayo.Close
    End If
    
    Call Conecta_Empresa
    
    WEnsayo.SetFocus
    
End Sub





Private Sub WEnsayo_Keypress(KeyAscii As Integer)
    If WEnsayo.Text <> "" Then
        If KeyAscii = 13 Then
        
            XEmpresa = WEmpresa
            
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            Rem spEnsayo = "ConsultaEnsayos " + "'" + WEnsayo.Text + "'"
            spEnsayo = "Select * FROM Ensayos Where Codigo = " + WEnsayo.Text
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                WDescri.Text = Trim(rstEnsayo!Descripcion)
                WDescriII.Text = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                WDescriII.Text = Trim(WDescriII.Text)
                Unidad.Text = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                Unidad.Text = Trim(Unidad.Text)
                rstEnsayo.Close
                WDescri.SetFocus
                    Else
                WDescri.Text = ""
                WDescriII.Text = ""
                Unidad.Text = ""
                WDescri.SetFocus
            End If
            
            Call Conecta_Empresa
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub EnsConsulta_Click()
    Dim IngresaItem As String
    Enspantalla.Clear
    EnsIndice.Clear
    
    XEmpresa = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    spEnsayo = "ListaEnsayos"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        With rstEnsayo
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(rstEnsayo!Codigo) + " " + rstEnsayo!Descripcion
                    Enspantalla.AddItem IngresaItem
                    IngresaItem = rstEnsayo!Codigo
                    EnsIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEnsayo.Close
    End If
    
    Call Conecta_Empresa
    
    Enspantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Enspantalla.Clear
    EnsIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    
    XEmpresa = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ensayos"
    spEnsayos = ZSql
    Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayos.RecordCount > 0 Then
        With rstEnsayos
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Da = Len(rstEnsayos!Descripcion) - WEspacios
                
                    For aa = 1 To Da
                        If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Descripcion), aa, WEspacios) Then
                            IngresaItem = Str$(rstEnsayos!Codigo) + " " + rstEnsayos!Descripcion
                            Enspantalla.AddItem IngresaItem
                            IngresaItem = rstEnsayos!Codigo
                            EnsIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                    
                        Else
                        
                    Exit Do
                
                End If
            Loop
        End With
        rstEnsayos.Close
    End If
    End If
    
    Call Conecta_Empresa

End Sub


Private Sub EnsList_Click()
    Desde.Text = 0
    Hasta.Text = 9999
    Pantalla.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Enspantalla_Click()
    Enspantalla.Visible = False
    Ayuda.Visible = False
    Dim Indice As Integer
    
    Indice = Enspantalla.ListIndex
    WEnsayo.Text = EnsIndice.List(Indice)
    
    XEmpresa = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    spEnsayo = "ConsultaEnsayos " + "'" + WEnsayo.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        WEnsayo.Text = rstEnsayo!Codigo
        WDescri.Text = Trim(rstEnsayo!Descripcion)
        WDescriII.Text = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
        WDescriII.Text = Trim(WDescriII.Text)
        Unidad.Text = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        Unidad.Text = Trim(Unidad.Text)
        rstEnsayo.Close
    End If
    
    Call Conecta_Empresa
    
    WEnsayo.SetFocus
    
End Sub

Private Sub EnsPrimer_Click()

    XEmpresa = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    spEnsayo = "ListaEnsayos"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        With rstEnsayo
            .MoveFirst
            WEnsayo.Text = rstEnsayo!Codigo
            WDescri.Text = Trim(rstEnsayo!Descripcion)
            WDescriII.Text = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
            WDescriII.Text = Trim(WDescriII.Text)
            Unidad.Text = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            Unidad.Text = Trim(Unidad.Text)
            WEnsayo.SetFocus
        End With
        rstEnsayo.Close
    End If
    
    Call Conecta_Empresa
    
End Sub

Private Sub EnsUltimo_Click()

    XEmpresa = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    spEnsayo = "ListaEnsayos"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        With rstEnsayo
            .MoveLast
            WEnsayo.Text = rstEnsayo!Codigo
            WDescri.Text = Trim(rstEnsayo!Descripcion)
            WDescriII.Text = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
            WDescriII.Text = Trim(WDescriII.Text)
            Unidad.Text = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            Unidad.Text = Trim(Unidad.Text)
            WEnsayo.SetFocus
        End With
        rstEnsayo.Close
    End If
    
    Call Conecta_Empresa
    
End Sub

Private Sub EnsSiguiente_Click()

    XEmpresa = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    spEnsayo = "PosteriorEnsayos " + "'" + WEnsayo.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    
    If rstEnsayo.RecordCount > 0 Then
        With rstEnsayo
            .MoveFirst
            WEnsayo.Text = rstEnsayo!Codigo
            WDescri.Text = Trim(rstEnsayo!Descripcion)
            WDescriII.Text = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
            WDescriII.Text = Trim(WDescriII.Text)
            Unidad.Text = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
            Unidad.Text = Trim(Unidad.Text)
        End With
        rstEnsayo.Close
    End If
    
    Call Conecta_Empresa
    
    WEnsayo.SetFocus

End Sub

Sub Ingresa_clave()
    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    XClave.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGraba = "N"
        ZGRABAII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZGRABAII = IIf(IsNull(rstOperador!GrabaII), "", rstOperador!GrabaII)
            rstOperador.Close
        End If
        
        If ZGRABAII = "S" Then
            WGraba = "S"
            XClave.Visible = False
            Select Case ZProceso
                Case 1
                    Call cmdAdd_Click
                Case Else
                    Call cmdDelete_Click
            End Select
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub


