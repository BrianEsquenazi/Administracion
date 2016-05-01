VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaEnsayoPt 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Ensayos en Producto Terminado"
   ClientHeight    =   2505
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   1
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   855
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
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
         Left            =   360
         TabIndex        =   6
         Top             =   1440
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
         Left            =   3360
         TabIndex        =   5
         Top             =   360
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
         Left            =   3360
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Ensayo"
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Ensayo"
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
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7560
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaEnsayoMp.rpt"
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
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaEnsayoPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim XParam As String
Dim EmpresaActual As String
Dim ZVector(10, 2) As String
Dim ZVectorII(10000, 5) As String

Private Sub Acepta_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            ZEmpresa = !Nombre
        End If
    End With
    
    
    ZSql = "DELETE VerificaEnsayo"
    spVerificaEnsayo = ZSql
    Set rstVerificaEnsayo = db.OpenRecordset(spVerificaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    
    
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
    
    
    Erase ZVectorII
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select producto,Ensayo1,valor1,ensayo2,valor2,ensayo3,valor3,ensayo4,valor4,ensayo5,valor5,ensayo6,valor6,ensayo7,valor7,ensayo8,valor8,ensayo9,valor9,ensayo10,valor10"
    ZSql = ZSql + " FROM EspecifUnifica"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        With rstEspecifUnifica
            .MoveFirst
            Do
            
                ZProducto = rstEspecifUnifica!Producto
            
                ZVector(1, 1) = rstEspecifUnifica!Ensayo1
                ZVector(1, 2) = rstEspecifUnifica!Valor1
            
                ZVector(2, 1) = rstEspecifUnifica!Ensayo2
                ZVector(2, 2) = rstEspecifUnifica!valor2
            
                ZVector(3, 1) = rstEspecifUnifica!Ensayo3
                ZVector(3, 2) = rstEspecifUnifica!Valor3
            
                ZVector(4, 1) = rstEspecifUnifica!Ensayo4
                ZVector(4, 2) = rstEspecifUnifica!valor4
            
                ZVector(5, 1) = rstEspecifUnifica!Ensayo5
                ZVector(5, 2) = rstEspecifUnifica!valor5
            
                ZVector(6, 1) = rstEspecifUnifica!Ensayo6
                ZVector(6, 2) = rstEspecifUnifica!valor6
            
                ZVector(7, 1) = rstEspecifUnifica!Ensayo7
                ZVector(7, 2) = rstEspecifUnifica!valor7
            
                ZVector(8, 1) = rstEspecifUnifica!Ensayo8
                ZVector(8, 2) = rstEspecifUnifica!valor8
            
                ZVector(9, 1) = rstEspecifUnifica!Ensayo9
                ZVector(9, 2) = rstEspecifUnifica!valor9
            
                ZVector(10, 1) = rstEspecifUnifica!Ensayo10
                ZVector(10, 2) = rstEspecifUnifica!valor10
                
                For Ciclo = 1 To 10
                    
                    If Val(Desde.Text) <= Val(ZVector(Ciclo, 1)) And Val(Hasta.Text) >= Val(ZVector(Ciclo, 1)) Then
                    
                        ZEnsayo = ZVector(Ciclo, 1)
                        ZResultado = ZVector(Ciclo, 2)
                        
                        Lugar = Lugar + 1
                        
                        ZVectorII(Lugar, 1) = ZEnsayo
                        ZVectorII(Lugar, 2) = ZProducto
                        ZVectorII(Lugar, 3) = ZDescripcion
                        ZVectorII(Lugar, 4) = ZResultado
                        ZVectorII(Lugar, 5) = ZEmpresa
                        
                    End If
                    
                Next Ciclo
                    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
        
    End If
    
    
    Call Conecta_Empresa
        
        
    For Ciclo = 1 To Lugar
    
        ZEnsayo = ZVectorII(Ciclo, 1)
        ZDesEnsayo = ""
        ZProducto = ZVectorII(Ciclo, 2)
        ZDescripcion = ZVectorII(Ciclo, 3)
        ZResultado = ZVectorII(Ciclo, 4)
        ZEmpresa = ZVectorII(Ciclo, 5)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Terminado"
        ZSql = ZSql + " WHERE Codigo = " + "'" + ZProducto + "'"
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZDescripcion = rstTerminado!Descripcion
            rstTerminado.Close
        End If
            
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
    
        spEnsayo = "Select * FROM Ensayos Where Codigo = " + ZEnsayo
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            ZDesEnsayo = Trim(rstEnsayo!Descripcion)
            rstEnsayo.Close
        End If
        
        Call Conecta_Empresa
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO VerificaEnsayo ("
        ZSql = ZSql + "Ensayo ,"
        ZSql = ZSql + "DesEnsayo ,"
        ZSql = ZSql + "Producto ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Resultado ,"
        ZSql = ZSql + "Empresa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZEnsayo + "',"
        ZSql = ZSql + "'" + ZDesEnsayo + "',"
        ZSql = ZSql + "'" + ZProducto + "',"
        ZSql = ZSql + "'" + ZDescripcion + "',"
        ZSql = ZSql + "'" + ZResultado + "',"
        ZSql = ZSql + "'" + ZEmpresa + "')"
        
        spVerificaEnsayo = ZSql
        Set rstVerificaEnsayo = db.OpenRecordset(spVerificaEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                    
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Ensayos en Producto Terminado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT VerificaEnsayo.Ensayo, VerificaEnsayo.Producto, VerificaEnsayo.Descripcion, VerificaEnsayo.Resultado, VerificaEnsayo.Empresa, VerificaEnsayo.DesEnsayo " _
                + "From " _
                + DSQ + ".dbo.VerificaEnsayo VerificaEnsayo " _
                + "Where " _
                + "VerificaEnsayo.Ensayo >= 0 AND " _
                + "VerificaEnsayo.Ensayo <= 9999"
    
    Listado.ReportFileName = "ListaEnsayoPt.rpt"
    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    Desde.SetFocus
    PrgListaEnsayoPt.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub




