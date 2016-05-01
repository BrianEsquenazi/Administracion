VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaEnsayoMp 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Ensayos en Materias Primas"
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
Attribute VB_Name = "PrgListaEnsayoMp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim XParam As String
Dim EmpresaActual As String
Dim ZVector(30, 2) As String
Dim ZVectorII(10000, 5) As String
Dim ZZEnsayo(1000) As String

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
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecificacionesUnifica"
    spEspecificacionesUnifica = ZSql
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        With rstEspecificacionesUnifica
            .MoveFirst
            Do
            
                ZProducto = rstEspecificacionesUnifica!Producto
            
                ZVector(1, 1) = rstEspecificacionesUnifica!Ensayo1
                ZVector(1, 2) = rstEspecificacionesUnifica!Valor1
            
                ZVector(2, 1) = rstEspecificacionesUnifica!Ensayo2
                ZVector(2, 2) = rstEspecificacionesUnifica!valor2
            
                ZVector(3, 1) = rstEspecificacionesUnifica!Ensayo3
                ZVector(3, 2) = rstEspecificacionesUnifica!Valor3
            
                ZVector(4, 1) = rstEspecificacionesUnifica!Ensayo4
                ZVector(4, 2) = rstEspecificacionesUnifica!valor4
            
                ZVector(5, 1) = rstEspecificacionesUnifica!Ensayo5
                ZVector(5, 2) = rstEspecificacionesUnifica!valor5
            
                ZVector(6, 1) = rstEspecificacionesUnifica!Ensayo6
                ZVector(6, 2) = rstEspecificacionesUnifica!valor6
            
                ZVector(7, 1) = rstEspecificacionesUnifica!Ensayo7
                ZVector(7, 2) = rstEspecificacionesUnifica!valor7
            
                ZVector(8, 1) = rstEspecificacionesUnifica!Ensayo8
                ZVector(8, 2) = rstEspecificacionesUnifica!valor8
            
                ZVector(9, 1) = rstEspecificacionesUnifica!Ensayo9
                ZVector(9, 2) = rstEspecificacionesUnifica!valor9
            
                ZVector(10, 1) = rstEspecificacionesUnifica!Ensayo10
                ZVector(10, 2) = rstEspecificacionesUnifica!valor10
                        
                ZVector(11, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
                ZVector(11, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
                        
                ZVector(12, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
                ZVector(12, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
                        
                ZVector(13, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
                ZVector(13, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
                        
                ZVector(14, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
                ZVector(14, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
                        
                ZVector(15, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
                ZVector(15, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
                        
                ZVector(16, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
                ZVector(16, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
                        
                ZVector(17, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
                ZVector(17, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
                        
                ZVector(18, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
                ZVector(18, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
                        
                ZVector(19, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
                ZVector(19, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
                        
                ZVector(20, 1) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
                ZVector(20, 2) = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
                
                For Ciclo = 1 To 20
                    
                    If Val(Desde.Text) <= Val(ZVector(Ciclo, 1)) And Val(Hasta.Text) >= Val(ZVector(Ciclo, 1)) Then
                    
                        If Val(ZVector(Ciclo, 1)) <> 0 Then
                        
                            ZEnsayo = ZVector(Ciclo, 1)
                            ZResultado = ZVector(Ciclo, 2)
                            
                            Lugar = Lugar + 1
                            
                            ZVectorII(Lugar, 1) = ZEnsayo
                            ZVectorII(Lugar, 2) = ZProducto
                            ZVectorII(Lugar, 3) = ZDescripcion
                            ZVectorII(Lugar, 4) = ZResultado
                            ZVectorII(Lugar, 5) = ZEmpresa
                            
                        End If
                    
                    End If
                    
                Next Ciclo
                    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstEspecificacionesUnifica.Close
        
    End If
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecificacionesUnificaIII"
    spEspecificacionesUnificaIII = ZSql
    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
        With rstEspecificacionesUnificaIII
            .MoveFirst
            Do
            
                ZProducto = rstEspecificacionesUnificaIII!Producto
            
                ZVector(1, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
                ZVector(1, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor21), "", rstEspecificacionesUnificaIII!Valor21)
                        
                ZVector(2, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
                ZVector(2, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor22), "", rstEspecificacionesUnificaIII!Valor22)
                        
                ZVector(3, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
                ZVector(3, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor23), "", rstEspecificacionesUnificaIII!Valor23)
                        
                ZVector(4, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
                ZVector(4, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor24), "", rstEspecificacionesUnificaIII!Valor24)
                        
                ZVector(5, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
                ZVector(5, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor25), "", rstEspecificacionesUnificaIII!Valor25)
                        
                ZVector(6, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
                ZVector(6, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor26), "", rstEspecificacionesUnificaIII!Valor26)
                        
                ZVector(7, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
                ZVector(7, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor27), "", rstEspecificacionesUnificaIII!Valor27)
                        
                ZVector(8, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
                ZVector(8, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor28), "", rstEspecificacionesUnificaIII!Valor28)
                        
                ZVector(9, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
                ZVector(9, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor29), "", rstEspecificacionesUnificaIII!Valor29)
                        
                ZVector(10, 1) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
                ZVector(10, 2) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor30), "", rstEspecificacionesUnificaIII!Valor30)
                
                For Ciclo = 1 To 10
                    
                    If Val(Desde.Text) <= Val(ZVector(Ciclo, 1)) And Val(Hasta.Text) >= Val(ZVector(Ciclo, 1)) Then
                    
                        If Val(ZVector(Ciclo, 1)) <> 0 Then
                        
                            ZEnsayo = ZVector(Ciclo, 1)
                            ZResultado = ZVector(Ciclo, 2)
                            
                            Lugar = Lugar + 1
                            
                            ZVectorII(Lugar, 1) = ZEnsayo
                            ZVectorII(Lugar, 2) = ZProducto
                            ZVectorII(Lugar, 3) = ZDescripcion
                            ZVectorII(Lugar, 4) = ZResultado
                            ZVectorII(Lugar, 5) = ZEmpresa
                            
                        End If
                    
                    End If
                    
                Next Ciclo
                    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstEspecificacionesUnificaIII.Close
        
    End If
    
    
    
    
    
    
    
    
    
    Erase ZZEnsayo
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ensayos"
    spEnsayos = ZSql
    Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayos.RecordCount > 0 Then
        With rstEnsayos
            .MoveFirst
            Do
                ZZEnsayo(rstEnsayos!Codigo) = Trim(rstEnsayos!Descripcion)
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstEnsayos.Close
    End If
    
    Call Conecta_Empresa
        
    For Ciclo = 1 To Lugar
    
        ZEnsayo = ZVectorII(Ciclo, 1)
        ZProducto = ZVectorII(Ciclo, 2)
        ZDescripcion = Trim(ZVectorII(Ciclo, 3))
        ZResultado = Left$(Trim(ZVectorII(Ciclo, 4)), 50)
        ZEmpresa = ZVectorII(Ciclo, 5)
        ZDesEnsayo = ZZEnsayo(Val(ZEnsayo))
        
        ZEntra = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select Codigo, Descripcion"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " WHERE Codigo = " + "'" + ZProducto + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZDescripcion = rstArticulo!Descripcion
            ZEntra = "S"
            rstArticulo.Close
        End If
                    
        If ZEntra = "N" Then
        
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
            ZSql = ZSql + "Select Codigo, Descripcion"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " WHERE Codigo = " + "'" + ZProducto + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZDescripcion = rstArticulo!Descripcion
                ZEntra = "S"
                rstArticulo.Close
            End If

            Call Conecta_Empresa
    
        End If
        
        If ZEntra = "S" Then
            
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
            
        End If
                    
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Ensayos en Materia Prima"
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
    
    Listado.ReportFileName = "ListaEnsayoMp.rpt"
    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    Desde.SetFocus
    PrgListaEnsayoMp.Hide
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





