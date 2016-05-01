VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgEnv 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Envases"
   ClientHeight    =   4965
   ClientLeft      =   3165
   ClientTop       =   1845
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   6615
   Begin VB.TextBox Peso 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   32
      Text            =   " "
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Tipo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   30
      Text            =   " "
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Kilos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   29
      Text            =   " "
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Abreviatura 
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   28
      Text            =   " "
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Envases 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   480
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   24
         Text            =   " "
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4440
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wenvases.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "PrgEnvase.frx":0000
      Left            =   480
      List            =   "PrgEnvase.frx":0007
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   1800
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   600
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles"
      Height          =   1335
      Left            =   4320
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   4
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "Peso"
      Height          =   375
      Left            =   3960
      TabIndex        =   33
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo"
      Height          =   375
      Left            =   2040
      TabIndex        =   31
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Kilos"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Abreviatura"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Envase"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnvases As Recordset
Dim spEnvases As String
Dim XParam As String
Dim Empe(12, 2) As String

Dim WEnvases As Integer
Dim WKilos As Integer

Private Sub Acepta_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Rem With rstAuxiliar
    Rem     .Index = "Clave"
    Rem     .Seek "=", 1
    Rem     If .NoMatch = False Then
    Rem         .Edit
    Rem         !Nombre = WAuxiliar
    Rem         .Update
    Rem     End If
    Rem End With
    
    Listado.WindowTitle = "Listado de Envases"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Envases.Envases} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Envases.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Envases.Envases , Envases.Descripcion, Envases.Abreviatura, Envases.Kilos " _
                        + "From " + DSQ + ".dbo.Envases Envases " _
                        + "Where Envases.Envases >= 0 AND Envases.Envases <= 9999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()

    WEnvases = Val(Envases.Text)
    WKilos = Val(Kilos.Text)
    Peso.Text = Pusing("###,###.##", Peso.Text)
    
    If WEnvases <> 0 Then
    
        XEmpresa = WEmpresa
    
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0002"
        Empe(2, 2) = "Empresa02"
        Empe(3, 1) = "0003"
        Empe(3, 2) = "Empresa03"
        Empe(4, 1) = "0004"
        Empe(4, 2) = "Empresa04"
        Empe(5, 1) = "0005"
        Empe(5, 2) = "Empresa05"
        Empe(6, 1) = "0006"
        Empe(6, 2) = "Empresa06"
        Empe(7, 1) = "0007"
        Empe(7, 2) = "Empresa07"
        Empe(8, 1) = "0008"
        Empe(8, 2) = "Empresa08"
        Empe(9, 1) = "0009"
        Empe(9, 2) = "Empresa09"
        Empe(10, 1) = "0010"
        Empe(10, 2) = "Empresa10"
        Empe(11, 1) = "0011"
        Empe(11, 2) = "Empresa11"
    
        For a = 1 To 11
    
            WEmpresa = Empe(a, 1)
            txtOdbc = Empe(a, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

            spEnvases = "ConsultaEnvases " + "'" + Envases.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
            
                rstEnvases.Close
                XParam = "'" + Envases.Text + "','" + Descripcion.Text + "','" + Abreviatura.Text + "','" _
                             + Kilos.Text + "'"
                Set rstEnvases = db.OpenRecordset("ModificaEnvases " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
                    
                XParam = "'" + Envases.Text + "','" + Descripcion.Text + "','" + Abreviatura.Text + "','" _
                             + Kilos.Text + "'"
                Set rstEnvases = db.OpenRecordset("AltaEnvase " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Envases SET "
            ZSql = ZSql + " Tipo = " + "'" + Tipo.Text + "',"
            ZSql = ZSql + " Peso = " + "'" + Peso.Text + "'"
            ZSql = ZSql + " Where Envases = " + "'" + Envases.Text + "'"
            spEnvases = ZSql
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
        Next a
        
        Call Conecta_Empresa
            
        Call CmdLimpiar_Click
        Envases.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    WEnvases = Envases.Text
    
    If WEnvases <> 0 Then
    
        spEnvases = "ConsultaEnvases " + "'" + Envases.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        If WPasa = "S" Then
            T$ = "Borrar Registro"
            Estilo = Estilo = vbYesNo + vbCritical + vbDefaultButton2
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spEnvases = "BorrarEnvases " + "'" + Envases.Text + "'"
                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    
    End If
    Envases.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Envases.Text = ""
    Descripcion.Text = ""
    Abreviatura.Text = ""
    Kilos.Text = ""
    Tipo.Text = ""
    Peso.Text = ""
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    Envases.SetFocus
    PrgEnv.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spEnvases = "AnteriorEnvases " + "'" + Envases.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstEnvases
        .MoveLast
        Envases.Text = rstEnvases!Envases
        Descripcion.Text = rstEnvases!Descripcion
        Abreviatura.Text = rstEnvases!Abreviatura
        Kilos.Text = rstEnvases!Kilos
        Tipo.Text = IIf(IsNull(rstEnvases!Tipo), "", rstEnvases!Tipo)
        Tipo.Text = Trim(Tipo.Text)
        Peso.Text = IIf(IsNull(rstEnvases!Peso), "", rstEnvases!Peso)
        Peso.Text = Trim(Peso.Text)
        rstEnvases.Close
    End With
    
    Envases.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Envases", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Envases.SetFocus
    
End Sub




Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spEnvases = "PosteriorEnvases " + "'" + Envases.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstEnvases
        .MoveFirst
        Envases.Text = rstEnvases!Envases
        Descripcion.Text = rstEnvases!Descripcion
        Abreviatura.Text = rstEnvases!Abreviatura
        Kilos.Text = rstEnvases!Kilos
        Tipo.Text = IIf(IsNull(rstEnvases!Tipo), "", rstEnvases!Tipo)
        Tipo.Text = Trim(Tipo.Text)
        Peso.Text = IIf(IsNull(rstEnvases!Peso), "", rstEnvases!Peso)
        Peso.Text = Trim(Peso.Text)
        rstEnvases.Close
    End With
    
    Envases.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Envases", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Envases.SetFocus
    
End Sub



Private Sub Lista_Click()
    Desde.Text = "0"
    Hasta.Text = "9999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Abreviatura.SetFocus
    End If
End Sub

Private Sub Abreviatura_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Kilos.SetFocus
    End If
End Sub

Private Sub Kilos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tipo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Peso.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Peso_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
Private Sub Envases_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Envases.Text) <> 0 Then
        
            WEnvases = Envases.Text
            spEnvases = "ConsultaEnvases " + "'" + Envases.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                Envases.Text = rstEnvases!Envases
                Descripcion.Text = Trim(rstEnvases!Descripcion)
                Abreviatura.Text = Trim(rstEnvases!Abreviatura)
                Kilos.Text = rstEnvases!Kilos
                Tipo.Text = IIf(IsNull(rstEnvases!Tipo), "", rstEnvases!Tipo)
                Tipo.Text = Trim(Tipo.Text)
                Peso.Text = IIf(IsNull(rstEnvases!Peso), "", rstEnvases!Peso)
                Peso.Text = Trim(Peso.Text)
                rstEnvases.Close
                    Else
                WEnvases = Envases.Text
                CmdLimpiar_Click
                Envases.Text = WEnvases
            End If
        
        End If
        Descripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
        
            spEnvases = "ListaEnvases"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstEnvases
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstEnvases!Envases
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEnvases.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Envases.Text = WIndice.List(Indice)
            Call Envases_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spEnvases = "ListaEnvases"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstEnvases
        .MoveFirst
        Envases.Text = rstEnvases!Envases
        Descripcion.Text = rstEnvases!Descripcion
        Abreviatura.Text = rstEnvases!Abreviatura
        Kilos.Text = rstEnvases!Kilos
        Tipo.Text = IIf(IsNull(rstEnvases!Tipo), "", rstEnvases!Tipo)
        Tipo.Text = Trim(Tipo.Text)
        Peso.Text = IIf(IsNull(rstEnvases!Peso), "", rstEnvases!Peso)
        Peso.Text = Trim(Peso.Text)
        rstEnvases.Close
    End With
    
    Envases.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Envases", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Envases.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spEnvases = "ListaEnvases"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstEnvases
        .MoveLast
        Envases.Text = rstEnvases!Envases
        Descripcion.Text = rstEnvases!Descripcion
        Abreviatura.Text = rstEnvases!Abreviatura
        Kilos.Text = rstEnvases!Kilos
        Tipo.Text = IIf(IsNull(rstEnvases!Tipo), "", rstEnvases!Tipo)
        Tipo.Text = Trim(Tipo.Text)
        Peso.Text = IIf(IsNull(rstEnvases!Peso), "", rstEnvases!Peso)
        Peso.Text = Trim(Peso.Text)
        rstEnvases.Close
        Envases.SetFocus
    End With
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Envases", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Envases.SetFocus
 End Sub


Private Sub Form_Load()

    Envases.Text = ""
    Descripcion.Text = ""
    Abreviatura.Text = ""
    Kilos.Text = ""
    Tipo.Text = ""
    Peso.Text = ""

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEnv.Caption = "Ingreso de Envases :  " + !Nombre
        End If
    End With
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

