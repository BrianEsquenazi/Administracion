VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListhoja 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Hojas de Produccion"
   ClientHeight    =   3345
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3345
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox Tipo 
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
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   4215
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
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
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   2160
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
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   2160
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
         Left            =   3480
         TabIndex        =   5
         Top             =   240
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
         Left            =   3480
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   840
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wlisthoja.rpt"
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
Attribute VB_Name = "PrgListhoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(5000, 3) As String

Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim ZReal As Double
Dim ZRealAnt As Double

Private Sub Acepta_Click()
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)

    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + " WImporte = " + " '" + "0" + "',"
    ZSql = ZSql + " Lista = " + " '" + "N" + "',"
    ZSql = ZSql + " Suma1 = " + " '" + "0" + "',"
    ZSql = ZSql + " Suma2 = " + " '" + "0" + "',"
    ZSql = ZSql + " Suma3 = " + " '" + "0" + "',"
    ZSql = ZSql + " Suma4 = " + " '" + "0" + "',"
    ZSql = ZSql + " Suma5 = " + " '" + "0" + "',"
    ZSql = ZSql + " Suma6 = " + " '" + "0" + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    ZLugar = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Fechaingord >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Fechaingord <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Renglon = 2"
    ZSql = ZSql + " Order by Clave"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    ZReal = IIf(IsNull(rstHoja!Real), "0", rstHoja!Real)
                    ZRealAnt = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
                    ZVector(ZLugar, 1) = Str$(rstHoja!Hoja)
                    If ZRealAnt > 0 Then
                        ZVector(ZLugar, 2) = Str$(ZRealAnt)
                            Else
                        ZVector(ZLugar, 2) = Str$(ZReal)
                    End If
                    ZVector(ZLugar, 3) = rstHoja!Producto
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If


    For Ciclo = 1 To ZLugar
    
        ZHoja = ZVector(Ciclo, 1)
        ZReal = Val(ZVector(Ciclo, 2))
        ZProducto = ZVector(Ciclo, 3)
        
        ZSuma1 = 0
        ZSuma2 = 0
        ZSuma3 = 0
        ZSuma4 = 0
        ZSuma5 = 0
        ZSuma6 = 0
        
        ZTipo = Left$(ZProducto, 2)
        Select Case ZTipo
            Case "RE"
                ZSuma2 = 1
                ZSuma5 = ZReal
            Case "NK"
                ZSuma3 = 1
                ZSuma6 = ZReal
            Case Else
                ZSuma1 = 1
                ZSuma4 = ZReal
        End Select
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Lista = " + " '" + "S" + "',"
        ZSql = ZSql + " Suma1 = " + " '" + Str$(ZSuma1) + "',"
        ZSql = ZSql + " Suma2 = " + " '" + Str$(ZSuma2) + "',"
        ZSql = ZSql + " Suma3 = " + " '" + Str$(ZSuma3) + "',"
        ZSql = ZSql + " Suma4 = " + " '" + Str$(ZSuma4) + "',"
        ZSql = ZSql + " Suma5 = " + " '" + Str$(ZSuma5) + "',"
        ZSql = ZSql + " Suma6 = " + " '" + Str$(ZSuma6) + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    Sql1 = "UPDATE Hoja SET "
    Sql2 = " WImporte = Real"
    Sql3 = " Where Fechaingord >= " + "'" + WDesde + "'"
    Sql4 = " and Fechaingord <= " + "'" + WHasta + "'"
    Rem Sql5 = " and Realant IS NULL"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Sql1 = "UPDATE Hoja SET "
    Sql2 = " WImporte = Realant"
    Sql3 = " Where Fechaingord >= " + "'" + WDesde + "'"
    Sql4 = " and Fechaingord <= " + "'" + WHasta + "'"
    Sql5 = " and Realant <> " + "'" + "0" + "'"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem With rstHoja
    Rem     .Index = "Clave"
    Rem     .MoveFirst1
    Rem     Do
    Rem         If .EOF = False Then
    Rem             .Edit
    Rem             !WImporte = 0
    Rem             If !Renglon = 1 Then
    Rem                 If !fechaingord >= WDesde And !fechaingord <= WHasta Then
    Rem                     !WImporte = !Real
    Rem                 End If
    Rem             End If
    Rem             .Update
    Rem            .MoveNext
    Rem                 Else
    Rem             Exit Do
    Rem         End If
    Rem     Loop
    Rem End With
    
    Listado.WindowTitle = "Listado de Hoja de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Hoja.fechaingord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34) + " and {Hoja.renglon} = 1"
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Select Case Tipo.ListIndex
        Case 0
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Renglon, Hoja.Producto, Hoja.FechaIng, Hoja.FechaIngOrd, Hoja.WImporte, " _
                                + "Terminado.Descripcion " _
                                + "From " _
                                + DSQ + ".dbo.Hoja Hoja, " _
                                + DSQ + ".dbo.Terminado Terminado " _
                                + "Where " _
                                + "Hoja.Producto = Terminado.Codigo AND " _
                                + "Hoja.Renglon = 1 AND " _
                                + "Hoja.FechaIngOrd >= '" + WDesde + "' AND " _
                                + "Hoja.FechaIngOrd <= '" + WHasta + "'"
    
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
            Listado.Connect = Connect()
            Listado.ReportFileName = "WListhojaProd.rpt"
            
        Case 1
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Renglon, Hoja.Producto, Hoja.FechaIng, Hoja.FechaIngOrd, Hoja.WImporte, Terminado.Descripcion " _
                        + "From " + DSQ + ".dbo.Hoja Hoja, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where Hoja.Producto = Terminado.Codigo AND Hoja.Renglon = 1 AND Hoja.FechaIngOrd >= '" + WDesde + "' AND Hoja.FechaIngOrd <= '" + WHasta + "'"
    
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
            Listado.Connect = Connect()
            Listado.ReportFileName = "WListhoja.rpt"
            
        Case Else
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Renglon, Hoja.Producto, Hoja.FechaIng, Hoja.FechaIngOrd, Hoja.WImporte, Hoja.Lista, Hoja.Suma1, Hoja.Suma2, Hoja.Suma3, Hoja.Suma4, Hoja.Suma5, Hoja.Suma6, " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Hoja Hoja, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "Hoja.Producto = Terminado.Codigo AND " _
                    + "Hoja.Renglon = 1 AND " _
                    + "Hoja.FechaIngOrd >= '" + WDesde + "' AND " _
                    + "Hoja.FechaIngOrd <= '" + WHasta + "' AND " _
                    + "Hoja.Lista = 'S'"
    
            Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
            Listado.Connect = Connect()
            Listado.ReportFileName = "WListHojaProductividad.rpt"
    
    End Select
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgListhoja.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Por Producto Terminado"
    Tipo.AddItem "Por Hoja de produccion"
    Tipo.AddItem "Por Hoja de produccion (Solo Produccion)"
    
    Tipo.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListhoja.Caption = "Listado de Hoja de Produuccion :  " + !Nombre
        End If
    End With
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


