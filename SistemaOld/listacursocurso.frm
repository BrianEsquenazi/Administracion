VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaCursoCurso 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cursos Ralizados por Curso"
   ClientHeight    =   6855
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6855
   ScaleWidth      =   8085
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5640
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   1560
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   5655
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
         Left            =   1920
         TabIndex        =   18
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Ano 
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
         TabIndex        =   2
         Top             =   1440
         Width           =   855
      End
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
         Top             =   960
         Width           =   975
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
         Top             =   480
         Width           =   975
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
         Left            =   3120
         TabIndex        =   8
         Top             =   2400
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
         Left            =   1440
         TabIndex        =   7
         Top             =   2400
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
         Left            =   4320
         TabIndex        =   6
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
         Left            =   4320
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Listado"
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
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Image Consulta 
         Height          =   480
         Left            =   4680
         MouseIcon       =   "listacursocurso.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "listacursocurso.frx":030A
         ToolTipText     =   "Consulta de Datos"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
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
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Curso"
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
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Curso"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaCursoLegajo.rpt"
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
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgListaCursoCurso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLegajo As Recordset
Dim spLegajo As String
Dim rstCurso As Recordset
Dim spCurso As String
Dim rstCursadas As Recordset
Dim spCursadas As String

Dim ZVector(10000, 5) As String
Dim XParam As String

Private Sub Acepta_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE Cronograma SET "
    ZSql = ZSql + " Realizado = 0"
    ZSql = ZSql + " Where Ano = " + "'" + Ano.Text + "'"
    spCronograma = ZSql
    Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)

    WRenglon = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cursadas"
    ZSql = ZSql + " Order by Cursadas.Clave"
    rsCursadas = ZSql
    Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
    If rstCursadas.RecordCount > 0 Then
        With rstCursadas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZFecha = Mid$(rstCursadas!Fecha, 7, 4)
                    
                    If Val(ZFecha) = Val(Ano.Text) Then
                
                        WRenglon = WRenglon + 1
                        
                        ZVector(WRenglon, 1) = rstCursadas!Curso
                        ZVector(WRenglon, 2) = rstCursadas!Legajo
                        ZVector(WRenglon, 3) = Str$(rstCursadas!Horas)
                        ZVector(WRenglon, 4) = rstCursadas!Fecha
                        ZVector(WRenglon, 5) = rstCursadas!Clave
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCursadas.Close
    End If
    
    For Ciclo = 1 To WRenglon
    
        ZCurso = ZVector(Ciclo, 1)
        ZLegajo = ZVector(Ciclo, 2)
        ZHoras = ZVector(Ciclo, 3)
        ZFecha = ZVector(Ciclo, 4)
        ZClave = ZVector(Ciclo, 5)
        
        ZAno = Mid$(ZFecha, 7, 4)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cronograma"
        ZSql = ZSql + " Where Ano = " + "'" + ZAno + "'"
        ZSql = ZSql + " and Legajo = " + "'" + ZLegajo + "'"
        ZSql = ZSql + " and Curso = " + "'" + ZCurso + "'"
        rsCursadas = ZSql
        Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
        If rstCursadas.RecordCount > 0 Then
            rstCursadas.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Cronograma SET "
            ZSql = ZSql + " Realizado = Realizado + " + "'" + ZHoras + "'"
            ZSql = ZSql + " Where Ano = " + "'" + ZAno + "'"
            ZSql = ZSql + " and Legajo = " + "'" + ZLegajo + "'"
            ZSql = ZSql + " and Curso = " + "'" + ZCurso + "'"
            spCronograma = ZSql
            Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cursadas SET "
            ZSql = ZSql + " TipoCursada = 0"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spCursadas = ZSql
            Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cursadas SET "
            ZSql = ZSql + " TipoCursada = 1"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spCursadas = ZSql
            Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo

    Listado.WindowTitle = "Listado de Cursos por Curso"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WDesdeFecha = Ano.Text + "0101"
    WHastaFecha = Ano.Text + "1231"
    
    Uno = "{Cursadas.Curso} in " + Desde.Text + " to " + Hasta.Text + " and "
    Select Case Tipo.ListIndex
         Case 1
            Dos = "{Cursadas.TipoCursada} in 0 to 0 and"
         Case 2
            Dos = "{Cursadas.TipoCursada} in 1 to 1 and"
         Case Else
            Dos = "{Cursadas.TipoCursada} in 0 to 9999 and"
    End Select
    Tres = "{Cursadas.OrdFecha} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
   
    If impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case Tipo.ListIndex
         Case 1
            Listado.SQLQuery = "SELECT Cursadas.Codigo, Cursadas.CursoCursadas.Fecha, Cursadas.OrdFecha, Cursadas.Horas, Cursadas.Legajo, Cursadas.DesLegajo, Cursadas.Observaciones, Cursadas.TipoCursada, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cursadas Cursadas, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cursadas.Curso = Curso.Codigo AND " _
                + "Cursadas.Curso >= " + Desde.Text + " AND " _
                + "Cursadas.Curso <= " + Hasta.Text + " AND " _
                + "Cursadas.OrdFecha >= '" + WDesdeFecha + "' AND " _
                + "Cursadas.OrdFecha <= '" + WHastaFecha + "' AND " _
                + "Cursadas.TipoCursada >= 0 AND " _
                + "Cursadas.TipoCursada <= 0"
         Case 2
            Listado.SQLQuery = "SELECT Cursadas.Codigo, Cursadas.CursoCursadas.Fecha, Cursadas.OrdFecha, Cursadas.Horas, Cursadas.Legajo, Cursadas.DesLegajo, Cursadas.Observaciones, Cursadas.TipoCursada, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cursadas Cursadas, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cursadas.Curso = Curso.Codigo AND " _
                + "Cursadas.Curso >= " + Desde.Text + " AND " _
                + "Cursadas.Curso <= " + Hasta.Text + " AND " _
                + "Cursadas.OrdFecha >= '" + WDesdeFecha + "' AND " _
                + "Cursadas.OrdFecha <= '" + WHastaFecha + "' AND " _
                + "Cursadas.TipoCursada >= 1 AND " _
                + "Cursadas.TipoCursada <= 1"
         Case Else
            Listado.SQLQuery = "SELECT Cursadas.Codigo, Cursadas.CursoCursadas.Fecha, Cursadas.OrdFecha, Cursadas.Horas, Cursadas.Legajo, Cursadas.DesLegajo, Cursadas.Observaciones, Cursadas.TipoCursada, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cursadas Cursadas, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cursadas.Curso = Curso.Codigo AND " _
                + "Cursadas.Curso >= " + Desde.Text + " AND " _
                + "Cursadas.Curso <= " + Hasta.Text + " AND " _
                + "Cursadas.OrdFecha >= '" + WDesdeFecha + "' AND " _
                + "Cursadas.OrdFecha <= '" + WHastaFecha + "' AND " _
                + "Cursadas.TipoCursada >= 0 AND " _
                + "Cursadas.TipoCursada <= 9999"
        End Select
    
    Listado.Connect = Connect()
    Listado.ReportFileName = "WListaCursoCurso.rpt"
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaCursoCurso.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Cursos"
     
     Opcion.ListIndex = 0

     Rem Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Curso"
            Sql3 = " Order by Curso.Codigo"
            spCurso = Sql1 + Sql2 + Sql3
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                With rstCurso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstCurso!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstCurso!Descripcion
                            IngresaItem = rstCurso!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCurso.Close
            End If
            
        Case Else
    End Select
    
    Pantalla.TopRow = 1
    Pantalla.Col = 1
    Pantalla.Row = 1
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.Row - 1
            Desde.Text = WIndice.List(Indice)
            Hasta.Text = WIndice.List(Indice)
            Call Desde_KeyPress(13)
            
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Planificado"
    Tipo.AddItem "No Planificado"
    
    Tipo.ListIndex = 0

    panta.Value = True
    impresora.Value = False
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            Pantalla.FixedCols = 1
            Pantalla.Cols = 3
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
    End Select
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    Select Case XIndice
        Case 0
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Codigo"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        Pantalla.Text = "Nombre"
                        Pantalla.ColWidth(Ciclo) = 6000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
        Case Else
            
    End Select
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = Pantalla.Text
        WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        WTitulo(Ciclo).Width = Pantalla.CellWidth
        WTitulo(Ciclo).Height = Pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.TopRow = 1
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Curso"
            Sql3 = " Order by Curso.Codigo"
            spCurso = Sql1 + Sql2 + Sql3
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                With rstCurso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstCurso!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstCurso!Descripcion, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstCurso!Codigo
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstCurso!Descripcion
                                    IngresaItem = rstCurso!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCurso.Close
            End If
            
            Pantalla.TopRow = 1
            Pantalla.Col = 1
            Pantalla.Row = 1
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

