VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgConsultaCursos 
   AutoRedraw      =   -1  'True
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   11880
   Begin Crystal.CrystalReport Listado 
      Left            =   1680
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
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
      Height          =   735
      Left            =   6720
      TabIndex        =   5
      Top             =   7080
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
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
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
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6120
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   6000
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgConsultaCursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCurso As Recordset
Dim spCurso As String
Dim XParam As String
Dim ZZLugar As Integer

Dim ZZAyuda(1000) As String

Private Sub Cancela_click()
    PrgConsultaCursos.Hide
    Unload Me
    PrgAgendaTotal.Show
End Sub

Sub Form_Load()

    Call Limpia_Ayuda
    
    ZLugar = 0

    For Ciclo = 1 To 1000
    
        If Val(ZZPasaDatos(Ciclo, 1)) <> 0 Then
        
            ZLegajo = ZZPasaDatos(Ciclo, 1)
            ZCurso = ZZPasaDatos(Ciclo, 2)
            
            ZLugar = ZLugar + 1
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Legajo"
            ZSql = ZSql + " Where Legajo.Codigo = " + "'" + ZLegajo + "'"
            spLegajo = ZSql
            Set rstlegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstlegajo.RecordCount > 0 Then
                Pantalla.TextMatrix(ZLugar, 1) = ZLegajo
                Pantalla.TextMatrix(ZLugar, 2) = rstlegajo!Descripcion
                rstlegajo.Close
            End If
            
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Curso"
            ZSql = ZSql + " Where Curso.Codigo = " + "'" + ZCurso + "'"
            spCurso = ZSql
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                Pantalla.TextMatrix(ZLugar, 3) = ZCurso
                Pantalla.TextMatrix(ZLugar, 4) = rstCurso!Descripcion
                rstCurso.Close
            End If
            
        End If
        
    Next Ciclo
    
End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla

    Pantalla.FixedCols = 1
    Pantalla.Cols = 5
    Pantalla.FixedRows = 1
    Pantalla.Rows = 10001
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        Select Case Ciclo
            Case 1
                Pantalla.Text = "Legajo"
                Pantalla.ColWidth(Ciclo) = 1000
                Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                Pantalla.Text = "Nombre"
                Pantalla.ColWidth(Ciclo) = 4000
                Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Pantalla.Text = "Curso"
                Pantalla.ColWidth(Ciclo) = 1000
                Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                Pantalla.Text = "Descripcion"
                Pantalla.ColWidth(Ciclo) = 4000
                Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        Rem WTitulo(Ciclo).Text = Pantalla.Text
        Rem WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        Rem WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        Rem WTitulo(Ciclo).Width = Pantalla.CellWidth
        Rem WTitulo(Ciclo).Height = Pantalla.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Rem Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub


Private Sub Impre_Click()


    Rem On Error GoTo WError
    
    
    ZSql = "DELETE ImpreCursos"
    ZSql = ZSql + " Where ImpreCursos.Codigo = " + "'" + ZZOperadorResponsable + "'"
    spImpreCursos = ZSql
    Set rstImpreCursos = db.OpenRecordset(spImpreCursos, dbOpenSnapshot, dbSQLPassThrough)
    
    For Ciclo = 1 To 1000
    
        If Val(ZZPasaDatos(Ciclo, 1)) <> 0 Then
        
            ZLegajo = Pantalla.TextMatrix(Ciclo, 1)
            ZNombre = Trim(Pantalla.TextMatrix(Ciclo, 2))
            ZCurso = Pantalla.TextMatrix(Ciclo, 3)
            ZDescripcion = Trim(Pantalla.TextMatrix(Ciclo, 4))
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreCursos ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Responsable ,"
            ZSql = ZSql + "Legajo ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Curso ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZOperadorResponsable + "',"
            ZSql = ZSql + "'" + ZZOperadorResponsableNombre + "',"
            ZSql = ZSql + "'" + ZLegajo + "',"
            ZSql = ZSql + "'" + ZNombre + "',"
            ZSql = ZSql + "'" + ZCurso + "',"
            ZSql = ZSql + "'" + ZDescripcion + "')"
            
            spImpreCursos = ZSql
            Set rstImpreCursos = db.OpenRecordset(spImpreCursos, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    Listado.WindowTitle = "Consulta de Cursos Pendientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ImpreCursos.Codigo, ImpreCursos.Responsable, ImpreCursos.Legajo, ImpreCursos.Nombre, ImpreCursos.Curso, ImpreCursos.Descripcion " _
                + "From " _
                + DSQ + ".dbo.ImpreCursos ImpreCursos " _
                + "Where " _
                + "ImpreCursos.Codigo >= " + ZZOperadorResponsable + " AND " _
                + "ImpreCursos.Codigo <= " + ZZOperadorResponsable
                    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{ImpreCursos.Codigo} in " + ZZOperadorResponsable + " to " + ZZOperadorResponsable
    Listado.SelectionFormula = "{ImpreCursos.Codigo} in " + ZZOperadorResponsable + " to " + ZZOperadorResponsable
    
    Listado.Destination = 1
    Listado.Destination = 0
    
    Listado.ReportFileName = "ImpreCursos.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
     Resume Next
    



End Sub
