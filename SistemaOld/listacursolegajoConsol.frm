VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaCursoLegajoConsol 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Temas por Legajo Consolidado"
   ClientHeight    =   6690
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6690
   ScaleWidth      =   8085
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   975
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
      Left            =   1800
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5880
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
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.TextBox AnoII 
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
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
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
         Left            =   2160
         TabIndex        =   14
         Top             =   840
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
         Left            =   2160
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
         Left            =   3240
         TabIndex        =   5
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
         Left            =   1560
         TabIndex        =   4
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
         Left            =   4560
         TabIndex        =   3
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
         Left            =   4560
         TabIndex        =   2
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
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1455
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
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   855
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
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4895
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgListaCursoLegajoConsol"
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

Dim ZVector(10000, 6) As String
Dim XParam As String

Private Sub Acepta_Click()

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
                    
                    If Val(ZFecha) >= Val(Ano.Text) And Val(ZFecha) <= Val(AnoII.Text) Then
                
                        WRenglon = WRenglon + 1
                        
                        ZVector(WRenglon, 1) = rstCursadas!Curso
                        ZVector(WRenglon, 2) = rstCursadas!Legajo
                        ZVector(WRenglon, 3) = Str$(rstCursadas!Horas)
                        ZVector(WRenglon, 4) = rstCursadas!Fecha
                        ZVector(WRenglon, 5) = rstCursadas!Clave
                        ZVector(WRenglon, 6) = rstCursadas!Tema
                        
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
        ZTema = ZVector(Ciclo, 6)
        
        ZAno = Mid$(ZFecha, 7, 4)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cronograma"
        ZSql = ZSql + " Where Ano = " + "'" + ZAno + "'"
        ZSql = ZSql + " and Legajo = " + "'" + ZLegajo + "'"
        ZSql = ZSql + " and Curso = " + "'" + ZCurso + "'"
        ZSql = ZSql + " and tema = " + "'" + ZTema + "'"
        rsCursadas = ZSql
        Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
        If rstCursadas.RecordCount > 0 Then
            rstCursadas.Close
        
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


    Listado.WindowTitle = "Listado de Cursos Realizados por Legajo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WDesdeFecha = Ano.Text + "0101"
    WHastaFecha = AnoII.Text + "1231"
    
    WDesde = "0"
    WHasta = "999999"
    
    Uno = "{Cursadas.Legajo} in " + WDesde + " to " + WHasta + " and "
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
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case Tipo.ListIndex
         Case 1
            Listado.SQLQuery = "SELECT Cursadas.Codigo, Cursadas.Curso, Cursadas.Fecha, Cursadas.OrdFecha, Cursadas.Horas, Cursadas.Legajo, Cursadas.DesLegajo, Cursadas.Observaciones, Cursadas.TipoCursada, Cursadas.Tema, Cursadas.DesTema, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cursadas Cursadas, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cursadas.Curso = Curso.Codigo AND " _
                + "Cursadas.OrdFecha >= '" + WDesdeFecha + "' AND " _
                + "Cursadas.OrdFecha <= '" + WHastaFecha + "' AND " _
                + "Cursadas.Legajo >= " + WDesde + " AND " _
                + "Cursadas.Legajo <= " + WHasta + " AND " _
                + "Cursadas.TipoCursada >= 0 AND " _
                + "Cursadas.TipoCursada <= 0"
                
         Case 2
            Listado.SQLQuery = "SELECT Cursadas.Codigo, Cursadas.Curso, Cursadas.Fecha, Cursadas.OrdFecha, Cursadas.Horas, Cursadas.Legajo, Cursadas.DesLegajo, Cursadas.Observaciones, Cursadas.TipoCursada, Cursadas.Tema, Cursadas.DesTema, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cursadas Cursadas, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cursadas.Curso = Curso.Codigo AND " _
                + "Cursadas.OrdFecha >= '" + WDesdeFecha + "' AND " _
                + "Cursadas.OrdFecha <= '" + WHastaFecha + "' AND " _
                + "Cursadas.Legajo >= " + WDesde + " AND " _
                + "Cursadas.Legajo <= " + WHasta + " AND " _
                + "Cursadas.TipoCursada >= 1 AND " _
                + "Cursadas.TipoCursada <= 1"
                
         Case Else
            Listado.SQLQuery = "SELECT Cursadas.Codigo, Cursadas.Curso, Cursadas.Fecha, Cursadas.OrdFecha, Cursadas.Horas, Cursadas.Legajo, Cursadas.DesLegajo, Cursadas.Observaciones, Cursadas.TipoCursada, Cursadas.Tema, Cursadas.DesTema, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cursadas Cursadas, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cursadas.Curso = Curso.Codigo AND " _
                + "Cursadas.OrdFecha >= '" + WDesdeFecha + "' AND " _
                + "Cursadas.OrdFecha <= '" + WHastaFecha + "' AND " _
                + "Cursadas.Legajo >= " + WDesde + " AND " _
                + "Cursadas.Legajo <= " + WHasta + " AND " _
                + "Cursadas.TipoCursada >= 0 AND " _
                + "Cursadas.TipoCursada <= 9999"
    End Select
        
    
    
    
    Listado.Connect = Connect()
    Listado.ReportFileName = "WListaCursoLegajoConsol.rpt"
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaCursoLegajoConsol.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Planificado"
    Tipo.AddItem "No Planificado"
    
    Tipo.ListIndex = 0
    
    Ano.Text = ""
    AnoII.Text = ""

    Panta.Value = True
    Impresora.Value = False
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AnoII.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub AnoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
    If KeyAscii = 27 Then
        AnoII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
