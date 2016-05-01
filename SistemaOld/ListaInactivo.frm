VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaInactivo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Horas Cursadas por Legajo"
   ClientHeight    =   3735
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   3735
   ScaleWidth      =   8085
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.TextBox Horas 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1080
         Width           =   855
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   0
         Top             =   600
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
         Left            =   2520
         TabIndex        =   5
         Top             =   1800
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
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
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
         Left            =   4320
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Horas"
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
         Left            =   720
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
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
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaInactivo.rpt"
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
Attribute VB_Name = "PrgListaInactivo"
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

    ZSql = ""
    ZSql = ZSql + "UPDATE Legajo SET "
    ZSql = ZSql + " Horas = 0 ,"
    ZSql = ZSql + " HorasTotal = 0"
    spLegajo = ZSql
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)

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
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Legajo SET "
        ZSql = ZSql + " Horas = Horas + " + "'" + ZHoras + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZLegajo + "'"
        spLegajo = ZSql
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        
        Sql1 = "Select *"
        Sql2 = " FROM Legajo"
        Sql3 = " Where Legajo.Codigo = " + "'" + ZLegajo + "'"
        spLegajo = Sql1 + Sql2 + Sql3
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
            ZDescripcion = rstLegajo!Descripcion
            rstLegajo.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Legajo SET "
        ZSql = ZSql + " HorasTotal = HorasTotal + " + "'" + ZHoras + "'"
        ZSql = ZSql + " Where Descripcion = " + "'" + ZDescripcion + "'"
        spLegajo = ZSql
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Horas Cursadas por Legajo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Legajo.Renglon} = 1"
    Dos = " and {Legajo.HorasTotal} in " + "0" + " to " + Horas.Text
    Tres = " and not ({Legajo.Descripcion} in " + Chr$(34) + "" + Chr$(34) + ")"
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Legajo.Codigo, Legajo.Renglon, Legajo.Descripcion, Legajo.Perfil, Legajo.ImprePerfil, Legajo.Horas, Legajo.HorasTotal " _
                + "From " _
                + DSQ + ".dbo.Legajo Legajo " _
                + "Where " _
                + "Legajo.Renglon = 1 AND " _
                + "Legajo.Descripcion <> '' AND " _
                + "Legajo.HorasTotal >= " + "0" + " AND " _
                + "Legajo.HorasTotal <= " + Horas.Text
    
    Listado.Connect = Connect()
    Listado.ReportFileName = "WListaInactivos.rpt"
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaInactivo.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Panta.Value = True
    Impresora.Value = False
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Horas.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Horas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
    If KeyAscii = 27 Then
        Horas.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

