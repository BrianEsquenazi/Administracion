VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaPlanificacion 
   AutoRedraw      =   -1  'True
   Caption         =   "Plan de Capacitacion Anual"
   ClientHeight    =   2925
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   2925
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5655
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   0
         Top             =   720
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1560
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
         Top             =   1560
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
         Left            =   1320
         TabIndex        =   6
         Top             =   720
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
End
Attribute VB_Name = "PrgListaPlanificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCurso As Recordset
Dim spCurso As String
Dim XParam As String
Dim ZAno As String
Dim ZMes As String
Dim XMes(20) As String
Dim ZVector(1000, 20) As String
Dim ZVectorII(1000, 20) As String

Private Sub Acepta_Click()

    WRenglon = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CronogramaII"
    ZSql = ZSql + " Where CronogramaII.Ano = " + "'" + Ano.Text + "'"
    
    rsCronogramaII = ZSql
    Set rstCronogramaII = db.OpenRecordset(rsCronogramaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCronogramaII.RecordCount > 0 Then
        With rstCronogramaII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    ZVector(WRenglon, 1) = Str$(rstCronogramaII!Curso)
                    ZVector(WRenglon, 2) = Trim(rstCronogramaII!Mes1)
                    ZVector(WRenglon, 3) = Trim(rstCronogramaII!Mes2)
                    ZVector(WRenglon, 4) = Trim(rstCronogramaII!Mes3)
                    ZVector(WRenglon, 5) = Trim(rstCronogramaII!Mes4)
                    ZVector(WRenglon, 6) = Trim(rstCronogramaII!Mes5)
                    ZVector(WRenglon, 7) = Trim(rstCronogramaII!Mes6)
                    ZVector(WRenglon, 8) = Trim(rstCronogramaII!Mes7)
                    ZVector(WRenglon, 9) = Trim(rstCronogramaII!Mes8)
                    ZVector(WRenglon, 10) = Trim(rstCronogramaII!Mes9)
                    ZVector(WRenglon, 11) = Trim(rstCronogramaII!Mes10)
                    ZVector(WRenglon, 12) = Trim(rstCronogramaII!Mes11)
                    ZVector(WRenglon, 13) = Trim(rstCronogramaII!Mes12)
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCronogramaII.Close
    End If
    
    ZAno = Ano.Text
    Call Ceros(ZAno, 4)
    
    For Ciclo = 1 To WRenglon
    
        ZCurso = ZVector(Ciclo, 1)
        XMes(1) = ZVector(Ciclo, 2)
        XMes(2) = ZVector(Ciclo, 3)
        XMes(3) = ZVector(Ciclo, 4)
        XMes(4) = ZVector(Ciclo, 5)
        XMes(5) = ZVector(Ciclo, 6)
        XMes(6) = ZVector(Ciclo, 7)
        XMes(7) = ZVector(Ciclo, 8)
        XMes(8) = ZVector(Ciclo, 9)
        XMes(9) = ZVector(Ciclo, 10)
        XMes(10) = ZVector(Ciclo, 11)
        XMes(11) = ZVector(Ciclo, 12)
        XMes(12) = ZVector(Ciclo, 13)
        
        ZPersonas = 0
        ZHoras = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cronograma"
        ZSql = ZSql + " Where Cronograma.Curso = " + "'" + ZCurso + "'"
        ZSql = ZSql + " and Cronograma.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " Order by Cronograma.Clave"
        rsCronograma = ZSql
        Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
        If rstCronograma.RecordCount > 0 Then
            With rstCronograma
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZPersonas = ZPersonas + 1
                        ZHoras = ZHoras + !Horas
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCronograma.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CronogramaII SET "
        ZSql = ZSql + " Personas = " + "'" + Str$(ZPersonas) + "',"
        ZSql = ZSql + " Horas = " + "'" + Str$(ZHoras) + "'"
        ZSql = ZSql + " Where Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and Curso = " + "'" + ZCurso + "'"
        spCronogramaII = ZSql
        Set rstCronogramaII = db.OpenRecordset(spCronogramaII, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Listado.WindowTitle = " Plan de Capacitacion Anual"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{CronogramaII.ano} in " + Ano.Text + " to " + Ano.Text
    Listado.SelectionFormula = "{CronogramaII.ano} in " + Ano.Text + " to " + Ano.Text

    Kill "dada.doc"

    Listado.Destination = crptToFile
    Listado.PrintFileName = "dada.doc"
    Listado.PrintFileType = crptWinWord
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CronogramaII.Ano, CronogramaII.Curso, CronogramaII.Mes1, CronogramaII.Mes2, CronogramaII.Mes3, CronogramaII.Mes4, CronogramaII.Mes5, CronogramaII.Mes6, CronogramaII.Mes7, CronogramaII.Mes8, CronogramaII.Mes9, CronogramaII.Mes10, CronogramaII.Mes11, CronogramaII.Mes12, CronogramaII.Personas, CronogramaII.Horas, CronogramaII.PorceII, CronogramaII.Mes, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.CronogramaII CronogramaII, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "CronogramaII.Curso = Curso.Codigo AND " _
                + "CronogramaII.Ano >= " + Ano.Text + " AND " _
                + "CronogramaII.Ano <= " + Ano.Text
    
    Listado.Connect = Connect()
    Listado.ReportFileName = "WListaPlanificacion.rpt"
    Listado.Action = 1
    
    
    If impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Listado.PrintFileName = ""
    Listado.PrintFileType = 0
    
    Listado.Action = 1
    
    
End Sub

Private Sub Cancela_click()
    PrgListaPlanificacion.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    panta.Value = True
    impresora.Value = False
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

