VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaPlanilla 
   AutoRedraw      =   -1  'True
   Caption         =   "Planilla de Cursos no Programados"
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Mes 
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
         Left            =   1320
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1920
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
         Left            =   1320
         TabIndex        =   4
         Top             =   1920
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
         TabIndex        =   3
         Top             =   600
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
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Mes"
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
         Top             =   600
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaPlanilla.rpt"
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
Attribute VB_Name = "PrgListaPlanilla"
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
Dim ZAno As String
Dim ZMes As String

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


    
    ZAno = Ano.Text
    ZMes = Mes.Text
    
    Call Ceros(ZAno, 4)
    Call Ceros(ZMes, 2)
    
    WDesdeFecha = ZAno + "0131"
    WHastaFecha = ZAno + ZMes + "31"
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Cursadas SET "
    ZSql = ZSql + " Ano = " + "'" + Ano.Text + "',"
    ZSql = ZSql + " Horas1 = 0,"
    ZSql = ZSql + " Horas2 = 0,"
    ZSql = ZSql + " Horas3 = 0,"
    ZSql = ZSql + " Horas4 = 0,"
    ZSql = ZSql + " Horas5 = 0,"
    ZSql = ZSql + " Horas6 = 0,"
    ZSql = ZSql + " Horas7 = 0,"
    ZSql = ZSql + " Horas8 = 0,"
    ZSql = ZSql + " Horas9 = 0,"
    ZSql = ZSql + " Horas10 = 0,"
    ZSql = ZSql + " Horas11 = 0,"
    ZSql = ZSql + " Horas12 = 0"
    spCursadas = ZSql
    Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
    
    
    For Ciclo = 1 To Val(ZMes)
        Select Case Ciclo
            Case 1
            
                ZDesde = ZAno + "0101"
                ZHasta = ZAno + "0131"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas1 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 2
            
                ZDesde = ZAno + "0201"
                ZHasta = ZAno + "0231"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas2 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 3
            
                ZDesde = ZAno + "0301"
                ZHasta = ZAno + "0331"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas3 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 4
            
                ZDesde = ZAno + "0401"
                ZHasta = ZAno + "0431"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas4 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 5
            
                ZDesde = ZAno + "0501"
                ZHasta = ZAno + "0531"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas5 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 6
            
                ZDesde = ZAno + "0601"
                ZHasta = ZAno + "0631"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas6 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 7
            
                ZDesde = ZAno + "0701"
                ZHasta = ZAno + "0731"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas7 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 8
            
                ZDesde = ZAno + "0801"
                ZHasta = ZAno + "0831"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas8 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 9
            
                ZDesde = ZAno + "0901"
                ZHasta = ZAno + "0931"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas9 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 10
            
                ZDesde = ZAno + "1001"
                ZHasta = ZAno + "1031"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas10 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 11
            
                ZDesde = ZAno + "1101"
                ZHasta = ZAno + "1131"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas11 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 12
            
                ZDesde = ZAno + "1201"
                ZHasta = ZAno + "1231"
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Cursadas SET "
                ZSql = ZSql + " Horas12 = Horas"
                ZSql = ZSql + " Where TipoCursada = 1"
                ZSql = ZSql + " and OrdFecha >= " + "'" + ZDesde + "'"
                ZSql = ZSql + " and OrdFecha <= " + "'" + ZHasta + "'"
                spCursadas = ZSql
                Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
            
            Case Else
        End Select
    Next Ciclo
    
    
    
    
    
    Listado.WindowTitle = "Planilla de Cursos no Programados"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Cursadas.OrdFecha} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34)
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno

    Kill "dada.doc"
    
    Listado.Destination = crptToFile
    Listado.PrintFileName = "dada.doc"
    Listado.PrintFileType = crptWinWord
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Cursadas.Curso, Cursadas.OrdFecha, Cursadas.TipoCursada, Cursadas.Ano, Cursadas.Horas1, Cursadas.Horas2, Cursadas.Horas3, Cursadas.Horas4, Cursadas.Horas5, Cursadas.Horas6, Cursadas.Horas7, Cursadas.Horas8, Cursadas.Horas9, Cursadas.Horas10, Cursadas.Horas11, Cursadas.Horas12, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cursadas Cursadas, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cursadas.Curso = Curso.Codigo AND " _
                + "Cursadas.OrdFecha >= '" + WDesdeFecha + "' AND " _
                + "Cursadas.OrdFecha <= '" + WHastaFecha + "' AND " _
                + "Cursadas.TipoCursada = 1"
    
    Listado.Connect = Connect()
    Listado.ReportFileName = "WListaPlanilla.rpt"
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
    PrgListaPlanilla.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    panta.Value = True
    impresora.Value = False
End Sub

Private Sub Mes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
    If KeyAscii = 27 Then
        Mes.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mes.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

