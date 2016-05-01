VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPlanificacion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Planificacion Anual de Capacitacion por Legajo"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Begin VB.TextBox Legajo 
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
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
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11160
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   2280
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "Planificacion.frx":0000
      Left            =   120
      List            =   "Planificacion.frx":0007
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8070
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Cartel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESTOS CURSOS NO ESTAN GRABADOS, LOS TRAE EN FORMA AUTOMATICA   DESDE EL MAESTRO DE LEGAJOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   14
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Legajo"
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
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label DesLegajo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   4455
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
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "Planificacion.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "Planificacion.frx":031F
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7320
      MouseIcon       =   "Planificacion.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "Planificacion.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "Planificacion.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "Planificacion.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "Planificacion.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "Planificacion.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
End
Attribute VB_Name = "PrgPlanificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCronograma As Recordset
Dim spCronograma As String
Dim rstLegajo As Recordset
Dim spLegajo As String
Dim rstCurso As Recordset
Dim spCurso As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim Ciclo As Integer
Dim ZVector(100, 3) As String
Dim ZControl As String
Dim ZControlII As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String


Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Legajo"
     Opcion.AddItem "Curso"
     Opcion.Visible = True
     
End Sub



Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Legajo"
            Sql3 = " Order by Codigo"
            spLegajo = Sql1 + Sql2 + Sql3
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLegajo.RecordCount > 0 Then
                With rstLegajo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstLegajo!Renglon = 1 Then
                                IngresaItem = Str$(rstLegajo!Codigo) + " " + rstLegajo!Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = Str$(rstLegajo!Codigo)
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLegajo.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Curso"
            Sql3 = " Order by Codigo"
            spCurso = Sql1 + Sql2 + Sql3
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                With rstCurso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstCurso!Codigo) + " " + rstCurso!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Str$(rstCurso!Codigo)
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
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Legajo"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLegajo.RecordCount > 0 Then
                With rstLegajo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstLegajo!Renglon = 1 Then
                                da = Len(rstLegajo!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstLegajo!Descripcion), aa, WEspacios) Then
                                        IngresaItem = Str$(rstLegajo!Codigo) + " " + rstLegajo!Descripcion
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstLegajo!Codigo
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next aa
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLegajo.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Curso"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spCurso = ZSql
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                With rstCurso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstCurso!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstCurso!Descripcion), aa, WEspacios) Then
                                    IngresaItem = Str$(rstCurso!Codigo) + " " + rstCurso!Descripcion
                                    Pantalla.AddItem IngresaItem
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
            
        Case Else
    End Select
            
    End If

End Sub



Private Sub cmdClose1_Click()
    Call Limpia_Click
    PrgPlanificacion.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Graba_Click()

    ZSql = ""
    ZSql = ZSql + "DELETE Cronograma"
    ZSql = ZSql + " Where Legajo = " + "'" + Legajo.Text + "'"
    ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
    rsCronograma = ZSql
    Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
    
    ZTarea = 0
    ZSector = 0
    ZDesSector = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Legajo"
    ZSql = ZSql + " Where Legajo.Codigo = " + "'" + Legajo.Text + "'"
    spLegajo = ZSql
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLegajo.RecordCount > 0 Then
        ZTarea = rstLegajo!Perfil
        rstLegajo.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Tarea"
    ZSql = ZSql + " Where Tarea.Codigo = " + "'" + Str$(ZTarea) + "'"
    spTarea = ZSql
    Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
    If rstTarea.RecordCount > 0 Then
        ZSector = rstTarea!Sector
        rstTarea.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Sector"
    Sql3 = " Where Sector.Codigo = " + "'" + Str$(ZSector) + "'"
    spSector = Sql1 + Sql2 + Sql3
    Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
    If rstSector.RecordCount > 0 Then
        ZDesSector = Trim(rstSector!Descripcion)
        rstSector.Close
    End If
    
    WRenglon = 0
    For IRow = 1 To 100
    
        ZCurso = WVector1.TextMatrix(IRow, 1)
        ZDescripcion = WVector1.TextMatrix(IRow, 2)
        ZObservaciones = WVector1.TextMatrix(IRow, 3)
        ZHoras = WVector1.TextMatrix(IRow, 4)
        ZRealizado = WVector1.TextMatrix(IRow, 5)
        ZObservacionesII = WVector1.TextMatrix(IRow, 6)
        
        If Val(ZCurso) <> 0 Then
            
        Auxi1 = Legajo.Text
        Call Ceros(Auxi1, 6)
        
        Auxi2 = Ano.Text
        Call Ceros(Auxi2, 4)
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        WClave = Auxi1 + Auxi2 + Auxi
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Cronograma ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Legajo ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Curso ,"
        ZSql = ZSql + "Horas ,"
        ZSql = ZSql + "Realizado ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "ObservacionesII ,"
        ZSql = ZSql + "DesSector ,"
        ZSql = ZSql + "DesLegajo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Legajo.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + ZCurso + "',"
        ZSql = ZSql + "'" + ZHoras + "',"
        ZSql = ZSql + "'" + ZRealizado + "',"
        ZSql = ZSql + "'" + ZObservaciones + "',"
        ZSql = ZSql + "'" + ZObservacionesII + "',"
        ZSql = ZSql + "'" + ZDesSector + "',"
        ZSql = ZSql + "'" + DesLegajo.Caption + "')"
        
        rsCronograma = ZSql
        Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next IRow
        
    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
        
    Legajo.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Legajo.Text = ""
    DesLegajo.Caption = ""
    Ano.Text = ""
    Cartel.Visible = False
    
    Renglon = 0
    
    Legajo.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Legajo.Text = WIndice.List(Indice)
            Call Legajo_KeyPress(13)
            
        Case 1
            WTexto1.Visible = False
            WTexto2.Visible = False
            Indice = Pantalla.ListIndex
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Curso"
            ZSql = ZSql + " Where Curso.Codigo = " + "'" + WIndice.List(Indice) + "'"
            spCurso = ZSql
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstCurso!Codigo
                WVector1.Col = 2
                WVector1.Text = rstCurso!Descripcion
                WVector1.Col = 4
                WVector1.Text = rstCurso!Horas
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                WVector1.Col = 3
                Call StartEdit
                rstCurso.Close
            End If
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Cartel.Visible = False

    Legajo.Text = ""
    DesLegajo.Caption = ""
    Ano.Text = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cronograma"
    ZSql = ZSql + " Where Cronograma.Legajo = " + "'" + Legajo.Text + "'"
    ZSql = ZSql + " and Cronograma.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " Order by Cronograma.Clave"
    
    rsCronograma = ZSql
    Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
    If rstCronograma.RecordCount > 0 Then
        With rstCronograma
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                    
                    aa = rstCronograma!Ano
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCronograma!Curso)
            
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCronograma!Observaciones)
            
                    WVector1.Col = 4
                    WVector1.Text = Trim(rstCronograma!Horas)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    
                    WVector1.Col = 5
                    WVector1.Text = Trim(rstCronograma!Realizado)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
            
                    WVector1.Col = 6
                    WVector1.Text = Trim(rstCronograma!observacionesii)
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCronograma.Close
    End If
    
    For Ciclo = 1 To WRenglon
        ZCurso = WVector1.TextMatrix(Ciclo, 1)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Curso"
        ZSql = ZSql + " Where Curso.Codigo = " + "'" + ZCurso + "'"
        spCurso = ZSql
        Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
        If rstCurso.RecordCount > 0 Then
            WVector1.TextMatrix(Ciclo, 2) = rstCurso!Descripcion
            rstCurso.Close
        End If
    Next Ciclo
    
End Sub

Private Sub ProcesoII_Click()

    Call Limpia_Vector
    Erase ZVector
    WRenglon = 0
    WRenglonII = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Legajo"
    ZSql = ZSql + " Where Legajo.Codigo = " + "'" + Legajo.Text + "'"
    spLegajo = ZSql
    Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLegajo.RecordCount > 0 Then
        ZPerfil = Str$(rstLegajo!Perfil)
        rstLegajo.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *, Curso.Descripcion as [WDesCurso]"
    ZSql = ZSql + " FROM Tarea, Curso"
    ZSql = ZSql + " Where Tarea.Codigo = " + "'" + ZPerfil + "'"
    ZSql = ZSql + " and Tarea.Curso = Curso.Codigo"
    ZSql = ZSql + " Order by Tarea.Curso"
    
    spTarea = ZSql
    Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
    If rstTarea.RecordCount > 0 Then
        With rstTarea
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglonII = WRenglonII + 1
                    
                    ZVector(WRenglonII, 1) = Str$(rstTarea!Curso)
                    ZVector(WRenglonII, 2) = rstTarea!WDesCurso
                    ZVector(WRenglonII, 3) = Str$(rstTarea!Horas)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTarea.Close
    End If
    
    For Ciclo = 1 To WRenglonII
    
        ZCurso = ZVector(Ciclo, 1)
        
        Sql1 = "Select *"
        Sql2 = " FROM Legajo"
        Sql3 = " Where Legajo.Codigo = " + "'" + Legajo.Text + "'"
        Sql4 = " and Legajo.Curso = " + "'" + ZCurso + "'"
        spLegajo = Sql1 + Sql2 + Sql3 + Sql4
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
        
            ZEstaCurso = Str$(rstLegajo!EstaCurso)
            
            Select Case Val(ZEstaCurso)
                Case 0, 3, 4, 5
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = ZVector(Ciclo, 1)
                    
                    WVector1.Col = 2
                    WVector1.Text = ZVector(Ciclo, 2)
            
                    WVector1.Col = 3
                    WVector1.Text = ""
            
                    WVector1.Col = 4
                    WVector1.Text = ZVector(Ciclo, 3)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    
                    WVector1.Col = 5
                    WVector1.Text = ""
            
                    WVector1.Col = 6
                    WVector1.Text = ""
                Case Else
            End Select
            rstLegajo.Close
        End If
        
    Next Ciclo
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1

End Sub

Private Sub Legajo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZControl = Legajo.Text
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Legajo"
        ZSql = ZSql + " Where Legajo.Codigo = " + "'" + Legajo.Text + "'"
        spLegajo = ZSql
        Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLegajo.RecordCount > 0 Then
            DesLegajo.Caption = Trim(rstLegajo!Descripcion)
            rstLegajo.Close
            Ano.SetFocus
                Else
            Legajo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Legajo.Text = ""
        DesLegajo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        ZControlII = Ano.Text
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cronograma"
        ZSql = ZSql + " Where Cronograma.Legajo = " + "'" + Legajo.Text + "'"
        ZSql = ZSql + " And Cronograma.Ano = " + "'" + Ano.Text + "'"
        rsCronograma = ZSql
        Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
        If rstCronograma.RecordCount > 0 Then
            rstCronograma.Close
            Cartel.Visible = False
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Cartel.Visible = True
            Call Limpia_Vector
            Call ProcesoII_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If
        
    End If
    
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If

End Sub

Private Sub Legajo_LostFocus()
    If Val(Legajo.Text) <> Val(ZControl) Then
        Call Legajo_KeyPress(13)
        If Val(Ano.Text) <> 0 Then
            Call Ano_KeyPress(13)
        End If
    End If
End Sub

Private Sub Legajo_GotFocus()
    ZControl = Legajo.Text
End Sub

Private Sub Ano_LostFocus()
    If Val(Ano.Text) <> Val(ZControlII) Then
        If Val(Ano.Text) <> 0 Then
            Call Ano_KeyPress(13)
        End If
    End If
End Sub

Private Sub Ano_GotFocus()
    ZControlII = Ano.Text
End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub

Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 3
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Val(WVector1.Text) <> 0 Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Curso"
                ZSql = ZSql + " Where Curso.Codigo = " + "'" + WVector1.Text + "'"
                spCurso = ZSql
                Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
                If rstCurso.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstCurso!Descripcion
                    WVector1.Col = 4
                    WVector1.Text = rstCurso!Horas
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    WVector1.Col = 2
                    rstCurso.Close
                End If
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For IRow = 100 To 1 Step -1
        
        Curso = WVector1.TextMatrix(IRow, 1)
            
        If Curso <> "" Then
            HastaRenglon = IRow
            Exit For
        End If
            
    Next IRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 0 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
    Opcion.AddItem "Legajos"
    Opcion.AddItem "Cursos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Curso"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 90
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 2700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Horas"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###.##"
            Case 5
                WVector1.Text = "Realizado"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###.##"
            Case 6
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 2700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

