VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCursadas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cursos Realizados"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
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
      Index           =   3
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3000
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2880
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Codigo 
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
      Width           =   1215
   End
   Begin VB.TextBox Actividad 
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
      Left            =   7440
      MaxLength       =   50
      TabIndex        =   7
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox Temas 
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
      MaxLength       =   100
      TabIndex        =   8
      Top             =   1200
      Width           =   10335
   End
   Begin VB.TextBox Instructor 
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
      MaxLength       =   50
      TabIndex        =   6
      Top             =   840
      Width           =   4695
   End
   Begin VB.ComboBox TipoII 
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
      Left            =   8280
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox TipoI 
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
      Left            =   6120
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
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
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Curso 
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
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   1
      Top             =   120
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
      TabIndex        =   16
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   6480
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
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   10
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
      ItemData        =   "Cursadas.frx":0000
      Left            =   120
      List            =   "Cursadas.frx":0007
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2280
      TabIndex        =   17
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
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7011
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label9 
      Caption         =   "Codigo"
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
      TabIndex        =   27
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Temas"
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
      TabIndex        =   26
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Actividad"
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
      Left            =   6240
      TabIndex        =   25
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Instructor"
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
      TabIndex        =   24
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Left            =   8640
      TabIndex        =   23
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Horas"
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
      Left            =   3000
      TabIndex        =   21
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha"
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
      TabIndex        =   20
      Top             =   480
      Width           =   855
   End
   Begin VB.Label DesCurso 
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
      Left            =   4800
      TabIndex        =   19
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   11160
      MouseIcon       =   "Cursadas.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "Cursadas.frx":031F
      ToolTipText     =   "Salida"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   11160
      MouseIcon       =   "Cursadas.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "Cursadas.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   11160
      MouseIcon       =   "Cursadas.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "Cursadas.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   11160
      MouseIcon       =   "Cursadas.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "Cursadas.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Curso"
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
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "PrgCursadas"
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

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim ZVector(100) As String
Dim ZControl As String
Dim ZControlII As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String


Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Curso"
     Opcion.AddItem "Legajo"
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
    
        Case 1
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
            
        Case 1
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
            
        Case Else
    End Select
            
    End If

End Sub

Private Sub cmdClose1_Click()
    Call Limpia_Click
    PrgCursadas.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Graba_Click()

    Horas.Text = Pusing("###,###.##", Horas.Text)
    
    Rem For Ciclo = 1 To 100
    Rem
    Rem     ZLegajo = WVector1.TextMatrix(Ciclo, 1)
    Rem     ZDesLegajo = WVector1.TextMatrix(Ciclo, 2)
    Rem     ZObservaciones = WVector1.TextMatrix(Ciclo, 3)
    Rem
    Rem     If Val(ZLegajo) <> 0 Then
    Rem         ZSql = ""
    Rem         ZSql = ZSql + "Select *"
    Rem         ZSql = ZSql + " FROM Cronograma"
    Rem         ZSql = ZSql + " Where Legajo = " + "'" + ZLegajo + "'"
    Rem         ZSql = ZSql + " and Curso = " + "'" + Curso.Text + "'"
    Rem         spCronograma = ZSql
    Rem         Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)
    Rem         If rstCronograma.RecordCount > 0 Then
    Rem             rstCronograma.Close
    Rem                 Else
    Rem             m$ = "El Legajo " + ZLegajo + " no tiene programado el curso"
    Rem             A% = MsgBox(m$, 0, "Ingreso de Cursos")
    Rem             Exit Sub
    Rem         End If
    Rem     End If
    Rem
    Rem Next Ciclo

    ZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
    Renglon = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cursadas"
    ZSql = ZSql + " Where Cursadas.Codigo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Cursadas.Clave"
    rsCursadas = ZSql
    Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
    If rstCursadas.RecordCount > 0 Then
        With rstCursadas
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    ZVector(Renglon) = Str$(rstCursadas!Legajo)
                    ZHoras = Str$(rstCursadas!Horas)
                    ZCurso = Str$(rstCursadas!Curso)
                    ZAno = Mid$(rstCursadas!Fecha, 7, 4)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCursadas.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        ZLegajo = ZVector(Ciclo)
                    
        ZSql = ""
        ZSql = ZSql + "UPDATE Cronograma SET "
        ZSql = ZSql + " Realizado = Realizado - " + "'" + ZHoras + "'"
        ZSql = ZSql + " Where Ano = " + "'" + ZAno + "'"
        ZSql = ZSql + " and Legajo = " + "'" + ZLegajo + "'"
        ZSql = ZSql + " and Curso = " + "'" + ZCurso + "'"
        spCronograma = ZSql
        Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "DELETE Cursadas"
    ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
    rsCursadas = ZSql
    Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
    
    WRenglon = 0
    For IRow = 1 To 100
    
        ZLegajo = WVector1.TextMatrix(IRow, 1)
        ZDesLegajo = WVector1.TextMatrix(IRow, 2)
        ZObservaciones = WVector1.TextMatrix(IRow, 3)
        
        If Val(ZLegajo) <> 0 Then
            
            Auxi1 = Codigo.Text
            Call Ceros(Auxi1, 6)
        
            WRenglon = WRenglon + 1
            Auxi = Str$(WRenglon)
            Call Ceros(Auxi, 2)
        
            WClave = Auxi1 + Auxi
            
            ZDesSector = ""
            ZSector = ""
            ZTarea = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Legajo"
            ZSql = ZSql + " Where Legajo.Codigo = " + "'" + ZLegajo + "'"
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLegajo.RecordCount > 0 Then
                ZTarea = Str$(rstLegajo!Perfil)
                rstLegajo.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Tarea"
            ZSql = ZSql + " Where Tarea.Codigo = " + "'" + ZTarea + "'"
            spTarea = ZSql
            Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
            If rstTarea.RecordCount > 0 Then
                ZSector = Str$(rstTarea!Sector)
                rstTarea.Close
            End If
            
            Sql1 = "Select *"
            Sql2 = " FROM Sector"
            Sql3 = " Where Sector.Codigo = " + "'" + ZSector + "'"
            spSector = Sql1 + Sql2 + Sql3
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
            If rstSector.RecordCount > 0 Then
                ZDesSector = Trim(rstSector!Descripcion)
                rstSector.Close
            End If
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Cursadas ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Curso ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Horas ,"
            ZSql = ZSql + "TipoI ,"
            ZSql = ZSql + "TipoII ,"
            ZSql = ZSql + "Instructor ,"
            ZSql = ZSql + "Actividad ,"
            ZSql = ZSql + "Temas ,"
            ZSql = ZSql + "Legajo ,"
            ZSql = ZSql + "DesLegajo ,"
            ZSql = ZSql + "DesSector ,"
            ZSql = ZSql + "Observaciones )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Str$(WRenglon) + "',"
            ZSql = ZSql + "'" + Curso.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + ZOrdFecha + "',"
            ZSql = ZSql + "'" + Horas.Text + "',"
            ZSql = ZSql + "'" + Str$(TipoI.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(TipoII.ListIndex) + "',"
            ZSql = ZSql + "'" + Instructor.Text + "',"
            ZSql = ZSql + "'" + Actividad.Text + "',"
            ZSql = ZSql + "'" + Temas.Text + "',"
            ZSql = ZSql + "'" + ZLegajo + "',"
            ZSql = ZSql + "'" + ZDesLegajo + "',"
            ZSql = ZSql + "'" + ZDesSector + "',"
            ZSql = ZSql + "'" + ZObservaciones + "')"
            
            rsCursadas = ZSql
            Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
            
            ZAno = Mid$(Fecha.Text, 7, 4)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Cronograma SET "
            ZSql = ZSql + " Realizado = Realizado + " + "'" + Horas.Text + "',"
            ZSql = ZSql + " ObservacionesII = " + "'" + ZObservaciones + "'"
            ZSql = ZSql + " Where Ano = " + "'" + ZAno + "'"
            ZSql = ZSql + " and Legajo = " + "'" + ZLegajo + "'"
            ZSql = ZSql + " and Curso = " + "'" + Curso.Text + "'"
            spCronograma = ZSql
            Set rstCronograma = db.OpenRecordset(spCronograma, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next IRow
        
    T$ = "Carga de Cursos Realizados"
    m$ = "Desea Imprimir la planilla"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        Listado.GroupSelectionFormula = "{Cursadas.Codigo} in " + Codigo.Text + " to " + Codigo.Text
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        Listado.SQLQuery = "SELECT Cursadas.Clave, Cursadas.Codigo, Cursadas.Curso, Cursadas.Fecha, Cursadas.Horas, Cursadas.TipoI, Cursadas.TipoII, Cursadas.Instructor, Cursadas.Actividad, Cursadas.Temas, Cursadas.Legajo, Cursadas.DesLegajo, Cursadas.DesSector, " _
                    + "Curso.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Cursadas Cursadas, " _
                    + DSQ + ".dbo.Curso Curso " _
                    + "Where " _
                    + "Cursadas.Curso = Curso.Codigo AND " _
                    + "Cursadas.Codigo >= " + Codigo.Text + " AND " _
                    + "Cursadas.Codigo <= " + Codigo.Text
                        
        Listado.Connect = Connect()
        Listado.ReportFileName = "PlanillaCursada.rpt"
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.CopiesToPrinter = 1
        Listado.Action = 1
        
    End If
        
    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
        
    Codigo.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Codigo.Text = ""
    Curso.Text = ""
    DesCurso.Caption = ""
    Fecha.Text = "  /  /    "
    Horas.Text = ""
    Instructor.Text = ""
    Actividad.Text = ""
    Temas.Text = ""
    
    TipoI.ListIndex = 0
    TipoII.ListIndex = 0
    
    Renglon = 0
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Cursadas"
    spCursadas = Sql1 + Sql2
    Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
    If rstCursadas.RecordCount > 0 Then
        rstCursadas.MoveLast
        ZUltimo = IIf(IsNull(rstCursadas!CodigoMayor), "0", rstCursadas!CodigoMayor)
        Codigo.Text = Mid$(Str$(ZUltimo + 1), 2, 8)
        rstCursadas.Close
            Else
        Codigo.Text = "1"
    End If
    
    Codigo.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Curso.Text = WIndice.List(Indice)
            Call Curso_KeyPress(13)
            
        Case 1
            WTexto1.Visible = False
            WTexto2.Visible = False
            Indice = Pantalla.ListIndex
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Legajo"
            ZSql = ZSql + " Where Legajo.Codigo = " + "'" + WIndice.List(Indice) + "'"
            spLegajo = ZSql
            Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLegajo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstLegajo!Codigo
                WVector1.Col = 2
                WVector1.Text = rstLegajo!Descripcion
                WVector1.Col = 3
                Call StartEdit
                rstLegajo.Close
            End If
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    TipoI.Clear
    
    TipoI.AddItem "Interna"
    TipoI.AddItem "Externa"
    
    TipoI.ListIndex = 0
    
    TipoII.Clear
    
    
    
    TipoII.AddItem "Programada"
    TipoII.AddItem "No Programada"
    
    TipoII.ListIndex = 0

    Codigo.Text = ""
    Curso.Text = ""
    DesCurso.Caption = ""
    Fecha.Text = "  /  /    "
    Horas.Text = ""
    Instructor.Text = ""
    Actividad.Text = ""
    Temas.Text = ""
    
    TipoI.ListIndex = 0
    TipoII.ListIndex = 0
    
    Renglon = 0
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Cursadas"
    spCursadas = Sql1 + Sql2
    Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
    If rstCursadas.RecordCount > 0 Then
        rstCursadas.MoveLast
        ZUltimo = IIf(IsNull(rstCursadas!CodigoMayor), "0", rstCursadas!CodigoMayor)
        Codigo.Text = Mid$(Str$(ZUltimo + 1), 2, 8)
        rstCursadas.Close
            Else
        Codigo.Text = "1"
    End If
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cursadas"
    ZSql = ZSql + " Where Cursadas.Codigo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Cursadas.Clave"
    
    rsCursadas = ZSql
    Set rstCursadas = db.OpenRecordset(rsCursadas, dbOpenSnapshot, dbSQLPassThrough)
    If rstCursadas.RecordCount > 0 Then
        With rstCursadas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCursadas!Legajo)
            
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstCursadas!DesLegajo)
                    
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCursadas!Observaciones)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCursadas.Close
    End If
    
End Sub

Private Sub ProcesoII_Click()

    Call Limpia_Vector
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cronograma"
    ZSql = ZSql + " Where Cronograma.Curso = " + "'" + Curso.Text + "'"
    ZSql = ZSql + " Order by Cronograma.Legajo"
    
    rsCronograma = ZSql
    Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
    If rstCronograma.RecordCount > 0 Then
        With rstCronograma
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstCronograma!Horas > rstCronograma!Realizado Then
                
                        WRenglon = WRenglon + 1
                        WVector1.Row = WRenglon
                        Renglon = WRenglon
                
                        WVector1.Col = 1
                        WVector1.Text = Trim(rstCronograma!Legajo)
            
                        WVector1.Col = 2
                        WVector1.Text = Trim(rstCronograma!DesLegajo)
                    
                        WVector1.Col = 3
                        WVector1.Text = ""
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
            End If
            Loop
        End With
        rstCronograma.Close
    End If
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Curso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZControlII = Curso.Text
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Curso"
        ZSql = ZSql + " Where Curso.Codigo = " + "'" + Curso.Text + "'"
        spCurso = ZSql
        Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
        If rstCurso.RecordCount > 0 Then
            DesCurso.Caption = rstCurso!Descripcion
            rstCurso.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cursadas"
            ZSql = ZSql + " Where Cursadas.Codigo = " + "'" + Codigo.Text + "'"
            spCursadas = ZSql
            Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
            If rstCursadas.RecordCount > 0 Then
                rstCursadas.Close
                    Else
                Call ProcesoII_Click
            End If
            
            Fecha.SetFocus
                Else
            Curso.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Curso.Text = ""
        DesCurso.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Horas.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Horas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Horas.Text = Pusing("###,###.##", Horas.Text)
        TipoI.SetFocus
    End If
    If KeyAscii = 27 Then
        Horas.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub TipoI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoII.SetFocus
    End If
End Sub

Private Sub TipoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Instructor.SetFocus
    End If
End Sub

Private Sub Instructor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Actividad.SetFocus
    End If
    If KeyAscii = 27 Then
        Instructor.Text = ""
    End If
End Sub

Private Sub Actividad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Temas.SetFocus
    End If
    If KeyAscii = 27 Then
        Actividad.Text = ""
    End If
End Sub

Private Sub Temas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Temas.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZControl = Codigo.Text
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cursadas"
        ZSql = ZSql + " Where Cursadas.Codigo = " + "'" + Codigo.Text + "'"
        spCursadas = ZSql
        Set rstCursadas = db.OpenRecordset(spCursadas, dbOpenSnapshot, dbSQLPassThrough)
        If rstCursadas.RecordCount > 0 Then
            Curso.Text = rstCursadas!Curso
            Fecha.Text = rstCursadas!Fecha
            Horas.Text = rstCursadas!Horas
            Horas.Text = Pusing("###,###.##", Horas.Text)
            TipoI.ListIndex = rstCursadas!TipoI
            TipoII.ListIndex = rstCursadas!TipoII
            Instructor.Text = rstCursadas!Instructor
            Actividad.Text = rstCursadas!Actividad
            Temas.Text = rstCursadas!Temas
            rstCursadas.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Curso"
            ZSql = ZSql + " Where Curso.Codigo = " + "'" + Curso.Text + "'"
            spCurso = ZSql
            Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            If rstCurso.RecordCount > 0 Then
                DesCurso.Caption = rstCurso!Descripcion
                rstCurso.Close
            End If
            
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
            
                Else
                
            Curso.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
                ZSql = ZSql + " FROM Legajo"
                ZSql = ZSql + " Where Legajo.Codigo = " + "'" + WVector1.Text + "'"
                spLegajo = ZSql
                Set rstLegajo = db.OpenRecordset(spLegajo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLegajo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstLegajo!Descripcion
                    WVector1.Col = 2
                    rstLegajo.Close
                        Else
                    WControl = "N"
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
        
        ZLegajo = WVector1.TextMatrix(IRow, 1)
            
        If ZLegajo <> "" Then
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
    
    Opcion.AddItem "Cursos"
    Opcion.AddItem "Legajos"

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
    WVector1.Cols = 4
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Legajo"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 4000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 4000
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
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

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

Private Sub Codigo_LostFocus()
    If Val(Codigo.Text) <> Val(ZControl) Then
        If Val(Codigo.Text) <> 0 Then
            Call Codigo_KeyPress(13)
        End If
    End If
End Sub

Private Sub Codigo_GotFocus()
    ZControl = Codigo.Text
End Sub

Private Sub Curso_LostFocus()
    If Val(Curso.Text) <> Val(ZControlII) Then
        If Val(Curso.Text) <> 0 Then
            Call Curso_KeyPress(13)
        End If
    End If
End Sub

Private Sub Curso_GotFocus()
    ZControlII = Curso.Text
End Sub

