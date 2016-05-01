VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaTareas 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Perfiles"
   ClientHeight    =   6645
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6645
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
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7440
      TabIndex        =   12
      Top             =   840
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
      Left            =   1680
      TabIndex        =   11
      Top             =   3360
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
      TabIndex        =   10
      Top             =   5760
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
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   6135
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1080
         Width           =   855
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
         Left            =   2040
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
      Begin VB.Image Consulta 
         Height          =   480
         Left            =   5040
         MouseIcon       =   "listatareas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "listatareas.frx":030A
         ToolTipText     =   "Consulta de Datos"
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta Perfil"
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
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Perfil"
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
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7680
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WListaTareas.rpt"
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
      Top             =   3360
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgListaTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Acepta_Click()

    Listado.WindowTitle = " Descripcion de Perfil de Empleado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Tarea.Codigo} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{Tarea.Codigo} in " + Desde.Text + " to " + Hasta.Text

    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Tarea.Codigo, Tarea.Descripcion, Tarea.Vigencia, Tarea.Sector, Tarea.TareasI, Tarea.TareasII, Tarea.TareasIII, Tarea.DescriI, Tarea.DescriII, Tarea.DescriIII, Tarea.DescriIV, Tarea.DescriV, Tarea.ObservaI, Tarea.ObservaII, Tarea.ObservaIII, Tarea.ObservaIV, Tarea.ObservaV, Tarea.NecesariaI, Tarea.NecesariaII, Tarea.NecesariaIII, Tarea.NecesariaIV, Tarea.NecesariaV, Tarea.DeseableI, Tarea.DeseableII, Tarea.DeseableIII, Tarea.DeseableIV, Tarea.DeseableV, Tarea.Equivalencias, Tarea.Fisica, Tarea.OtrosI, Tarea.OtrosII, Tarea.Curso, Tarea.NecesariaCurso, Tarea.DeseableCurso, " _
                + "Curso.Descripcion, " _
                + "Sector.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Tarea Tarea, " _
                + DSQ + ".dbo.Curso Curso, " _
                + DSQ + ".dbo.Sector Sector " _
                + "Where " _
                + "Tarea.Curso = Curso.Codigo AND " _
                + "Tarea.Sector = Sector.Codigo AND " _
                + "Tarea.Codigo >= " + Desde.Text + " AND " _
                + "Tarea.Codigo <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    If WEmpresa = "0001" Then
    Listado.ReportFileName = "ImpreTarea.rpt"
        Else
     Listado.ReportFileName = "ImpreTareapelli.rpt"
    End If
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgListaTareas.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Panta.Value = True
    Impresora.Value = False
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
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Tareas"
     
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
            Sql2 = " FROM Tarea"
            Sql3 = " Order by Tarea.Codigo"
            spTarea = Sql1 + Sql2 + Sql3
            Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
            If rstTarea.RecordCount > 0 Then
                With rstTarea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstTarea!Renglon = 1 Then
                                LugarAyuda = LugarAyuda + 1
                                Pantalla.Row = LugarAyuda
                                Pantalla.Col = 1
                                Pantalla.Text = rstTarea!Codigo
                                Pantalla.Col = 2
                                Pantalla.Text = rstTarea!Descripcion
                                IngresaItem = rstTarea!Codigo
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTarea.Close
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
            Sql2 = " FROM Tarea"
            Sql3 = " Order by Tarea.Codigo"
            spTarea = Sql1 + Sql2 + Sql3
            Set rstTarea = db.OpenRecordset(spTarea, dbOpenSnapshot, dbSQLPassThrough)
            If rstTarea.RecordCount > 0 Then
                With rstTarea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstTarea!Renglon = 1 Then
                                da = Len(rstTarea!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstTarea!Descripcion, aa, WEspacios) Then
                                        LugarAyuda = LugarAyuda + 1
                                        Pantalla.Row = LugarAyuda
                                        Pantalla.Col = 1
                                        Pantalla.Text = rstTarea!Codigo
                                        Pantalla.Col = 2
                                        Pantalla.Text = rstTarea!Descripcion
                                        IngresaItem = rstTarea!Codigo
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
                rstTarea.Close
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





