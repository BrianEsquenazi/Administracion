VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCronogramaII 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cronograma de Capacitacion del Personal"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
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
      TabIndex        =   0
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
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
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
      ItemData        =   "CronogramaII.frx":0000
      Left            =   120
      List            =   "CronogramaII.frx":0007
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
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
      Height          =   5175
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
      _Version        =   327680
      BackColor       =   16777152
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
      Top             =   120
      Width           =   855
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "CronogramaII.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "CronogramaII.frx":031F
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7320
      MouseIcon       =   "CronogramaII.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "CronogramaII.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "CronogramaII.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "CronogramaII.frx":19B7
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "CronogramaII.frx":21F9
      MousePointer    =   99  'Custom
      Picture         =   "CronogramaII.frx":2503
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
End
Attribute VB_Name = "PrgCronogramaII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCronogramaII As Recordset
Dim spCronogramaII As String
Dim rstLegajo As Recordset
Dim spLegajo As String
Dim rstCurso As Recordset
Dim spCurso As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer

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

Private Sub cmdClose1_Click()
    Call Limpia_Click
    PrgCronograma.Hide
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
    
    WRenglon = 0
    For IRow = 1 To 100
    
        ZCurso = WVector1.TextMatrix(IRow, 1)
        ZDescripcion = WVector1.TextMatrix(IRow, 2)
        ZHoras = WVector1.TextMatrix(IRow, 3)
        ZRealizado = WVector1.TextMatrix(IRow, 4)
        ZObservaciones = WVector1.TextMatrix(IRow, 5)
        
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
        ZSql = ZSql + "Observaciones )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Legajo.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + ZCurso + "',"
        ZSql = ZSql + "'" + ZHoras + "',"
        ZSql = ZSql + "'" + ZRealizado + "',"
        ZSql = ZSql + "'" + ZObservaciones + "')"
        
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
                WVector1.Col = 3
                WVector1.Text = rstCurso!Horas
                WVector1.Col = 4
                Call StartEdit
                rstCurso.Close
            End If
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

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
                
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCronograma!curso)
            
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCronograma!Horas)
                    
                    WVector1.Col = 4
                    WVector1.Text = Trim(rstCronograma!Realizado)
            
                    WVector1.Col = 5
                    WVector1.Text = Trim(rstCronograma!Observaciones)
            
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

Private Sub Legajo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
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
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cronograma"
        ZSql = ZSql + " Where Cronograma.Legajo = " + "'" + Legajo.Text + "'"
        ZSql = ZSql + " And Cronograma.Ano = " + "'" + Ano.Text + "'"
        rsCronograma = ZSql
        Set rstCronograma = db.OpenRecordset(rsCronograma, dbOpenSnapshot, dbSQLPassThrough)
        If rstCronograma.RecordCount > 0 Then
            rstCronograma.Close
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Call Limpia_Vector
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If
        
    End If
    
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If

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
        Case 5
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
                    WVector1.Col = 3
                    WVector1.Text = rstCurso!Horas
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
        
        curso = WVector1.TextMatrix(IRow, 1)
            
        If curso <> "" Then
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
    WVector1.Cols = 6
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
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 90
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Horas"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Realizado"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
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

