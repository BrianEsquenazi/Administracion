VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCronograma 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cronograma de Capacitacion del Personal"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   480
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
      ItemData        =   "Cronograma.frx":0000
      Left            =   120
      List            =   "Cronograma.frx":0007
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
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
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
      MouseIcon       =   "Cronograma.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "Cronograma.frx":031F
      ToolTipText     =   "Salida"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   8160
      MouseIcon       =   "Cronograma.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "Cronograma.frx":0E6B
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   9000
      MouseIcon       =   "Cronograma.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "Cronograma.frx":19B7
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
      Width           =   480
   End
End
Attribute VB_Name = "PrgCronograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCronogramaII As Recordset
Dim spCronogramaII As String
Dim rstCronograma As Recordset
Dim spCronograma As String
Dim rstLegajo As Recordset
Dim spLegajo As String
Dim rstCurso As Recordset
Dim spCurso As String
Dim ZVector(1000, 20) As String
Dim ZControl As String

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

Private Sub cmdClose1_Click()
    Call Limpia_Click
    PrgCronograma.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Graba_Click()

    ZSql = ""
    ZSql = ZSql + "DELETE CronogramaII"
    ZSql = ZSql + " Where Ano = " + "'" + Ano.Text + "'"
    rsCronogramaII = ZSql
    Set rstCronogramaII = db.OpenRecordset(rsCronogramaII, dbOpenSnapshot, dbSQLPassThrough)
    
    WRenglon = 0
    For IRow = 1 To 100
    
        ZCurso = WVector1.TextMatrix(IRow, 1)
        ZDescripcion = WVector1.TextMatrix(IRow, 2)
        ZPersonas = WVector1.TextMatrix(IRow, 3)
        ZHoras = WVector1.TextMatrix(IRow, 4)
        ZTotal = WVector1.TextMatrix(IRow, 5)
        ZCursadas = WVector1.TextMatrix(IRow, 6)
        ZResta = WVector1.TextMatrix(IRow, 7)
        ZMes1 = WVector1.TextMatrix(IRow, 8)
        ZMes2 = WVector1.TextMatrix(IRow, 9)
        ZMes3 = WVector1.TextMatrix(IRow, 10)
        ZMes4 = WVector1.TextMatrix(IRow, 11)
        ZMes5 = WVector1.TextMatrix(IRow, 12)
        ZMes6 = WVector1.TextMatrix(IRow, 13)
        ZMes7 = WVector1.TextMatrix(IRow, 14)
        ZMes8 = WVector1.TextMatrix(IRow, 15)
        ZMes9 = WVector1.TextMatrix(IRow, 16)
        ZMes10 = WVector1.TextMatrix(IRow, 17)
        ZMes11 = WVector1.TextMatrix(IRow, 18)
        ZMes12 = WVector1.TextMatrix(IRow, 19)
        
        If Val(ZCurso) <> 0 Then
            
            Auxi1 = Ano.Text
            Call Ceros(Auxi1, 4)
            
            Auxi2 = ZCurso
            Call Ceros(Auxi2, 4)
            
            WClave = Auxi1 + Auxi2
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CronogramaII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Curso ,"
            ZSql = ZSql + "Mes1 ,"
            ZSql = ZSql + "Mes2 ,"
            ZSql = ZSql + "Mes3 ,"
            ZSql = ZSql + "Mes4 ,"
            ZSql = ZSql + "Mes5 ,"
            ZSql = ZSql + "Mes6 ,"
            ZSql = ZSql + "Mes7 ,"
            ZSql = ZSql + "Mes8 ,"
            ZSql = ZSql + "Mes9 ,"
            ZSql = ZSql + "Mes10 ,"
            ZSql = ZSql + "Mes11 ,"
            ZSql = ZSql + "Mes12 )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + ZCurso + "',"
            ZSql = ZSql + "'" + ZMes1 + "',"
            ZSql = ZSql + "'" + ZMes2 + "',"
            ZSql = ZSql + "'" + ZMes3 + "',"
            ZSql = ZSql + "'" + ZMes4 + "',"
            ZSql = ZSql + "'" + ZMes5 + "',"
            ZSql = ZSql + "'" + ZMes6 + "',"
            ZSql = ZSql + "'" + ZMes7 + "',"
            ZSql = ZSql + "'" + ZMes8 + "',"
            ZSql = ZSql + "'" + ZMes9 + "',"
            ZSql = ZSql + "'" + ZMes10 + "',"
            ZSql = ZSql + "'" + ZMes1 + "',"
            ZSql = ZSql + "'" + ZMes12 + "')"
        
            rsCronogramaII = ZSql
            Set rstCronogramaII = db.OpenRecordset(rsCronogramaII, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next IRow
        
    Call Limpia_Click

    WVector1.Col = 1
    WVector1.Row = 1
        
    Ano.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector

    Ano.Text = ""
    Renglon = 0
    
    Ano.SetFocus

End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Ano.Text = ""
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    WRenglon = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Curso"
    ZSql = ZSql + " Order by Curso.Codigo"
    
    rsCurso = ZSql
    Set rstCurso = db.OpenRecordset(rsCurso, dbOpenSnapshot, dbSQLPassThrough)
    If rstCurso.RecordCount > 0 Then
        With rstCurso
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    ZVector(WRenglon, 1) = Str$(rstCurso!Codigo)
                    ZVector(WRenglon, 2) = Trim(rstCurso!Descripcion)
                    ZVector(WRenglon, 4) = Str$(rstCurso!Horas)
            
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCurso.Close
    End If
    
    
    For Ciclo = 1 To WRenglon
    
        ZCurso = ZVector(Ciclo, 1)
        ZCantidad = 0
        ZRealizado = 0
        
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
                        ZCantidad = ZCantidad + 1
                        ZRealizado = ZRealizado + !Realizado
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCronograma.Close
        End If
        
        ZVector(Ciclo, 3) = Str$(ZCantidad)
        ZVector(Ciclo, 5) = Str$(Val(ZVector(Ciclo, 3)) * Val(ZVector(Ciclo, 4)))
        If ZRealizado > Val(ZVector(Ciclo, 5)) Then
            ZRealizado = Val(ZVector(Ciclo, 5))
        End If
        ZVector(Ciclo, 6) = Str$(ZRealizado)
        ZVector(Ciclo, 7) = Str$(Val(ZVector(Ciclo, 5)) - Val(ZVector(Ciclo, 6)))
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CronogramaII"
        ZSql = ZSql + " Where CronogramaII.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CronogramaII.Curso = " + "'" + ZCurso + "'"
        spCronogramaII = ZSql
        Set rstCronogramaII = db.OpenRecordset(spCronogramaII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCronogramaII.RecordCount > 0 Then
            ZVector(Ciclo, 8) = Trim(rstCronogramaII!Mes1)
            ZVector(Ciclo, 9) = Trim(rstCronogramaII!Mes2)
            ZVector(Ciclo, 10) = Trim(rstCronogramaII!Mes3)
            ZVector(Ciclo, 11) = Trim(rstCronogramaII!Mes4)
            ZVector(Ciclo, 12) = Trim(rstCronogramaII!Mes5)
            ZVector(Ciclo, 13) = Trim(rstCronogramaII!Mes6)
            ZVector(Ciclo, 14) = Trim(rstCronogramaII!Mes7)
            ZVector(Ciclo, 15) = Trim(rstCronogramaII!Mes8)
            ZVector(Ciclo, 16) = Trim(rstCronogramaII!Mes9)
            ZVector(Ciclo, 17) = Trim(rstCronogramaII!Mes10)
            ZVector(Ciclo, 18) = Trim(rstCronogramaII!Mes11)
            ZVector(Ciclo, 19) = Trim(rstCronogramaII!Mes12)
            rstCronogramaII.Close
        End If
        
    Next Ciclo
    
    ZLugar = 0
    For Ciclo = 1 To WRenglon
    
        If Val(ZVector(Ciclo, 3)) <> 0 Then
            
            Lugar = Lugar + 1
                
            WVector1.TextMatrix(Lugar, 1) = ZVector(Ciclo, 1)
            WVector1.TextMatrix(Lugar, 2) = ZVector(Ciclo, 2)
            WVector1.TextMatrix(Lugar, 3) = ZVector(Ciclo, 3)
            WVector1.TextMatrix(Lugar, 4) = ZVector(Ciclo, 4)
            WVector1.TextMatrix(Lugar, 5) = ZVector(Ciclo, 5)
            WVector1.TextMatrix(Lugar, 6) = ZVector(Ciclo, 6)
            WVector1.TextMatrix(Lugar, 7) = ZVector(Ciclo, 7)
            WVector1.TextMatrix(Lugar, 8) = ZVector(Ciclo, 8)
            WVector1.TextMatrix(Lugar, 9) = ZVector(Ciclo, 9)
            WVector1.TextMatrix(Lugar, 10) = ZVector(Ciclo, 10)
            WVector1.TextMatrix(Lugar, 11) = ZVector(Ciclo, 11)
            WVector1.TextMatrix(Lugar, 12) = ZVector(Ciclo, 12)
            WVector1.TextMatrix(Lugar, 13) = ZVector(Ciclo, 13)
            WVector1.TextMatrix(Lugar, 14) = ZVector(Ciclo, 14)
            WVector1.TextMatrix(Lugar, 15) = ZVector(Ciclo, 15)
            WVector1.TextMatrix(Lugar, 16) = ZVector(Ciclo, 16)
            WVector1.TextMatrix(Lugar, 17) = ZVector(Ciclo, 17)
            WVector1.TextMatrix(Lugar, 18) = ZVector(Ciclo, 18)
            WVector1.TextMatrix(Lugar, 19) = ZVector(Ciclo, 19)
            
        End If
            
    Next Ciclo
    
End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        ZControl = Ano.Text
        Call Proceso_Click
        WVector1.Col = 8
        WVector1.Row = 1
        Call StartEdit
    End If
    
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If

End Sub

Private Sub Ano_LostFocus()
    If Val(Ano.Text) <> Val(ZControl) Then
        If Val(Ano.Text) <> 0 Then
            Call Ano_KeyPress(13)
        End If
    End If
End Sub

Private Sub Ano_GotFocus()
    ZControl = Ano.Text
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
        Case 19
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 8
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
        Case 99
            Rem If Val(WVector1.Text) <> 0 Then
            Rem     ZSql = ""
            Rem     ZSql = ZSql + "Select *"
            Rem     ZSql = ZSql + " FROM Curso"
            Rem     ZSql = ZSql + " Where Curso.Codigo = " + "'" + WVector1.Text + "'"
            Rem     spCurso = ZSql
            Rem     Set rstCurso = db.OpenRecordset(spCurso, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstCurso.RecordCount > 0 Then
            Rem         WVector1.Col = 2
            Rem         WVector1.Text = rstCurso!Descripcion
            Rem         WVector1.Col = 3
            Rem         WVector1.Text = rstCurso!Horas
            Rem         rstCurso.Close
            Rem     End If
            Rem End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
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
    WVector1.Cols = 20
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
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
    
    WVector1.ColWidth(0) = 100
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Curso"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 90
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Pers."
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Horas"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Total"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Curs."
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Resta"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Ene"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Feb"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Mar"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = "Abr"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 12
                WVector1.Text = "May"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 13
                WVector1.Text = "Jun"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 14
                WVector1.Text = "Jul"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 15
                WVector1.Text = "Ago"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 16
                WVector1.Text = "Sep"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 17
                WVector1.Text = "Oct"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 18
                WVector1.Text = "Nov"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 19
                WVector1.Text = "Dic"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
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

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Or WVector1.Col = 2 Or WVector1.Col = 3 Then
    
        ZCurso = WVector1.TextMatrix(WVector1.Row, 1)
        
        Listado.WindowTitle = "Listado de Cursos por Curso"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
    
        Uno = "{Cronograma.Curso} in " + ZCurso + " to " + ZCurso + " and "
        Dos = "{Cronograma.Ano} in " + Ano.Text + " to " + Ano.Text
    
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
   
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT Cronograma.Legajo, Cronograma.Ano, Cronograma.Curso, Cronograma.Horas, Cronograma.Realizado, Cronograma.DesLegajo, Cronograma.ObservacionesII, " _
                + "Curso.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Cronograma Cronograma, " _
                + DSQ + ".dbo.Curso Curso " _
                + "Where " _
                + "Cronograma.Curso = Curso.Codigo AND " _
                + "Cronograma.Ano >= " + Ano.Text + " AND " _
                + "Cronograma.Ano <= " + Ano.Text + " AND " _
                + "Cronograma.Curso >= " + ZCurso + " AND " _
                + "Cronograma.Curso <= " + ZCurso
    
        Listado.Destination = 0
        Listado.Connect = Connect()
        Listado.ReportFileName = "WListaCursoPlani.rpt"
        Listado.Action = 1
        
    End If
    
End Sub

