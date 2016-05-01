VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgaprobacion 
   Caption         =   "aprobacion de proyectos"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tipo 
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      MouseIcon       =   "aprobacion.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "aprobacion.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Menu Principal"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      MouseIcon       =   "aprobacion.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "aprobacion.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      MouseIcon       =   "aprobacion.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "aprobacion.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   360
      Width           =   855
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   6
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   375
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
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
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
      Left            =   2880
      TabIndex        =   2
      Top             =   2400
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
      Index           =   3
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2880
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6015
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10610
      _Version        =   393216
      BackColor       =   16777152
      FormatString    =   ""
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Top             =   2400
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
End
Attribute VB_Name = "Prgaprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstProyecto As Recordset
Dim spProyecto As String
Dim rstAsigna As Recordset
Dim spAsigna As String
Dim filastotales As Integer

Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private XIndice As Single
Dim vector(100, 10) As String
Private Auxi As String
Private XColor As String
Private XArticulo As String
Private WTipopro As Integer
Private RenglonConjunto As String
Private ZDetalle As String
Dim ZVector(1000) As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Consulta_Click()

    Opcion.Clear
    Opcion.AddItem "Proyecto"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    Call Opcion_Click
     
End Sub


Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Order by Proyecto.Codigo"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Rem If rstProyecto!Ano = 0 Then
                                IngresaItem = Str$(rstProyecto!Codigo) + " " + rstProyecto!Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstProyecto!Codigo
                                WIndice.AddItem IngresaItem
                            Rem End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
            Ayuda.SetFocus
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose_Click()
    PrgAsigna.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Graba_Click()
Dia = Date
    
    For i = 1 To filastotales
    
       WVector1.Row = i
       WVector1.Col = 1
       
    ver = WVector1.Text
        WVector1.Col = 3
    ver2 = WVector1.Text
    
    If ver2 = "aprobado" Then
       Estado = 2
           
    ZSql = ""
        ZSql = ZSql + "UPDATE Proyecto SET "
        ZSql = ZSql + " estado = " + "'" + Str(Estado) + "',"
        ZSql = ZSql + " FechaAprobado = " + "'" + Str(Dia) + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ver + "'"
        spProyecto = ZSql
        Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    
       
    End If
    
        
    Next i
    
    
    
    
    
    
    
    
    
    
    
 Rem   Call Limpia_Click
    
        
End Sub

Private Sub Limpia_Click()
    Call Limpia_Vector
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            WTexto1.Visible = False
            WTexto2.Visible = False
            Indice = Pantalla.ListIndex
            WVector1.TextMatrix(WVector1.Row, 1) = WIndice.List(Indice)
            WTexto1.Text = WIndice.List(Indice)
            Call WTexto1_KeyDown(13, 0)
            Ayuda.Visible = False
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()
    Call Limpia_Vector
Rem    Ano.Text = ""

Rem ZSql = ""
Rem        ZSql = ZSql + "Select *"
Rem        ZSql = ZSql + " FROM Asigna"
Rem ZSql = ZSql + " Where Asigna.Ano = " + "'" + Ano.Text + "'"
Rem        spAsigna = ZSql
Rem        Set rstAsigna = db.OpenRecordset(spAsigna, dbOpenSnapshot, dbSQLPassThrough)
Rem        If rstAsigna.RecordCount > 0 Then
Rem            rstAsigna.Close
            Call proceso_click
Rem        End If
            
      WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
        
  Rem  End If
    
 Rem   If KeyAscii = 27 Then
  Rem      Ano.Text = ""
 Rem   End If


















End Sub

Private Sub proceso_click()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
  
     ZSql = ""
 ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proyecto"
    ZSql = ZSql + " Where Proyecto.estado = 1 "
    ZSql = ZSql + " Order by Proyecto.Codigo"
    spProyecto = ZSql
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        With rstProyecto
            .MoveFirst
               Do
                If .EOF = False Then

                    Renglon = Renglon + 1
                    WVector1.Row = Renglon

                    WVector1.Col = 1
                    WVector1.Text = rstProyecto!Codigo
                    
                    WVector1.Col = 2
                    WVector1.Text = rstProyecto!Descripcion
                    
                    WVector1.Col = 3
                    prestado = rstProyecto!Estado
                    
                    WVector1.Text = "Pendiente"
                  Rem  WVector1.Text = Str$(rstAsigna!Importe)
                  Rem  WVector1.Text = Pusing("###,###.##", WVector1.Text)
                
                    .MoveNext
                    
                        Else
                    filastotales = Renglon
                    Exit Do
                
                End If
            
            Loop
        End With
        rstProyecto.Close
    End If
    tipo.Text = "Pendiente"
                    
    WVector1.Col = 1
    WVector1.Row = 1
    Call StartEdit

End Sub

Private Sub Ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Asigna"
        ZSql = ZSql + " Where Asigna.Ano = " + "'" + Ano.Text + "'"
        spAsigna = ZSql
        Set rstAsigna = db.OpenRecordset(spAsigna, dbOpenSnapshot, dbSQLPassThrough)
        If rstAsigna.RecordCount > 0 Then
            rstAsigna.Close
            Call proceso_click
        End If
            
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
        
    End If
    
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    Pantalla.Clear
    WIndice.Clear
    
    If KeyAscii > 31 Then
        ZAyuda = Ayuda.Text + Chr$(KeyAscii)
            Else
        Select Case KeyAscii
            Case 27
                Ayuda.Text = ""
                ZAyuda = ""
            Case 8
                WEspacios = Len(Ayuda.Text)
                If WEspacios > 0 Then
                    WEspacios = WEspacios - 1
                    ZAyuda = Left$(Ayuda.Text, WEspacios)
                End If
            Case Else
                ZAyuda = Ayuda.Text
        End Select
    End If
    WEspacios = Len(ZAyuda)
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Where Proyecto.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Proyecto.Codigo"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    
    
    Xcolumna = WVector1.Col
    XTipoDato = WParametros(3, Xcolumna)

If Xcolumna <> 1 Then
    
    
    
    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, Xcolumna)
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
          Rem  WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
              
                
                Case 0 To Asc(" ")
                   
                          WTexto2.Text = WVector1.Text
                            If WTexto2.Text = tipo.Text Then
                              WTexto2.Text = "aprobado"
                             Else
                                If WTexto2.Text = "aprobado" Then
                              WTexto2.Text = "Pendiente"
                                     Else
                             If WTexto2.Text = "Pendiente" Then
                                WTexto2.Text = "aprobado"
                             End If
                             End If
                             End If
                             
                     
                     
                     WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
          Rem  WTexto2.SetFocus
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
End If
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
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

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

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

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
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

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
    Xcolumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, Xcolumna) = 0 Then
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
  Rem by nan
     If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc("  ")
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
    Xcolumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case Xcolumna
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + WVector1.Text + "'"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstProyecto!Descripcion
                WVector1.Col = 3
                WVector1.Text = Str$(rstProyecto!Presupuesto)
                WVector1.Col = 2
                rstProyecto.Close
                    Else
                WControl = "N"
            End If
            
        Case Else
            WVector1.Col = Xcolumna
    End Select
End Sub

Private Sub WVector1_DblClick()

Rem    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
Rem    WTexto1.Visible = False
Rem    WTexto2.Visible = False
Rem    WTexto3.Visible = False

Rem    For Ciclo = 1 To WVector1.Cols - 1
Rem        WVector1.Col = Ciclo
Rem        WVector1.Text = ""
Rem    Next Ciclo
    
Rem    Erase WBorra
Rem    EntraVector = 0
    
Rem    For Ciclo = 1 To WVector1.Rows - 1
Rem        WVector1.Row = Ciclo
Rem        WVector1.Col = 1
Rem        WAuxi1 = WVector1.Text
Rem        If WAuxi1 <> "" Then
Rem            EntraVector = EntraVector + 1
Rem            For Ciclo1 = 1 To WVector1.Cols - 1
Rem                WVector1.Col = Ciclo1
Rem                WBorra(EntraVector, Ciclo1) = WVector1.Text
Rem            Next Ciclo1
Rem        End If
Rem    Next Ciclo
    
Rem    Call Limpia_Vector
    
Rem    For Ciclo = 1 To EntraVector
Rem        WVector1.Row = Ciclo
Rem        For da = 1 To WVector1.Cols - 1
Rem            WVector1.Col = da
Rem            WVector1.Text = WBorra(Ciclo, da)
Rem        Next da
Rem    Next Ciclo
    
Rem    End If
    
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Proyecto"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Estado"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
               Rem WFormato(Ciclo) = "###,###.##"
                
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

Private Sub WTexto1_DblClick()

  Rem  If WVector1.Col = 1 Then

 Rem       Opcion.Clear
    
  Rem      Opcion.AddItem "Proyecto"

        Rem Opcion.Visible = False
    
  Rem      Opcion.ListIndex = 0
    
  Rem      Call Opcion_Click
    
  Rem  End If
    
End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Ano_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Graba_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub
























