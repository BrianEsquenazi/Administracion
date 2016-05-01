VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Rem para el vector V

Dim CargaEnsayoV(1000, 20) As String
Dim WParametrosV(10, 20) As Double
Dim WFormatoV(20) As String
Dim WControlV As String




















Rem
Rem Controles de la WVector5
Rem

Private Sub GridEditTextV(ByVal KeyAscii As Integer)

    XColumna = WVector5.Col
    XTipoDato = WParametrosV(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto15.Left = WVector5.CellLeft + WVector5.Left
            WTexto15.Top = WVector5.CellTop + WVector5.Top
            WTexto15.Width = WVector5.CellWidth
            WTexto15.Height = WVector5.CellHeight
            WTexto15.MaxLength = WParametrosV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto15.Text = WVector5.Text
                    WTexto15.SelStart = Len(WTexto15.Text)
                Case Else
                    WTexto15.Text = Chr$(KeyAscii)
                    WTexto15.SelStart = 1
            End Select
            WTexto15.Visible = True
            WTexto15.SetFocus
        Case 1
            WTexto25.Left = WVector5.CellLeft + WVector5.Left
            WTexto25.Top = WVector5.CellTop + WVector5.Top
            WTexto25.Width = WVector5.CellWidth
            WTexto25.Height = WVector5.CellHeight
            WTexto25.MaxLength = WParametrosV(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto25.Text = WVector5.Text
                    Rem WTexto25.SelStart = Len(WTexto25.Text)
                    WTexto25.SelStart = 0
                Case Else
                    WTexto25.Text = Chr$(KeyAscii)
                    WTexto25.SelStart = 1
            End Select
            WTexto25.Visible = True
            WTexto25.SetFocus
        Case 2
            WTexto35.Left = WVector5.CellLeft + WVector5.Left
            WTexto35.Top = WVector5.CellTop + WVector5.Top
            WTexto35.Width = WVector5.CellWidth
            WTexto35.Height = WVector5.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector5.Text) = 10 Then
                        WTexto35.Text = WVector5.Text
                            Else
                        WTexto35.Text = "  /  /    "
                    End If
                    WTexto35.SelStart = 0
                Case Else
                    WTexto35.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto35.SelStart = 1
            End Select
            WTexto35.Visible = True
            WTexto35.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditV()
    Pasa = 0
    If WCombo15.Visible Then
        Pasa = 0
        WVector5.Text = WCombo15.Text
        WCombo15.Visible = False
            Else
        If WTexto15.Visible Then
            Pasa = 1
            WVector5.Text = WTexto15.Text
            WTexto15.Visible = False
                Else
            If WTexto25.Visible Then
                Pasa = 1
                WVector5.Text = WTexto25.Text
                WTexto25.Visible = False
                    Else
                If WTexto35.Visible Then
                    Pasa = 1
                    WVector5.Text = WTexto35.Text
                    WTexto35.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoV(WVector5.Col) <> "" Then
            WVector5.Text = Pusing(WFormatoV(WVector5.Col), WVector5.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboV()
    ' Position the ComboBox over the cell.
    WCombo15.Left = WVector5.CellLeft + WVector5.Left
    WCombo15.Top = WVector5.CellTop + WVector5.Top
    WCombo15.Width = WVector5.CellWidth
    WCombo15.Visible = True
    WCombo15.SetFocus
End Sub

Private Sub WTexto15_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto15.Text = ""
            
        Rem F1
        Case 113
            WTexto15.Text = WVector5.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector5.SetFocus
            DoEvents
            Call Control_CampoV
            If WControlV = "S" Then
                Call Control_WVectorV
            End If
            Call StartEditV

        Case vbKeyDown
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row < WVector5.Rows - 1 Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row + 1
                End If
            End If
            Call StartEditV

        Case vbKeyUp
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row > WVector5.FixedRows Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row - 1
                End If
            End If
            Call StartEditV
        Case 34
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow < WVector5.Rows - 12 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow + 12
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 33
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow - 12 > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow - 12
                    WVector5.Row = WVector5.TopRow
                        Else
                    WVector5.TopRow = 1
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 123
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Col > 1 Then
                WVector5.Col = WVector5.Col - 1
            End If
            Call StartEditV

    End Select
End Sub

Private Sub WTexto25_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto25.Text = ""
            
        Rem F1
        Case 113
            WTexto25.Text = WVector5.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector5.SetFocus
            DoEvents
            Call Control_CampoV
            If WControlV = "S" Then
                Call Control_WVectorV
            End If
            Call StartEditV
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row < WVector5.Rows - 1 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.Row = WVector5.Row + 1
                Rem End If
            End If
            Call StartEditV

        Case vbKeyUp
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.Row = WVector5.Row - 1
                Rem End If
            End If
            Call StartEditV
        Case 34
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow < WVector5.Rows - 12 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow + 12
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 33
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow - 12 > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow - 12
                    WVector5.Row = WVector5.TopRow
                        Else
                    WVector5.TopRow = 1
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV

    End Select
End Sub

Private Sub WTexto35_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto35.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto35.Text = WVector5.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector5.SetFocus
            Call Control_CampoV
            If WControlV = "S" Then
                Call Control_WVectorV
            End If
            Call StartEditV

        Case vbKeyDown
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row < WVector5.Rows - 1 Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row + 1
                End If
            End If
            Call StartEditV

        Case vbKeyUp
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.Row > WVector5.FixedRows Then
                Call Control_CampoV
                If WControlV = "S" Then
                    WVector5.Row = WVector5.Row - 1
                End If
            End If
            Call StartEditV
        Case 34
            ' Move down 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow < WVector5.Rows - 12 Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow + 12
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV
            
        Case 33
            ' Move up 1 row.
            WVector5.SetFocus
            DoEvents
            If WVector5.TopRow - 12 > WVector5.FixedRows Then
                Rem Call Control_CampoV
                Rem If WControlV = "S" Then
                    WVector5.TopRow = WVector5.TopRow - 12
                    WVector5.Row = WVector5.TopRow
                        Else
                    WVector5.TopRow = 1
                    WVector5.Row = WVector5.TopRow
                Rem End If
            End If
            Call StartEditV

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto15_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto25_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActVeControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto35_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo15_Click()
    WVector5.SetFocus
End Sub


Private Sub WVector5_Click()
    StartEditV
End Sub

Private Sub WVector5_LeaveCell()
    EndEditV
End Sub

Private Sub WVector5_GotFocus()
    EndEditV
End Sub

Private Sub WVector5_KeyPress(KeyAscii As Integer)
    XColumna = WVector5.Col
    Select Case WParametrosV(4, WVector5.Col)
        Case 1
        Case Else
            If WParametrosV(2, XColumna) = 0 Then
                GridEditTextV KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditV()
    Select Case WParametrosV(4, WVector5.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo15.Clear
            WCombo15.AddItem "Campo1"
            WCombo15.AddItem "Campo2"
            On Error Resume Next
            WCombo15.Text = WVector5.Text
            On Error GoTo 0
            GridEditComboV
        Case Else
            If WParametrosV(2, WVector5.Col) = 0 Then
                GridEditTextV Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVectorV()
    Select Case WVector5.Col
        Case 4
            If WVector5.Row < WVector5.Rows - 1 Then
                WVector5.Row = WVector5.Row + 1
            End If
            Rem WVector5.Col = 1
        Case Else
            If WVector5.Col < WVector5.Cols - 1 Then
                WVector5.Col = WVector5.Col + 1
            End If
    End Select
    WVector5.SetFocus
    GridEditTextV KeyAscii
End Sub

Private Sub Control_CampoV()
    XColumna = WVector5.Col
    XFila = WVector5.Row
    WControlV = "S"
End Sub



Private Sub WVector5_DblClick()

    If WVector5.Col = 0 Or WVector5.Col = 1 Then
    
    WTexto15.Visible = False
    WTexto25.Visible = False
    WTexto35.Visible = False
    
    RenglonAuxiliar = WVector5.Row

    For Ciclo = 1 To WVector5.Cols - 1
        WVector5.Col = Ciclo
        WVector5.Text = ""
    Next Ciclo
    
    Erase WBorraV
    EntraVector = 0
    
    HastaRenglon = 0
    For IRow = 100 To 1 Step -1
        
        Etapa = WVector5.TextMatrix(IRow, 1)
        Fecha = WVector5.TextMatrix(IRow, 2)
        Participantes = WVector5.TextMatrix(IRow, 3)
        Resultados = WVector5.TextMatrix(IRow, 4)
        Acciones = WVector5.TextMatrix(IRow, 5)
        Responsables = WVector5.TextMatrix(IRow, 6)
        Estado = WVector5.TextMatrix(IRow, 7)
            
        If Etapa <> "" Or Fecha <> "" Or Participantes <> "" Or Resultados <> "" Or Acciones <> "" Or Responsables <> "" Or Estado <> "" Then
            HastaRenglon = IRow
            Exit For
        End If
            
    Next IRow
    
    For Ciclo = 1 To HastaRenglon
        WVector5.Row = Ciclo
        WVector5.Col = 1
        WAuxi1 = WVector5.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector5.Cols - 1
                WVector5.Col = Ciclo1
                WBorraV(EntraVector, Ciclo1) = WVector5.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorV
    
    For Ciclo = 1 To EntraVector
        WVector5.Row = Ciclo
        For da = 0 To WVector5.Cols - 1
            WVector5.Col = da
            WVector5.Text = WBorraV(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub AgregaRenglonV_Click()

    Hasta = WVector5.Row

    For IRow = 100 To Hasta Step -1
        WVector5.TextMatrix(IRow, 0) = WVector5.TextMatrix(IRow - 1, 0)
        WVector5.TextMatrix(IRow, 1) = WVector5.TextMatrix(IRow - 1, 1)
        WVector5.TextMatrix(IRow, 2) = WVector5.TextMatrix(IRow - 1, 2)
        WVector5.TextMatrix(IRow, 3) = WVector5.TextMatrix(IRow - 1, 3)
        WVector5.TextMatrix(IRow, 4) = WVector5.TextMatrix(IRow - 1, 4)
        WVector5.TextMatrix(IRow, 5) = WVector5.TextMatrix(IRow - 1, 5)
        WVector5.TextMatrix(IRow, 6) = WVector5.TextMatrix(IRow - 1, 6)
        WVector5.TextMatrix(IRow, 7) = WVector5.TextMatrix(IRow - 1, 7)
    Next IRow

    WVector5.TextMatrix(Hasta, 0) = ""
    WVector5.TextMatrix(Hasta, 1) = ""
    WVector5.TextMatrix(Hasta, 2) = ""
    WVector5.TextMatrix(Hasta, 3) = ""
    WVector5.TextMatrix(Hasta, 4) = ""
    WVector5.TextMatrix(Hasta, 5) = ""
    WVector5.TextMatrix(Hasta, 6) = ""
    WVector5.TextMatrix(Hasta, 7) = ""
    
    WTexto15.Text = ""
    WTexto25.Text = ""

End Sub




Private Sub Limpia_VectorV()

    WVector5.Clear

    Rem ponga la WVector5 en negritas
    WVector5.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto15.FontName = WVector5.FontName
    WTexto15.FontSize = WVector5.FontSize
    WTexto15.Visible = False
    WTexto25.FontName = WVector5.FontName
    WTexto25.FontSize = WVector5.FontSize
    WTexto25.Visible = False
    WTexto35.FontName = WVector5.FontName
    WTexto35.FontSize = WVector5.FontSize
    WTexto35.Visible = False
    WCombo15.FontName = WVector5.FontName
    WCombo15.FontSize = WVector5.FontSize
    WCombo15.Visible = False

    ' Establesco loa Valores de la WVector5
    
    WVector5.FixedCols = 1
    WVector5.Cols = 8
    WVector5.FixedRows = 1
    WVector5.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector5.Text = "Articulo"
    
    Rem Longitud
    Rem WVector5.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector5.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosV(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosV(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosV(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosV(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector5.ColWidth(0) = 400
    WVector5.Row = 0
    For Ciclo = 1 To WVector5.Cols - 1
        WVector5.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector5.Text = "Etapa"
                WVector5.ColWidth(Ciclo) = 1200
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 20
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 2
                WVector5.Text = "Fecha"
                WVector5.ColWidth(Ciclo) = 1000
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 10
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 3
                WVector5.Text = "Participantes"
                WVector5.ColWidth(Ciclo) = 2000
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 50
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 4
                WVector5.Text = "Resultados"
                WVector5.ColWidth(Ciclo) = 3000
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 80
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 5
                WVector5.Text = "Acciones"
                WVector5.ColWidth(Ciclo) = 2000
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 40
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 6
                WVector5.Text = "Responsables"
                WVector5.ColWidth(Ciclo) = 2000
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 20
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
            Case 7
                WVector5.Text = "Estado"
                WVector5.ColWidth(Ciclo) = 1500
                WVector5.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosV(1, Ciclo) = 30
                WParametrosV(2, Ciclo) = 0
                WParametrosV(3, Ciclo) = 0
                WParametrosV(4, Ciclo) = 0
                WFormatoV(Ciclo) = ""
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector5.Row = 0
    For Ciclo = 1 To WVector5.Cols - 1
        WVector5.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector5.Text
        Rem WTitulo(Ciclo).Left = WVector5.CellLeft + WVector5.Left
        Rem WTitulo(Ciclo).Top = WVector5.CellTop + WVector5.Top
        Rem WTitulo(Ciclo).Width = WVector5.CellWidth
        Rem WTitulo(Ciclo).Height = WVector5.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector5
    
    WAncho = 400
    For Ciclo = 0 To WVector5.Cols - 1
        WAncho = WAncho + WVector5.ColWidth(Ciclo)
    Next Ciclo
    WVector5.Width = WAncho

    ' Size the columns.
    Font.Name = WVector5.Font.Name
    Font.Size = WVector5.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector5.AllowUserResizing = flexResizeBoth
    
    WVector5.Col = 1
    WVector5.Row = 1
    
End Sub

Private Sub WVector5_Scroll()
    WTexto15.Visible = False
    WTexto25.Visible = False
    WTexto35.Visible = False
End Sub













