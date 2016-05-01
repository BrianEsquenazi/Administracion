VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgpagovar 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Pagos a Varios"
   ClientHeight    =   6780
   ClientLeft      =   1335
   ClientTop       =   975
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   ScaleHeight     =   6780
   ScaleWidth      =   9480
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion"
      Height          =   300
      Left            =   8400
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin Crystal.CrystalReport listado 
      Left            =   3840
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ordpag2.rpt"
      WindowTitle     =   "Orden de Pago"
      CopiesToPrinter =   2
      WindowState     =   2
   End
   Begin VB.TextBox Importe 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   3
      Text            =   " "
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1680
      MaxLength       =   45
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   3735
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   6840
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSDBGrid.DBGrid DbGrid1 
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "pagovar.frx":0000
      TabIndex        =   4
      Top             =   2400
      Width           =   9255
   End
   Begin VB.TextBox Orden 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "pagovar.frx":09C2
      Left            =   5520
      List            =   "pagovar.frx":09C9
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7320
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   7320
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   8400
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Grabar"
      Height          =   300
      Left            =   6240
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Importe"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label sfd 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.   2) Bco.  3) Ch. Terc."
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro. Orden de Pago"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "Prgpagovar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.

Private Sub Suma_Datos()
    Creditos.Caption = ""
    
    For iRow = 0 To 9
        DBGrid1.Col = 5
        DBGrid1.Row = iRow
        If Val(DBGrid1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(DBGrid1.Text))
        End If
    Next iRow
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
    DBGrid1.Col = 0
    DBGrid1.Row = 0
End Sub

Private Sub Lee_Datos()
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        With rstPagos
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", Orden.Text + Auxi1
             If .NoMatch = False Then
                Credito = Credito + 1
                DBGrid1.Row = Credito - 1
                DBGrid1.Col = 0
                DBGrid1.Text = !Tipo2
                DBGrid1.Col = 1
                DBGrid1.Text = !Numero2
                DBGrid1.Col = 2
                DBGrid1.Text = !Fecha2
                DBGrid1.Col = 3
                DBGrid1.Text = !banco2
                DBGrid1.Col = 4
                If !Observaciones <> "" Then
                    DBGrid1.Text = !Observaciones
                End If
                DBGrid1.Col = 5
                DBGrid1.Text = !Importe2
                DBGrid1.Text = Pusing("###,###.##", DBGrid1.Text)
                    Else
                Exit Do
            End If
        End With
    Loop
End Sub
Sub Verifica_datos()
    If Importe.Text = 0 Then
        Importe.Text = "0"
    End If
End Sub
Sub Format_datos()
    Importe.Text = Pusing("###,###.##", Importe.Text)
End Sub

Sub Imprime_Datos()

End Sub

Private Sub cmdAdd_Click()

    If Orden.Text <> "" And Fecha.Text <> "" Then
    
    Auxi1 = Orden.Text
    Call Ceros(Auxi1, 6)
    Orden.Text = Auxi1
        
    With rstPagos
        Existe = "N"
        .Index = "Clave"
        Claveven$ = Orden.Text + "01"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Existe = "S"
        End If
    End With
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        Debito = 0
        Credito = 0
        
        If Val(Importe.Text) <> 0 Then
            Debito = Val(Importe.Text)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = Val(Creditos.Caption)
        End If
        
        If Debito = Credito Then
    
        With rstPagos
            Renglon = 0
            .Index = "Clave"
            For iRow = 0 To 9
                WRow = iRow
                DBGrid1.Col = 5
                DBGrid1.Row = iRow
                If Val(DBGrid1.Text) <> 0 Then
                    .AddNew
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Orden = Orden.Text
                    !Renglon = Auxi1
                    !Proveedor = 0
                    !Concepto = 0
                    !Importe = Val(Importe.Text)
                    !Observaciones = Observaciones.Text
                    !Fecha = Fecha.Text
                    !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !TipoOrd = "3"
                    !Tiporeg = "2"
                    !Tipo1 = ""
                    !Letra1 = ""
                    !Punto1 = ""
                    !Numero1 = ""
                    !Importe1 = 0
                    DBGrid1.Col = 0
                    !Tipo2 = Left$(DBGrid1.Text, 2)
                    DBGrid1.Col = 1
                    !Numero2 = Left$(DBGrid1.Text, 8)
                    DBGrid1.Col = 2
                    !Fecha2 = Left$(DBGrid1.Text, 10)
                    !FechaOrd2 = Right$(!Fecha2, 4) + Mid$(!Fecha2, 4, 2) + Left$(!Fecha2, 2)
                    DBGrid1.Col = 3
                    !banco2 = Val(DBGrid1.Text)
                    DBGrid1.Col = 4
                    !Observaciones2 = Left$(DBGrid1.Text, 20)
                    DBGrid1.Col = 5
                    !Importe2 = Val(DBGrid1.Text)
                    !Empresa = 1
                    !Clave = !Orden + !Renglon
                    .Update
                    .Bookmark = .LastModified
                    
                    With rstRecibos
                        .Index = "Clave"
                        DBGrid1.Col = 6
                        Claveven$ = DBGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            .Edit
                            !Estado2 = "X"
                            !Destino = Observaciones.Text
                            .Update
                        End If
                    End With
                    
                End If
                
            Next iRow
        End With
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = "1"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                WCtaProveedor = !CtaProveedores
                WCtaEfectivo = !CtaEfectivo
                WCtaCheques = !CtaCheque
            End If
        End With
        
        With rstImputac
            Renglon = 0
            .Index = "Clave"
            
            .AddNew
            !Tipomovi = "3"
            !Proveedor = "000000"
            !TipoComp = "01"
            !LetraComp = "A"
            !PuntoComp = "0000"
            !NroComp = Orden.Text
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            !Renglon = Auxi1$
            !Fecha = Fecha.Text
            !Observaciones = Left$(Observaciones.Text, 30)
            !Debito = Val(Importe.Text)
            !Credito = 0
            !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            !Titulo = "Pagos"
            !Empresa = 1
            !Cuenta = WCuenta
            !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
            .Update
            
            For iRow = 0 To 9
                WRow = iRow
                DBGrid1.Col = 5
                DBGrid1.Row = iRow
                If Val(DBGrid1.Text) <> 0 Then
                    .AddNew
                    !Tipomovi = "3"
                    !Proveedor = "000000"
                    !TipoComp = "01"
                    !LetraComp = "A"
                    !PuntoComp = "0000"
                    !NroComp = Orden.Text
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Renglon = Auxi1$
                    !Fecha = Fecha.Text
                    !Observaciones = Left$(Observaciones.Text, 30)
                    !Credito = Val(DBGrid1.Text)
                    !Debito = 0
                    !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Titulo = "Pagos"
                    !Empresa = 1
                    DBGrid1.Col = 0
                    Select Case Val(DBGrid1.Text)
                        Case 2
                            !Cuenta = "999999"
                            With rstBanco
                                DBGrid1.Col = 3
                                .Index = "Banco"
                                .Seek "=", Val(DBGrid1.Text)
                                If .NoMatch = False Then
                                    WCuenta = !Cuenta
                                End If
                            End With
                            !Cuenta = WCtaCheques
                        Case 3
                            !Cuenta = WCtaCheques
                        Case Else
                            !Cuenta = WCtaEfectivo
                    End Select
                    !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                    .Update
                End If
                
            Next iRow
        End With
        
        Listado.GroupSelectionFormula = "{Pagos.Orden} in " + Chr$(34) + Orden.Text + Chr$(34) + " to " + Chr$(34) + Orden.Text + Chr$(34)
        Listado.Destination = 1
        Listado.Action = 1

        Call CmdLimpiar_Click
        Orden.SetFocus
        
        End If
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Orden.Text <> "" Then
            Rem Borro los datos anteriores
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Orden.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow
    End If
    Orden.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    For iCol = 0 To 6
        For iRow = 0 To 9
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iRow
    Next iCol
    Orden.Text = ""
    Observaciones.Text = ""
    Importe.Text = ""
    Fecha.Text = "  /  /    "
    Creditos.Caption = ""
    Orden.SetFocus
    
    With rstPagos
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Orden.Text = !Orden + 1
                Else
            Orden.Text = ""
        End If
    End With

End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstRecibos
        .Close
    End With
    With rstPagos
        .Close
    End With
    With rstImputac
        .Close
    End With
    With rstBanco
        .Close
    End With
    DbsAdminis.Close
    Orden.SetFocus
    Prgpagovar.Hide
    Menu.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Impresion_Click()
    Listado.GroupSelectionFormula = "{Pagos.Orden} in " + Chr$(34) + Orden.Text + Chr$(34) + " to " + Chr$(34) + Orden.Text + Chr$(34)
    Listado.Destination = 1
    Listado.Action = 1
End Sub

Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        With rstPagos
            Existe = "N"
            .Index = "Clave"
            Claveven$ = Orden.Text + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Existe = "S"
                If !Importe <> "" Then
                    Importe.Text = !Importe
                    Importe.Text = Pusing("###,###.##", Importe.Text)
                End If
                If !Observaciones <> "" Then
                    Observaciones.Text = !Observaciones
                End If
                Fecha.Text = !Fecha
            End If
        End With
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            DBGrid1.Col = 0
            DBGrid1.Row = 0
            DBGrid1.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Importe.SetFocus
    End If
End Sub

Private Sub Importe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Importe.Text = Pusing("###,###.##", Importe.Text)
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        DBGrid1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     XRow = DBGrid1.Row
     XCol = DBGrid1.Col

     Opcion.Clear

     Opcion.AddItem "Cheques terceros"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstRecibos
                .Index = "Clave"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(!Tiporeg) = 2 Then
                            If Val(!Tipo2) = 2 And !Estado2 <> "X" Then
                                Auxi$ = Str$(!Importe2)
                                Auxi$ = Pusing("###,###.##", Auxi$)
                                IngresaItem = !Numero2 + " " + !banco2 + " " + !Fecha2 + " " + Auxi$
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
     
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            With rstRecibos
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Clave"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DBGrid1.Col = 1
                    DBGrid1.Text = !Numero2
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = ""
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = !Fecha2
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = !banco2
                
                    DBGrid1.Col = 5
                    DBGrid1.Text = !Importe2
                    DBGrid1.Text = Pusing("###,###.##", DBGrid1.Text)
                    
                    DBGrid1.Col = 6
                    DBGrid1.Text = Claveven$
                    
                    Call Suma_Datos
                    
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 0
                    
                End If
                If DBGrid1.Row < 10 Then
                    DBGrid1.Row = DBGrid1.Row + 1
                    DBGrid1.Col = 0
                    KeyCode = 0
                            Else
                    DBGrid1.Col = 0
                    KeyCode = 0
                End If
            End With
                
        Case Else
    End Select
    
End Sub
Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0
                If KeyCode = 13 Then
                    If Val(DBGrid1.Text) = 1 Or Val(DBGrid1.Text) = 2 Or Val(DBGrid1.Text) = 3 Then
                        Auxi$ = Str$(Val(DBGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DBGrid1.Text = Auxi$
                        
                        Select Case Val(DBGrid1.Text)
                            Case 1
                                DBGrid1.Col = 1
                                DBGrid1.Text = ""
                                DBGrid1.Col = 2
                                DBGrid1.Text = ""
                                DBGrid1.Col = 3
                                DBGrid1.Text = ""
                                DBGrid1.Col = 4
                                DBGrid1.Text = ""
                                DBGrid1.Col = 5
                                KeyCode = 0
                                
                            Case 3
                                Call Consulta_Click
                                
                            Case Else
                                DBGrid1.Col = 1
                                KeyCode = 0
                                
                        End Select
                        
                            Else
                            
                        DBGrid1.Col = 0
                        KeyCode = 0
                        
                    End If
                End If
                
            Case 1
                If KeyCode = 13 Then
                    DBGrid1.Col = 0
                    If Val(DBGrid1.Text) = 3 Then
                        DBGrid1.Col = 1
                        KeyCode = 0
                            Else
                        DBGrid1.Col = 1
                        Auxi$ = Str$(Val(DBGrid1.Text))
                        Call Ceros(Auxi$, 8)
                        DBGrid1.Text = Auxi$
                        DBGrid1.Col = 2
                        KeyCode = 0
                    End If
                End If
                
            Case 2
                If KeyCode = 13 Then
                    DBGrid1.Col = 2
                    
                    Call Valida_fecha1(DBGrid1.Text, Auxi)
                    If Auxi <> "S" Then
                        DBGrid1.Col = 2
                        KeyCode = 0
                                Else
                        DBGrid1.Col = 3
                        KeyCode = 0
                    End If
                End If
                
            Case 3
                If KeyCode = 13 Then
                    With rstBanco
                        .Index = "Banco"
                        DBGrid1.Col = 3
                        Claveven$ = DBGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            DBGrid1.Col = 4
                            DBGrid1.Text = !Nombre
                            DBGrid1.Col = 5
                            KeyCode = 0
                                Else
                            DBGrid1.Col = 3
                            KeyCode = 0
                        End If
                    End With
                End If

            Case 5
                If KeyCode = 13 Then
                    iRow = DBGrid1.Row
                    DBGrid1.Col = 5
                    DBGrid1.Text = Pusing("###,###.##", DBGrid1.Text)
                    Call Suma_Datos
                    DBGrid1.Row = iRow
                    If DBGrid1.Row < 10 Then
                        DBGrid1.Row = DBGrid1.Row + 1
                        DBGrid1.Col = 0
                        KeyCode = 0
                            Else
                        DBGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If

            Case Else
                
    End Select
    
End Sub
Private Sub DbGrid1_Keypress(KeyAscii As Integer)

    Select Case DBGrid1.Col
            Case 0
                Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Case 1
                Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Case 3
                Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Case 5
                Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Case Else
                
    End Select
    
End Sub

' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

ReDim UserData(0 To 6, 0 To 9)

mTotalRows& = 10

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
     DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Numero"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Fecha"
             DBGrid1.Columns(newcnt).Width = 1150
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Banco"
             DBGrid1.Columns(newcnt).Width = 500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Nombre"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Importe"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 6
             DBGrid1.Columns(newcnt).Caption = ""
             DBGrid1.Columns(newcnt).Width = 1
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Orden.Text = ""
    Observaciones.Text = ""
    Importe.Text = ""
    Fecha.Text = "  /  /    "
    Creditos.Caption = ""
    
    With rstPagos
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Orden.Text = !Orden + 1
                Else
            Orden.Text = ""
        End If
    End With
     
End Sub
