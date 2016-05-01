VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSalTer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Salida de Productos Terminados"
   ClientHeight    =   7095
   ClientLeft      =   630
   ClientTop       =   1140
   ClientWidth     =   11070
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11070
   Visible         =   0   'False
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   450
      Left            =   1560
      TabIndex        =   18
      Top             =   6360
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6000
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   6360
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   7080
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   14
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Codigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   12
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   450
      Left            =   240
      TabIndex        =   10
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   450
      Left            =   5760
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   450
      Left            =   3120
      TabIndex        =   7
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   6855
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   17
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   450
      Left            =   4440
      TabIndex        =   3
      Top             =   6360
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "SalTer.frx":0000
      TabIndex        =   2
      Top             =   1080
      Width           =   6615
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8880
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   5715
      ItemData        =   "SalTer.frx":09F2
      Left            =   6960
      List            =   "SalTer.frx":09F9
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Movimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "PrgSalTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 3 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WLinea.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstTerminado
        .Close
    End With
    With rstSalTer
        .Close
    End With
    
    DbsVentas.Close
    DbsAdminis.Close
    DbsCotiza.Close
    
    PrgSalTer.Hide
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"

     Opcion.Visible = True
     
 End Sub



 Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            With rstTerminado
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Codigo + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Codigo
                        WIndice.AddItem IngresaItem
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

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -     -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    WCantidad.Text = DBGrid1.Text
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""

        For da = 1 To 40

        With rstSalTer
    
            .Index = "Clave"
            Auxi = Codigo.Text
            Call Ceros(Auxi, 6)
            Auxi1 = da
            Call Ceros(Auxi1, 2)
            .Seek "=", Auxi + Auxi1
            If .NoMatch = False Then
                Cantidad = !Cantidad
                Articulo = !Articulo
                With rstTerminado
                    .Index = "Codigo"
                    .Seek "=", Articulo
                    If .NoMatch = False Then
                        .Edit
                        !Salidas = !Salidas - Val(Cantidad)
                        .Update
                    End If
                End With
                .Delete
            End If
        
        End With
        Next da

        Renglon = 0
        
        DBGrid1.Refresh
        
        With rstSalTer
        
            Renglon = 0
            .Index = "Clave"
                                        
            For A = 0 To 3
        
                Suma = A * 10
                DBGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Articulo = DBGrid1.Text
                    
                    DBGrid1.Col = 2
                    Cantidad = DBGrid1.Text
                    
                    If Articulo <> "" Then
                        
                        Renglon = Renglon + 1
                        Auxi = Str$(Renglon)
                        Call Ceros(Auxi, 2)
                        
                        Auxi1 = Str$(Codigo.Text)
                        Call Ceros(Auxi1, 6)
                    
                        .AddNew
                        
                        !Codigo = Codigo.Text
                        !Renglon = Renglon
                        !Fecha = Fecha.Text
                        !Articulo = Articulo
                        !Cantidad = Val(Cantidad)
                        !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        !Clave = Auxi1 + Auxi
                        .Update
                        
                        With rstTerminado
                            .Index = "Codigo"
                            .Seek "=", Articulo
                            If .NoMatch = False Then
                                .Edit
                                !Salidas = !Salidas + Val(Cantidad)
                                .Update
                            End If
                        End With
                        
                    End If
                                        
                Next iRow
            
            Next A
            
        End With
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
    
        Codigo.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""

    Codigo.Text = ""
    Fecha.Text = "  /  /    "
    
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 2
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    With rstSalTer
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Codigo.Text = !Codigo + 1
                Else
            Codigo.Text = ""
        End If
    End With

    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Codigo.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstTerminado
            .Index = "Codigo"
            .Seek "=", WArticulo.Text
            If .NoMatch = False Then
                WDescripcion.Caption = !Descripcion
                WCantidad.SetFocus
                    Else
                WArticulo.SetFocus
            End If
        End With
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = PUsing("###,###.##", WCantidad.Text)
        Call Alta_Vector
        Call Ingresa_Click
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            With rstTerminado
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Codigo"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WArticulo.Text = !Codigo
                    WDescripcion.Caption = !Descripcion
                    
                    DBGrid1.Col = 0
                    DBGrid1.Text = !Codigo
                    DBGrid1.Col = 1
                    DBGrid1.Text = !Descripcion
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DBGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DBGrid1.SetFocus
                    WCantidad.SetFocus
                    
                End If
            End With
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 2, 0 To 40)

mTotalRows& = 40

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
For i = 0 To 2
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3620
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    With rstSalTer
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Codigo.Text = !Codigo + 1
                Else
            Codigo.Text = ""
        End If
    End With

 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 2
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    
    For WRenglon = 1 To 40
    
    With rstSalTer
    
        Auxi = Codigo.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        .Index = "Clave"
        .Seek "=", Auxi + Auxi1
        If .NoMatch = False Then
        
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 10) * 10
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                
                DBGrid1.Col = 0
                DBGrid1.Text = !Articulo
                Auxi1 = !Articulo
                
                DBGrid1.Col = 2
                DBGrid1.Text = PUsing("###,###.##", Val(!Cantidad))
                
                With rstTerminado
                    .Index = "Codigo"
                    .Seek "=", Auxi1
                    If .NoMatch = False Then
                        DBGrid1.Col = 1
                        DBGrid1.Text = !Descripcion
                        WArticulo.SetFocus
                    End If
                End With
                
        End If
        
    End With
    
    Next WRenglon

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    WArticulo.SetFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = PUsing("###,###.##", WCantidad.Text)
                
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = PUsing("###,###.##", WCantidad.Text)
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstSalTer
            Auxi = Codigo.Text
            Call Ceros(Auxi, 6)
            .Index = "Clave"
            .Seek "=", Auxi + "01"
            If .NoMatch = False Then
                Fecha.Text = !Fecha
                Call Proceso_Click
                    Else
                WCodigo = Codigo.Text
                Call Limpia_Click
                Codigo.Text = WCodigo
                Fecha.SetFocus
            End If
        End With
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            WArticulo.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

