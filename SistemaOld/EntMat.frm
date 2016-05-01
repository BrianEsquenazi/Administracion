VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEntMat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Entradas de Materias Primas"
   ClientHeight    =   8175
   ClientLeft      =   210
   ClientTop       =   555
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.TextBox Remito 
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Proveedor 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   6600
      TabIndex        =   19
      Top             =   6720
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   8280
      TabIndex        =   18
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   15
      Top             =   120
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
      TabIndex        =   13
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   450
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   450
      Left            =   5280
      TabIndex        =   10
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   450
      Left            =   2640
      TabIndex        =   8
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   11895
      Begin VB.TextBox WPrecio 
         Height          =   300
         Left            =   6480
         TabIndex        =   24
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WCantidad 
         Height          =   300
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   450
      Left            =   3960
      TabIndex        =   4
      Top             =   6720
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4095
      Left            =   0
      OleObjectBlob   =   "EntMat.frx":0000
      TabIndex        =   3
      Top             =   1680
      Width           =   11895
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "EntMat.frx":09EE
      Left            =   7680
      List            =   "EntMat.frx":09F5
      TabIndex        =   1
      Top             =   6600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   450
      Left            =   1320
      TabIndex        =   0
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Remito"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DesProveedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Informe de Recepcion"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgEntmat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 4 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer

Private Sub Borra_Click()

    DbGrid1.Col = 0
    DbGrid1.Text = ""
    
    DbGrid1.Col = 1
    DbGrid1.Text = ""

    DbGrid1.Col = 2
    DbGrid1.Text = ""
    
    DbGrid1.Col = 3
    DbGrid1.Text = ""
    
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Text = ""
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
    With rstProveedor
        .Close
    End With
    With rstArticulo
        .Close
    End With
    With rstEntMat
        .Close
    End With
    
    DbsVentas.Close
    DbsAdminis.Close
    DbsCotiza.Close
    
    PrgEntmat.Hide
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Articulos"

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
            With rstProveedor
                .Index = "Proveedor"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Proveedor + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 1
            With rstArticulo
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

    DbGrid1.Col = 0
    If Len(DbGrid1.Text) = 10 Then
        WLinea.Text = DbGrid1.Row + 1
        WArticulo.Text = DbGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    DbGrid1.Col = 1
    WDescripcion.Caption = DbGrid1.Text

    DbGrid1.Col = 2
    WCantidad.Text = DbGrid1.Text
    
    DbGrid1.Col = 3
    WPrecio.Text = DbGrid1.Text
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DbGrid1.FirstRow = Lugar1
    DbGrid1.Row = Lugar2 - 1
    
    DbGrid1.Col = 0
    DbGrid1.Text = ""

        For da = 1 To 40

        With rstEntMat
    
            .Index = "Clave"
            Auxi = Codigo.Text
            Call Ceros(Auxi, 6)
            Auxi1 = da
            Call Ceros(Auxi1, 2)
            .Seek "=", Auxi + Auxi1
            If .NoMatch = False Then
                Cantidad = !Cantidad
                Articulo = !Articulo
                With rstArticulo
                    .Index = "Codigo"
                    .Seek "=", Articulo
                    If .NoMatch = False Then
                        .Edit
                        !Entradas = !Entradas - Val(Cantidad)
                        .Update
                    End If
                End With
                .Delete
            End If
        
        End With
        Next da

        Renglon = 0
        
        DbGrid1.Refresh
        
        With rstEntMat
        
            Renglon = 0
            .Index = "Clave"
                                        
            For A = 0 To 3
        
                Suma = A * 10
                DbGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRow = iRow
                    DbGrid1.Row = WRow
                    
                    DbGrid1.Col = 0
                    Articulo = DbGrid1.Text
                    
                    DbGrid1.Col = 2
                    Cantidad = DbGrid1.Text
                    
                    DbGrid1.Col = 3
                    Precio = Val(DbGrid1.Text)
                    
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
                        !Proveedor = Proveedor.Text
                        !Remito = Remito.Text
                        !Articulo = Articulo
                        !Cantidad = Val(Cantidad)
                        !Precio = Precio
                        !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        !Clave = Auxi1 + Auxi
                        .Update
                        
                        With rstArticulo
                            .Index = "Codigo"
                            .Seek "=", Articulo
                            If .NoMatch = False Then
                                .Edit
                                !Entradas = !Entradas + Val(Cantidad)
                                .Update
                            End If
                        End With
                        
                    End If
                                        
                Next iRow
            
            Next A
            
        End With
        
        Call Limpia_Click

        DbGrid1.FirstRow = 0
        DbGrid1.Col = 0
        DbGrid1.Row = 0
    
        Codigo.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Text = ""

    Codigo.Text = ""
    Fecha.Text = "  /  /    "
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Remito.Text = ""
    
    For A = 0 To 3
        Suma = A * 10
        DbGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 3
                DbGrid1.Col = iCol
                DbGrid1.Row = iRow
                DbGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    With rstEntMat
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Codigo.Text = !Codigo + 1
                Else
            Codigo.Text = ""
        End If
    End With

    
    DbGrid1.FirstRow = 0
    Renglon = 0

    Codigo.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstArticulo
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
        WPrecio.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WPrecio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPrecio.Text = PUsing("###,###.##", WPrecio.Text)
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
            With rstProveedor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Proveedor"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Proveedor.Text = Claveven$
                    DesProveedor.Caption = !Nombre
                End If
            End With
            
        Case 1
            With rstArticulo
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Codigo"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WArticulo.Text = !Codigo
                    WDescripcion.Caption = !Descripcion
                    
                    DbGrid1.Col = 0
                    DbGrid1.Text = !Codigo
                    DbGrid1.Col = 1
                    DbGrid1.Text = !Descripcion
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DbGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DbGrid1.SetFocus
                    WCantidad.SetFocus
                    
                End If
            End With
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
            Case 0, 1, 2, 3
                Select Case KeyCode
                    Case 13
                        If DbGrid1.Row < 40 Then
                            DbGrid1.Row = DbGrid1.Row + 1
                            DbGrid1.Col = 0
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
        UserData(iCol, mTotalRows - 1) = DbGrid1.Columns(iCol).DefaultValue
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
ReDim UserData(0 To 3, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DbGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DbGrid1.Columns.Count - 1 To 0 Step -1
      DbGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 3
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Producto"
             DbGrid1.Columns(newcnt).Width = 1400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Descripcion"
             DbGrid1.Columns(newcnt).Width = 3620
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Cantidad"
             DbGrid1.Columns(newcnt).Width = 1100
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Precio"
             DbGrid1.Columns(newcnt).Width = 1100
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
             DbGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    With rstEntMat
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Codigo.Text = !Codigo + 1
                Else
            Codigo.Text = ""
        End If
    End With

 
    DbGrid1.FirstRow = 0
    DbGrid1.Col = 0
    DbGrid1.Row = 0
    
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
    Suma = A * 10
    DbGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 3
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    
    For WRenglon = 1 To 40
    
    With rstEntMat
    
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
                
                DbGrid1.FirstRow = Lugar1
                DbGrid1.Row = Lugar2 - 1
                
                DbGrid1.Col = 0
                DbGrid1.Text = !Articulo
                Auxi1 = !Articulo
                
                DbGrid1.Col = 2
                DbGrid1.Text = PUsing("###,###.##", Val(!Cantidad))
                
                DbGrid1.Col = 3
                DbGrid1.Text = PUsing("###,###.##", Val(!Precio))
                
                With rstArticulo
                    .Index = "Codigo"
                    .Seek "=", Auxi1
                    If .NoMatch = False Then
                        DbGrid1.Col = 1
                        DbGrid1.Text = !Descripcion
                        WArticulo.SetFocus
                    End If
                End With
                
        End If
        
    End With
    
    Next WRenglon

    DbGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DbGrid1.FirstRow = Lugar1
    DbGrid1.Row = Lugar2 - 1
    
    DbGrid1.Col = 0
    DbGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DbGrid1.FirstRow = Lugar1
    DbGrid1.Row = Lugar2 - 1
    
    WArticulo.SetFocus

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DbGrid1.FirstRow = Lugar1
            DbGrid1.Row = Lugar2 - 1
                
            WAnterior = DbGrid1.Row
            
            DbGrid1.Col = 0
            DbGrid1.Text = WArticulo.Text
            
            DbGrid1.Col = 1
            DbGrid1.Text = WDescripcion.Caption
                
            DbGrid1.Col = 2
            DbGrid1.Text = PUsing("###,###.##", WCantidad.Text)
                
            DbGrid1.Col = 3
            DbGrid1.Text = PUsing("###,###.##", WPrecio.Text)
            
            DbGrid1.Row = Renglon
            DbGrid1.Col = 0
            
                Else
                
            DbGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DbGrid1.Row
            
            DbGrid1.Col = 0
            DbGrid1.Text = WArticulo.Text
            
            DbGrid1.Col = 1
            DbGrid1.Text = WDescripcion.Caption
                
            DbGrid1.Col = 2
            DbGrid1.Text = PUsing("###,###.##", WCantidad.Text)
            
            DbGrid1.Col = 3
            DbGrid1.Text = PUsing("###,###.##", WPrecio.Text)
            
            DbGrid1.Row = Renglon
            DbGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEntMat
            Auxi = Codigo.Text
            Call Ceros(Auxi, 6)
            .Index = "Clave"
            .Seek "=", Auxi + "01"
            If .NoMatch = False Then
                Fecha.Text = !Fecha
                Proveedor.Text = !Proveedor
                Remito.Text = !Remito
                
                With rstProveedor
                    .Index = "Proveedor"
                    Claveven$ = Proveedor.Text
                    .Seek "=", Proveedor.Text
                    If .NoMatch = False Then
                        Proveedor.Text = !Proveedor
                        DesProveedor.Caption = !Nombre
                    End If
                End With
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
            Proveedor.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            With rstProveedor
                .Index = "Proveedor"
                Claveven$ = Proveedor.Text
                .Seek "=", Proveedor.Text
                If .NoMatch Then
                    Proveedor.Text = Claveven$
                    Proveedor.SetFocus
                        Else
                    Proveedor.Text = !Proveedor
                    DesProveedor.Caption = !Nombre
                End If
            End With
        End If
        Remito.SetFocus
    End If
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

