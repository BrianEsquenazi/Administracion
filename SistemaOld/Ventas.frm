VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgPlan 
   Caption         =   "Ingreso de Plan de ventas"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6120
   ScaleWidth      =   9480
   Begin VB.TextBox Mes 
      Height          =   285
      Left            =   600
      MaxLength       =   2
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin MSDBGrid.DBGrid DbGrid1 
      Height          =   6015
      Left            =   0
      OleObjectBlob   =   "Ventas.frx":0000
      TabIndex        =   11
      Top             =   480
      Width           =   6615
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4680
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   " "
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8280
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "Ventas.frx":09D2
      Left            =   5760
      List            =   "Ventas.frx":09D9
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
      Height          =   300
      Left            =   8400
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   8400
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   8400
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   8400
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   8400
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   8400
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1620
      Left            =   6240
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Mes"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Año"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 5 ' Número máximo de campos del conjunto de registros.

Private Sub Lee_Datos()
    If Ano.text <> "" Then
    
        Rem Guardo los datos
        
        WAno = Str$(Ano.text)
        Call Ceros(WAno, 2)
        
        WMes = Str$(Mes.text)
        Call Ceros(WMes, 2)
        
        With rstPlan
        
            .Index = "Clave"
            
            Rem Borro los datos anteriores
            
            For iRow = 0 To 20
                WRow = iRow
                Auxi1 = Str$(iRow)
                Call Ceros(Auxi1, 2)
                .Seek "=", WAno + WMes + Auxi1
                If .NoMatch = False Then
                    DbGrid1.Col = 0
                    DbGrid1.Row = iRow
                    DbGrid1.text = !Cliente
                    DbGrid1.Col = 2
                    DbGrid1.Row = iRow
                    DbGrid1.text = !Modelo
                    DbGrid1.Col = 4
                    DbGrid1.Row = iRow
                    DbGrid1.text = !Cantidad
                    
                    With rstCliente
                        .Index = "Cliente"
                        DbGrid1.Col = 0
                        DbGrid1.Row = WRow
                        .Seek "=", DbGrid1.text
                        If .NoMatch = False Then
                            DbGrid1.Col = 1
                            DbGrid1.Row = WRow
                            DbGrid1.text = !Razon
                        End If
                    End With
                    
                    With rstModelo
                        .Index = "Modelo"
                        DbGrid1.Col = 2
                        DbGrid1.Row = WRow
                        .Seek "=", DbGrid1.text
                        If .NoMatch = False Then
                            DbGrid1.Col = 3
                            DbGrid1.Row = WRow
                            DbGrid1.text = !Descripcion
                        End If
                    End With
                    
                End If
            Next iRow
        End With
        
    End If
End Sub
Sub Verifica_datos()
    Rem If Val(Comision.text) = 0 Then
    Rem     Comision.text = "0"
    Rem End If
End Sub
Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()
    Rem With rstVendedor
    Rem     .Index = "Vendedor"
    Rem     .Seek "=", Vendedor.text
    Rem     If .NoMatch = False Then
    Rem         Vendedor.text = !Vendedor
    Rem         DesVendedor.Caption = !Nombre
    Rem         Call Format_datos
    Rem     End If
    Rem End With
End Sub


Private Sub cmdAdd_Click()
    If Ano.text <> "" Then
    
        Rem Guardo los datos
        
        WAno = Str$(Ano.text)
        Call Ceros(WAno, 2)
        
        WMes = Str$(Mes.text)
        Call Ceros(WMes, 2)
        
        With rstPlan
        
            .Index = "Clave"
            
            Rem Borro los datos anteriores
            
            For iRow = 0 To 20
                Auxi1 = Str$(iRow)
                Call Ceros(Auxi1, 2)
                .Seek "=", WAno + WMes + Auxi1
                If .NoMatch = False Then
                    .Delete
                End If
            Next iRow
            
            Rem Grago los datos actuales
            
            .Index = "Clave"
            
            For iRow = 0 To 20
                WRow = iRow
                DbGrid1.Col = 0
                DbGrid1.Row = iRow
                If DbGrid1.text <> "" Then
                    .AddNew
                    Auxi1 = Str$(iRow)
                    Call Ceros(Auxi1, 2)
                    !Ano = WAno
                    !Mes = WMes
                    !Renglon = Auxi1
                    DbGrid1.Col = 0
                    !Cliente = DbGrid1.text
                    WCliente = !Cliente
                    WCiudad = 1
                    With rstCliente
                        .Index = "Cliente"
                        .Seek "=", WCliente
                        If .NoMatch = False Then
                            WCiudad = !Ciudad
                        End If
                    End With
                    DbGrid1.Col = 2
                    !Modelo = DbGrid1.text
                    DbGrid1.Col = 4
                    !Cantidad = DbGrid1.text
                    !Ciudad = WCiudad
                    !Clave = !Ano + !Mes + !Renglon
                    WCliente = !Cliente
                    WModelo = !Modelo
                    Call Ceros(WCliente, 4)
                    Call Ceros(WModelo, 4)
                    !Claveventa = !Ano + !Mes + WCliente + WModelo
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                    
                End If
            Next iRow
        End With
        Call CmdLimpiar_Click
        Mes.SetFocus
    End If
End Sub


Private Sub CmdLimpiar_Click()
    For iCol = 0 To 4
        For iRow = 0 To 20
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.text = ""
        Next iRow
    Next iCol
    Ano.text = ""
    Mes.text = ""
    Rem Comision.text = ""
    Mes.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Mes.SetFocus
    PrgPlan.Hide
    Menu.SetFocus
End Sub


Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Mes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Mes.text) >= 1 And Val(Mes.text) <= 12 Then
            Ano.SetFocus
        End If
    End If
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Lee_Datos
        DbGrid1.Col = 0
        DbGrid1.Row = 0
        DbGrid1.SetFocus
    End If
End Sub


Private Sub Consulta_Click()

    XRow = DbGrid1.Row
    XCol = DbGrid1.Col


     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Modelos"

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
            With rstCliente
                .Index = "Cliente"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Cliente) + " " + !Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 1
            With rstModelo
                .Index = "Modelo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Modelo) + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Modelo
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

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
        
            If XCol = 0 Then
        
                With rstCliente

                    Indice = Pantalla.ListIndex
                    Claveven$ = WIndice.List(Indice)
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 0
                    DbGrid1.text = Claveven$
                    .Index = "Cliente"
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 1
                        DbGrid1.text = !Razon
                            Else
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 0
                        DbGrid1.text = ""
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 1
                        DbGrid1.text = ""
                    End If
                    
                End With
                
                DbGrid1.Row = XRow
                DbGrid1.Col = 0
                DbGrid1.SetFocus
                
                    Else
                    
                DbGrid1.Row = XRow
                DbGrid1.Col = XCol
                DbGrid1.SetFocus

            End If
            
        Case 1
        
            If XCol = 2 Then
        
                With rstModelo

                    Indice = Pantalla.ListIndex
                    Claveven$ = WIndice.List(Indice)
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 2
                    DbGrid1.text = Claveven$
                    .Index = "Modelo"
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 3
                        DbGrid1.text = !Descripcion
                            Else
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 2
                        DbGrid1.text = ""
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 3
                        DbGrid1.text = ""
                    End If
                    
                End With
                
                DbGrid1.Row = XRow
                DbGrid1.Col = 2
                DbGrid1.SetFocus
                
                    Else
                    
                DbGrid1.Row = XRow
                DbGrid1.Col = XCol
                DbGrid1.SetFocus

            End If
    
        Case Else
    End Select
    
End Sub
Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
            Case 0
                If KeyCode = 13 Then
                    With rstCliente
                        .Index = "Cliente"
                        .Seek "=", Val(DbGrid1.text)
                        If .NoMatch = False Then
                            DbGrid1.Col = 1
                            DbGrid1.text = !Razon
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End With
                End If
                
            Case 2
                If KeyCode = 13 Then
                    With rstModelo
                        .Index = "Modelo"
                        .Seek "=", Val(DbGrid1.text)
                        If .NoMatch = False Then
                            DbGrid1.Col = 3
                            DbGrid1.text = !Descripcion
                                Else
                            DbGrid1.Col = 2
                            KeyCode = 0
                        End If
                    End With
                End If
                
            Case 4
                If KeyCode = 13 Then
                    DbGrid1.Col = 4
                    Rem DbGrid1.text = Str$(Val(DbGrid1.text))
                    If DbGrid1.Row < 21 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If
                
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
ReDim UserData(0 To 4, 0 To 20)

' Fila 1
Rem UserData(0, 0) = "Cristina"
Rem UserData(1, 0) = "Martínez"
Rem UserData(2, 0) = "19"

' Fila 2
Rem UserData(0, 1) = "Juan"
Rem UserData(1, 1) = "Pardo"
Rem UserData(2, 1) = "12"
' Fila 3
Rem UserData(0, 2) = "Mara"
Rem UserData(1, 2) = "Belén"
Rem UserData(2, 2) = "33"
' Fila 4
Rem UserData(0, 3) = "Daniel"
Rem UserData(1, 3) = "Rendich"
Rem UserData(2, 3) = "40"
' Fila 5
Rem UserData(0, 4) = "Federico"
Rem UserData(1, 4) = "Couto"
Rem UserData(2, 4) = "42"
' Fila 6
Rem UserData(0, 5) = "María Asunción"
Rem UserData(1, 5) = "Rodríguez"
Rem UserData(2, 5) = "28"
' Fila 7
Rem UserData(0, 6) = "Agustina"
Rem UserData(1, 6) = "Rivera"
Rem UserData(2, 6) = "29"
' Fila 8
Rem UserData(0, 7) = "Enrique"
Rem UserData(1, 7) = "Ballinea"
Rem UserData(2, 7) = "42"
' Fila 9
Rem UserData(0, 8) = "Susana"
Rem UserData(1, 8) = "Medrano"
Rem UserData(2, 8) = "31"
' Fila 10
Rem UserData(0, 9) = "Guillermo"
Rem UserData(1, 9) = "Alzaga"
Rem UserData(2, 9) = "30"
' Fila 11
Rem UserData(0, 10) = "Luis"
Rem UserData(1, 10) = "Romero"
Rem UserData(2, 10) = "28"
' Fila 12
Rem UserData(0, 11) = "Ernesto"
Rem UserData(1, 11) = "Méndez"
Rem UserData(2, 11) = "60"
' Fila 13
Rem UserData(0, 12) = "Antonio"
Rem UserData(1, 12) = "Abelló"
Rem UserData(2, 12) = "33"
' Fila 14
Rem UserData(0, 13) = "Ana"
Rem UserData(1, 13) = "Martínez"
Rem UserData(2, 13) = "30"
' Fila 15
Rem UserData(0, 14) = "Francisco"
Rem UserData(1, 14) = "Colusso"
Rem UserData(2, 14) = "60"

mTotalRows& = 21

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
For i = 0 To 4
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Cliente"
             DbGrid1.Columns(newcnt).Width = 600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Razon Social"
             DbGrid1.Columns(newcnt).Width = 2000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Modelo"
             DbGrid1.Columns(newcnt).Width = 800
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Descricion"
             DbGrid1.Columns(newcnt).Width = 2000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Cantidad"
             DbGrid1.Columns(newcnt).Width = 850
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
 Mes.text = ""
 Ano.text = ""

End Sub
