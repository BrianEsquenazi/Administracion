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
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Leecli 
      Caption         =   "Clientes"
      Height          =   300
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Left            =   720
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   615
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
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox Opcion 
      Height          =   1620
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "Plan.frx":0000
      Left            =   5760
      List            =   "Plan.frx":0007
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSDBGrid.DBGrid DbGrid1 
      Height          =   6015
      Left            =   0
      OleObjectBlob   =   "Plan.frx":0015
      TabIndex        =   7
      Top             =   480
      Width           =   9375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Año"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
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
Private Const MAXCOLS = 11 ' Número máximo de campos del conjunto de registros.

Private Sub Lee_Datos()
    If Val(Ano.text) > 1900 Then
    
        Rem Guardo los datos
        
        WAno = Str$(Ano.text)
        Call Ceros(WAno, 4)
        WAno = Right$(WAno, 2)
        
        With rstPlan
        
            .Index = "Clave"
            
            Rem Borro los datos anteriores
            
            Salida = "N"
            ClienteAnte = 0
            
            For A = 0 To 49
            Suma = A * 10
            DbGrid1.FirstRow = Suma
            For iRow = 0 To 9
                WRow = iRow
                Auxi1 = Str$(iRow + Suma)
                Call Ceros(Auxi1, 4)
                .Seek "=", WAno + Auxi1
                If .NoMatch = False Then
                    DbGrid1.Col = 0
                    DbGrid1.Row = iRow
                    If !Cliente <> ClienteAnte Then
                        DbGrid1.text = !Cliente
                        ClienteAnte = !Cliente
                            Else
                        DbGrid1.text = ""
                    End If
                    DbGrid1.Col = 7
                    DbGrid1.Row = iRow
                    DbGrid1.text = !Cantidad
                    DbGrid1.Col = 8
                    DbGrid1.Row = iRow
                    DbGrid1.text = !Mes
                    DbGrid1.Col = 9
                    DbGrid1.Row = iRow
                    DbGrid1.text = !Modelo
                    
                    With rstCliente
                        .Index = "Cliente"
                        DbGrid1.Col = 0
                        DbGrid1.Row = WRow
                        .Seek "=", Val(DbGrid1.text)
                        If .NoMatch = False Then
                        
                            DbGrid1.Row = WRow
                            DbGrid1.Col = 1
                            DbGrid1.text = !Razon
                            DbGrid1.Col = 2
                            DbGrid1.text = !Flota
                            DbGrid1.Col = 3
                            DbGrid1.text = !Productos
                            DbGrid1.Col = 4
                            DbGrid1.text = !Norandon
                            DbGrid1.Col = 5
                            
                            Rem WTipoModelo = !TipoModelo
                            Rem With rstModelo
                            Rem     .Index = "Modelo"
                            Rem     .Seek "=", WTipoModelo
                            Rem     If .NoMatch = False Then
                            Rem         DbGrid1.text = !Descripcion
                            Rem     End If
                            Rem End With
                            
                            DbGrid1.text = !Transporte
                            DbGrid1.Col = 6
                            If Val(!Flota) <> 0 Then
                                DbGrid1.text = Val(!Productos) / Val(!Flota) * 100
                                    Else
                                DbGrid1.text = 0
                            End If
                            
                        End If
                    End With
                    
                    With rstModelo
                        .Index = "Modelo"
                        DbGrid1.Col = 9
                        DbGrid1.Row = WRow
                        .Seek "=", DbGrid1.text
                        If .NoMatch = False Then
                            DbGrid1.Col = 10
                            DbGrid1.Row = WRow
                            DbGrid1.text = !Descripcion
                        End If
                    End With
                    
                        Else
                        
                    Salida = "S"
                    Exit For
                    
                End If
            Next iRow
            
            If Salida = "S" Then
                Exit For
            End If
            
            Next A
            
            Rem DbGrid1.FirstRow = 0
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
            
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

    If Val(Ano.text) <= 1900 Then
        WPasa = "N"
        M$ = "Codigo de Año Incorrecto"
        AA% = MsgBox(M$, 0, "Archivo de Plan de Ventas")
    End If

    If Val(Ano.text) > 1900 Then
    
        WPasa = "S"

        For A = 0 To 49
        
            Suma = A * 10
            DbGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DbGrid1.Col = 7
                DbGrid1.Row = iRow
                
                If DbGrid1.text <> "" Then
                
                    DbGrid1.Col = 0
                    DbGrid1.Row = iRow
                    If DbGrid1.text <> "" Then
                        With rstCliente
                            .Index = "Cliente"
                            .Seek "=", DbGrid1.text
                            If .NoMatch = True Then
                                WPasa = "N"
                                M$ = "Codigo de Cliente Incorrecto"
                                AA% = MsgBox(M$, 0, "Archivo de Plan de Ventas")
                            End If
                        End With
                    End If
                    
                    DbGrid1.Col = 8
                    DbGrid1.Row = iRow
                    WMes = DbGrid1.text
                    If WMes <= 0 Or WMes > 12 Then
                        WPasa = "N"
                        M$ = "Codigo de Mes Incorrecto"
                        AA% = MsgBox(M$, 0, "Archivo de Plan de Ventas")
                    End If

                    DbGrid1.Col = 9
                    DbGrid1.Row = iRow
                    With rstModelo
                        .Index = "Modelo"
                        .Seek "=", DbGrid1.text
                        If .NoMatch = True Then
                            WPasa = "N"
                            M$ = "Codigo de Modelo Incorrecto"
                            AA% = MsgBox(M$, 0, "Archivo de Plan de Ventas")
                        End If
                    End With
                
                End If
                    
            Next iRow
            
        Next A
        
        If WPasa = "N" Then
            DbGrid1.FirstRow = 0
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
            Exit Sub
        End If

        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", 1
            If .NoMatch = False Then
                WNroEmpresa = !NroEmpresa
                    Else
                WNroEmpresa = 0
            End If
        End With
    
        Rem Guardo los datos
        
        WAno = Str$(Ano.text)
        Call Ceros(WAno, 4)
        WAno = Right$(WAno, 2)
        
        With rstPlan
        
            .Index = "Clave"
            
            Rem Borro los datos anteriores
            
            For iRow = 0 To 500
                Auxi1 = Str$(iRow)
                Call Ceros(Auxi1, 4)
                .Seek "=", WAno + Auxi1
                If .NoMatch = False Then
                    .Delete
                End If
            Next iRow
            
            Rem Grago los datos actuales
            
            .Index = "Clave"
            Salida = "N"
            
            For A = 0 To 49
            Suma = A * 10
            DbGrid1.FirstRow = Suma
            For iRow = 0 To 9
                
                WRow = iRow
                DbGrid1.Col = 7
                DbGrid1.Row = iRow
                If DbGrid1.text <> "" Then
                    DbGrid1.Col = 0
                    DbGrid1.Row = iRow
                    .AddNew
                    Auxi1 = Str$(iRow + Suma)
                    Call Ceros(Auxi1, 4)
                    !Ano = WAno
                    !Renglon = Right$(Auxi1, 2)
                    DbGrid1.Col = 0
                    If Val(DbGrid1.text) <> 0 Then
                        !Cliente = DbGrid1.text
                            Else
                        !Cliente = ClienteAnte
                    End If
                    ClienteAnte = !Cliente
                    WCliente = !Cliente
                    WCiudad = 1
                    With rstCliente
                        .Index = "Cliente"
                        .Seek "=", WCliente
                        If .NoMatch = False Then
                            WCiudad = !Ciudad
                        End If
                    End With
                    DbGrid1.Col = 7
                    !Cantidad = Val(DbGrid1.text)
                    DbGrid1.Col = 8
                    WMes = DbGrid1.text
                    Call Ceros(WMes, 2)
                    !Mes = WMes
                    DbGrid1.Col = 9
                    !Modelo = Val(DbGrid1.text)
                    !Ciudad = WCiudad
                    !Clave = !Ano + "00" + !Renglon
                    WCliente = !Cliente
                    WModelo = !Modelo
                    Call Ceros(WCliente, 4)
                    Call Ceros(WModelo, 4)
                    !Claveventa = !Ano + !Mes + WCliente + WModelo
                    !Empresa = 1
                    !NroEmpresa = WNroEmpresa
                    .Update
                    .Bookmark = .LastModified
                    
                            Else
                            
                    Salida = "S"
                    Exit For
                    
                End If
            Next iRow
            If Salida = "S" Then
                Exit For
            End If
            Next A
        End With
        Call cmdClose_Click
    End If
End Sub

Private Sub CmdLimpiar_Click()
    For A = 0 To 3
    Suma = A * 10
    DbGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 10
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.text = ""
        Next iCol
    Next iRow
    Next A
    DbGrid1.FirstRow = 0
    Ano.text = ""
    Rem Comision.text = ""
    Ano.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Ano.text = ""
    Ano.SetFocus
    PrgPlan.Hide
    Menu.SetFocus
End Sub


Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Ano.text) > 1900 Then
            Call Lee_Datos
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
        End If
    End If
End Sub


Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

    XRow = DbGrid1.Row
    XCol = DbGrid1.Col


     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Modelos"

     Opcion.Visible = True
     
End Sub

Private Sub Leecli_Click()

    Fila = 0

    With rstCliente
        .Index = "Razon"
        .MoveFirst
        Do
            If .EOF = False Then
                DbGrid1.Col = 0
                DbGrid1.Row = Fila
                DbGrid1.text = !Cliente
                DbGrid1.Col = 1
                DbGrid1.text = !Razon
                DbGrid1.Col = 2
                DbGrid1.text = !Flota
                DbGrid1.Col = 3
                DbGrid1.text = !Productos
                DbGrid1.Col = 4
                DbGrid1.text = !Norandon
                DbGrid1.Col = 5
                Rem WTipoModelo = ""
                Rem With rstModelo
                Rem     .Index = "Modelo"
                Rem     .Seek "=", WTipoModelo
                Rem    If .NoMatch = False Then
                Rem         DbGrid1.text = !Descripcion
                Rem     End If
                Rem End With
                DbGrid1.text = !Transporte
                DbGrid1.Col = 6
                If Val(!Flota) <> 0 Then
                    DbGrid1.text = Val(!Productos) / Val(!Flota) * 100
                        Else
                    DbGrid1.text = 0
                End If
                DbGrid1.Col = 7
                DbGrid1.text = ""
                DbGrid1.Col = 8
                DbGrid1.text = ""
                DbGrid1.Col = 9
                DbGrid1.text = ""
                DbGrid1.Col = 10
                DbGrid1.text = ""
                Fila = Fila + 1
                .MoveNext
                        Else
                Exit Do
            End If
        Loop
    End With
    
    DbGrid1.Col = 0
    DbGrid1.Row = Fila
    DbGrid1.SetFocus
    
    
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
                .Index = "Razon"
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
                        DbGrid1.Col = 2
                        DbGrid1.text = !Flota
                        DbGrid1.Col = 3
                        DbGrid1.text = !Productos
                        DbGrid1.Col = 4
                        DbGrid1.text = !Norandon
                        DbGrid1.Col = 5
                        Rem WTipoModelo = !TipoModelo
                        Rem With rstModelo
                        Rem     .Index = "Modelo"
                        Rem     .Seek "=", WTipoModelo
                        Rem     If .NoMatch = False Then
                        Rem         DbGrid1.text = !Descripcion
                        Rem     End If
                        Rem End With
                        DbGrid1.text = !Transporte
                        DbGrid1.Col = 6
                        If Val(!Flota) <> 0 Then
                            DbGrid1.text = Val(!Productos) / Val(!Flota) * 100
                                Else
                            DbGrid1.text = 0
                        End If
                            Else
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 0
                        DbGrid1.text = ""
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 1
                        DbGrid1.text = ""
                        DbGrid1.Col = 2
                        DbGrid1.text = ""
                        DbGrid1.Col = 3
                        DbGrid1.text = ""
                        DbGrid1.Col = 4
                        DbGrid1.text = ""
                        DbGrid1.Col = 5
                        DbGrid1.text = ""
                        DbGrid1.Col = 6
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
        
            If XCol = 9 Then
        
                With rstModelo

                    Indice = Pantalla.ListIndex
                    Claveven$ = WIndice.List(Indice)
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 9
                    DbGrid1.text = Claveven$
                    .Index = "Modelo"
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 10
                        DbGrid1.text = !Descripcion
                            Else
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 9
                        DbGrid1.text = ""
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 10
                        DbGrid1.text = ""
                    End If
                    
                End With
                
                DbGrid1.Row = XRow
                DbGrid1.Col = 9
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
                            DbGrid1.Col = 2
                            DbGrid1.text = !Flota
                            DbGrid1.Col = 3
                            DbGrid1.text = !Productos
                            DbGrid1.Col = 4
                            DbGrid1.text = !Norandon
                            DbGrid1.Col = 5
                            Rem WTipoModelo = !TipoModelo
                            Rem With rstModelo
                            Rem     .Index = "Modelo"
                            Rem     .Seek "=", WTipoModelo
                            Rem     If .NoMatch = False Then
                            Rem         DbGrid1.text = !Descripcion
                            Rem     End If
                            Rem End With
                            DbGrid1.text = !Transporte
                            DbGrid1.Col = 6
                            If Val(!Flota) <> 0 Then
                                DbGrid1.text = Val(!Productos) / Val(!Flota) * 100
                                    Else
                                DbGrid1.text = 0
                            End If
                            DbGrid1.Col = 7
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End With
                End If
                
            Case 7
                If KeyCode = 13 Then
                    DbGrid1.Col = 8
                    KeyCode = 0
                    Rem DbGrid1.text = Str$(Val(DbGrid1.text))
                End If
                
            Case 8
                If KeyCode = 13 Then
                    DbGrid1.Col = 8
                    If Val(DbGrid1.text) > 0 And Val(DbGrid1.text) < 13 Then
                            DbGrid1.Col = 9
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 8
                            KeyCode = 0
                    End If
                End If
                
            Case 9
                If KeyCode = 13 Then
                    With rstModelo
                        .Index = "Modelo"
                        .Seek "=", Val(DbGrid1.text)
                        If .NoMatch = False Then
                            DbGrid1.Col = 10
                            DbGrid1.text = !Descripcion
                            DbGrid1.Col = 0
                            If DbGrid1.Row < 21 Then
                                DbGrid1.Row = DbGrid1.Row + 1
                                DbGrid1.Col = 7
                            End If
                            KeyCode = 0
                            
                                Else
                            DbGrid1.Col = 9
                            KeyCode = 0
                        End If
                    End With
                End If
                
            Case 4
                If KeyCode = 13 Then
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


Private Sub Form_load()

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 10, 0 To 499)

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

mTotalRows& = 500

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
For i = 0 To 10
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Cliente"
             DbGrid1.Columns(newcnt).Width = 600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Razon Social"
             DbGrid1.Columns(newcnt).Width = 1500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Flota"
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Randon"
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 4
             DbGrid1.Columns(newcnt).Caption = "No Randon"
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 5
             DbGrid1.Columns(newcnt).Caption = "Productos"
             DbGrid1.Columns(newcnt).Width = 2000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 6
             DbGrid1.Columns(newcnt).Caption = "% Pen."
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 7
             DbGrid1.Columns(newcnt).Caption = "Cantidad"
             DbGrid1.Columns(newcnt).Width = 850
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 8
             DbGrid1.Columns(newcnt).Caption = "Mes"
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 9
             DbGrid1.Columns(newcnt).Caption = "Modelo"
             DbGrid1.Columns(newcnt).Width = 800
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 10
             DbGrid1.Columns(newcnt).Caption = "Descricion"
             DbGrid1.Columns(newcnt).Width = 1600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
 Ano.text = ""

End Sub
