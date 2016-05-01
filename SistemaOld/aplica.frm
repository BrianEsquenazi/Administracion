VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PrgAplica 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Aplicacion de Cuenta Corriente de Proveedores"
   ClientHeight    =   6330
   ClientLeft      =   1890
   ClientTop       =   1455
   ClientWidth     =   8250
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   8250
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   300
      Left            =   6000
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Proveedor 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   0
      Text            =   " "
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
      Height          =   300
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "aplica.frx":0000
      TabIndex        =   5
      Top             =   2280
      Width           =   7815
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "aplica.frx":09D6
      Left            =   120
      List            =   "aplica.frx":09DD
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6000
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label DesProveedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "PrgAplica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 8 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Auxi As String
Private dada As String
Private Importe As Double
Private WSuma As Double
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim RstAplicaProve As Recordset
Dim spAplicaProve As String
Dim cParam As String
Dim XParam As String
Private WSaldo As Double

Dim ZZClave As String
Dim ZZCodigo As String
Dim ZZRenglon As String
Dim ZZFecha As String
Dim ZZOrdFecha As String
Dim ZZProveedor As String
Dim ZZTipo As String
Dim ZZLetra As String
Dim ZZPunto As String
Dim ZZNumero As String
Dim ZZImporte As String


Private Sub cmdClose_Click()
    PrgAplica.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spProveedor = "ListaProveedoresOrdConsulta"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Proveedor
                    Call Ceros(Auxi, 11)
                    IngresaItem = Auxi + "      " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
    
    Pantalla.Visible = True

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Graba_Click()

        WPasa = "S"
        
        WSuma = 0
        
        For A = 0 To 18
        
            Suma = A * 10
            DbGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                        
                DbGrid1.Col = 0
                DbGrid1.Row = iRow
                Tipo = DbGrid1.Text
                
                DbGrid1.Col = 6
                DbGrid1.Row = iRow
                ImporteIni = Val(DbGrid1.Text)
                
                DbGrid1.Col = 7
                DbGrid1.Row = iRow
                Importe = Val(DbGrid1.Text)
               
                If Val(Tipo) <> 0 Then
                    If Val(Tipo) = 1 Or Val(Tipo) = 2 Then
                        WSuma = WSuma + Importe
                            Else
                        WSuma = WSuma - Importe
                    End If
                End If
        
                If Importe > ImporteIni Then
                    WPasa = "N"
                    m$ = "Importe a aplicar mayor al original"
                    Call Errores(coderr, "Archivo de Aplicacion de Comprobantes", m$)
                    Exit Sub
                End If
                        
            Next iRow
            
        Next A
        
        Call Redondeo(WSuma)
        
        If WSuma <> 0 Then
            WPasa = "N"
            m$ = "Importe a aplicar no balancea"
            Call Errores(coderr, "Archivo de Aplicacion de Comprobantes", "No balancea los importes")
            Exit Sub
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
        ZSql = ZSql + " FROM AplicaProve"
        spAplicaProve = ZSql
        Set RstAplicaProve = db.OpenRecordset(spAplicaProve, dbOpenSnapshot, dbSQLPassThrough)
        If RstAplicaProve.RecordCount > 0 Then
            RstAplicaProve.MoveLast
            ZUltimo = IIf(IsNull(RstAplicaProve!CodigoMayor), "0", RstAplicaProve!CodigoMayor)
            WCodigo = Mid$(Str$(ZUltimo + 1), 2, 8)
            RstAplicaProve.Close
                Else
            WCodigo = "1"
        End If
        
        ZLugar = 0
        DbGrid1.FirstRow = 0
        DbGrid1.Col = 0
        DbGrid1.Row = 0
        
        If WPasa = "S" Then
        
            Renglon = 0
                                        
            For A = 0 To 18
        
                Suma = A * 10
                DbGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRow = iRow
                    DbGrid1.Row = WRow
                    DbGrid1.Col = 7
                    Importe = Val(DbGrid1.Text)
                    
                    If Importe <> 0 Then
                    
                        DbGrid1.Col = 0
                        Tipo = DbGrid1.Text
                        
                        DbGrid1.Col = 1
                        Letra = DbGrid1.Text
                        
                        DbGrid1.Col = 2
                        Punto = DbGrid1.Text
                        
                        DbGrid1.Col = 3
                        Numero = DbGrid1.Text
                        
                        DbGrid1.Col = 7
                        Importe = Val(DbGrid1.Text)
                        
                        Clave = Proveedor.Text + Letra + Tipo + Punto + Numero
                        cParam = "'" & Clave & "'"
                        spCtaprv = "ConsultaCtaCtePrv "
                        Set RstCtaPrv = db.OpenRecordset(spCtaprv + cParam, dbOpenSnapshot, dbSQLPassThrough)
                        If RstCtaPrv.RecordCount > 0 Then
                            WClave = RstCtaPrv!Clave
                            If Val(RstCtaPrv!Tipo) = 1 Or Val(RstCtaPrv!Tipo) = 2 Then
                                XSaldo = Str$(RstCtaPrv!Saldo - Importe)
                            Else
                                XSaldo = Str$(RstCtaPrv!Saldo + Importe)
                            End If
                            XParam = "'" + WClave + "','" _
                                    + XSaldo + "'"
                            spCtaprv = "ActualizaSaldoCtaCtePrv " + XParam
                            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        ZLugar = ZLugar + 1
                        
                        ZZCodigo = WCodigo
                        ZZRenglon = Str$(ZLugar)
                        ZZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                        ZZOrdFecha = Right$(ZZFecha, 4) + Mid$(ZZFecha, 4, 2) + Left$(ZZFecha, 2)
                        ZZProveedor = Proveedor.Text
                        ZZTipo = Tipo
                        ZZLetra = Letra
                        ZZPunto = Punto
                        ZZNumero = Numero
                        If Val(Tipo) = 1 Or Val(Tipo) = 2 Then
                            ZZImporte = Str$(Importe)
                                Else
                            ZZImporte = Str$(Importe * -1)
                        End If
                        
                        Auxi = ZZCodigo
                        Call Ceros(Auxi, 8)
                        Auxi1 = ZZRenglon
                        Call Ceros(Auxi1, 3)
                        ZZClave = Auxi + Auxi1
                        
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO AplicaProve ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "Codigo ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "OrdFecha ,"
                        ZSql = ZSql + "Proveedor ,"
                        ZSql = ZSql + "Tipo ,"
                        ZSql = ZSql + "Letra ,"
                        ZSql = ZSql + "Punto ,"
                        ZSql = ZSql + "Numero ,"
                        ZSql = ZSql + "Importe )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + ZZClave + "',"
                        ZSql = ZSql + "'" + ZZCodigo + "',"
                        ZSql = ZSql + "'" + ZZRenglon + "',"
                        ZSql = ZSql + "'" + ZZFecha + "',"
                        ZSql = ZSql + "'" + ZZOrdFecha + "',"
                        ZSql = ZSql + "'" + ZZProveedor + "',"
                        ZSql = ZSql + "'" + ZZTipo + "',"
                        ZSql = ZSql + "'" + ZZLetra + "',"
                        ZSql = ZSql + "'" + ZZPunto + "',"
                        ZSql = ZSql + "'" + ZZNumero + "',"
                        ZSql = ZSql + "'" + ZZImporte + "')"
                            
                        spAplicaProve = ZSql
                        Set RstAplicaProve = db.OpenRecordset(spAplicaProve, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                                        
                Next iRow
            
            Next A
            
        End If

        For A = 0 To 18
        Suma = A * 10
        DbGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 7
                DbGrid1.Col = iCol
                DbGrid1.Row = iRow
                DbGrid1.Text = ""
            Next iCol
        Next iRow
        Next A
        
        Proveedor.Text = ""
        Desproveedor.Caption = ""

        DbGrid1.FirstRow = 0
        DbGrid1.Col = 0
        DbGrid1.Row = 0
    
        Proveedor.SetFocus
        
End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        Proveedor.Text = RstProveedor!Proveedor
        Desproveedor.Caption = RstProveedor!Nombre
        RstProveedor.Close
        Call Proceso_Click
        Rem DBGrid1.FirstRow = 0
        Rem DBGrid1.Col = 0
        Rem DBGrid1.Row = 0
        DbGrid1.SetFocus
            Else
        Proveedor.Text = Claveven$
    End If
    Proveedor.SetFocus
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
            Case 7
                If KeyCode = 13 Then
                    DbGrid1.Text = Pusing("###,###,###.##", DbGrid1.Text)
                    If DbGrid1.Row < 100 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 7
                        KeyCode = 0
                    End If
                End If
                
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
ReDim UserData(0 To 7, 0 To 1000)

mTotalRows& = 1000

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
For i = 0 To 7
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Letra"
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Punto"
             DbGrid1.Columns(newcnt).Width = 500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Fecha"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 5
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1100
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 6
             DbGrid1.Columns(newcnt).Caption = "Saldo"
             DbGrid1.Columns(newcnt).Width = 1100
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 7
             DbGrid1.Columns(newcnt).Caption = "Aplica"
             DbGrid1.Columns(newcnt).Width = 1100
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Proveedor.Text = ""
    Desproveedor.Caption = ""

    DbGrid1.FirstRow = 0
    DbGrid1.Col = 0
    DbGrid1.Row = 0
    
    Proveedor.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 18
    Suma = A * 10
    DbGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 7
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    DbGrid1.FirstRow = 0
    
    Renglon = 0
    
    XParam = "'" + Proveedor.Text + "','" _
                 + Proveedor.Text + "'"
    spCtaprv = "ListaCtaprvDesdeHasta " + XParam
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                WSaldo = !Saldo
                Call Redondeo(WSaldo)
            
                If WSaldo <> 0 Then
            
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DbGrid1.FirstRow = Lugar1
                    DbGrid1.Row = Lugar2 - 1
                
                    DbGrid1.Col = 0
                    DbGrid1.Text = !Tipo
                
                    DbGrid1.Col = 1
                    DbGrid1.Text = !Letra
                
                    DbGrid1.Col = 2
                    DbGrid1.Text = !Punto
                
                    DbGrid1.Col = 3
                    DbGrid1.Text = !Numero
                
                    DbGrid1.Col = 4
                    DbGrid1.Text = !Fecha
                
                    DbGrid1.Col = 5
                    DbGrid1.Text = Pusing("###,###,###.##", Str$(Abs(!Total)))
                
                    DbGrid1.Col = 6
                    DbGrid1.Text = Pusing("###,###,###.##", Str$(Abs(!Saldo)))
                
                    DbGrid1.Col = 7
                    DbGrid1.Text = ""
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    End If

    Rem DbGrid1.FirstRow = 0
    Rem DbGrid1.Col = 7
    Rem DbGrid1.Row = 0

End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WProveedor = Proveedor.Text
        Proveedor.Text = WProveedor
        spProveedor = "ConsultaProveedores "
        cParam = "'" & Proveedor.Text & "'"
        Set RstProveedor = db.OpenRecordset(spProveedor + cParam, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            Desproveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
            Call Proceso_Click
            Rem DBGrid1.FirstRow = 0
            Rem DBGrid1.Col = 0
            Rem DBGrid1.Row = 0
            Rem DbGrid1.SetFocus
                Else
            Proveedor.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

