VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsFicTerAnt 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Ficha de Stock de Producto Terminado Historico"
   ClientHeight    =   7920
   ClientLeft      =   195
   ClientTop       =   735
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   11655
   Begin Crystal.CrystalReport Listado 
      Left            =   5640
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wloteter.rpt"
   End
   Begin VB.TextBox RE 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Nk 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   " "
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox StkProceso 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   " "
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Deposito 
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   " "
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Unidad 
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   " "
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox XStock 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox XSalidas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox XEntradas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   " "
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox XInicial 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327680
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####-###"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
      Height          =   300
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4815
      Left            =   0
      OleObjectBlob   =   "consficterAnt.frx":0000
      TabIndex        =   4
      Top             =   3000
      Width           =   11535
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "consficterAnt.frx":09D6
      Left            =   120
      List            =   "consficterAnt.frx":09DD
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3480
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Stock Re"
      Height          =   255
      Left            =   8160
      TabIndex        =   25
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Stock NK"
      Height          =   255
      Left            =   8160
      TabIndex        =   20
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Stock en Proceso"
      Height          =   255
      Left            =   8160
      TabIndex        =   19
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Deposito"
      Height          =   255
      Left            =   8160
      TabIndex        =   18
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Unidad de Medida"
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo Final"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Salidas"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Entradas"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Saldo Inicial"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label DesTerminado 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Articulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "PrgConsFicTerAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 9 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Vector(8000, 11) As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim XParam As String
Dim Auxiliar(10000, 6) As String
Dim XLote(12, 2) As String
Dim WXEntrada As Double
Dim WXSalida As Double
Dim WXInicial As Double
Dim WXStock As Double
Dim WSaldo As Double

Private Sub cmdClose_Click()
    Terminado.SetFocus
    PrgConsFicTerAnt.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstTerminado
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstTerminado!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstTerminado.Close
            
    Pantalla.Visible = True

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaTer
End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Terminado.Text = rstTerminado!Codigo
        DesTerminado.Caption = rstTerminado!Descripcion
        Unidad.Text = rstTerminado!Unidad
        Deposito.Text = rstTerminado!Deposito
        StkProceso.Text = Pusing("###,###.##", Str$(rstTerminado!Proceso))
        rstTerminado.Close
        Call Proceso_Click
        Rem DBGrid1.FirstRow = 0
        Rem DBGrid1.Col = 0
        Rem DBGrid1.Row = 0
        DBGrid1.SetFocus
            Else
        Terminado.Text = Claveven$
    End If
    Terminado.SetFocus
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 1, 2, 3, 4, 5
                Rem If KeyCode = 13 Then
                Rem    If DBGrid1.Row < 100 Then
                Rem        DBGrid1.Row = DBGrid1.Row + 1
                Rem        DBGrid1.Col = 7
                Rem        KeyCode = 0
                Rem    End If
                Rem End If
                
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 8, 0 To 8000)

mTotalRows& = 8000

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
For i = 0 To 8
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Fecha"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Numero"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Observaciones"
             DBGrid1.Columns(newcnt).Width = 2800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Entrada"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Salida"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Remito"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 7
             DBGrid1.Columns(newcnt).Caption = "Partida"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 8
             DBGrid1.Columns(newcnt).Caption = "Saldo"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True

         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgConsFicTerAnt.Caption = "Consulta de Ficha de Stock de Producto Terminado Historico :  " + !Nombre
        End If
    End With
    
    
    Terminado.SetFocus
    
End Sub

Private Sub Proceso_Click()
        
    Terminado.Text = UCase(Terminado.Text)
    
    WXInicial = 0
    WXEntradas = 0
    WXSalidas = 0
    WXStock = 0

    For a = 0 To 50
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        DBGrid1.Col = 0
        DBGrid1.Row = iRow
        If DBGrid1.Text = "" Then
            Salida = "S"
            Exit For
        End If
        For iCol = 0 To 8
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    If Salida = "S" Then
        Exit For
    End If
    Next a
    DBGrid1.FirstRow = 0
    
    Renglon = 0
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        WTerminado = rstTerminado!Codigo
        WInicial = rstTerminado!Inicial
        WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
                                        
        Renglon = Renglon + 1
                                
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = "14/12/2000"
                        
        DBGrid1.Col = 1
        DBGrid1.Text = ""
                        
        DBGrid1.Col = 2
        DBGrid1.Text = ""
                        
        DBGrid1.Col = 3
        DBGrid1.Text = "Saldo Inicial"
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Pusing("###,###.##", Str$(rstTerminado!Inicial))
                
        DBGrid1.Col = 5
        DBGrid1.Text = ""
                
        DBGrid1.Col = 6
        DBGrid1.Text = ""
        
        DBGrid1.Col = 7
        DBGrid1.Text = ""
        
        DBGrid1.Col = 8
        DBGrid1.Text = ""
                
        WXInicial = rstTerminado!Inicial
        
        rstTerminado.Close
        
    End If
    
    Rem PROCESA LAS ESTADISTICAS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spEstadistica = "ListaEstadisticaDesdeHasta" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WFec = IIf(IsNull(rstEstadistica!Fecha), "", rstEstadistica!Fecha)
                
                If WFec <> "" Then

                
                XFec = Right$(rstEstadistica!Fecha, 4) + Mid$(rstEstadistica!Fecha, 4, 2) + Left$(rstEstadistica!Fecha, 2)
                If XFec < "20001218" Or XFec > WOrdFechaCierre Then
                
                    Else
                
                WTipo = rstEstadistica!Tipo
                WTerminado = rstEstadistica!Articulo
                WSalida = rstEstadistica!Cantidad
                WFecha = rstEstadistica!Fecha
                WNumero = rstEstadistica!Numero
                WImpre1 = rstEstadistica!Cliente
                WCliente = rstEstadistica!Cliente
                
                Erase XLote
                
                XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                
                WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                            
                If Len(Trim(WLoteAdicional)) = 98 Then
                    XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                    XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                    XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                    XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                    XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                    XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                    XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                    XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                    XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                    XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                    XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                    XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                    XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                    XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                        Else
                    XLote(6, 1) = "0"
                    XLote(6, 2) = "0"
                    XLote(7, 1) = "0"
                    XLote(7, 2) = "0"
                    XLote(8, 1) = "0"
                    XLote(8, 2) = "0"
                    XLote(9, 1) = "0"
                    XLote(9, 2) = "0"
                    XLote(10, 1) = "0"
                    XLote(10, 2) = "0"
                    XLote(11, 1) = "0"
                    XLote(11, 2) = "0"
                    XLote(12, 1) = "0"
                    XLote(12, 2) = "0"
                End If
                
                If XLote(1, 2) = 0 Then
                    XLote(1, 2) = rstEstadistica!Cantidad
                End If
                
                For x = 1 To 12
            
                    If XLote(x, 2) <> 0 Then
                
                        WSalida = XLote(x, 2)
                        Lugar = Lugar + 1
                    
                        Vector(Lugar, 1) = WFecha
                        If Val(WTipo) = 1 Then
                            Vector(Lugar, 2) = "Fac"
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(WSalida))
                            WXSalidas = WXSalidas + WSalida
                                Else
                            Vector(Lugar, 2) = "Dev"
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(Abs(WSalida)))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + Abs(WSalida)
                        End If
                        Vector(Lugar, 3) = WNumero
                        Vector(Lugar, 4) = WImpre2
                        If Left$(rstEstadistica!Remito, 1) = "C" Then
                            Vector(Lugar, 7) = Str$(Val(Mid$(rstEstadistica!Remito, 2, 9)))
                                Else
                            Vector(Lugar, 7) = Str$(Val(rstEstadistica!Remito))
                        End If
                        Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        Vector(Lugar, 9) = WImpre1
                        Vector(Lugar, 10) = XLote(x, 1)
                        Vector(Lugar, 11) = ""
                    End If
                
                Next x
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
        
        rstEstadistica.Close
        
    End If
    
    For Ciclo = 1 To Lugar
    
        WImpre1 = Vector(Ciclo, 9)
        
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spCliente = "ConsultaCliente" + "'" + WImpre1 + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WImpre2 = rstCliente!Razon
            rstCliente.Close
                Else
            WImpre2 = ""
        End If
        
        Call Conecta_Empresa
    
        Vector(Ciclo, 4) = WImpre2
    
    Next Ciclo
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla
    
    
    Rem PROCESA LAS HOJAS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spHoja = "ListaHojaTerminadoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If XFec < WOrdFechaCierre Then
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                WMarcaAnt = IIf(IsNull(rstHoja!MarcaAnt), "", rstHoja!MarcaAnt)
                Rem If WMarcaAnt = "X" Or XFec < "20001218" Then
                If XFec < "20001218" Then
                
                    Else
                            
                If rstHoja!Tipo = "T" Then
                
                    WTerminado = rstHoja!Terminado
                    WCantidad = rstHoja!Cantidad
                    WFecha = rstHoja!Fecha
                    WHoja = rstHoja!Hoja
                    
                    Erase XLote
                
                    XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                    XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                    XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                    XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                
                    If XLote(1, 2) = 0 Then
                        XLote(1, 2) = rstHoja!Cantidad
                    End If
                
                    For x = 1 To 3
                
                        If XLote(x, 2) <> 0 Then
                    
                            WCantidad = XLote(x, 2)
                            Lugar = Lugar + 1
                    
                            Vector(Lugar, 1) = WFecha
                            Vector(Lugar, 2) = "Hoja"
                            Vector(Lugar, 3) = WHoja
                            Vector(Lugar, 4) = ""
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                            Vector(Lugar, 7) = ""
                            Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            Vector(Lugar, 9) = ""
                            Vector(Lugar, 10) = XLote(x, 1)
                            Vector(Lugar, 11) = ""

                            WXSalidas = WXSalidas + WCantidad
                            
                        End If
                        
                    Next x
                
                End If
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla
    
    Rem PROCESA LAS HOJAS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
            
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If XFec < WOrdFechaCierre Then

            
                Rem WMarcaAnt = IIf(IsNull(rstHoja!MarcaAnt), "", rstHoja!MarcaAnt)
                Rem If WMarcaAnt = "X" And rstHoja!Saldo = 0 Then
                Rem
                Rem     Else
                
                If Val(rstHoja!Renglon) = 1 Then
                Rem And rstHoja!Realant <> 0 Then
                 
                    WProducto = rstHoja!Producto
                    WCantidad = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
                    WFecha = rstHoja!Fecha
                    WHoja = rstHoja!Hoja
                    WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(WSaldo)
                    
                    Lugar = Lugar + 1
                    
                    Vector(Lugar, 1) = WFecha
                    Vector(Lugar, 2) = "Hoja"
                    Vector(Lugar, 3) = WHoja
                    Vector(Lugar, 4) = ""
                    Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                    Vector(Lugar, 6) = ""
                    Vector(Lugar, 7) = ""
                    Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    Vector(Lugar, 10) = WHoja
                    Vector(Lugar, 11) = Pusing("###,###.##", Str$(WSaldo))
                    
                    WXEntradas = WXEntradas + WCantidad
                    
                End If
                
                Rem End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spMovvar = "ListaMovvarTerminadoDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstMovvar!Fecha, 4) + Mid$(rstMovvar!Fecha, 4, 2) + Left$(rstMovvar!Fecha, 2)
                If XFec < "20001218" Or XFec > WOrdFechaCierre Then

                
                        Else
                
                If rstMovvar!Tipo = "T" Then
                
                    WTerminado = rstMovvar!Terminado
                    WCantidad = rstMovvar!Cantidad
                    WFecha = rstMovvar!Fecha
                    WCodigo = rstMovvar!Codigo
                    WMovi = rstMovvar!Movi
                    WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                    
                    Lugar = Lugar + 1
                    
                    Vector(Lugar, 1) = WFecha
                    If Val(rstMovvar!Tipomov) = 0 Or Val(rstMovvar!Tipomov) = 1 Then
                        Vector(Lugar, 2) = "Mov.Var"
                            Else
                        Vector(Lugar, 2) = "Guia In"
                    End If
                    Vector(Lugar, 3) = WCodigo
                    Vector(Lugar, 4) = rstMovvar!Observaciones
                    If WMovi = "E" Then
                        Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                        Vector(Lugar, 6) = ""
                        WXEntradas = WXEntradas + WCantidad
                            Else
                        Vector(Lugar, 5) = ""
                        Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                        WXSalidas = WXSalidas + WCantidad
                    End If
                    Vector(Lugar, 7) = ""
                    Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    Vector(Lugar, 10) = WLote
                    Vector(Lugar, 11) = ""
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        rstMovvar.Close
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNO
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                If XFec < WOrdFechaCierre Then
                
                WMarcaAnt = IIf(IsNull(rstMovguia!MarcaAnt), "", rstMovguia!MarcaAnt)
                If WMarcaAnt = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    WTerminado = rstMovguia!Terminado
                    WCantidad = IIf(IsNull(rstMovguia!Cantidadant), "0", rstMovguia!Cantidadant)
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov
                    WLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    WPartida = IIf(IsNull(rstMovguia!Partida), "", rstMovguia!Partida)
                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(WSaldo)

                    Lugar = Lugar + 1
                    
                    Vector(Lugar, 1) = WFecha
                    If Val(WCodigo) > 900000 Then
                        Vector(Lugar, 2) = "Prestamo"
                        Vector(Lugar, 3) = WCodigo - 900000
                            Else
                        Vector(Lugar, 2) = "Guia In"
                        Vector(Lugar, 3) = WCodigo
                    End If
                    Rem Vector(Lugar, 4) = rstMovguia!Observaciones
                    
                    If WMovi = "E" Then
                    
                        Select Case WTipomov
                            Case 1
                                Vector(Lugar, 4) = "Recepcion de Surfactan"
                            Case 2
                                Vector(Lugar, 4) = "Recepcion de Pellital"
                            Case 3
                                Vector(Lugar, 4) = "Recepcion de Surfactan II"
                            Case 4
                                Vector(Lugar, 4) = "Recepcion de Pellital II"
                            Case 5
                                Vector(Lugar, 4) = "Recepcion de Surfactan III"
                            Case 6
                                Vector(Lugar, 4) = "Recepcion de Surfactan IV"
                            Case 7
                                Vector(Lugar, 4) = "Recepcion de Surfactan V"
                            Case 8
                                Vector(Lugar, 4) = "Recepcion de Pellital V"
                            Case 9
                                Vector(Lugar, 4) = "Recepcion de Pellital IV"
                            Case 10
                                Vector(Lugar, 4) = "Recepcion de Surfactan VI"
                            Case 11
                                Vector(Lugar, 4) = "Recepcion de Surfactan VII"
                            Case Else
                        End Select
                        Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                        Vector(Lugar, 6) = ""
                        Vector(Lugar, 10) = WLote
                        Vector(Lugar, 11) = Pusing("###,###.##", Str$(WSaldo))
                        WXEntradas = WXEntradas + WCantidad
                        
                            Else
                            
                        Select Case WDestino
                            Case 1
                                Vector(Lugar, 4) = "Envio a Surfactan"
                            Case 2
                                Vector(Lugar, 4) = "Envio a Pellital"
                            Case 3
                                Vector(Lugar, 4) = "Envio a Surfactan II"
                            Case 4
                                Vector(Lugar, 4) = "Envio a Pellital II"
                            Case 5
                                Vector(Lugar, 4) = "Envio a Surfactan III"
                            Case 6
                                Vector(Lugar, 4) = "Envio a Surfactan IV"
                            Case 7
                                Vector(Lugar, 4) = "Envio a Surfactan V"
                            Case 8
                                Vector(Lugar, 4) = "Envio a Pellital V"
                            Case 9
                                Vector(Lugar, 4) = "Envio a Pellital IV"
                            Case 10
                                Vector(Lugar, 4) = "Envio a Surfactan VI"
                            Case 11
                                Vector(Lugar, 4) = "Envio a Surfactan VII"
                            Case Else
                        End Select
                        Vector(Lugar, 5) = ""
                        Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                        Vector(Lugar, 10) = WPartida
                        Vector(Lugar, 11) = ""
                        WXSalidas = WXSalidas + WCantidad
                        
                    End If
                    Vector(Lugar, 7) = ""
                    Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                
                End If
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        rstMovguia.Close
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla
    
    
    
    Rem PROCESA LOS MOVIMIENTOS DE LABORATORIO
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spMovlab = "ListaMovlabTerminadoDesdeHasta" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstMovlab!Fecha, 4) + Mid$(rstMovlab!Fecha, 4, 2) + Left$(rstMovlab!Fecha, 2)
                If XFec < "20001218" Or XFec > WOrdFechaCierre Then
                
                    Else
                
                If rstMovlab!Tipo = "T" Then
                
                    WTerminado = rstMovlab!Terminado
                    WCantidad = rstMovlab!Cantidad
                    WFecha = rstMovlab!Fecha
                    WCodigo = rstMovlab!Codigo
                    WMovi = rstMovlab!Movi
                    WLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)

                    Lugar = Lugar + 1
                    
                    Vector(Lugar, 1) = WFecha
                    Vector(Lugar, 2) = "Mov.Lab"
                    Vector(Lugar, 3) = WCodigo
                    Vector(Lugar, 4) = rstMovlab!Observaciones
                    
                    If WMovi = "E" Then
                        Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                        Vector(Lugar, 6) = ""
                        WXEntradas = WXEntradas + WCantidad
                                Else
                        Vector(Lugar, 5) = ""
                        Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                        WXSalidas = WXSalidas + WCantidad
                    End If
                    Vector(Lugar, 7) = ""
                    Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    Vector(Lugar, 10) = WLote
                    Vector(Lugar, 11) = ""
                    
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        
        rstMovlab.Close
        
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla
    
    
    Rem REMITOS EN CONSIGNACION
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spConsig = "ListaConsigTerminado" + XParam
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount > 0 Then
    
        With rstConsig
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WTerminado = rstConsig!Terminado
                WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                WFecha = rstConsig!Fecha
                WCodigo = rstConsig!Numero
                WLote = IIf(IsNull(rstConsig!Lote), "", rstConsig!Lote)
                    
                If WCantidad <> 0 Then
                
                    Lugar = Lugar + 1
                    
                    Vector(Lugar, 1) = WFecha
                    Vector(Lugar, 2) = "Rem.Con."
                    Vector(Lugar, 3) = WCodigo
                    Vector(Lugar, 4) = rstConsig!Observaciones
                    Vector(Lugar, 5) = ""
                    Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                    WXSalidas = WXSalidas + WCantidad
                    Vector(Lugar, 7) = ""
                    Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    Vector(Lugar, 9) = rstConsig!Cliente
                    Vector(Lugar, 10) = WLote
                    Vector(Lugar, 11) = ""
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstConsig.Close
        
    End If
    
    For Ciclo = 1 To Lugar
    
        WImpre1 = Vector(Ciclo, 9)
        
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spCliente = "ConsultaCliente" + "'" + WImpre1 + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WImpre2 = rstCliente!Razon
            rstCliente.Close
                Else
            WImpre2 = ""
        End If
        
        Call Conecta_Empresa
    
        Vector(Ciclo, 4) = WImpre2
    
    Next Ciclo
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla
    
        
    Rem PROCESA LOS las devoluciones de mercaderia
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstEntdev!Fecha, 4) + Mid$(rstEntdev!Fecha, 4, 2) + Left$(rstEntdev!Fecha, 2)
                If XFec < "20001218" Or XFec > WOrdFechaCierre Then
                
                        Else
                
                WTerminado = rstEntdev!Terminado
                WCantidad = rstEntdev!Cantidad
                WFecha = rstEntdev!Fecha
                WCodigo = rstEntdev!Codigo
                WLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                WSaldo = rstEntdev!Saldo
                    
                Lugar = Lugar + 1
                    
                Vector(Lugar, 1) = WFecha
                Vector(Lugar, 2) = "Ent.Dev"
                Vector(Lugar, 3) = WCodigo
                Vector(Lugar, 4) = rstEntdev!Observaciones
                Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                Vector(Lugar, 6) = ""
                WXEntradas = WXEntradas + WSaldo
                Vector(Lugar, 7) = ""
                Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                Vector(Lugar, 10) = WLote
                Vector(Lugar, 11) = Pusing("###,###.##", Str$(WSaldo))
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstEntdev.Close
        
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                Vector(Ciclo, 11) = Vector(dada, 11)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10
                Vector(dada, 11) = Auxi11

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
                
        DBGrid1.Col = 0
        DBGrid1.Text = Vector(Cicla, 1)
                        
        DBGrid1.Col = 1
        DBGrid1.Text = Vector(Cicla, 2)
                                               
        DBGrid1.Col = 2
        DBGrid1.Text = Vector(Cicla, 3)
                        
        DBGrid1.Col = 3
        DBGrid1.Text = Vector(Cicla, 4)
                        
        DBGrid1.Col = 4
        DBGrid1.Text = Vector(Cicla, 5)
                
        DBGrid1.Col = 5
        DBGrid1.Text = Vector(Cicla, 6)
        
        DBGrid1.Col = 6
        DBGrid1.Text = Vector(Cicla, 7)
        
        DBGrid1.Col = 7
        DBGrid1.Text = Vector(Cicla, 10)
        
        DBGrid1.Col = 8
        DBGrid1.Text = Vector(Cicla, 11)
    
    Next Cicla

    WXStock = WXInicial + WXEntradas - WXSalidas
    
    XInicial.Text = Pusing("###,###.##", Str$(WXInicial))
    XEntradas.Text = Pusing("###,###.##", Str$(WXEntradas))
    XSalidas.Text = Pusing("###,###.##", Str$(WXSalidas))
    XStock.Text = Pusing("###,###.##", Str$(WXStock))
    
    spTerminado = "ConsultaTerminado " + "'" + "NK" + Right$(Terminado.Text, 10) + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Nk.Text = Pusing("###,###.##", rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
        rstTerminado.Close
            Else
        Nk.Text = "0.00"
    End If
    
    spTerminado = "ConsultaTerminado " + "'" + "RE" + Right$(Terminado.Text, 10) + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        RE.Text = Pusing("###,###.##", rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
        rstTerminado.Close
            Else
        RE.Text = "0.00"
    End If

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0

End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        WTerminado = Terminado.Text
        Terminado.Text = WTerminado
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = rstTerminado!Descripcion
            Unidad.Text = rstTerminado!Unidad
            Deposito.Text = rstTerminado!Deposito
            StkProceso.Text = Pusing("###,###.##", Str$(rstTerminado!Proceso))
            rstTerminado.Close
            Call Proceso_Click
            Rem DBGrid1.FirstRow = 0
            Rem DBGrid1.Col = 0
            Rem DBGrid1.Row = 0
            DBGrid1.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
End Sub




