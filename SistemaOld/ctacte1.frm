VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCtaCte1 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Clientes"
   ClientHeight    =   7320
   ClientLeft      =   570
   ClientTop       =   1155
   ClientWidth     =   10995
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   10995
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3480
      TabIndex        =   23
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   3255
   End
   Begin VB.CommandButton datoscli 
      Caption         =   "Datos del Cliente"
      Height          =   735
      Left            =   9960
      TabIndex        =   21
      Top             =   720
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.CommandButton Listar 
      Caption         =   "Listar"
      Height          =   300
      Left            =   3480
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Datos"
      Height          =   1335
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Todos 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Pendiente 
         Caption         =   "Pendiente"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Listado"
      Height          =   1335
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Total 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Documentos 
         Caption         =   "Documentos"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton CtaCte 
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   1335
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Dolares 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Pesos 
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Cliente 
      Height          =   375
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   6
      Text            =   " "
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      Height          =   300
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "ctacte1.frx":0000
      TabIndex        =   3
      Top             =   2280
      Width           =   10815
   End
   Begin VB.ListBox Pantalla 
      Height          =   1230
      ItemData        =   "ctacte1.frx":09CE
      Left            =   120
      List            =   "ctacte1.frx":09D5
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCtaCte1"
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
Private Importe1 As Double
Private Importe2 As Double
Private Importe3 As Double
Private WTipo As Integer
Private WSaldo As Double
Private Acumula As Double
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstCtacte As Recordset
Dim spCtecte As String
Dim XParam As String
Private WNume As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgCtaCte1.Hide
    Unload Me
    Menu.Show
End Sub



Private Sub Command1_Click()

    Sql1 = "UPDATE CtaCte SET "
    Sql2 = " Saldo  = 0"
    Sql3 = " SaldoUs = 0"
    Sql4 = " Where Cliente = " + "'" + Cliente.Text + "'"
                     
    spCtacte = Sql1 + Sql2 + Sql3 + Sql4
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCliente!Cliente + "     " + rstCliente!Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCliente!Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
        End With
        rstCliente.Close
    End If
            
    Rem Pantalla.Visible = True

End Sub

Private Sub datoscli_Click()
    PCliente = Cliente.Text
    prgcli.Show
End Sub

Private Sub pantalla_Click()
    Rem Pantalla.Visible = False
       
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    Cliente.Text = Claveven$
    
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Cliente.Text = rstCliente!Cliente
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
        Call Proceso_Click
            Else
        Cliente.Text = Claveven$
    End If
    Cliente.SetFocus
    
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
ReDim UserData(0 To 7, 0 To 2000)

mTotalRows& = 2000

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

    Call Limpia_Vector
 
    Cliente.Text = ""
    DesCliente.Caption = ""

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Pesos.Value = True
    CtaCte.Value = True
    Pendiente.Value = True
    
    Call Consulta_Click
    
    Cliente.Text = PCliente
    
    Cliente.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Cliente.Text = UCase(Cliente.Text)
    
    WSalida = "N"
        
    For a = 0 To 100
    
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    
    For iRow = 0 To 9
        For iCol = 0 To 7
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            If iCol = 0 Then
                Auxi = DBGrid1.Text
                If Auxi = "" Then
                    WSalida = "S"
                    Exit For
                End If
            End If
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    
    If WSalida = "S" Then Exit For
    
    Next a

    DBGrid1.FirstRow = 0
    Renglon = 0
    WSaldo = 0
    
    XParam = "'" + Cliente.Text + "'"
    spCtacte = "ListaCtacteCliente " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
        With rstCtacte
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    WPasa = "N"
                    
                    If CtaCte.Value = True Then
                        If !Tipo < 50 Then
                            WPasa = "S"
                        End If
                    End If
                    
                    If Documentos.Value = True Then
                        If !Tipo >= 50 Then
                            WPasa = "S"
                        End If
                    End If
                    
                    If Total.Value = True Then
                        WPasa = "S"
                    End If
                        
                    If WPasa = "S" Then
                        If Pesos.Value = True Then
                            If !Total > 0 Then
                                Importe1 = !Total
                                Importe2 = 0
                                    Else
                                Importe1 = 0
                                Importe2 = !Total
                            End If
                            Importe3 = !Saldo
                                Else
                            If !Totalus > 0 Then
                                Importe1 = !Totalus
                                Importe2 = 0
                                    Else
                                Importe1 = 0
                                Importe2 = !Totalus
                            End If
                            Importe3 = !Saldous
                        End If
                        
                        Call Redondeo(Importe3)
                    
                        If Importe3 <> 0 Or Todos.Value = True Then
                    
                            Renglon = Renglon + 1
                
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                    
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                    
                            Rem DBGrid1.Col = 0
                            Rem DBGrid1.Text = !Tipo
                            
                            Select Case !Tipo
                                Case 1
                                    DBGrid1.Col = 0
                                    DBGrid1.Text = "Fac"
                                Case 2
                                    DBGrid1.Col = 0
                                    DBGrid1.Text = "Dev"
                                Case 3
                                    DBGrid1.Col = 0
                                    DBGrid1.Text = "Fac"
                                Case 4
                                    DBGrid1.Col = 0
                                    Select Case Left$(!Impre, 2)
                                        Case "DC"
                                            DBGrid1.Text = "D.C"
                                        Case "CH"
                                            DBGrid1.Text = "CHR"
                                        Case Else
                                            DBGrid1.Text = "N/D"
                                    End Select
                                Case 5
                                    DBGrid1.Col = 0
                                    Select Case Left$(!Impre, 2)
                                        Case "DC"
                                            DBGrid1.Text = "D.C"
                                        Case "CH"
                                            DBGrid1.Text = "CHR"
                                        Case Else
                                            DBGrid1.Text = "N/C"
                                    End Select
                                Case 6
                                    DBGrid1.Col = 0
                                    DBGrid1.Text = "Rec"
                                Case 7
                                    DBGrid1.Col = 0
                                    DBGrid1.Text = "Ant"
                                Case 10
                                    DBGrid1.Col = 0
                                    DBGrid1.Text = "FCR"
                                Case 50
                                    DBGrid1.Col = 0
                                    DBGrid1.Text = "Doc"
                                Case Else
                            End Select
                            
                            DBGrid1.Col = 1
                            DBGrid1.Text = Pusing("######", Str$(!Numero))
                    
                            DBGrid1.Col = 2
                            DBGrid1.Text = !Fecha
                    
                            If Importe1 <> 0 Then
                                DBGrid1.Col = 3
                                DBGrid1.Text = Pusing("###,###,###.##", Str$(Importe1))
                                    Else
                                DBGrid1.Col = 3
                                DBGrid1.Text = ""
                            End If
                    
                            If Importe2 <> 0 Then
                                DBGrid1.Col = 4
                                DBGrid1.Text = Pusing("###,###,###.##", Str$(Importe2))
                                    Else
                                DBGrid1.Col = 4
                                DBGrid1.Text = ""
                            End If
                    
                            If Importe3 <> 0 Then
                                DBGrid1.Col = 5
                                DBGrid1.Text = Pusing("###,###,###.##", Str$(Importe3))
                                    Else
                                DBGrid1.Col = 5
                                DBGrid1.Text = ""
                            End If
                            
                            WSaldo = WSaldo + Importe3
                    
                            DBGrid1.Col = 6
                            DBGrid1.Text = !Vencimiento
                            
                            DBGrid1.Col = 7
                            DBGrid1.Text = !Vencimiento1
                        
                        End If
                        
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        rstCtacte.Close
    
    End If
    
    Renglon = Renglon + 1
    
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Saldo.Caption = Pusing("###,###,###.##", Str$(WSaldo))
    
    DBGrid1.Refresh
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        Cliente.Text = WCliente
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            rstCliente.Close
            Call Proceso_Click
            Rem DBGrid1.FirstRow = 0
            Rem DBGrid1.Col = 0
            Rem DBGrid1.Row = 0
            DBGrid1.SetFocus
                Else
            Cliente.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Vector()
' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
     DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 7
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Numero"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Fecha"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Debito"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Credito"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Saldo"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Vencimiento"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 7
             DBGrid1.Columns(newcnt).Caption = "Vencimiento"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i

End Sub

Private Sub Listar_Click()

    WDesde = UCase(Cliente.Text)
    WHasta = UCase(Cliente.Text)
    
    spCtacte = "ModificaCtacteTipo1"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    spCtacte = "ModificaCtacteTipo2"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    spCtacte = "ModificaCtacteImporte0"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

    If CtaCte.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte1 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte2 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte3 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte4 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Documentos.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte5 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte6 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte7 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte8 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Total.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte9 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte10 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte11 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte12 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
    
    DA = ""
    With rstImpCtaCte
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    XParam = "'" + WDesde + "','" _
            + WHasta + "'"
    spCtacte = "ListaCtacteDesdeHasta " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then

        With rstCtacte
            .MoveFirst
            Do
            
                WPasa = "N"
                If CtaCte.Value = True Then
                    If !Tipo < 50 Then
                        WPasa = "S"
                    End If
                End If
                
                If Documentos.Value = True Then
                    If !Tipo >= 50 Then
                        WPasa = "S"
                    End If
                End If
                
                If Total.Value = True Then
                    WPasa = "S"
                End If
                    
                If WPasa = "S" Then
            
                If Todos.Value = True Or !Importe3 <> 0 Then
            
                WTipo = !Tipo
                WImpre = !Impre
                WNumero = !Numero
                WRenglon = !Renglon
                WCliente = !Cliente
                WFecha = !Fecha
                WEstado = !Estado
                Wvencimiento = !Vencimiento
                WVencimiento1 = !Vencimiento1
                WTotal = !Total
                WTotalUs = !Totalus
                WSaldo = !Saldo
                WSaldoUs = !Saldous
                WNeto = !Neto
                WIva1 = !Iva1
                WWIva2 = !Iva2
                WOrdFecha = !OrdFecha
                WOrdVencimiento = !OrdVencimiento
                WOrdVencimiento1 = !OrdVencimiento1
                WPedido = !Pedido
                WRemito = !Remito
                WOrden = !Orden
                WParidad = !Paridad
                WProvincia = !Provincia
                WVendedor = !vendedor
                WRubro = !Rubro
                WCcomprobante = !Comprobante
                WAceptada = !Aceptada
                WCosto = !Costo
                WImporte1 = !Importe1
                WImporte2 = !Importe2
                WImporte3 = !Importe3
                WImporte4 = !Importe4
                WImporte5 = !Importe5
                WImporte6 = !Importe6
                WImporte7 = !Importe7
                WClave = !Clave
                
                With rstImpCtaCte
        
                    .Index = "Clave"
                                            
                    .AddNew
                    
                    !Tipo = WTipo
                    !Impre = WImpre
                    !Numero = WNumero
                    !Renglon = WRenglon
                    !Cliente = WCliente
                    !Fecha = WFecha
                    !Estado = WEstado
                    !Vencimiento = Wvencimiento
                    !Vencimiento1 = WVencimiento1
                    !Total = WTotal
                    !Totalus = WTotalUs
                    !Saldo = WSaldo
                    !Saldous = WSaldoUs
                    !Neto = WNeto
                    !Iva1 = WIva1
                    !Iva2 = WIva2
                    !OrdFecha = WOrdFecha
                    !OrdVencimiento = WOrdVencimiento
                    !OrdVencimiento1 = WOrdVencimiento1
                    !Pedido = WPedido
                    !Remito = WRemito
                    !Orden = WOrden
                    !Paridad = WParidad
                    !Provincia = WProvincia
                    !vendedor = WVendedor
                    !Rubro = WRubro
                    !Comprobante = WComprobante
                    !Aceptada = WAceptada
                    !Costo = WCosto
                    !Importe1 = WImporte1
                    !Importe2 = WImporte2
                    !Importe3 = WImporte3
                    !Importe4 = WImporte4
                    !Importe5 = WImporte5
                    !Importe6 = WImporte6
                    !Importe7 = WImporte7
                    !Clave = WClave
                    WNume = Str$(!Numero)
                    Call Ceros(WNume, 8)
                    !ClaveImpre = !Cliente + !OrdFecha + !Tipo + WNume
                    !Empresa = Val(WEmpresa)
                    
                    .Update
                    
                End With
                
                End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstCtacte.Close
    
        With rstImpCtaCte
                .Index = "ClaveImpre"
                .MoveFirst
                Do
                
                    WRazon = ""
                    spCliente = "ConsultaCliente " + !Cliente
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WRazon = rstCliente!Razon
                    End If
                
                    If Pasa = 0 Then
                        Pasa = 1
                        Acumula = 0
                        corte = !Cliente
                    End If
                    If corte <> !Cliente Then
                        Acumula = 0
                        corte = !Cliente
                    End If
                    .Edit
                    Acumula = Acumula + !Importe3
                    Call Redondeo(Acumula)
                    !Importe4 = Acumula
                    !Razon = WRazon
                    .Update
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
        End With
    
    End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = ""
    
    If CtaCte.Value = True Then
        WTitulo = "Cuenta Corriente - "
    End If
    If Documentos.Value = True Then
        WTitulo = "Documentos - "
    End If
    If Total.Value = True Then
        WTitulo = "Total - "
    End If
    
    If Pesos.Value = True Then
        WTitulo = WTitulo + "Pesos"
    End If
    If Dolares.Value = True Then
        WTitulo = WTitulo + "Dolares"
    End If
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Cuenta Corriente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{impCtaCte.Cliente} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.ReportFileName = "wimpctacte.rpt"
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Rem Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Rem Listado.Connect = Connect()
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.Action = 1
    
End Sub


Private Sub DBGrid1_DblClick()

    DBGrid1.Col = 0
    Tipo = DBGrid1.Text
    
    If Tipo = "Rec" Then
        DBGrid1.Col = 1
        WRecibo = DBGrid1.Text
        PrgRec.Show
    End If

    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                
                    DA = Len(!Razon) - WEspacios
                    
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = !Cliente
                            IngresaItem = Auxi + "    " + !Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Cliente
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next aa
                    .MoveNext
                        
                            Else
                            
                    Exit Do
                    
                End If
            Loop
        End With
        rstCliente.Close
    
    End If
    
    End If

End Sub



Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_ImpCtacte
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

