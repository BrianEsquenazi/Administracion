VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRec 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Recibos"
   ClientHeight    =   8250
   ClientLeft      =   690
   ClientTop       =   420
   ClientWidth     =   10665
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   10665
   Begin VB.TextBox RetSuss 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9240
      MaxLength       =   15
      TabIndex        =   23
      Text            =   " "
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cierre de Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   22
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox RetOtra 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6480
      MaxLength       =   15
      TabIndex        =   5
      Text            =   " "
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox RetIva 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      MaxLength       =   15
      TabIndex        =   7
      Text            =   " "
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Retganancias 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   4
      Text            =   " "
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Recibos"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   5295
      Begin VB.OptionButton Tipo1 
         Caption         =   "Cobro de Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Anticipos"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.TextBox Clientes 
      Height          =   285
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   735
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
      Height          =   5415
      Left            =   0
      OleObjectBlob   =   "reci.frx":0000
      TabIndex        =   6
      Top             =   2280
      Width           =   9975
   End
   Begin VB.TextBox Recibo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Ret. Suss"
      Height          =   255
      Left            =   7920
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Debitos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.   2) Ch.   3) Doc."
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Ret.I.B."
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Ret.Iva"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Rte.Ganancias"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label DesClientes 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod. Cilente"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero de Recibo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "PrgRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 10 ' Número máximo de campos del conjunto de registros.
Private Auxi As String
Private Auxi1 As String
Private WSaldo As Double
Private Vector(20, 10) As String
Private Provincia(100) As String
Private m(20) As String
Private Impre1(100) As Single
Private Impre2(100) As Single
Private ImpreTipo(100) As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WPostal As String
Private WProvincia As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim XParam As String
Dim Auxiliar(100, 2) As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Suma_Datos()
    Debitos.Caption = ""
    Creditos.Caption = ""
    
    Creditos.Caption = Str$(Val(Retganancias.Text) + Val(RetIva.Text) + Val(RetOtra.Text) + Val(RetSuss.Text))
    For iRow = 0 To 19
        DBGrid1.Col = 4
        DBGrid1.Row = iRow
        If Val(DBGrid1.Text) <> 0 Then
            Debitos.Caption = Str$(Val(Debitos.Caption) + Val(DBGrid1.Text))
        End If
        DBGrid1.Col = 9
        DBGrid1.Row = iRow
        If Val(DBGrid1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(DBGrid1.Text))
        End If
    Next iRow
    Debitos.Caption = Pusing("###,###.##", Debitos.Caption)
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
    DBGrid1.Col = 0
    DBGrid1.Row = 0
End Sub

Private Sub Lee_Datos()
    For iRow = 0 To 19
        For iCol = 0 To 9
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow

    Erase Auxiliar
    WAA = 0
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        With rstRecibos
        
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            ClaveRecibo = Recibo.Text + Auxi1
        
            spRecibos = "ConsultaRecibosClave " + "'" + ClaveRecibo + "'"
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                Select Case Val(rstRecibos!Tiporeg)
                    Case 1
                        Debito = Debito + 1
                        DBGrid1.Row = Debito - 1
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstRecibos!Tipo1
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstRecibos!Letra1
                        DBGrid1.Col = 2
                        DBGrid1.Text = rstRecibos!Punto1
                        DBGrid1.Col = 3
                        DBGrid1.Text = rstRecibos!Numero1
                        DBGrid1.Col = 4
                        DBGrid1.Text = rstRecibos!Importe1
                        DBGrid1.Text = Alinea("###,###.##", DBGrid1.Text)
                        WAA = WAA + 1
                        Auxiliar(WAA, 1) = rstRecibos!Tipo1
                        Auxiliar(WAA, 2) = rstRecibos!Numero1
                    Case 2
                        Credito = Credito + 1
                        DBGrid1.Row = Credito - 1
                        DBGrid1.Col = 5
                        DBGrid1.Text = rstRecibos!Tipo2
                        DBGrid1.Col = 6
                        DBGrid1.Text = rstRecibos!Numero2
                        DBGrid1.Col = 7
                        DBGrid1.Text = rstRecibos!Fecha2
                        DBGrid1.Col = 8
                        DBGrid1.Text = rstRecibos!Banco2
                        DBGrid1.Col = 9
                        DBGrid1.Text = rstRecibos!Importe2
                        DBGrid1.Text = Alinea("###,###.##", DBGrid1.Text)
                    Case Else
                End Select
                rstRecibos.Close
                    Else
                Exit Do
            End If
        End With
    Loop
    
    For XAA = 1 To WAA
    
        WTipo = Auxiliar(XAA, 1)
        WNumero = Auxiliar(XAA, 2)
        
        With rstCtacte
            ClaveCtacte = WTipo + WNumero + "01"
            spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
                DBGrid1.Row = XAA - 1
                DBGrid1.Col = 2
                DBGrid1.Text = rstCtacte!Fecha
                rstCtacte.Close
            End If
        End With
    Next XAA
    
End Sub
Sub Verifica_datos()
    If Val(Retganancias.Text) = 0 Then
        Retganancias.Text = "0"
    End If
    If Val(RetIva.Text) = 0 Then
        RetIva.Text = "0"
    End If
    If Val(RetOtra.Text) = 0 Then
        RetOtra.Text = "0"
    End If
End Sub
Sub Format_datos()
    Retganancias.Text = Pusing("###,###.##", Retganancias.Text)
    RetIva.Text = Pusing("###,###.##", RetIva.Text)
    RetOtra.Text = Pusing("###,###.##", RetOtra.Text)
    RetSuss.Text = Pusing("###,###.##", RetSuss.Text)
End Sub

Sub Imprime_Datos()
    spClientes = "ConsultaClientes " + "'" + Clientes.Text + "'"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        Clientes.Text = rstClientes!Cliente
        DesClientes.Caption = rstClientes!Razon
        WRazon = rstClientes!Razon
        WDireccion = rstClientes!Direccion
        WLocalidad = rstClientes!Localidad
        WPostal = rstClientes!Postal
        WProvincia = Provincia(rstClientes!Provincia)
        WProv = rstClientes!Provincia
        rstClientes.Close
        Call Format_datos
    End If
End Sub


Private Sub CmdLimpiar_Click()
    For iRow = 0 To 19
        For iCol = 0 To 9
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    
End Sub

Private Sub Cerrar_Click()
    Call CmdLimpiar_Click
    Recibo.SetFocus
    PrgRec.Hide
    Unload Me
    PrgCtaCte2.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Command1_Click()
    Call Lee_Datos
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        With rstRecibos
            Existe = "N"
            ClaveRecibo = Recibo.Text + "01"
            spRecibos = "ConsultaRecibos " + "'" + ClaveRecibo + "'"
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                Existe = "S"
                Clientes.Text = rstRecibos!Cliente
                Observaciones.Text = rstRecibos!Observaciones
                Fecha.Text = rstRecibos!Fecha
                Retganancias.Text = rstRecibos!Retganancias
                RetIva.Text = rstRecibos!RetIva
                RetOtra.Text = rstRecibos!RetOtra
                RetSuss.Text = rstRecibos!RetSuss
                Tipo1.Value = True
                Tipo2.Value = False
                Select Case Val(rstRecibos!TipoRec)
                    Case 1
                        Tipo1.Value = True
                    Case 2
                        Tipo2.Value = True
                    Case Else
                End Select
                rstRecibos.Close
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


Private Sub Recibo_GotFocus()

        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        With rstRecibos
            Existe = "N"
            ClaveRecibo = Recibo.Text + "01"
            spRecibos = "ConsultaRecibos " + "'" + ClaveRecibo + "'"
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                Existe = "S"
                Clientes.Text = rstRecibos!Cliente
                Observaciones.Text = rstRecibos!Observaciones
                Fecha.Text = rstRecibos!Fecha
                Retganancias.Text = rstRecibos!Retganancias
                RetIva.Text = rstRecibos!RetIva
                RetOtra.Text = rstRecibos!RetOtra
                RetSuss.Text = rstRecibos!RetSuss
                Tipo1.Value = True
                Tipo2.Value = False
                Select Case Val(rstRecibos!TipoRec)
                    Case 1
                        Tipo1.Value = True
                    Case 2
                        Tipo2.Value = True
                    Case Else
                End Select
                rstRecibos.Close
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

ReDim UserData(0 To 9, 0 To 19)

mTotalRows& = 20

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
For i = 0 To 9
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Letra"
             DBGrid1.Columns(newcnt).Width = 10
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Fecha"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Numero"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Importe"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Numero"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 7
             DBGrid1.Columns(newcnt).Caption = "Fecha"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 8
             DBGrid1.Columns(newcnt).Caption = "Banco"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
         Case 9
             DBGrid1.Columns(newcnt).Caption = "Importe"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
 
    Provincia$(0) = "Capital Federal"
    Provincia$(1) = "Buenos Aires"
    Provincia$(2) = "Catamarca"
    Provincia$(3) = "Cordoba"
    Provincia$(4) = "Corrientes"
    Provincia$(5) = "Chaco"
    Provincia$(6) = "Chubut"
    Provincia$(7) = "Entre Rios"
    Provincia$(8) = "Formosa"
    Provincia$(9) = "Jujuy"
    Provincia$(10) = "La Pampa"
    Provincia$(11) = "La Rioja"
    Provincia$(12) = "Mendoza"
    Provincia$(13) = "Misiones"
    Provincia$(14) = "Neuquen"
    Provincia$(15) = "Rio Negro"
    Provincia$(16) = "Salta"
    Provincia$(17) = "San Juan"
    Provincia$(18) = "San Luis"
    Provincia$(19) = "Santa Cruz"
    Provincia$(20) = "Santa Fe"
    Provincia$(21) = "Santiago del Estero"
    Provincia$(22) = "Tucuman"
    Provincia$(23) = "Tierra del Fuego"
    Provincia$(24) = "Exterior"
    Provincia$(25) = ""
     
    ImpreTipo$(1) = "FC"
     
    Tipo1.Value = True
    Tipo2.Value = False
    
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"

    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    RetSuss.Text = "0"
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Observaciones.Text = ""
    
    Recibo.Text = WRecibo
    Recibo.SetFocus

End Sub


