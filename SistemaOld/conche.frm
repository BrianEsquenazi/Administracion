VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrgConche 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cehques"
   ClientHeight    =   7890
   ClientLeft      =   15
   ClientTop       =   480
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   7890
   ScaleWidth      =   11880
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   7215
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12726
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.ComboBox Tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Cheque 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro de Cheque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "PrgConche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Auxi As String
Private dada As String
Private WCheque As String
Private WCompara As String
Private WPasa  As String
Private WProveedor As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim rstDepositos As Recordset
Dim spDepositos As String
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
Dim Auxiliar(10000, 3) As String

Private Sub cmdClose_Click()
    PrgConche.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Tipo.Clear
    
    Tipo.AddItem "Cheque Tercero"
    Tipo.AddItem "Cheque Propio"
    
    Tipo.ListIndex = 0
 
    Cheque.Text = ""

    Muestra.TopRow = 1
    Muestra.Col = 1
    Muestra.Row = 1
    
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError

    WCheque = Cheque.Text
    Canti = Len(WCheque)
    Salida = ""
    
    Call Limpia_Vector

    Renglon = 0
    Erase Auxiliar
    
    If Tipo.ListIndex = 0 Then
    
    spRecibos = "ListaRecibosBusqueda"
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    With rstRecibos
        .MoveFirst
        If .NoMatch = False Then
            Do
                WPasa = "N"
                WChequeo = "S"
                WCompara = IIf(IsNull(rstRecibos!Numero2), "", rstRecibos!Numero2)
                Call Ceros(WCompara, 10)
                If WCheque = Right$(WCompara, Canti) Then
                    WPasa = "S"
                End If
                If WPasa = "S" And WChequeo = "S" Then
                    Renglon = Renglon + 1
            
                    Muestra.TextMatrix(Renglon, 1) = IIf(IsNull(rstRecibos!Numero2), "", rstRecibos!Numero2)
                    Muestra.TextMatrix(Renglon, 2) = IIf(IsNull(rstRecibos!Banco2), "", rstRecibos!Banco2)
                    WImporte = IIf(IsNull(rstRecibos!Importe2), "", rstRecibos!Importe2)
                    Muestra.TextMatrix(Renglon, 3) = Mascara("###,###.##", Str$(Abs(WImporte)))
                    Muestra.TextMatrix(Renglon, 4) = IIf(IsNull(rstRecibos!Fecha), "", rstRecibos!Fecha)
                    Muestra.TextMatrix(Renglon, 5) = IIf(IsNull(rstRecibos!Fecha2), "", rstRecibos!Fecha2)
                    Muestra.TextMatrix(Renglon, 6) = "Rec:" + IIf(IsNull(rstRecibos!Recibo), "", rstRecibos!Recibo)
                    
                    Auxiliar(Renglon, 1) = "1"
                    Auxiliar(Renglon, 2) = IIf(IsNull(rstRecibos!Cliente), "", rstRecibos!Cliente)
                    Auxiliar(Renglon, 3) = ""
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
    End With
    rstRecibos.Close
    End If
    
    
    spDepositos = "ListaDepositosConsulta"
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
    With rstDepositos
        .MoveFirst
        If .NoMatch = False Then
            Do
                WPasa = "N"
                WChequeo = "S"
                WCompara = IIf(IsNull(rstDepositos!Numero2), "", rstDepositos!Numero2)
                Call Ceros(WCompara, 10)
                If WCheque = Right$(WCompara, Canti) Then
                    WPasa = "S"
                End If
                If WPasa = "S" And WChequeo = "S" Then
                    Renglon = Renglon + 1
            
                    Muestra.TextMatrix(Renglon, 1) = IIf(IsNull(rstDepositos!Numero2), "", rstDepositos!Numero2)
                    Muestra.TextMatrix(Renglon, 2) = IIf(IsNull(rstDepositos!Observaciones2), "", rstDepositos!Observaciones2)
                    WImporte = IIf(IsNull(rstDepositos!Importe2), "", rstDepositos!Importe2)
                    Muestra.TextMatrix(Renglon, 3) = Mascara("###,###.##", Str$(Abs(WImporte)))
                    Muestra.TextMatrix(Renglon, 4) = IIf(IsNull(rstDepositos!Fecha), "", rstDepositos!Fecha)
                    Muestra.TextMatrix(Renglon, 5) = IIf(IsNull(rstDepositos!Fecha2), "", rstDepositos!Fecha2)
                    Muestra.TextMatrix(Renglon, 6) = "Dep:" + IIf(IsNull(rstDepositos!Deposito), "", rstDepositos!Deposito)
                    
                    Auxiliar(Renglon, 1) = "2"
                    Auxiliar(Renglon, 2) = IIf(IsNull(rstDepositos!Banco), "", rstDepositos!Banco)
                    Auxiliar(Renglon, 3) = ""
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        rstDepositos.Close
    End With
    End If
    
    spPagos = "ListaPagosConsulta"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
    With rstPagos
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                WPasa = "N"
                WChequeo = "S"
                WCompara = IIf(IsNull(rstPagos!Numero2), "", rstPagos!Numero2)
                Call Ceros(WCompara, 10)
                If WCheque = Right$(WCompara, Canti) Then
                    WPasa = "S"
                End If
                If WPasa = "S" And WChequeo = "S" Then
                    Renglon = Renglon + 1
            
                    Muestra.TextMatrix(Renglon, 1) = IIf(IsNull(rstPagos!Numero2), "", rstPagos!Numero2)
                    Muestra.TextMatrix(Renglon, 2) = IIf(IsNull(rstPagos!Observaciones2), "", rstPagos!Observaciones2)
                    WImporte = IIf(IsNull(rstPagos!Importe2), "", rstPagos!Importe2)
                    Muestra.TextMatrix(Renglon, 3) = Mascara("###,###.##", Str$(Abs(WImporte)))
                    Muestra.TextMatrix(Renglon, 4) = IIf(IsNull(rstPagos!Fecha), "", rstPagos!Fecha)
                    Muestra.TextMatrix(Renglon, 5) = IIf(IsNull(rstPagos!Fecha2), "", rstPagos!Fecha2)
                    Muestra.TextMatrix(Renglon, 6) = "O.P.:" + IIf(IsNull(rstPagos!Orden), "", rstPagos!Orden)
                    
                    Auxiliar(Renglon, 1) = "3"
                    Auxiliar(Renglon, 2) = IIf(IsNull(rstPagos!Proveedor), "", rstPagos!Proveedor)
                    Auxiliar(Renglon, 3) = IIf(IsNull(rstPagos!Observaciones), "", rstPagos!Observaciones)
                    
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        rstPagos.Close
    End With
    End If
    
    End If
    
    If Tipo.ListIndex = 1 Then
    
    spPagos = "ListaPagosConsultaII"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
    With rstPagos
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                WPasa = "N"
                WChequeo = "S"
                WCompara = IIf(IsNull(rstPagos!Numero2), "", rstPagos!Numero2)
                Call Ceros(WCompara, 10)
                If WCheque = Right$(WCompara, Canti) Then
                    WPasa = "S"
                End If
                If WPasa = "S" And WChequeo = "S" Then
                    Renglon = Renglon + 1
            
                    Muestra.TextMatrix(Renglon, 1) = IIf(IsNull(rstPagos!Numero2), "", rstPagos!Numero2)
                    Muestra.TextMatrix(Renglon, 2) = IIf(IsNull(rstPagos!Observaciones2), "", rstPagos!Observaciones2)
                    WImporte = IIf(IsNull(rstPagos!Importe2), "", rstPagos!Importe2)
                    Muestra.TextMatrix(Renglon, 3) = Mascara("###,###.##", Str$(Abs(WImporte)))
                    Muestra.TextMatrix(Renglon, 4) = IIf(IsNull(rstPagos!Fecha), "", rstPagos!Fecha)
                    Muestra.TextMatrix(Renglon, 5) = IIf(IsNull(rstPagos!Fecha2), "", rstPagos!Fecha2)
                    Muestra.TextMatrix(Renglon, 6) = "O.P.:" + IIf(IsNull(rstPagos!Orden), "", rstPagos!Orden)
                    
                    Auxiliar(Renglon, 1) = "3"
                    Auxiliar(Renglon, 2) = IIf(IsNull(rstPagos!Proveedor), "", rstPagos!Proveedor)
                    Auxiliar(Renglon, 3) = IIf(IsNull(rstPagos!Observaciones), "", rstPagos!Observaciones)
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        rstPagos.Close
    End With
    End If
    
    End If
    
    
    WRenglon = Renglon
    Renglon = 0
    
    For da = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Select Case Auxiliar(Renglon, 1)
            Case "1"
                WCliente = Auxiliar(Renglon, 2)
                spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WDescri = rstCliente!Razon
                    rstCliente.Close
                            Else
                    WDescri = ""
                End If
    
                Muestra.TextMatrix(Renglon, 7) = WCliente + " " + WDescri
    
            Case "2"
                WBanco = Auxiliar(Renglon, 2)
                spBanco = "ConsultaBanco " + "'" + WBanco + "'"
                Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                If rstBanco.RecordCount > 0 Then
                    WDescri = rstBanco!Nombre
                    rstBanco.Close
                        Else
                    WDescri = ""
                End If
                
                Muestra.TextMatrix(Renglon, 7) = Str$(WBanco) + " " + WDescri
                    
            Case "3"
                WProveedor = Auxiliar(Renglon, 2)
                WObservaciones = Auxiliar(Renglon, 3)
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WDescri = WProveedor + " " + RstProveedor!Nombre
                    RstProveedor.Close
                            Else
                    WDescri = WObservaciones
                End If
                
                Muestra.TextMatrix(Renglon, 7) = WDescri
                
            Case Else
        End Select
    Next da
    
    Muestra.TopRow = 1
    Muestra.Row = 1
    Muestra.Col = 1
    
    Muestra.SetFocus
    
    Exit Sub
    
WError:

    WChequeo = "N"
    Resume Next

End Sub

Private Sub Cheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCheque = Cheque.Text
        Call Proceso_Click
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Limpia_Vector()

    Muestra.Clear

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Muestra.FixedCols = 1
    Muestra.Cols = 8
    Muestra.FixedRows = 1
    Muestra.Rows = 10000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Cheque"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Banco"
                Muestra.ColWidth(Ciclo) = 2000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Importe"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Fecha Comp."
                Muestra.ColWidth(Ciclo) = 1300
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Muestra.Text = "Fecha Cheque"
                Muestra.ColWidth(Ciclo) = 1300
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 6
                Muestra.Text = "Comprobante"
                Muestra.ColWidth(Ciclo) = 1200
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 7
                Muestra.Text = "Observaicones"
                Muestra.ColWidth(Ciclo) = 2600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
End Sub

