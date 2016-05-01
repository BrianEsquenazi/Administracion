VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModpedExp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Actualizacion de Pedidos de Exportacion"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "modpedexp.frx":0000
      Left            =   3360
      List            =   "modpedexp.frx":0007
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partida"
      Height          =   2655
      Left            =   5400
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox WCanti5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox WCanti4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox WLote5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox WLote4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox WCanti3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WCanti2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox WCanti1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Wlote3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox WLote2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox WLote1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Partida"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   10
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   10200
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   6120
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Cliente 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   5
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   500
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   500
      Left            =   9000
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "modpedexp.frx":0015
      TabIndex        =   25
      Top             =   1320
      Width           =   11415
   End
   Begin VB.Label Label11 
      Caption         =   "Pedido"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PrgModpedExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WAceptada As String
Private WDirentrega As String
Private WFecEntrega As String
Private WDespago As String
Private WObservaciones As String

Private Auxiliar(100, 14) As String
Private ClavePedido(100) As String
Private XLote(100, 12) As String

Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstPago As Recordset
Dim spPago As String

Dim XParam As String
Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim WSaldo4 As Double
Dim WSaldo5 As Double
Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim XSaldo4 As String
Dim XSaldo5 As String
Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer
Dim XCantidad1 As String
Dim xCantidad2 As String
Dim XLote1 As String
Dim XCantiLote1 As String
Dim XLote2 As String
Dim XCantiLote2 As String
Dim XLote3 As String
Dim XCantiLote3 As String
Dim XLote4 As String
Dim XCantiLote4 As String
Dim XLote5 As String
Dim XCantiLote5 As String
Dim WLugar As Integer

Private Sub Borra_Click()

    Rem DBGrid1.Col = 0
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 1
    Rem DBGrid1.Text = ""

    Rem DBGrid1.Col = 2
    Rem DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = "S"
    
    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
    
    XLote(WLugar, 1) = ""
    XLote(WLugar, 2) = ""
    XLote(WLugar, 3) = ""
    XLote(WLugar, 4) = ""
    XLote(WLugar, 5) = ""
    XLote(WLugar, 6) = ""
    XLote(WLugar, 7) = ""
    XLote(WLugar, 8) = ""
    XLote(WLugar, 9) = ""
    XLote(WLugar, 10) = ""
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    
    PrgModpedExp.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    WRenglon = 0
    DBGrid1.Refresh
        
    For A = 0 To 7
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            For iRow = 0 To 9
            
                WRenglon = WRenglon + 1
                WRow = iRow
                DBGrid1.Row = WRow
                
                DBGrid1.Col = 3
                Cantidad = Val(DBGrid1.Text)
                
                DBGrid1.Col = 4
                Resta = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Or Resta <> 0 Then
                    DBGrid1.Col = 5
                    If DBGrid1.Text <> "S" Then
                        m$ = "No asigno las partidas a todos los productos"
                        A = MsgBox(m$, 0, "MODULO DE FACTURACION")
                        DBGrid1.Refresh
                        Exit Sub
                    End If
                End If
                
            Next iRow
    Next A
    
    Erase Auxiliar
    Auxi = 0
        
    Suma = 0
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    For A = 0 To 7
        
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                Suma = Suma + 1
                WRenglon = WRenglon + 1
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text
                    
                DBGrid1.Col = 3
                Cantidad = DBGrid1.Text
                    
                DBGrid1.Col = 4
                Resta = Val(DBGrid1.Text)
                    
                Auxi = Pedido.Text
                Call Ceros(Auxi, 6)
        
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
            
                XPedido = Left$(ClavePedido(WRenglon), 6)
                XRenglon = Right$(ClavePedido(WRenglon), 2)
            
                XParam = "'" + XPedido + "','" _
                        + XRenglon + "'"
                WClavePedido = ClavePedido(WRenglon)
                spPedido = "ConsultaPedido2 " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                If rstPedido.RecordCount > 0 Then
                
                    XCantidad1 = Cantidad
                    xCantidad2 = Cantidad
                    
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    
                    XLote1 = XLote(WLugar, 1)
                    XLote2 = XLote(WLugar, 3)
                    XLote3 = XLote(WLugar, 5)
                    XLote4 = XLote(WLugar, 7)
                    XLote5 = XLote(WLugar, 9)
                    XCantiLote1 = XLote(WLugar, 2)
                    XCantiLote2 = XLote(WLugar, 4)
                    XCantiLote3 = XLote(WLugar, 6)
                    XCantiLote4 = XLote(WLugar, 8)
                    XCantiLote5 = XLote(WLugar, 10)
                
                    XParam = "'" + WClavePedido + "','" _
                            + XCantidad1 + "','" + xCantidad2 + "','" _
                            + XLote1 + "','" + XCantiLote1 + "','" _
                            + XLote2 + "','" + XCantiLote2 + "','" _
                            + XLote3 + "','" + XCantiLote3 + "','" _
                            + XLote4 + "','" + XCantiLote4 + "','" _
                            + XLote5 + "','" + XCantiLote5 + "'"
                                           
                    spPedido = "ModificaPedidoActualizaExpo " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                                        
            Next iRow
            
        Next A
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
    Exit Sub

WError:
     Resume Next
        
End Sub


Private Sub Limpia_Click()

    CargaLote.Visible = False
    Erase XLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    
    For A = 0 To 7
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Pedido.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 3
                Select Case KeyCode
                    Case 13
                        DBGrid1.Col = 3
                        DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = 4
                        KeyCode = 0
                    Case Else
                End Select
                        
            Case 4
                Select Case KeyCode
                    Case 13
                        DBGrid1.Col = 4
                        DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = 0
                        XTerminado = DBGrid1.Text
                        DBGrid1.Col = 3
                        XCantidad = Val(DBGrid1.Text)
                        WRow = DBGrid1.Row
                        
                        CargaLote.Visible = True
                        WLote1.Text = ""
                        WCanti1.Text = ""
                        WLote2.Text = ""
                        WCanti2.Text = ""
                        WLote3.Text = ""
                        WCanti3.Text = ""
                        WLote4.Text = ""
                        WCanti4.Text = ""
                        WLote5.Text = ""
                        WCanti5.Text = ""
                        
                        WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                            
                        If Val(XLote(WLugar, 1)) <> 0 Then
                            WLote1.Text = XLote(WLugar, 1)
                            WCanti1.Text = XLote(WLugar, 2)
                        End If
                        If Val(XLote(WLugar, 3)) <> 0 Then
                            WLote2.Text = XLote(WLugar, 3)
                            WCanti2.Text = XLote(WLugar, 4)
                        End If
                        If Val(XLote(WLugar, 5)) <> 0 Then
                            WLote3.Text = XLote(WLugar, 5)
                            WCanti3.Text = XLote(WLugar, 6)
                        End If
                        If Val(XLote(WLugar, 7)) <> 0 Then
                            WLote4.Text = XLote(WLugar, 7)
                            WCanti4.Text = XLote(WLugar, 8)
                        End If
                        If Val(XLote(WLugar, 9)) <> 0 Then
                            WLote5.Text = XLote(WLugar, 9)
                            WCanti5.Text = XLote(WLugar, 10)
                        End If
                            
                        WLote1.SetFocus
                        
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = WLote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Emision de facturas")
            End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WLote2.SetFocus
                Else
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo1
            G% = MsgBox(m$, 0, "Emiison de facturas")
            WLote1.SetFocus
        End If
        Rem WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
        Rem WLote2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = WLote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Emision de Facturas")
            End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            WLote3.SetFocus
                Else
            XSaldo2 = WSaldo2
            XSaldo2 = Pusing("###,###.##", XSaldo2)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo2
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote2.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub



Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = WLote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Emision de Facturas")
            End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
            WLote4.SetFocus
                Else
            XSaldo3 = WSaldo3
            XSaldo3 = Pusing("###,###.##", XSaldo3)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo3
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote3.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote4.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo4 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote4.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo4 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote4.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = WLote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote4.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti4.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote4.Text + " inexistente"
                G% = MsgBox(m$, 0, "Emision de Facturas")
            End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo4 >= Val(WCanti4.Text) Then
            WCanti4.Text = Pusing("###,###.##", WCanti4.Text)
            WLote5.SetFocus
                Else
            XSaldo4 = WSaldo4
            XSaldo4 = Pusing("###,###.##", XSaldo4)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo4
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote4.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote5.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo5 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote5.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo5 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote5.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = WLote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    DBGrid1.SetFocus
                    Exit Sub
                        Else
                    WLote5.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti5.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote5.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo5 >= Val(WCanti5.Text) Then
            WCanti5.Text = Pusing("###,###.##", WCanti5.Text)
            Call Verifica_Lote
            If WEstado = "S" Then
                WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                XLote(WLugar, 1) = WLote1.Text
                XLote(WLugar, 2) = WCanti1.Text
                XLote(WLugar, 3) = WLote2.Text
                XLote(WLugar, 4) = WCanti2.Text
                XLote(WLugar, 5) = WLote3.Text
                XLote(WLugar, 6) = WCanti3.Text
                XLote(WLugar, 7) = WLote4.Text
                XLote(WLugar, 8) = WCanti4.Text
                XLote(WLugar, 9) = WLote5.Text
                XLote(WLugar, 10) = WCanti5.Text
                CargaLote.Visible = False
                DBGrid1.Col = 5
                DBGrid1.Text = "S"
                If DBGrid1.Row < 40 Then
                    DBGrid1.Row = DBGrid1.Row + 1
                    WRow = DBGrid1.Row
                    XRow = DBGrid1.Row
                    DBGrid1.Col = 3
                    KeyCode = 0
                End If
                DBGrid1.Row = XRow
                DBGrid1.Col = 3
                KeyCode = 0
                Exit Sub
            End If
                Else
            XSaldo5 = WSaldo5
            XSaldo5 = Pusing("###,###.##", XSaldo5)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo5
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote5.SetFocus
        End If
        
        Rem WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
        Rem Call Verifica_Lote
        Rem If WEstado = "S" Then
        Rem     XLote(WRow, 1) = WLote1.Text
        Rem     XLote(WRow, 2) = WCanti1.Text
        Rem     XLote(WRow, 3) = WLote2.Text
        Rem     XLote(WRow, 4) = WCanti2.Text
        Rem     XLote(WRow, 5) = Wlote3.Text
        Rem     XLote(WRow, 6) = WCanti3.Text
        Rem     CargaLote.Visible = False
        Rem     DBGrid1.Col = 5
        Rem     DBGrid1.Text = "S"
        Rem     If DBGrid1.Row < 40 Then
        Rem         DBGrid1.Row = DBGrid1.Row + 1
        Rem         WRow = DBGrid1.Row
        Rem         XRow = DBGrid1.Row
        Rem         DBGrid1.Col = 3
        Rem         KeyCode = 0
        Rem     End If
        Rem     DBGrid1.Row = XRow
        Rem     DBGrid1.Col = 3
        Rem     KeyCode = 0
        Rem     Exit Sub
        Rem End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
ReDim UserData(0 To 5, 0 To 80)

mTotalRows& = 80

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
For i = 0 To 5
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Cantidad a Entregar"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad a Restar"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "OK"
             DBGrid1.Columns(newcnt).Width = 300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

    Erase XLote
    
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Renglon = 0
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Pedido.SetFocus
     
End Sub

Private Sub Proceso_Click()
        
    For A = 0 To 7
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    WNeto = 0
    
    Erase Auxiliar
    Erase ClavePedido
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Canti = !Cantidad - !Facturado
                    
                    If Canti > 0 Then
                
                        Renglon = Renglon + 1
                
                        Lugar1 = Int((Renglon - 1) / 10) * 10
                        Lugar2 = Renglon - Lugar1
                
                        DBGrid1.FirstRow = Lugar1
                        DBGrid1.Row = Lugar2 - 1
                
                        DBGrid1.Col = 0
                        DBGrid1.Text = !Terminado
                        Auxi1 = !Terminado
                
                        DBGrid1.Col = 2
                        DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                
                        Cantidad = IIf(IsNull(rstPedido!Cantidad1), "0", rstPedido!Cantidad1)
                        DBGrid1.Col = 3
                        DBGrid1.Text = Pusing("###,###.##", Str$(Cantidad))
                
                        Resta = IIf(IsNull(rstPedido!Cantidad2), "0", rstPedido!Cantidad2)
                        DBGrid1.Col = 4
                        DBGrid1.Text = Pusing("###,###.##", Str$(Resta))
                    
                        If Resta <> 0 Or Cantidad <> 0 Then
                            DBGrid1.Col = 5
                            DBGrid1.Text = "S"
                        End If
                    
                        
                        
                        WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        
                        XLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "0", rstPedido!lote1)
                        XLote(WLugar, 2) = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                        XLote(WLugar, 3) = IIf(IsNull(rstPedido!lote2), "0", rstPedido!lote2)
                        XLote(WLugar, 4) = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                        XLote(WLugar, 5) = IIf(IsNull(rstPedido!lote3), "0", rstPedido!lote3)
                        XLote(WLugar, 6) = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                        XLote(WLugar, 7) = IIf(IsNull(rstPedido!lote4), "0", rstPedido!lote4)
                        XLote(WLugar, 8) = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                        XLote(WLugar, 9) = IIf(IsNull(rstPedido!lote5), "0", rstPedido!lote5)
                        XLote(WLugar, 10) = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                    
                        Auxiliar(Renglon, 1) = Auxi1
                        Auxiliar(Renglon, 2) = Canti
                        
                        ClavePedido(Renglon) = rstPedido!Clave
                    
                    End If
        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(Da, 1)
        Canti = Auxiliar(Da, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
        
            DBGrid1.Col = 1
            DBGrid1.Text = rstPrecios!Descripcion
            
            Rem DBGrid1.Col = 3
            Rem DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
            
            Precio = rstPrecios!Precio
            rstPrecios.Close
        End If

        If Val(Canti) <> 0 Then
            WNeto = WNeto + (Val(Canti) * Precio)
        End If
        
    Next Da
    
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
    
    Graba.Enabled = True

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            If rstPedido!Autorizo <> "X" Then
                rstPedido.Close
                m$ = "EL PEDIDO NO FUE AUTORIZADO"
                A% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                    Else
                Cliente.Text = rstPedido!Cliente
                Fecha.Text = rstPedido!Fecha
                WFecEntrega = rstPedido!FecEntrega
                WObservaciones = rstPedido!Observaciones
                rstPedido.Close
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    WDirentrega = rstCliente!DirEntrega
                    WPago = Str$(rstCliente!Pago1)
                    rstCliente.Close
                    spPago = "ConsultaPago " + "'" + WPago + "'"
                    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPago.RecordCount > 0 Then
                        WDespago = rstPago!Nombre
                        rstPago.Close
                    End If
                End If
                Call Proceso_Click
                DBGrid1.FirstRow = 0
                DBGrid1.Row = 0
                DBGrid1.Col = 3
                DBGrid1.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    
    If Val(WLote1.Text) <> 0 Then
        Suma = Suma + Val(WCanti1.Text)
    End If
    If Val(WLote2.Text) <> 0 Then
        Suma = Suma + Val(WCanti2.Text)
    End If
    If Val(WLote3.Text) <> 0 Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    If Val(WLote4.Text) <> 0 Then
        Suma = Suma + Val(WCanti4.Text)
    End If
    If Val(WLote5.Text) <> 0 Then
        Suma = Suma + Val(WCanti5.Text)
    End If
    
    If Suma = XCantidad Then
        WEstado = "S"
    End If
    
End Sub

