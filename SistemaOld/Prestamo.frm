VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPrestamo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Prestamos entre Plantas"
   ClientHeight    =   8595
   ClientLeft      =   210
   ClientTop       =   405
   ClientWidth     =   11835
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11835
   Visible         =   0   'False
   Begin VB.CommandButton AvisoError 
      Caption         =   "No se puede grabar el prestamo. El sistema se encuentra sin conexion con las demas plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6120
      Picture         =   "Prestamo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Pass 
      Height          =   1575
      Left            =   3840
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton WCancela 
         Caption         =   "Cancela"
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Ingrese se Password"
         Height          =   255
         Left            =   720
         TabIndex        =   37
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.ComboBox Destino 
      Height          =   315
      Left            =   8880
      TabIndex        =   33
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   480
      Width           =   5055
   End
   Begin VB.ComboBox Tipomov 
      Height          =   315
      Left            =   8880
      TabIndex        =   29
      Text            =   " "
      Top             =   120
      Width           =   2415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11040
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   500
      Left            =   2520
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   4680
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   5160
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
      Height          =   500
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      Height          =   500
      Left            =   1320
      TabIndex        =   10
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   500
      Left            =   2520
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   11175
      Begin VB.TextBox WLote 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9960
         MaxLength       =   6
         TabIndex        =   35
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox WMovi 
         Height          =   285
         Left            =   8760
         MaxLength       =   1
         TabIndex        =   21
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox WCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7440
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin MSMask.MaskEdBox WTerminado 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WTipo 
         Height          =   285
         Left            =   360
         MaxLength       =   1
         TabIndex        =   18
         Text            =   " "
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         Height          =   255
         Left            =   9960
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E/S"
         Height          =   255
         Left            =   8760
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   7440
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto Terminado"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M/T"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label WDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   300
         Left            =   3840
         TabIndex        =   6
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.CommandButton Graba1 
      Caption         =   "Graba"
      Height          =   500
      Left            =   120
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3735
      Left            =   240
      OleObjectBlob   =   "Prestamo.frx":0742
      TabIndex        =   3
      Top             =   960
      Width           =   11295
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10560
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2205
      ItemData        =   "Prestamo.frx":112C
      Left            =   3840
      List            =   "Prestamo.frx":1133
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   500
      Left            =   1320
      TabIndex        =   0
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destno"
      Height          =   255
      Left            =   7080
      TabIndex        =   32
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Movimiento"
      Height          =   285
      Left            =   7080
      TabIndex        =   28
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Movimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Tipo As String
Private Articulo As String
Private Terminado As String
Private WTipomov As String
Private WDestino As String
Private Auxiliar(100, 6) As String
Private WCodigoMov As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPrestamo As Recordset
Dim spPrestamo As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim XParam As String
Dim ACodigo As String
Dim ADescripcion As String
Dim Alineas As String
Dim AUnidad As String
Dim AInicial  As String
Dim AEntradas As String
Dim ASalidas As String
Dim AMinimo As String
Dim ADeposito As String
Dim APedido As String
Dim AEnvase1 As String
Dim AEnvase2 As String
Dim AEnvase3 As String
Dim AEnvase4 As String
Dim AEnvase5 As String
Dim AEnvase6 As String
Dim AProceso As String
Dim ACosto As String
Dim AFactor As String
Dim AImpreadi As String
Dim AClase As String
Dim AIntervencion As String
Dim ANaciones As String
Dim AEmbalaje As String
Dim Aversion As String
Dim AFechaVersion As String
Dim AControla As String
Dim AEscrito As String
Dim AObservaciones As String
Dim ATipoeti As String
Dim ADate As String
Dim XSaldo As Double
Dim XLaudo As String
Dim XLote As String
Dim XCodigo As String
Dim XRenglon As String
Dim XTerminado As String
Private Producto As String
Private Costo As Double
Private XAuxiliar(100, 7) As String
Dim CargaEmpresa(12, 2) As String

Private Sub AvisoError_Click()
    AvisoError.Visible = False
End Sub

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = ""
    
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""
    WLinea.Text = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click
    PrgPrestamo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima"
     Opcion.AddItem "Lote de Producto Terminado"
     Opcion.AddItem "Lote de Materia Prima"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    Rem OPEN_FILE_Movguia
    Rem OPEN_FILE_TERMINADO
    Rem OPEN_FILE_Articulo
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
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstTerminado.RecordCount > 0 Then
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
            End If
            
        Case 1
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case 2
            If WTipo.Text = "T" Then
            
                XParam = "'" + WTerminado.Text + "','" _
                            + WTerminado.Text + "'"
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
                
                                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                                        Else
                                        
                                    XSaldo = rstHoja!Saldo
                                    Call Redondeo(XSaldo)
                    
                                    If Val(rstHoja!Renglon) = 1 And XSaldo <> 0 Then
                                
                                        XLaudo = rstHoja!Hoja
                                        Call Ceros(XLaudo, 6)
                                
                                        IngresaItem = "Lote : " + XLaudo + " Saldo : " + Str$(XSaldo)
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = XLaudo
                                        WIndice.AddItem IngresaItem
                                    
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
                
                Rem guias
                
                XParam = "'" + WTerminado.Text + "','" _
                            + WTerminado.Text + "'"
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
                
                                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                                        Else
                
                                    If rstMovguia!Tipo = "T" And rstMovguia!Movi = "E" Then
                
                                        XLaudo = rstMovguia!Lote
                                        XSaldo = rstMovguia!Saldo
                                        Call Redondeo(XSaldo)
                                        
                                        If XSaldo <> 0 Then
                                        
                                            IngresaItem = "Lote : " + XLaudo + " Saldo : " + Str$(XSaldo)
                                            Pantalla.AddItem IngresaItem
                                            IngresaItem = XLaudo
                                            WIndice.AddItem IngresaItem
                                        
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
                
            End If
            
        Case 3
            If WTipo.Text = "M" Then
            
                XParam = "'" + WArticulo.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
    
                    With rstLaudo
        
                        .MoveFirst
            
                        If .NoMatch = False Then
                            Do
                    
                                If .EOF = True Then
                                    Exit Do
                                End If
                
                                XSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                Call Redondeo(XSaldo)
                            
                                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                                        Else
                    
                                    If rstLaudo!Articulo = WArticulo.Text And XSaldo <> 0 Then
                                
                                        WLaudo = rstLaudo!Laudo
                                        WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                                        XLaudo = WLaudo
                                        Call Ceros(XLaudo, 6)
                        
                                        If WLiberada <> 0 Then
                                    
                                            IngresaItem = "Lote : " + XLaudo + " Saldo : " + Str$(XSaldo)
                                            Pantalla.AddItem IngresaItem
                                            IngresaItem = WLaudo
                                            WIndice.AddItem IngresaItem
                            
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
                
                    rstLaudo.Close
                
                End If
    
                XParam = "'" + WArticulo.Text + "','" _
                        + WArticulo.Text + "'"
                spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then

                    With rstMovguia
    
                        .MoveFirst
            
                        If .NoMatch = False Then
                            Do
            
                                If .EOF = True Then
                                    Exit Do
                                End If
                
                                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                                        Else
                        
                                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo.Text Then
                    
                                        XSaldo = rstMovguia!Saldo
                                        Call Redondeo(XSaldo)
                                        WLaudo = rstMovguia!Lote
                                        XLote = WLaudo
                                        Call Ceros(XLote, 6)
                        
                                        If rstMovguia!Movi = "E" And XSaldo <> 0 Then
                                    
                                            IngresaItem = "Lote : " + XLaudo + " Saldo : " + Str$(XSaldo)
                                            Pantalla.AddItem IngresaItem
                                            IngresaItem = WLaudo
                                            WIndice.AddItem IngresaItem
                                    
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
                
            End If
            
        Case Else
        
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 1 Then
        WLinea.Text = DBGrid1.Row + 1
        WTipo.Text = DBGrid1.Text
            Else
        WTipo.Text = ""
        WLinea.Text = ""
    End If

    DBGrid1.Col = 1
    If Len(DBGrid1.Text) = 12 Then
        WTerminado.Text = DBGrid1.Text
            Else
        WTerminado.Text = "  -     -   "
    End If

    DBGrid1.Col = 2
    If Len(DBGrid1.Text) = 10 Then
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
    End If
    
    DBGrid1.Col = 3
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 4
    WCantidad.Text = DBGrid1.Text
    
    DBGrid1.Col = 5
    WMovi.Text = DBGrid1.Text
    
    DBGrid1.Col = 6
    WLote.Text = DBGrid1.Text
    
    WTipo.SetFocus

End Sub

Private Sub Graba_Click()


    If Val(Codigo.Text) = 0 Then
        Exit Sub
    End If
        

    Auxi = Codigo.Text
    Call Ceros(Auxi, 5)
    WCodigoMov = "9" + Auxi

    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    WDestino = Str$(Destino.ListIndex)
    Call Ceros(WDestino, 2)
    
    If Val(WTipomov) <> 0 Then
        Exit Sub
    End If
    
    If Val(WDestino) = 0 Then
        Exit Sub
    End If
    
    If Val(WCodigoMov) < 900000 Then
        Exit Sub
    End If
    
    
    
    Select Case Val(WEmpresa)
        Case 1
            If Val(WEmpresa) = Val(WDestino) Or Val(WDestino) <> 2 Then
                Exit Sub
            End If
        Case 2
            If Val(WEmpresa) = Val(WDestino) Or Val(WDestino) <> 1 Then
                Exit Sub
            End If
        Case 3, 5, 6
            If Val(WEmpresa) = Val(WDestino) Or Val(WDestino) <> 4 Then
                Exit Sub
            End If
        Case 4
            If Val(WEmpresa) = Val(WDestino) Then
                Exit Sub
            End If
            If Val(WDestino) <> 3 And Val(WDestino) <> 5 And Val(WDestino) <> 6 Then
                Exit Sub
            End If
        Case 7
            If Val(WEmpresa) = Val(WDestino) Or Val(WDestino) <> 8 Then
                Exit Sub
            End If
         Case 8
            If Val(WEmpresa) = Val(WDestino) Or Val(WDestino) <> 7 Then
                Exit Sub
            End If
        Case 9, 10
            Exit Sub
        Case Else
    End Select
    
    
    Rem
    Rem verifica conexciones con las otras plantas
    Rem
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
    WDestino = Str$(Destino.ListIndex)
    Call Ceros(WDestino, 1)
    XDestino = WDestino
        
    CargaEmpresa(1, 1) = "0001"
    CargaEmpresa(1, 2) = "Empresa01"
    CargaEmpresa(2, 1) = "0002"
    CargaEmpresa(2, 2) = "Empresa02"
    CargaEmpresa(3, 1) = "0003"
    CargaEmpresa(3, 2) = "Empresa03"
    CargaEmpresa(4, 1) = "0004"
    CargaEmpresa(4, 2) = "Empresa04"
    CargaEmpresa(5, 1) = "0005"
    CargaEmpresa(5, 2) = "Empresa05"
    CargaEmpresa(6, 1) = "0006"
    CargaEmpresa(6, 2) = "Empresa06"
    CargaEmpresa(7, 1) = "0007"
    CargaEmpresa(7, 2) = "Empresa07"
    CargaEmpresa(8, 1) = "0008"
    CargaEmpresa(8, 2) = "Empresa08"
    CargaEmpresa(9, 1) = "0009"
    CargaEmpresa(9, 2) = "Empresa09"
    CargaEmpresa(10, 1) = "0010"
    CargaEmpresa(10, 2) = "Empresa10"
    CargaEmpresa(11, 1) = "0011"
    CargaEmpresa(11, 2) = "Empresa11"

    WEmpresa = CargaEmpresa(Val(XDestino), 1)
    txtOdbc = CargaEmpresa(Val(XDestino), 2)
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Call Conecta_Empresa
    
    On Error GoTo 0
    If WSalidaError = "N" Then Exit Sub
    
    
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    
    XParam = "'" + WTipomov + "','" _
                + WCodigoMov + "'"
    spMovguia = "ListaMovguia " + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        rstMovguia.Close
        Exit Sub
    End If
    
    
    

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Erase Auxiliar
    Renglon = 0
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 2)
    
    XParam = "'" + WTipomov + "','" _
                + WCodigoMov + "'"
    spMovguia = "ListaMovguia " + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Auxiliar(Renglon, 1) = rstMovguia!Tipo
                    Auxiliar(Renglon, 2) = rstMovguia!Terminado
                    Auxiliar(Renglon, 3) = rstMovguia!Articulo
                    Auxiliar(Renglon, 4) = rstMovguia!Cantidad
                    Auxiliar(Renglon, 5) = rstMovguia!Movi
                    Auxiliar(Renglon, 6) = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
    End If
    
    For Da = 1 To Renglon
    
        Tipo = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Movi = Auxiliar(Da, 5)
        Lote = Auxiliar(Da, 6)
        
        Select Case Tipo
            Case "M"
                WControla = 0
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    WCodigo = Articulo
                    If Movi = "E" Then
                        WEntradas = Str$(rstArticulo!Entradas - Val(Cantidad))
                        WSalidas = Str$(rstArticulo!Salidas)
                            Else
                        WSalidas = Str$(rstArticulo!Salidas - Val(Cantidad))
                        WEntradas = Str$(rstArticulo!Entradas)
                    End If
                    WDate = Date$
                    rstArticulo.Close
                
                    XParam = "'" + WCodigo + "','" _
                                 + WEntradas + "','" _
                                 + WSalidas + "','" _
                                 + WDate + "'"
                                           
                    spArticulo = "ModificaArticuloMovimientos " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Articulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WClave = rstLaudo!Clave
                            If Movi = "S" Then
                                WSaldo = Str$(rstLaudo!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstLaudo.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Articulo + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "S" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                    
                End If
                
            Case "T"
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    WCodigo = Terminado
                    If Movi = "E" Then
                        WEntradas = Str$(rstTerminado!Entradas - Val(Cantidad))
                        WSalidas = Str$(rstTerminado!Salidas)
                            Else
                        WSalidas = Str$(rstTerminado!Salidas - Val(Cantidad))
                        WEntradas = Str$(rstTerminado!Entradas)
                    End If
                    WDate = Date$
                    rstTerminado.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Terminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                        
                            WClave = rstHoja!Clave
                            If Movi = "S" Then
                                WSaldo = Str$(rstHoja!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstHoja.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spHoja = "ModificaHojaSaldo " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Terminado + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "S" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                    
                    
                End If
            
            Case Else
        End Select
        
    Next Da
    
    XParam = "'" + WTipomov + "','" _
                + WCodigoMov + "'"
    spMovguia = "BorrarMovguia " + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenDynaset, dbSQLPassThrough)
    
    Renglon = 0
    Erase Auxiliar
    DBGrid1.Refresh
                
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Tipo = DBGrid1.Text
                                       
            DBGrid1.Col = 1
            Terminado = DBGrid1.Text
                    
            DBGrid1.Col = 2
            Articulo = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 5
            Movi = DBGrid1.Text
            
            DBGrid1.Col = 6
            Lote = DBGrid1.Text
                    
            If Tipo <> "" Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = WCodigoMov
                Call Ceros(Auxi1, 6)
                
                WTipomov = Str$(Tipomov.ListIndex)
                Call Ceros(WTipomov, 1)
                WDestino = Str$(Destino.ListIndex)
                Call Ceros(WDestino, 1)
                WCodigo = WCodigoMov
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WTipo = Tipo
                WArticulo = Articulo
                WTerminado = Terminado
                WCantidad = Cantidad
                WMovi = Movi
                WObservaciones = Observaciones.Text
                WClave = Trim(Str$(Val(WTipomov))) + Auxi1 + Auxi
                WDate = Date$
                WMarca = ""
                WPartida = Lote
                WLote = ""
                WSaldo = "0"
                
                Auxiliar(Renglon, 1) = WTipo
                Auxiliar(Renglon, 2) = WTerminado
                Auxiliar(Renglon, 3) = WArticulo
                Auxiliar(Renglon, 4) = WCantidad
                Auxiliar(Renglon, 5) = WMovi
                Auxiliar(Renglon, 6) = WPartida

                XParam = "'" + WClave + "','" _
                         + WTipomov + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaord + "','" _
                         + WMovi + "','" _
                         + WObservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WDestino + "','" _
                         + WLote + "','" _
                         + WSaldo + "','" _
                         + WPartida + "'"
                         
                spMovguia = "AltaMovguia " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                
                Rem Da de alta el prestamo
                
                If WArticulo <> "AA-000-100" Then
                
                    XCodigo = WCodigo
                    XRenglon = "01"
                    XFecha = WFecha
                    XOrdFecha = WFechaord
                    XObservaciones = WObservaciones
                    XTipo = WTipo
                    XArticulo = WArticulo
                    XTerminado = WTerminado
                    XCantidad = WCantidad
                    XCosto = ""
                    XDestino = ""
                
                    Call Ceros(XCodigo, 6)
                    Call Ceros(XRenglon, 2)
                
                    XClave = XCodigo + XRenglon
                
                    XEmpresa = WEmpresa
                    Select Case Val(XEmpresa)
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
            
                    If XTipo = "M" Then
                        spArticulo = "ConsultaArticulo" + "'" + XArticulo + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            XCosto = Str$(rstArticulo!Costo1)
                            rstArticulo.Close
                        End If
                            Else
                        Call Calcula_Costo(XTerminado, Costo)
                        XCosto = Str$(Costo)
                    End If
            
                    Call Conecta_Empresa
                                        
                        Else
                        
                    XCodigo = WCodigo
                    XRenglon = "01"
                    XFecha = WFecha
                    XOrdFecha = WFechaord
                    XObservaciones = WObservaciones
                    XTipo = ""
                    XArticulo = ""
                    XTerminado = ""
                    XCantidad = WCantidad
                    XCosto = "1"
                    XDestino = ""
                
                    Call Ceros(XCodigo, 6)
                    Call Ceros(XRenglon, 2)
                
                    XClave = XCodigo + XRenglon
                
                End If
                    
                XParam = "'" + XClave + "','" _
                         + XCodigo + "','" _
                         + XRenglon + "','" _
                         + XFecha + "','" _
                         + XOrdFecha + "','" _
                         + XObservaciones + "','" _
                         + XTipo + "','" _
                         + XArticulo + "','" _
                         + XTerminado + "','" _
                         + XCantidad + "','" _
                         + XCosto + "','" _
                         + XDestino + "'"
                                         
                Set rstPrestamo = db.OpenRecordset("AltaPrestamo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
        Next iRow
            
    Next a
                
    For Da = 1 To Renglon
    
        Tipo = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Movi = Auxiliar(Da, 5)
        Lote = Auxiliar(Da, 6)
        
        Select Case Tipo
            Case "M"
                WControla = 0
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                    WCodigo = Articulo
                    If Movi = "E" Then
                        WEntradas = Str$(rstArticulo!Entradas + Val(Cantidad))
                        WSalidas = Str$(rstArticulo!Salidas)
                            Else
                        WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                        WEntradas = Str$(rstArticulo!Entradas)
                    End If
                    rstArticulo.Close
                    WDate = Date$
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spArticulo = "ModificaArticuloMovimientos " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Articulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WClave = rstLaudo!Clave
                            If Movi = "E" Then
                                WSaldo = Str$(rstLaudo!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstLaudo.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Articulo + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "E" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                End If
                
            Case "T"
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
        
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    WCodigo = Terminado
                    If Movi = "E" Then
                        WEntradas = Str$(rstTerminado!Entradas + Val(Cantidad))
                        WSalidas = Str$(rstTerminado!Salidas)
                            Else
                        WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
                        WEntradas = Str$(rstTerminado!Entradas)
                    End If
                    WDate = Date$
                    rstTerminado.Close
                
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + Terminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WClave = rstHoja!Clave
                            If Movi = "E" Then
                                WSaldo = Str$(rstHoja!Saldo + Val(Cantidad))
                                    Else
                                WSaldo = Str$(rstHoja!Saldo - Val(Cantidad))
                            End If
                            WDate = Date$
                            rstHoja.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spHoja = "ModificahojaSaldo " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Terminado + "','" _
                                    + Lote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                If Movi = "E" Then
                                    WSaldo = Str$(rstMovguia!Saldo + Val(Cantidad))
                                        Else
                                    WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad))
                                End If
                                WDate = Date$
                                rstMovguia.Close
                            
                                XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                    
                End If
            
            Case Else
        End Select
        
    Next Da
        
    WPasa = "S"
    If WPasa = "S" Then
    
        XEmpresa = WEmpresa
        XDestino = WDestino
        
        Select Case Val(WDestino)
            Case 1
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                WEmpresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                WEmpresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        Renglon = 0
        Erase Auxiliar
        DBGrid1.Refresh
                
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Tipo = DBGrid1.Text
                                       
                DBGrid1.Col = 1
                Terminado = DBGrid1.Text
                    
                DBGrid1.Col = 2
                Articulo = DBGrid1.Text
                    
                DBGrid1.Col = 4
                Cantidad = DBGrid1.Text
                    
                DBGrid1.Col = 5
                Movi = DBGrid1.Text
                
                DBGrid1.Col = 6
                Lote = DBGrid1.Text
                    
                If Tipo <> "" Then
                    
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = WCodigoMov
                    Call Ceros(Auxi1, 6)
                
                    WTipomov = Str$(Val(XEmpresa))
                    Call Ceros(WTipomov, 1)
                    WDestino = "0"
                    Call Ceros(WDestino, 1)
                    WCodigo = WCodigoMov
                    WRenglon = Str$(Renglon)
                    WFecha = Fecha.Text
                    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    WTipo = Tipo
                    WArticulo = Articulo
                    WTerminado = Terminado
                    WCantidad = Cantidad
                    If Movi = "E" Then
                        WMovi = "S"
                            Else
                        WMovi = "E"
                    End If
                    WObservaciones = Observaciones.Text
                    WClave = Trim(Str$(Val(WTipomov))) + Auxi1 + Auxi
                    WDate = Date$
                    WMarca = ""
                    WLote = Lote
                    WSaldo = Cantidad
                    WPartida = ""
                    
                    Select Case WTipo
                        Case "M"
                            WEntra = "N"
        
                            XParam = "'" + WLote + "','" _
                                         + WArticulo + "'"
                            spLaudo = "ListaLaudoArticulo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                WEntra = "S"
                                XClave = rstLaudo!Clave
                                WSaldo = Str$(rstLaudo!Saldo + Val(WCantidad))
                                WDate = Date$
                                rstLaudo.Close
                            
                                XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                spLaudo = "ModificaLaudoSaldo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                
                            If WEntra = "N" Then
                                XParam = "'" + WArticulo + "','" _
                                            + WLote + "'"
                                spMovguia = "ListaMovguiaLote " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstMovguia.Close
                            
                                    XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                    spMovguia = "ModificaMovguiaSaldo " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                            End If
                            
                        Case "T"
                            WEntra = "N"
                            
                            Select Case Val(WEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    If Left$(WTerminado, 2) = "SU" Then
                                        WTerminado = "PT" + Mid$(WTerminado, 3, 10)
                                            Else
                                        WTerminado = "PE" + Mid$(WTerminado, 3, 10)
                                    End If
                                Case Else
                                    If Left$(WTerminado, 2) = "PE" Then
                                        WTerminado = "PT" + Mid$(WTerminado, 3, 10)
                                            Else
                                        WTerminado = "SU" + Mid$(WTerminado, 3, 10)
                                    End If
                            End Select
            
                            XParam = "'" + WLote + "','" _
                                        + WTerminado + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                WEntra = "S"
                                XClave = rstHoja!Clave
                                WSaldo = Str$(rstHoja!Saldo + Val(WCantidad))
                                WDate = Date$
                                rstHoja.Close
                            
                                XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                spHoja = "ModificaHojaSaldo " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                
                            If WEntra = "N" Then
                                XParam = "'" + WTerminado + "','" _
                                            + WLote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstMovguia.Close
                            
                                    XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                    spMovguia = "ModificaMovguiaSaldo " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                            End If
                        Case Else
                    End Select
                        
                    If WEntra = "S" Then
                        WSaldo = "0"
                    End If
                
                    XParam = "'" + WClave + "','" _
                                + WTipomov + "','" _
                                + WCodigo + "','" _
                                + WRenglon + "','" _
                                + WFecha + "','" _
                                + WTipo + "','" _
                                + WArticulo + "','" _
                                + WTerminado + "','" _
                                + WCantidad + "','" _
                                + WFechaord + "','" _
                                + WMovi + "','" _
                                + WObservaciones + "','" _
                                + WDate + "','" _
                                + WMarca + "','" _
                                + WDestino + "','" _
                                + WLote + "','" _
                                + WSaldo + "','" _
                                + WPartida + "'"
                                
                    spMovguia = "AltaMovguia " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
                    Select Case WTipo
                        Case "M"
                            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                                WCodigo = WArticulo
                                If WMovi = "E" Then
                                    WEntradas = Str$(rstArticulo!Entradas + Val(WCantidad))
                                    WSalidas = Str$(rstArticulo!Salidas)
                                        Else
                                    WSalidas = Str$(rstArticulo!Salidas + Val(WCantidad))
                                    WEntradas = Str$(rstArticulo!Entradas)
                                End If
                                WDate = Date$
                                rstArticulo.Close
                    
                                XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                spArticulo = "ModificaArticuloMovimientos " + XParam
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                                            
                        Case "T"
                            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstTerminado.RecordCount > 0 Then
        
                                WCodigo = WTerminado
                                If WMovi = "E" Then
                                    WEntradas = Str$(rstTerminado!Entradas + Val(WCantidad))
                                    WSalidas = Str$(rstTerminado!Salidas)
                                        Else
                                    WSalidas = Str$(rstTerminado!Salidas + Val(WCantidad))
                                    WEntradas = Str$(rstTerminado!Entradas)
                                End If
                                WDate = Date$
                                rstTerminado.Close
                
                                XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                spTerminado = "ModificaTerminadoMovimientos " + XParam
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                
                                    Else
                                    
                                Alta = "N"
                                    
                                Call Conecta_Empresa
                                    
                                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                
                                    Alta = "S"
                                
                                    ACodigo = WTerminado
                                    ADescripcion = rstTerminado!Descripcion
                                    Alineas = rstTerminado!Linea
                                    AUnidad = rstTerminado!Unidad
                                    AInicial = ""
                                    AEntradas = ""
                                    ASalidas = ""
                                    AMinimo = ""
                                    ADeposito = rstTerminado!Deposito
                                    AEnvase1 = rstTerminado!Envase1
                                    AEnvase2 = rstTerminado!Envase2
                                    AEnvase3 = rstTerminado!Envase3
                                    AEnvase4 = rstTerminado!Envase4
                                    AEnvase5 = rstTerminado!Envase5
                                    AEnvase6 = rstTerminado!Envase6
                                    AProceso = ""
                                    AImpreadi = ""
                                    AClase = ""
                                    AIntervencion = ""
                                    ANaciones = ""
                                    AEmbalaje = ""
                                    AImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
                                    AClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                                    AIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                                    ANaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                                    AEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
                                    Aversion = ""
                                    AFechaVersion = "  /  /    "
                                    AControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                                    AObservaciones = IIf(IsNull(rstTerminado!Observaciones), "", rstTerminado!Observaciones)
                                    ATipoeti = IIf(IsNull(rstTerminado!Tipoeti), "", rstTerminado!Tipoeti)
                                    ACosto = ""
                                    AFactor = ""
                                    AEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
                                    
                                    rstTerminado.Close
                                    
                                End If
                                
                                Select Case Val(XDestino)
                                    Case 1
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 2
                                        WEmpresa = "0002"
                                        txtOdbc = "Empresa02"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 3
                                        WEmpresa = "0003"
                                        txtOdbc = "Empresa03"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 4
                                        WEmpresa = "0004"
                                        txtOdbc = "Empresa04"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 5
                                        WEmpresa = "0005"
                                        txtOdbc = "Empresa05"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 6
                                        WEmpresa = "0006"
                                        txtOdbc = "Empresa06"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 7
                                        WEmpresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 8
                                        WEmpresa = "0008"
                                        txtOdbc = "Empresa08"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 9
                                        WEmpresa = "0009"
                                        txtOdbc = "Empresa09"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 10
                                        WEmpresa = "0010"
                                        txtOdbc = "Empresa10"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case 11
                                        WEmpresa = "0011"
                                        txtOdbc = "Empresa11"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                End Select
                                
                                If Alta = "S" Then
                                
                                    If WMovi = "E" Then
                                        AEntradas = Str$(Val(WCantidad))
                                            Else
                                        ASalidas = Str$(Val(WCantidad))
                                    End If
                            
                                    XParam = "'" + ACodigo + "','" _
                                            + ADescripcion + "','" _
                                            + Alineas + "','" _
                                            + AUnidad + "','" _
                                            + AInicial + "','" + AEntradas + "','" _
                                            + ASalidas + "','" + AMinimo + "','" _
                                            + ADeposito + "','" + APedido + "','" _
                                            + AEnvase1 + "','" + AEnvase2 + "','" _
                                            + AEnvase3 + "','" + AEnvase4 + "','" _
                                            + AEnvase5 + "','" + AEnvase6 + "','" _
                                            + AProceso + "','" _
                                            + ACosto + "','" _
                                            + AFactor + "','" _
                                            + ADate + "','" _
                                            + AImpreadi + "','" _
                                            + AClase + "','" _
                                            + AIntervencion + "','" _
                                            + ANaciones + "','" _
                                            + AEmbalaje + "','" _
                                            + Aversion + "','" _
                                            + AFechaVersion + "','" _
                                            + AControla + "','" _
                                            + AObservaciones + "','" _
                                            + ATipoeti + "','" _
                                            + AEscrito + "'"

                                    Set rstTerminado = db.OpenRecordset("AltaTerminado " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                End If
                                
                            End If
            
                        Case Else
                    End Select
                End If
                    
            Next iRow
            
        Next a
    
        Call Conecta_Empresa
    
    End If
    
    T$ = "Guias de Traslado Interno"
    m$ = "Desea Imprimir la guia de traslado interno"
    Rem Respuesta% = MsgBox(m$, 32 + 4, T$)
    Rem If Respuesta% = 6 Then
    Rem     Call Impresion
    Rem End If
    
    Pass.Visible = False
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Codigo.SetFocus
        
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""
    
    WTipo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WTipo.Text = ""
    WTerminado.Text = "  -     -   "
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WMovi.Text = ""
    WLote.Text = ""

    Codigo.Text = ""
    WCodigoMov = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    Rem With rstMovguia
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Codigo.Text = !Codigo + 1
    Rem             Else
    Rem         Codigo.Text = ""
    Rem     End If
    Rem End With
    
    Tipomov.ListIndex = 0
    Codigo.Text = ""
    WCodigoMov = ""
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 1)
    Destino.ListIndex = 0
    
    Rem spMovguia = "ListaMovguiaNumero " + "'" + WTipomov + "'"
    Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMovguia.RecordCount > 0 Then
    Rem     With rstMovguia
    Rem         .MoveLast
    Rem         Codigo.Text = rstMovguia!Codigo + 1
    Rem         If Val(Codigo.Text) < 900000 Then
    Rem             Codigo.Text = "1"
    Rem                 Else
    Rem             Codigo.Text = Str$(Val(Right$(Codigo.Text, 5)))
    Rem         End If
    Rem     End With
    Rem     rstMovguia.Close
    Rem         Else
    Rem     Codigo.Text = "1"
    Rem End If
    
    Codigo.Text = ""
    
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Graba1.Enabled = True
    Pass.Visible = False

    Codigo.SetFocus

End Sub


Private Sub WTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "M" Or WTipo.Text = "T" Then
            If WTipo.Text = "M" Then
                WArticulo.SetFocus
                    Else
                WTerminado.SetFocus
            End If
                Else
            WTipo.SetFocus
        End If
    End If
End Sub

Private Sub WTerminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDescripcion.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            WCantidad.SetFocus
                Else
            WTerminado.SetFocus
        End If
    End If
End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
                WDescripcion.Caption = rstArticulo!Descripcion
                WCantidad.SetFocus
                    Else
                WArticulo.SetFocus
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        Select Case Tipomov.ListIndex
            Case 0
                WMovi.Text = "S"
            Case Else
                WMovi.Text = "E"
        End Select
        WLote.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If WTipo.Text = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            WArticulo.Text = UCase(WArticulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
            
                WCanti = 0
                XParam = "'" + WLote.Text + "','" _
                            + WArticulo.Text + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WCanti = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WArticulo.Text + "','" _
                            + WLote.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                If WEntra = "S" Then
                    If WCanti >= Val(WCantidad.Text) Or WMovi.Text = "E" Then
                        Call Alta_Vector
                        Call Ingresa_Click
                        WTipo.SetFocus
                            Else
                        m$ = WArticulo.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                        G% = MsgBox(m$, 0, "Guias de Traslado Internos")
                    End If
                        Else
                    m$ = WArticulo.Text + " Articulo inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Guias de Traslado Internos")
                End If
                
                    Else
                    
                Call Alta_Vector
                Call Ingresa_Click
                WTipo.SetFocus
                
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote.Text + "','" _
                        + WTerminado.Text + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WCanti = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WTerminado.Text + "','" _
                            + WLote.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WCanti = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra = "S" Then
                If WCanti >= Val(WCantidad.Text) Or WMovi.Text = "E" Then
                    Call Alta_Vector
                    Call Ingresa_Click
                    WTipo.SetFocus
                        Else
                    m$ = WTerminado.Text + " Stock Insufucuente. Cantidad:" + Str$(WCanti)
                    G% = MsgBox(m$, 0, "Guias de Traslado Internos")
                End If
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                G% = MsgBox(m$, 0, "Guias de Traslado Internos")
            End If
            
        End If
    
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub pantalla_Click()
    Codigo.SetFocus
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WTipo.Text = "T"
                WTerminado.Text = Claveven$
                WDescripcion.Caption = rstTerminado!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = "T"
                DBGrid1.Col = 1
                DBGrid1.Text = rstTerminado!Codigo
                DBGrid1.Col = 3
                DBGrid1.Text = rstTerminado!Descripcion
                rstTerminado.Close
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
            End If
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WTipo.Text = "M"
                WArticulo.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                    
                DBGrid1.Col = 0
                DBGrid1.Text = "M"
                DBGrid1.Col = 2
                DBGrid1.Text = rstArticulo!Codigo
                DBGrid1.Col = 3
                DBGrid1.Text = rstArticulo!Descripcion
                rstArticulo.Close
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                                        
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
                    
            End If
            
        Case 2
            Indice = Pantalla.ListIndex
            WLote.Text = Str$(Val(WIndice.List(Indice)))
            WLote.SetFocus
            
        Case 3
            Indice = Pantalla.ListIndex
            WLote.Text = WIndice.List(Indice)
            WLote.SetFocus
            
        Case Else
    End Select
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5
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
ReDim UserData(0 To 6, 0 To 40)

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
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Prod.Terminado"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Materia Prima"
             DBGrid1.Columns(newcnt).Width = 1500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3620
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Movimiento"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
             
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 2
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Codigo.Text = ""
    WCodigoMov = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
 
    Rem With rstMovguia
    Rem     .Index = "Clave"
    Rem    Claveven$ = "99999999"
    Rem    .Seek "<=", Claveven$
    Rem    If .NoMatch = False Then
    Rem        Codigo.Text = !Codigo + 1
    Rem            Else
    Rem        Codigo.Text = ""
    Rem    End If
    Rem End With
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgPrestamo.Caption = "Ingreso de Prestamos entre Plantas :  " + !Nombre
        End If
    End With
    
    Tipomov.Clear
    
    Tipomov.AddItem "Prestamos entre Plantas"
    
    Tipomov.ListIndex = 0
    
    Destino.Clear
    
    Destino.AddItem ""
    Destino.AddItem "Prestamo a Surfactan"
    Destino.AddItem "Prestamo a Pellital"
    Destino.AddItem "Prestamo a Surfactan II"
    Destino.AddItem "Prestamo a Pellital II"
    Destino.AddItem "Prestamo a Surfactan III"
    Destino.AddItem "Prestamo a Surfactan IV"
    Destino.AddItem "Prestamo a Surfactan V"
    Destino.AddItem "Prestamo a Pellital V"
    Destino.AddItem "Prestamo a Pellital IV"
    Destino.AddItem "Prestamo a Surfactan VI"
    Destino.AddItem "Prestamo a Surfactan VII"
    
    Destino.ListIndex = 0
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 1)
    
    
    
    Rem spMovguia = "ListaMovguiaNumero " + "'" + WTipomov + "'"
    Rem Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstMovguia.RecordCount > 0 Then
    Rem     With rstMovguia
    Rem         .MoveLast
    Rem         Codigo.Text = rstMovguia!Codigo + 1
    Rem         If Val(Codigo.Text) < 900000 Then
    Rem             Codigo.Text = "1"
    Rem                 Else
    Rem             Codigo.Text = Str$(Val(Right$(Codigo.Text, 5)))
    Rem         End If
    Rem     End With
    Rem     rstMovguia.Close
    Rem         Else
    Rem     Codigo.Text = "1"
    Rem End If
    
    Codigo.Text = ""
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Graba1.Enabled = True
    Pass.Visible = False
    
 Rem BY NAN AGREGADO
    
   ZSql = ""
    ZSql = ZSql + "Select prestamo.Codigo"
    ZSql = ZSql + " FROM prestamo"
    ZSql = ZSql + " Order by prestamo.codigo"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveLast
            Codigo.Text = Val(rstOrden!Codigo) + 1
        End With
        rstOrden.Close
    End If
    
    
    
    
    
Rem FIN BY NAN
    Codigo.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Erase Auxiliar
    Renglon = 0
    
    WTipomov = Str$(Tipomov.ListIndex)
    Call Ceros(WTipomov, 1)
    
    XParam = "'" + WTipomov + "','" _
                + WCodigoMov + "'"
    spMovguia = "ListaMovguia " + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstMovguia!Tipo
                
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstMovguia!Terminado
                    Auxi1 = rstMovguia!Terminado
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = rstMovguia!Articulo
                    Auxi2 = rstMovguia!Articulo
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", (rstMovguia!Cantidad))
                
                    DBGrid1.Col = 5
                    DBGrid1.Text = rstMovguia!Movi
                    
                    DBGrid1.Col = 6
                    DBGrid1.Text = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                    
                    Tipomov.ListIndex = Val(rstMovguia!Tipomov)
                    Destino.ListIndex = Val(rstMovguia!Destino)
                    Observaciones.Text = rstMovguia!Observaciones
                    Fecha.Text = rstMovguia!Fecha
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = Auxi2
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0

    For Da = 1 To WRenglon
    
        Auxi1 = Auxiliar(Da, 1)
        Auxi2 = Auxiliar(Da, 2)
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DBGrid1.Col = 3
            DBGrid1.Text = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 3
            DBGrid1.Text = rstArticulo!Descripcion
            rstArticulo.Close
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
    
    WTipo.SetFocus

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
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
            DBGrid1.Col = 5
            DBGrid1.Text = WMovi.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WTipo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WTerminado.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 4
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
            DBGrid1.Col = 5
            DBGrid1.Text = WMovi.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WTipomov = Str$(Tipomov.ListIndex)
        Call Ceros(WTipomov, 1)
        
        Auxi = Codigo.Text
        Call Ceros(Auxi, 5)
        WCodigoMov = "9" + Auxi
    
        XParam = "'" + WTipomov + "','" _
                + WCodigoMov + "'"
        spMovguia = "ListaMovguia " + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
        If rstMovguia.RecordCount > 0 Then
            Fecha.Text = rstMovguia!Fecha
            rstMovguia.Close
            Graba1.Enabled = False
            Call Proceso_Click
                Else
            WCodigo = Codigo.Text
            WCodigo1 = WCodigoMov
            Call Limpia_Click
            Codigo.Text = WCodigo
            WCodigoMov = WCodigo1
            Graba1.Enabled = True
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub



Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo.SetFocus
    End If
End Sub

Sub Impresion()

        If Val(WEmpresa) = 1 Then
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
                Else
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        End If
  
        Rem  #1, 255

        For FF = 1 To 2

        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Print #1, ""
        Print #1, Tab(48); "GUIA DE TRASLADO INTERNO"
        Print #1, ""
        Print #1, ""
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Select Case Val(WDestino)
            Case 1, 3, 5, 6, 7, 10, 11
                Print #1, Tab(7); "Surfactan"
                Print #1, Tab(7); "Malvinas Argentinas"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As. C.P.1414"
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-11111111-2"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega";
                Print #1, ""
            Case Else
                Print #1, Tab(7); "Pellital"
                Print #1, Tab(7); "Malvinas Argentinas"
                Print #1, Tab(7); "Victoria"
                Print #1, Tab(7); "Pcia. Bs.As. C.P.1414"
                Print #1, ""
                Print #1, Tab(7); "Inscripto";
                Print #1, Tab(48); "30-11111111-2"
                Print #1, ""
                Print #1, Tab(30); "Direccion Entrega";
                Print #1, ""
        End Select
                
        If FF = 1 Then
            Print #1, Tab(60); "ORIGINAL"
                Else
            Print #1, Tab(60); "DUPLICADO"
        End If
        Print #1, ""
        
        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 4
                Descri = DBGrid1.Text
            
                DBGrid1.Col = 3
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                        
                        Print #1, Tab(14); Left$(Descri, 40);
                        Print #1, Tab(58); Alinea("#####.##", Str$(Cantidad));
                        Print #1, " Kg";
                        Print #1, Tab(71); "Netos"
                        Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a
        
        For aa = Impre To 22
                Print #1, ""
        Next aa
        
        Print #1, ""
        Print #1, Tab(10); "Lugar de Pago : Ayacucho 1231 5to Piso Dto. 'A' Capital Federal"
        Print #1, ""
        
        For Da = 1 To 9
                Print #1, ""
        Next Da
        
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
        For xda = 2 To 4
                Print #1, ""
                Print #1, ""
        Next xda
        
        Print #1, ""
        Select Case XX
                Case 1
                        Print #1, Tab(10); "ORIGINAL";
                Case 2
                        Print #1, Tab(10); "DUPLICADO";
                Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                Case Else
        End Select
        Print #1, Tab(10); "Nro. Control : "; Codigo.Text
        Print #1, Chr$(12)

        Next FF

        Close #1

End Sub

Private Sub Graba1_Click()
    If Tipomov.ListIndex = 0 Or Tipomov.ListIndex = 2 Then
        WClave.Text = ""
        Pass.Visible = True
        WClave.SetFocus
            Else
        Call Graba_Click
    End If
End Sub

Private Sub WCancela_Click()
    Pass.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WClave.Text = "CLAVE04" Then
            Call Graba_Click
        End If
    End If
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Rem If Left$(Articulo1, 2) = "DW" Then
                        Rem  Tipo = "T"
                        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        Rem End If
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
            Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
            Rem     Renglon = Renglon + 1
            Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
            Rem     Auxiliar(Renglon, 2) = 1
            Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Auxiliar(Da, 2)
        XVector = Auxiliar(Da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
            rstArticulo.Close
        End If
    Next Da
    
End Sub


