VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgpag2 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Pagos a Proveedores"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   1155
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6750
   ScaleWidth      =   11880
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   7200
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton calcret 
      Caption         =   "Calc.Ret."
      Height          =   300
      Left            =   6000
      TabIndex        =   37
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Retencion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      TabIndex        =   36
      Text            =   " "
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Banco 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   33
      Text            =   " "
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame IngreCuenta 
      Caption         =   "Cuenta Contable"
      Height          =   855
      Left            =   2880
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Cuenta 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      Text            =   " "
      Top             =   720
      Width           =   5415
   End
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion"
      Height          =   300
      Left            =   6000
      TabIndex        =   23
      Top             =   1920
      Width           =   975
   End
   Begin Crystal.CrystalReport LISTADO 
      Left            =   6000
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ordpago.rpt"
      WindowTitle     =   "Orden de Pago"
      CopiesToPrinter =   2
      WindowState     =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Orden de Pago"
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   1440
      Width           =   3735
      Begin VB.OptionButton Tipo5 
         Caption         =   "Cheques Rechazados"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Tipo4 
         Caption         =   "Transferencias"
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Tipo3 
         Caption         =   "Pagos Varios"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Pagos de Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Anticipos"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Proveedor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      MaxLength       =   11
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   8160
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
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
      Height          =   3375
      Left            =   0
      OleObjectBlob   =   "pag2.frx":0000
      TabIndex        =   3
      Top             =   2640
      Width           =   11775
   End
   Begin VB.TextBox Orden 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2010
      ItemData        =   "pag2.frx":09C2
      Left            =   7200
      List            =   "pag2.frx":09C9
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4920
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   3840
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4920
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   6000
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Grabar"
      Height          =   300
      Left            =   3840
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retencion"
      Height          =   255
      Left            =   3840
      TabIndex        =   35
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label DesBanco 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2520
      TabIndex        =   34
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label bjm 
      Caption         =   "Banco"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   9960
      TabIndex        =   21
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Debitos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.   2) Bco.  3) Ch. Terc.  4) Documentos"
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Label DesProveedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   " "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro. Orden de Pago"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "Prgpag2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 10  ' Número máximo de campos del conjunto de registros.
Private Debito As Double
Private Credito As Double
Private WImpresion(12, 10) As String
Private WImpre2(12, 10) As String
Private WDebito(12, 2) As String
Private WCredito(12, 4) As String
Private WCuenta(12) As String
Private WCuentaBco As String
Private Numero As String
Private WNumero As String
Private WSaldo As Double
Private WRetencion As Double
Private WCuatri  As String
Private WEmpNombre As String
Private WEmpDirecion As String
Private WEmpLocalidad As String
Private WEmpCuit As String
Private WPrvDireccion As String
Private WPrvCuit As String
Private WLeyenda(10) As String
Private WTipo As String
Private WTipoprv As Single
Private WTipoiva As Single
Private WNeto As Double
Private WAnticipo As Double
Private WBruto As Double
Private WIva As Double
Private WRetenido As Double
Private WFecha As String
Private XNeto As Double
Private XBruto As Double
Private XIva As Double
Private XTBase As Double
Private XImpor As Double
Private WParametro(0 To 10) As Double
Private WTasa1(10) As Double
Private WAuxi As Double
Private Total As Double

Private Sub Suma_Datos()
    Debitos.Caption = ""
    Creditos.Caption = ""
    
    For iRow = 0 To 9
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Debitos.Caption = Str$(Val(Debitos.Caption) + Val(DbGrid1.Text))
        End If
        DbGrid1.Col = 11
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(DbGrid1.Text))
        End If
    Next iRow
    
    Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Retencion.Text))
    
    Debitos.Caption = Pusing("###,###.##", Debitos.Caption)
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
    DbGrid1.Col = 0
    DbGrid1.Row = 0
    
End Sub

Private Sub Lee_Datos()
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        With rstPagos
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", Orden.Text + Auxi1
            If .NoMatch = False Then
                Select Case Val(!Tiporeg)
                    Case 1
                        Debito = Debito + 1
                        DbGrid1.Row = Debito - 1
                        DbGrid1.Col = 0
                        DbGrid1.Text = !Tipo1
                        DbGrid1.Col = 1
                        DbGrid1.Text = !Letra1
                        DbGrid1.Col = 2
                        DbGrid1.Text = !Punto1
                        DbGrid1.Col = 3
                        DbGrid1.Text = !Numero1
                        DbGrid1.Col = 4
                        DbGrid1.Text = !Importe1
                        DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                        DbGrid1.Col = 5
                        DbGrid1.Text = !Observaciones2
                    Case 2
                        Credito = Credito + 1
                        DbGrid1.Row = Credito - 1
                        DbGrid1.Col = 6
                        DbGrid1.Text = !Tipo2
                        DbGrid1.Col = 7
                        DbGrid1.Text = !Numero2
                        DbGrid1.Col = 8
                        DbGrid1.Text = !Fecha2
                        DbGrid1.Col = 9
                        DbGrid1.Text = !Banco2
                        DbGrid1.Col = 10
                        If !Observaciones2 <> "" Then
                            DbGrid1.Text = !Observaciones2
                        End If
                        DbGrid1.Col = 11
                        DbGrid1.Text = !Importe2
                        DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    Case Else
                End Select
                    Else
                Exit Do
            End If
        End With
    Loop
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
    Rem Retganancias.text = PUsing("###,###.##", Retganancias.text)
End Sub

Sub Imprime_Datos()
    With rstProveedor
        .Index = "Proveedor"
        .Seek "=", Proveedor.Text
        If .NoMatch = False Then
            Proveedor.Text = !Proveedor
            DesProveedor.Caption = !Nombre
            WPrvDireccion = !Direccion
            WPrvCuit = !Cuit
            WTipoprv = Val(!Tipo) + 1
            WTipoiva = Val(!Iva)
            Call Format_datos
        End If
    End With
End Sub

Private Sub cmdAdd_Click()

    If Orden.Text <> "" And Fecha.Text <> "" Then
    
    If Proveedor.Text <> "" Or Tipo3.Value = True Or Tipo4.Value = True Or Tipo5.Value = True Then
    
    Auxi1 = Orden.Text
    Call Ceros(Auxi1, 6)
    Orden.Text = Auxi1
        
    With rstPagos
        Existe = "N"
        .Index = "Clave"
        Claveven$ = Orden.Text + "01"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Existe = "S"
        End If
    End With
    
    If Existe <> "Sd" Then
    
        Call Suma_Datos
        
        Debito = 0
        Credito = 0
        If Val(Debitos.Caption) <> 0 Then
            Debito = Val(Debitos.Caption)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = Val(Creditos.Caption)
        End If
        
        If Debito = Credito Then
    
        With rstPagos
            Renglon = 0
            .Index = "Clave"
            For iRow = 0 To 9
                WRow = iRow
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Then
                
                    DbGrid1.Col = 0
                    WTipo = Left$(DbGrid1.Text, 2)
                    DbGrid1.Col = 1
                    WLetra = Left$(DbGrid1.Text, 1)
                    DbGrid1.Col = 2
                    WPunto = Left$(DbGrid1.Text, 4)
                    DbGrid1.Col = 3
                    WNumero = Left$(DbGrid1.Text, 8)
                    DbGrid1.Col = 4
                    WImporte = Val(DbGrid1.Text)
                    
                    If Tipo1.Value = True Then
                        With rstCtaCtePrv
                            .Index = "CtaCte"
                            Claveven$ = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
                            .Seek "=", Claveven$
                            If .NoMatch = False Then
                                .Edit
                                !Saldo = !Saldo - WImporte
                                .Update
                            End If
                        End With
                    End If
                    
                End If
                
            Next iRow
        End With
        
        
        If Tipo1.Value = True Then
        
            WLetra = "A"
            WTipo = "04"
            WPunto = "0000"
            WNumero = Orden.Text
            WProveedor = Proveedor.Text
        
            Call Ceros(WNumero, 8)
            Rem Call Ceros(WProveedor, 6)
        
            With rstCtaCtePrv
                .Index = "CtaCte"
                .Seek "=", WProveedor + WLetra + WTipo + WPunto + WNumero
                If .NoMatch Then
                    .AddNew
                    !Proveedor = Proveedor.Text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = "  /  /    "
                    !Total = Debito * -1
                    !Saldo = 0
                    !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = "00000000"
                    !Impre = "OP"
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Proveedor = Proveedor.Text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = "  /  /    "
                    !Total = Debito * -1
                    !Saldo = 0
                    !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = "00000000"
                    !Impre = "OP"
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
        End If
        
        If Tipo2.Value = True Then
        
            WLetra = "A"
            WTipo = "05"
            WPunto = "0000"
            WNumero = Orden.Text
            WProveedor = Proveedor.Text
        
            Call Ceros(WNumero, 8)
            Rem Call Ceros(WProveedor, 6)
        
            With rstCtaCtePrv
                .Index = "CtaCte"
                .Seek "=", WProveedor + WLetra + WTipo + WPunto + WNumero
                If .NoMatch Then
                    .AddNew
                    !Proveedor = Proveedor.Text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = "  /  /    "
                    !Total = Debito * -1
                    !Saldo = Debito * -1
                    !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = "00000000"
                    !Impre = "AN"
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Proveedor = Proveedor.Text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = "  /  /    "
                    !Total = Debito * -1
                    !Saldo = Debito * -1
                    !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = "00000000"
                    !Impre = "AN"
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
        End If
        
        Orden.SetFocus
        Call CmdLimpiar_Click
        
        End If
        
        End If
        
    End If
    End If
End Sub

Private Sub cmdDelete_Click()
    If Orden.Text <> "" Then
                
            Rem Borro los datos anteriores
            
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Orden.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow

    End If
    Proveedor.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    For iCol = 0 To 12
        For iRow = 0 To 9
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iRow
    Next iCol
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    Debitos.Caption = ""
    Creditos.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    Orden.SetFocus
    
    With rstPagos
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Orden.Text = !Orden + 1
                Else
            Orden.Text = ""
        End If
    End With
    
    Pantalla.Visible = False
    Opcion.Visible = False
    IngreCuenta.Visible = False
    Erase WCuenta
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstProveedor
        .Close
    End With
    With rstRecibos
        .Close
    End With
    With rstPagos
        .Close
    End With
    With rstCtaCtePrv
        .Close
    End With
    Rem  With rstImputac
    Rem     .Close
    Rem  End With
    With rstBanco
        .Close
    End With
    With rstCuenta
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    DbsAdminis.Close
    Orden.SetFocus
    Prgpago.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub


Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        With rstPagos
            Existe = "N"
            .Index = "Clave"
            Claveven$ = Orden.Text + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Existe = "S"
                Proveedor.Text = !Proveedor
                Fecha.Text = !Fecha
                Retencion.Text = !Retencion
                Retencion.Text = Pusing("###,###.##", Retencion.Text)
                Tipo1.Value = False
                Tipo2.Value = False
                Tipo3.Value = False
                Tipo4.Value = False
                Tipo5.Value = False
                Select Case Val(!TipoOrd)
                    Case 1
                        Tipo1.Value = True
                    Case 2
                        Tipo2.Value = True
                    Case 3
                        Tipo3.Value = True
                    Case 4
                        Tipo4.Value = True
                    Case 5
                        Tipo5.Value = True
                    Case Else
                End Select
                Observaciones.Text = !Observaciones
                
            End If
        End With
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            If Tipo3.Value = True Or Tipo4.Value = True Or Tipo5.Value = True Then
                Observaciones.SetFocus
                    Else
                Proveedor.SetFocus
            End If
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
                    WPrvDireccion = !Direccion
                    WPrvCuit = !Cuit
                    WTipoprv = Val(!Tipo) + 1
                    WTipoiva = Val(!Iva)
                    Observaciones.SetFocus
                End If
            End With
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Tipo4.Value = True Then
            Banco.SetFocus
                Else
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
        End If
    End If
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Banco.Text) <> 0 Then
            With rstBanco
                .Index = "Banco"
                Claveven$ = Val(Banco.Text)
                .Seek "=", Val(Banco.Text)
                If .NoMatch Then
                    Banco.Text = Claveven$
                    Banco.SetFocus
                        Else
                    Banco.Text = !Banco
                    DesBanco.Caption = !Nombre
                    WCtabanco = !Cuenta
                    DbGrid1.Col = 0
                    DbGrid1.Row = 0
                    DbGrid1.SetFocus
                End If
            End With
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
    
     XRow = DbGrid1.Row
     XCol = DbGrid1.Col

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Cuenta Corrientes"
     Opcion.AddItem "Cheques terceros"
     Opcion.AddItem "Documentos"
     Opcion.AddItem "Cuentas Contables"

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
            Ayuda.Visible = True
            Ayuda.Text = ""
            With rstProveedor
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi$ = Mascara("###########", Str$(!Proveedor))
                        IngresaItem = Auxi$ + " " + !Nombre
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
            With rstCtaCtePrv
                .Index = "ClaveImpre"
                Auxi = Proveedor.Text
                .Seek ">", Auxi + Space$(100)
                If .NoMatch = False Then
                Do
                    If .EOF = False Then
                        If Proveedor.Text = !Proveedor Then
                            If !Saldo <> 0 Then
                                Auxi$ = Str$(!Saldo)
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                IngresaItem = !Impre + " " + !Letra + " " + !Punto + " " + !Numero + " " + !Fecha + " " + Auxi$
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                                Else
                        Exit Do
                    End If
                Loop
                End If
            End With
            
        Case 2
            With rstRecibos
                .Index = "Fecha2"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(!Tiporeg) = 2 Then
                            If Val(!Tipo2) = 2 And !Estado2 <> "X" Then
                                Auxi$ = Str$(!Importe2)
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                Numero = Str$(Val(!Numero2))
                                Call Ceros(Numero, 6)
                                IngresaItem = Numero + "  " + !Fecha2 + "  " + Auxi$ + "  " + !Banco2
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 3
            With rstCtaCte
                .Index = "Clave"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(!Tipo) = 50 Then
                            WSaldo = !Saldo
                            Call Redondeo(WSaldo)
                            Rem If !Numero = 5604 Then Stop
                            If WSaldo <> 0 And !Cliente <> Space$(6) Then
                                Auxi$ = Str$(Abs(!Saldo))
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                WNumero = !Numero
                                Call Ceros(WNumero, 6)
                                IngresaItem = WNumero + "  " + !Vencimiento1 + "  " + Auxi$ + "  " + !Cliente
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 4
            With rstCuenta
                .Index = "Cuenta"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Cuenta + "  " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cuenta
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

Private Sub Pantalla_Click()
    Select Case XIndice
        Case 0
            With rstProveedor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Proveedor.Text = Claveven$
                .Index = "Proveedor"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DesProveedor.Caption = !Nombre
                    WPrvDireccion = !Direccion
                    WPrvCuit = !Cuit
                    WTipoprv = Val(!Tipo) + 1
                    WTipoiva = Val(!Iva)
                            Else
                    Proveedor.Text = ""
                End If
            End With
                
            Ayuda.Visible = False
            Pantalla.Visible = False
            Proveedor.SetFocus
            
        Case 1
        
            If Tipo1.Value = True Then
            
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For iRow = 0 To 9
                DbGrid1.Row = iRow
                DbGrid1.Col = 1
                Compara2 = Proveedor.Text + DbGrid1.Text
                DbGrid1.Col = 0
                Compara2 = Compara2 + DbGrid1.Text
                DbGrid1.Col = 2
                Compara2 = Compara2 + DbGrid1.Text
                DbGrid1.Col = 3
                Compara2 = Compara2 + DbGrid1.Text
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next iRow
            
            If Entra = "S" Then
            
            For iRow = 0 To 9
                DbGrid1.Row = iRow
                DbGrid1.Col = 0
                If DbGrid1.Text = "" Then
                    XRow = DbGrid1.Row
                    Exit For
                End If
            Next iRow
        
            With rstCtaCtePrv

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "CtaCte"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 0
                    DbGrid1.Text = !Tipo
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 1
                    DbGrid1.Text = !Letra
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 2
                    DbGrid1.Text = !Punto
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 3
                    DbGrid1.Text = !Numero
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    DbGrid1.Text = !Saldo
                    DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 5
                    DbGrid1.Text = "Pago factura nro. " + Str$(!Numero)
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                End If
            End With
            
            End If
            
            End If
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 0
            DbGrid1.SetFocus
            
        Case 2
        
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For iRow = 0 To 9
                DbGrid1.Row = iRow
                DbGrid1.Col = 12
                Compara2 = DbGrid1.Text
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next iRow
            
            If Entra = "S" Then
            
            For iRow = 0 To 9
                DbGrid1.Row = iRow
                DbGrid1.Col = 6
                If DbGrid1.Text = "" Then
                    XRow = DbGrid1.Row
                    Exit For
                End If
            Next iRow
        
        
            With rstRecibos
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Clave"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                
                    DbGrid1.Col = 6
                    If XIndice = 2 Then
                        DbGrid1.Text = "3"
                            Else
                        DbGrid1.Text = "4"
                    End If
                    
                    DbGrid1.Col = 7
                    DbGrid1.Text = !Numero2
                
                    DbGrid1.Col = 8
                    DbGrid1.Text = !Fecha2
                
                    DbGrid1.Col = 9
                    DbGrid1.Text = ""
                
                    DbGrid1.Col = 10
                    DbGrid1.Text = !Banco2
                
                    DbGrid1.Col = 11
                    DbGrid1.Text = !Importe2
                    DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    
                    DbGrid1.Col = 12
                    DbGrid1.Text = Claveven$
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                    Pantalla.List(Indice) = ""
                    
                End If
                If DbGrid1.Row < 10 Then
                    DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 6
                    KeyCode = 0
                            Else
                    DbGrid1.Col = 6
                    KeyCode = 0
                End If
            End With
            
            End If
            
        Case 3
        
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For iRow = 0 To 9
                DbGrid1.Row = iRow
                DbGrid1.Col = 12
                Compara2 = DbGrid1.Text
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next iRow
            
            If Entra = "S" Then
            
            For iRow = 0 To 9
                DbGrid1.Row = iRow
                DbGrid1.Col = 6
                If DbGrid1.Text = "" Then
                    XRow = DbGrid1.Row
                    Exit For
                End If
            Next iRow
        
        
            With rstCtaCte
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Clave"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                
                    DbGrid1.Col = 6
                    DbGrid1.Text = "4"
                    
                    DbGrid1.Col = 7
                    DbGrid1.Text = !Numero
                
                    DbGrid1.Col = 8
                    DbGrid1.Text = !Vencimiento1
                
                    DbGrid1.Col = 9
                    DbGrid1.Text = ""
                
                    DbGrid1.Col = 10
                    DbGrid1.Text = ""
                
                    DbGrid1.Col = 11
                    DbGrid1.Text = !Saldo
                    DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    
                    DbGrid1.Col = 12
                    DbGrid1.Text = Claveven$
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                    Pantalla.List(Indice) = ""
                    
                End If
                If DbGrid1.Row < 10 Then
                    DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 6
                    KeyCode = 0
                            Else
                    DbGrid1.Col = 6
                    KeyCode = 0
                End If
            End With
            
            End If
            
        Case 4
            With rstCuenta
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Cuenta.Text = Claveven$
                .Index = "Cuenta"
                .Seek "=", Claveven$
            End With
                
            Pantalla.Visible = False
            Rem Cuenta.SetFocus
                
        Case Else
    End Select
    
End Sub
Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
    
            Case 0
                If KeyCode = 13 Then
                    If Tipo1.Value = True Then
                        If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Or Val(DbGrid1.Text) = 0 Then
                            Auxi$ = Str$(Val(DbGrid1.Text))
                            Call Ceros(Auxi$, 2)
                            DbGrid1.Text = Auxi$
                            DbGrid1.Col = 4
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                            Else
                        If Val(DbGrid1.Text) = 0 Then
                            Auxi$ = Str$(Val(DbGrid1.Text))
                            Call Ceros(Auxi$, 2)
                            DbGrid1.Text = Auxi$
                            DbGrid1.Col = 4
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End If
                End If
                
            Case 1
                If KeyCode = 13 Then
                    DbGrid1.Text = Left$(DbGrid1.Text, 1)
                    If Tipo1.Value = True Then
                        If DbGrid1.Text = "A" Or DbGrid1.Text = "C" Or DbGrid1.Text = "X" Then
                            DbGrid1.Col = 2
                            KeyCode = 0
                            Rem no hago anda
                                Else
                            DbGrid1.Col = 1
                            KeyCode = 0
                        End If
                            Else
                        DbGrid1.Col = 2
                        KeyCode = 0
                    End If
                End If
                
            Case 2
                If KeyCode = 13 Then
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 4)
                    DbGrid1.Text = Auxi$
                    DbGrid1.Col = 3
                    KeyCode = 0
                End If
                
            Case 3
                If KeyCode = 13 Then
                
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.Text = Auxi$
                
                    If Tipo1.Value = True Then
                        With rstCtaCtePrv
                            .Index = "CtaCte"
                            Claveven$ = Proveedor.Text
                            DbGrid1.Col = 1
                            Claveven$ = Claveven$ + DbGrid1.Text
                            DbGrid1.Col = 0
                            Claveven$ = Claveven$ + DbGrid1.Text
                            DbGrid1.Col = 2
                            Claveven$ = Claveven$ + DbGrid1.Text
                            DbGrid1.Col = 3
                            Claveven$ = Claveven$ + DbGrid1.Text
                        
                            .Seek "=", Claveven$
                            If .NoMatch = False Then
                                DbGrid1.Col = 4
                                XRow = DbGrid1.Row
                                If Val(DbGrid1.Text) = 0 Then
                                    DbGrid1.Text = !Saldo
                                    Call Suma_Datos
                                    DbGrid1.Col = 4
                                    DbGrid1.Row = XRow
                                End If
                                DbGrid1.Col = 4
                                KeyCode = 0
                                    Else
                                DbGrid1.Col = 0
                                KeyCode = 0
                            End If
                        End With
                            Else
                        DbGrid1.Col = 4
                        KeyCode = 0
                    End If
                    
                End If
                
            Case 4
                If KeyCode = 13 Then
                
                    If Tipo1.Value = True Then
                        With rstCtaCtePrv
                            .Index = "CtaCte"
                            Claveven$ = Proveedor.Text
                            DbGrid1.Col = 1
                            Claveven$ = Claveven$ + DbGrid1.Text
                            DbGrid1.Col = 0
                            Claveven$ = Claveven$ + DbGrid1.Text
                            DbGrid1.Col = 2
                            Claveven$ = Claveven$ + DbGrid1.Text
                            DbGrid1.Col = 3
                            Claveven$ = Claveven$ + DbGrid1.Text
                            .Seek "=", Claveven$
                            If .NoMatch = False Then
                                Saldo = !Saldo
                                    Else
                                Saldo = 0
                            End If
                        End With
                
                        DbGrid1.Col = 4
                        If Val(DbGrid1.Text) > Saldo Then
                            DbGrid1.Text = ""
                            DbGrid1.Col = 4
                            KeyCode = 0
                                Else
                            DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                            Call Suma_Datos
                            DbGrid1.Col = 5
                            KeyCode = 0
                        End If
                            Else
                        columna = DbGrid1.Row
                        DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                        Call Suma_Datos
                        DbGrid1.Col = 5
                        DbGrid1.Row = columna
                        KeyCode = 0
                    End If
                End If
                
            Case 5
                If KeyCode = 13 Then
                    If Tipo3.Value = True Then
                        Cuenta.Text = WCuenta(DbGrid1.Row)
                        IngreCuenta.Visible = True
                        Cuenta.SetFocus
                    End If
                    If Tipo4.Value = True Then
                        With rstBanco
                            .Index = "Banco"
                            Claveven$ = Val(Banco.Text)
                            .Seek "=", Val(Banco.Text)
                            If .NoMatch = False Then
                                WCuenta(DbGrid1.Row) = !Cuenta
                                    Else
                                WCuenta(DbGrid1.Row) = "999999"
                            End If
                        End With
                    End If
                    If Tipo5.Value = True Then
                        WCuenta(DbGrid1.Row) = "111"
                    End If
                    If DbGrid1.Row < 10 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 0
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If
                
            Case 6
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Or Val(DbGrid1.Text) = 4 Then
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.Text = Auxi$
                        
                        Select Case Val(DbGrid1.Text)
                        
                            Case 1
                                DbGrid1.Col = 7
                                DbGrid1.Text = ""
                                DbGrid1.Col = 8
                                DbGrid1.Text = ""
                                DbGrid1.Col = 9
                                DbGrid1.Text = ""
                                DbGrid1.Col = 10
                                DbGrid1.Text = ""
                                DbGrid1.Col = 11
                                DbGrid1.Columns(11).Locked = False
                                KeyCode = 0
                                
                            Case 3, 4
                                Call Consulta_Click
                                
                            Case Else
                                DbGrid1.Col = 7
                                KeyCode = 0
                                DbGrid1.Columns(7).Locked = False
                                
                        End Select
                        
                            Else
                            
                        DbGrid1.Col = 6
                        KeyCode = 0
                        
                    End If
                End If
                
            Case 7
                If KeyCode = 13 Then
                    DbGrid1.Col = 6
                    If Val(DbGrid1.Text) = 3 Or Val(DbGrid1.Text) = 4 Then
                        DbGrid1.Col = 7
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 7
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 8)
                        DbGrid1.Text = Auxi$
                        DbGrid1.Col = 8
                        KeyCode = 0
                        DbGrid1.Columns(7).Locked = True
                        DbGrid1.Columns(8).Locked = False
                    End If
                End If
                
            Case 8
                If KeyCode = 13 Then
                    DbGrid1.Col = 8
                    
                    Call Valida_fecha1(DbGrid1.Text, Auxi)
                    If Auxi <> "S" Then
                        DbGrid1.Col = 8
                        KeyCode = 0
                                Else
                        DbGrid1.Col = 9
                        KeyCode = 0
                        DbGrid1.Columns(8).Locked = True
                        DbGrid1.Columns(9).Locked = False
                    End If
                End If
                
            Case 9
                If KeyCode = 13 Then
                    With rstBanco
                        .Index = "Banco"
                        DbGrid1.Col = 9
                        Claveven$ = DbGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            DbGrid1.Col = 10
                            DbGrid1.Text = !Nombre
                            DbGrid1.Col = 11
                            KeyCode = 0
                            DbGrid1.Columns(9).Locked = True
                            DbGrid1.Columns(11).Locked = False
                                Else
                            DbGrid1.Col = 9
                            KeyCode = 0
                        End If
                    End With
                End If

            Case 11
                If KeyCode = 13 Then
                    iRow = DbGrid1.Row
                    DbGrid1.Col = 11
                    DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    Call Suma_Datos
                    DbGrid1.Row = iRow
                    If DbGrid1.Row < 10 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 6
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 6
                        KeyCode = 0
                    End If
                    DbGrid1.Columns(11).Locked = True
                End If

            Case Else
                
    End Select
    
End Sub
Private Sub DbGrid1_Keypress(KeyAscii As Integer)

    Select Case DbGrid1.Col
            Case 0, 2, 3, 4, 6, 7, 9, 11
                Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Case Else
                
    End Select
    
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta.Text <> "" Then
            With rstCuenta
                .Index = "Cuenta"
                Claveven$ = Cuenta.Text
                .Seek "=", Cuenta.Text
                If .NoMatch Then
                    Cuenta.SetFocus
                        Else
                    WCuenta(DbGrid1.Row - 1) = Cuenta.Text
                    IngreCuenta.Visible = False
                    Rem DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 0
                    KeyCode = 0
                    DbGrid1.SetFocus
                End If
            End With
        End If
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

ReDim UserData(0 To 9, 0 To 12)

mTotalRows& = 13

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
For i = 0 To 12
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Letra"
             DbGrid1.Columns(newcnt).Width = 450
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Punto"
             DbGrid1.Columns(newcnt).Width = 600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
        Case 5
             DbGrid1.Columns(newcnt).Caption = "Descripcion"
             DbGrid1.Columns(newcnt).Width = 2000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 6
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 7
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 8
             DbGrid1.Columns(newcnt).Caption = "Fecha"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 9
             DbGrid1.Columns(newcnt).Caption = "Banco"
             DbGrid1.Columns(newcnt).Width = 700
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 10
             DbGrid1.Columns(newcnt).Caption = "Nombre"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 11
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 12
             DbGrid1.Columns(newcnt).Caption = ""
             DbGrid1.Columns(newcnt).Width = 10
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
        Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
     
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    Debitos.Caption = ""
    Creditos.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    
    WLeyenda(1) = "Compra de Bienes"
    WLeyenda(2) = "Ejericio Prof. Lib. c/Aj.Inf."
    WLeyenda(3) = "Alquileres y Arrendamientos"
    
    WParametro(0) = 0
    WParametro(1) = 10000
    WParametro(2) = 20000
    WParametro(3) = 40000
    WParametro(4) = 100000
    WParametro(5) = 10000000
    
    WTasa1(1) = 0.1
    WTasa1(2) = 0.135
    WTasa1(3) = 0.17
    WTasa1(4) = 0.205
    WTasa1(5) = 0.24
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    With rstProveedor
        .Index = "Nombre"
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                    
                    
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
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
    
    End If

End Sub



