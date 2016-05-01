VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgpago 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Pagos a Proveedores"
   ClientHeight    =   7290
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7290
   ScaleWidth      =   11880
   Begin VB.TextBox RetIva 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7800
      TabIndex        =   55
      Text            =   " "
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Paridad 
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
      Height          =   300
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   " "
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame IngreCarpeta 
      Caption         =   "Ingreso de Importes para Carpetas"
      Height          =   3255
      Left            =   8040
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton GrabaCarpeta 
         Caption         =   "Confirma"
         Height          =   375
         Left            =   840
         TabIndex        =   50
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Carpeta4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   48
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Carpeta3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   47
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Carpeta2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   46
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Carpeta1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   45
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Carpeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   44
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   960
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton CargaCarpeta 
      Caption         =   "Carpetas"
      Height          =   300
      Left            =   3840
      TabIndex        =   42
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox RetIb 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10560
      TabIndex        =   40
      Text            =   " "
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Limpia1 
      Caption         =   "Limpia Renglon"
      Height          =   300
      Left            =   2040
      TabIndex        =   39
      Top             =   2640
      Width           =   1695
   End
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
      Left            =   4920
      TabIndex        =   37
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Retencion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7800
      TabIndex        =   36
      Text            =   " "
      Top             =   2640
      Width           =   1215
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
      Top             =   3960
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
      Left            =   3840
      TabIndex        =   23
      Top             =   2640
      Width           =   975
   End
   Begin Crystal.CrystalReport LISTADO 
      Left            =   6720
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
      Begin VB.OptionButton Tipo6 
         Caption         =   "Aplic.Pago Impo."
         Height          =   255
         Left            =   2040
         TabIndex        =   52
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Tipo5 
         Caption         =   "Cheques Rechazados"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1935
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
      Height          =   3255
      Left            =   0
      OleObjectBlob   =   "pago.frx":0000
      TabIndex        =   3
      Top             =   3360
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
      ItemData        =   "pago.frx":09C2
      Left            =   7200
      List            =   "pago.frx":09C9
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
      Left            =   240
      TabIndex        =   8
      Top             =   2640
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
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ret.Iva"
      Height          =   255
      Left            =   6120
      TabIndex        =   56
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Paridad"
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
      Left            =   4920
      TabIndex        =   54
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Dife 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   10320
      TabIndex        =   51
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ret.Ing.Brutos"
      Height          =   255
      Left            =   9120
      TabIndex        =   41
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retencion Ganan."
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   2640
      Width           =   1575
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
      Left            =   10320
      TabIndex        =   21
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Debitos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"pago.frx":09D7
      Height          =   495
      Left            =   4200
      TabIndex        =   19
      Top             =   6720
      Width           =   6135
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
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Prgpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 12  ' Número máximo de campos del conjunto de registros.
Private Debito As Double
Private Credito As Double
Private WImpresion(12, 10) As String
Private WImpre2(12, 10) As String
Private WDebito(12, 2) As String
Private WCredito(12, 4) As String
Private WCuenta(12, 2) As String
Private WCuentaBco As String
Private Numero As String
Private WNumero As String
Private WSaldo As Double
Private WSaldoUs As Double
Private WRetencion As Double
Private WRetIb As Double
Private WRetIva As Double
Private WSumaIva As Double
Private WDife As Double
Private WCuatri  As String
Private WEmpNombre As String
Private WEmpDirecion As String
Private WEmpLocalidad As String
Private WEmpCuit As String
Private WPrvDireccion As String
Private WPrvCuit As String
Private WPrvIb As String
Private WLeyenda(10) As String
Private WTipo As String
Private WTipoprv As Single
Private WTipoiva As Single
Private WTipoIb As Single
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
Private WAuxi1 As Double
Private Total As Double
Private WRete1 As Double
Private WRete2 As Double
Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim rstCtaCte As Recordset
Dim spCtaCte As String
Dim rstPagos As Recordset
Dim spPagos As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim rstRetencion As Recordset
Dim spRetencion As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstNumero As Recordset
Dim spNumero As String
Dim XParam As String
Dim WProceso As Integer
Dim WCerti As String
Dim WCerificado As Integer
Dim ImpreCopia(10) As String
Dim WRete As Double
Dim WImpoRetenido As Double
Dim XImpre1 As String
Dim XImpre2 As String
Dim XImpre3 As String
Dim XImpre4 As String
Dim WImpre4 As Double
Dim SumaCarpeta As Double
Dim WImpo1 As Double
Dim WImpo2 As Double
Dim WCertificadoGan As Integer
Dim WCertificadoIb As Integer
Dim WCertificadoIva As Integer
Dim Deuda(1000, 10) As String
Dim XNroInterno As String
Dim WTipoDife As String
Dim WLetraDife As String
Dim WPuntoDife As String
Dim WNumeroDife As String
Dim WNetoDife As Double
Dim WIvaDife As Double
Dim RenglonDife As Integer
Dim ParidadTotal As Double
Dim ZFecha As String


Private Sub Suma_Datos()
    Debitos.Caption = ""
    Creditos.Caption = ""
    Dife.Caption = ""
    
    For iRow = 0 To 9
        DbGrid1.Col = 0
        DbGrid1.Row = iRow
        WTipo = DbGrid1.Text
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            If Tipo1.Value = True Then
                If Val(WTipo) <> 0 Then
                    Debitos.Caption = Str$(Val(Debitos.Caption) + Val(DbGrid1.Text))
                End If
                    Else
                Debitos.Caption = Str$(Val(Debitos.Caption) + Val(DbGrid1.Text))
            End If
        End If
        DbGrid1.Col = 11
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(DbGrid1.Text))
        End If
    Next iRow
    
    If Existe <> "S" Then
        Call calcret_Click
        Call CalcRetIb
    End If
    Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Retencion.Text) + Val(RetIb.Text) + Val(RetIva.Text))
    
    WDife = Val(Debitos.Caption) - Val(Creditos.Caption)
    Dife.Caption = Str$(WDife)
    
    Debitos.Caption = Pusing("#,###,###.##", Debitos.Caption)
    Creditos.Caption = Pusing("#,###,###.##", Creditos.Caption)
    Dife.Caption = Pusing("#,###,###.##", Dife.Caption)
    
    DbGrid1.Col = 0
    DbGrid1.Row = 0
    
End Sub

Private Sub Lee_Datos()
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        ClavePagos = Orden.Text + Auxi1
    
        spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
            Select Case Val(rstPagos!Tiporeg)
                Case 1
                    Debito = Debito + 1
                    DbGrid1.Row = Debito - 1
                    DbGrid1.Col = 0
                    DbGrid1.Text = rstPagos!Tipo1
                    DbGrid1.Col = 1
                    DbGrid1.Text = rstPagos!Letra1
                    DbGrid1.Col = 2
                    DbGrid1.Text = rstPagos!Punto1
                    DbGrid1.Col = 3
                    DbGrid1.Text = rstPagos!Numero1
                    DbGrid1.Col = 4
                    DbGrid1.Text = rstPagos!Importe1
                    DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    DbGrid1.Col = 5
                    DbGrid1.Text = rstPagos!Observaciones2
                Case 2
                    Credito = Credito + 1
                    DbGrid1.Row = Credito - 1
                    DbGrid1.Col = 6
                    DbGrid1.Text = rstPagos!Tipo2
                    DbGrid1.Col = 7
                    DbGrid1.Text = rstPagos!Numero2
                    DbGrid1.Col = 8
                    DbGrid1.Text = rstPagos!Fecha2
                    DbGrid1.Col = 9
                    DbGrid1.Text = rstPagos!Banco2
                    DbGrid1.Col = 10
                    If rstPagos!Observaciones2 <> "" Then
                        DbGrid1.Text = rstPagos!Observaciones2
                    End If
                    DbGrid1.Col = 11
                    DbGrid1.Text = rstPagos!Importe2
                    DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                Case Else
            End Select
            rstPagos.Close
                Else
            Exit Do
        End If
    Loop
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
    Rem Retganancias.text = PUsing("#,###,###.##", Retganancias.text)
End Sub

Sub Imprime_Datos()

    If Val(Banco.Text) <> 0 Then
        spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            DesBanco.Caption = rstBanco!Nombre
            rstBanco.Close
        End If
    End If

    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        Proveedor.Text = RstProveedor!Proveedor
        DesProveedor.Caption = RstProveedor!Nombre
        WPrvDireccion = RstProveedor!Direccion
        WPrvCuit = RstProveedor!Cuit
        WPrvIb = RstProveedor!NroIb
        WTipoprv = Val(RstProveedor!Tipo) + 1
        WTipoiva = Val(RstProveedor!Iva)
        WTipoIb = RstProveedor!CodIb
        RstProveedor.Close
        Call Format_datos
    End If
    
End Sub

Private Sub cmdAdd_Click()


    If Fecha.Text <> "" Then
    
    If Proveedor.Text <> "" Or Tipo3.Value = True Or Tipo4.Value = True Or Tipo5.Value = True Then
    
    If Tipo4.Value = True And Val(Banco.Text) = 0 Then
        m$ = "No se ha informado el banco al cual se va a realizar al transferencia"
        A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
        Exit Sub
    End If
    
    If Tipo1.Value = False And Tipo2.Value = False Then
        If Val(Proveedor.Text) <> 0 Then
            m$ = "Solo se puede informar proveedor en las ordenes de pago de Pagos o Anticipos de Proveedores"
            A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
            Exit Sub
        End If
    End If
    
    Auxi1 = Orden.Text
    Call Ceros(Auxi1, 6)
    Orden.Text = Auxi1
    
    Existe = "N"
    
    ClavePagos = Orden.Text + "01"
    spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        Existe = "S"
        rstPagos.Close
    End If
    
    If Existe <> "S" Then
    
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
        
            If Val(Orden.Text) = 0 Then
                spPagos = "ListaPagosNumero"
                Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                If rstPagos.RecordCount > 0 Then
                    With rstPagos
                        .MoveLast
                        Orden.Text = rstPagos!Orden + 1
                        Auxi1 = Orden.Text
                        Call Ceros(Auxi1, 6)
                        Orden.Text = Auxi1
                    End With
                    rstPagos.Close
                End If
            End If
        
            For iRow = 0 To 9
            
                WRow = iRow
                
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Then
                    If Tipo3.Value = True Then
                        If WCuenta(iRow, 1) = "" Then
                            m$ = "No se ha imputado correctamente el concepto del pago"
                            A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos Varios")
                            Exit Sub
                        End If
                    End If
                End If
                
                DbGrid1.Col = 6
                DbGrid1.Row = iRow
                ZTipo = DbGrid1.Text
                If Val(ZTipo) = 2 Then
                    DbGrid1.Col = 8
                    DbGrid1.Row = iRow
                    ZFecha = DbGrid1.Text
                    Call Valida_fecha1(ZFecha, Auxi)
                    If Auxi = "S" And Len(ZFecha) = 10 Then
                        ZOrdFecha1 = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        ZOrdFecha2 = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
                        If ZOrdFecha2 < ZOrdFecha1 Then
                            m$ = "La Fecha de los valores informados no puede ser menor a la fecha de emision de la orden de pago"
                            A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos")
                            Exit Sub
                        End If
                            Else
                        m$ = "La Fecha de los valores informados es incorrecta"
                        A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos")
                        Exit Sub
                    End If
                End If
                
            Next iRow
            
            Rem If WTipoprv = 5 Then
            Rem
            Rem     SumaCarpeta = 0
            Rem
            Rem     spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
            Rem     Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstMovgas.RecordCount > 0 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta.Text)
            Rem         rstMovgas.Close
            Rem     End If
            Rem
            Rem     If Val(Carpeta.Text) = 999999 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta.Text)
            Rem     End If
            Rem
            Rem     spMovgas = "ListaMovgas " + "'" + Carpeta1.Text + "'"
            Rem     Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstMovgas.RecordCount > 0 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta1.Text)
            Rem         rstMovgas.Close
            Rem     End If
            Rem
            Rem     If Val(Carpeta1.Text) = 999999 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta1.Text)
            Rem     End If
            Rem
            Rem     spMovgas = "ListaMovgas " + "'" + Carpeta2.Text + "'"
            Rem     Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstMovgas.RecordCount > 0 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta2.Text)
            Rem         rstMovgas.Close
            Rem     End If
            Rem
            Rem     If Val(Carpeta2.Text) = 999999 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta2.Text)
            Rem     End If
            Rem
            Rem     spMovgas = "ListaMovgas " + "'" + Carpeta3.Text + "'"
            Rem     Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstMovgas.RecordCount > 0 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta3.Text)
            Rem         rstMovgas.Close
            Rem     End If
            Rem
            Rem     If Val(Carpeta3.Text) = 999999 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta3.Text)
            Rem     End If
            Rem
            Rem     spMovgas = "ListaMovgas " + "'" + Carpeta4.Text + "'"
            Rem     Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstMovgas.RecordCount > 0 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta4.Text)
            Rem         rstMovgas.Close
            Rem     End If
            Rem
            Rem     If Val(Carpeta4.Text) = 999999 Then
            Rem         SumaCarpeta = SumaCarpeta + Val(ImpoCarpeta4.Text)
            Rem     End If
            Rem
            Rem     WImpo1 = SumaCarpeta
            Rem     WImpo2 = Debito
            Rem     Call Redondeo(WImpo1)
            Rem     Call Redondeo(WImpo2)
            Rem
            Rem     If WImpo1 <> WImpo2 Then
            Rem         If Tipo6.Value <> True Then
            Rem             m$ = "Error en la asignacion de la orden de pago a carpetas de importacion"
            Rem             A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos")
            Rem             Exit Sub
            Rem         End If
            Rem     End If
            Rem
            Rem End If
            
            WCertificadoGan = 0
            WCertificadoIb = 0
            WCertificadoIva = 0
            
            If Val(Retencion.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Numero"
                ZSql = ZSql + " Where Numero.Codigo = " + "'" + "91" + "'"
                spNumero = ZSql
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                If rstNumero.RecordCount > 0 Then
                    WCertificadoGan = rstNumero!Numero + 1
                    rstNumero.Close
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Numero SET "
                    ZSql = ZSql + " Numero = Numero + 1"
                    ZSql = ZSql + " Where Codigo = " + "'" + "91" + "'"
                    spNumero = ZSql
                    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                Rem  Open WEmpresa + "nro.txt" For Input As #10
                Rem Input #10, WCerti
                Rem Input #10, WCerti2
                Rem  Input #10, WCerti3
                Rem  Close #10
    
                Rem  WCertificadoGan = Val(WCerti) + 1
                Rem Open WEmpresa + "nro.txt" For Output As #10
                Rem  Print #10, WCertificadoGan
                Rem   Print #10, WCerti2
                Rem   Print #10, WCerti3
                Rem    Close #10
            End If
            
            If Val(RetIb.Text) <> 0 Then
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Numero"
                ZSql = ZSql + " Where Numero.Codigo = " + "'" + "92" + "'"
                spNumero = ZSql
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                If rstNumero.RecordCount > 0 Then
                    WCertificadoIb = rstNumero!Numero + 1
                    rstNumero.Close
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Numero SET "
                    ZSql = ZSql + " Numero = Numero + 1"
                    ZSql = ZSql + " Where Codigo = " + "'" + "92" + "'"
                    spNumero = ZSql
                    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                Rem Open WEmpresa + "nro.txt" For Input As #10
                Rem Input #10, WCerti1
                Rem Input #10, WCerti2
                Rem Input #10, WCerti3
                Rem Close #10
    
                Rem WCertificadoIb = Val(WCerti2) + 1
                Rem Open WEmpresa + "nro.txt" For Output As #10
                Rem Print #10, WCerti1
                Rem Print #10, WCertificadoIb
                Rem Print #10, WCerti3
                Rem Close #10
            End If
            
            If Val(RetIva.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Numero"
                ZSql = ZSql + " Where Numero.Codigo = " + "'" + "93" + "'"
                spNumero = ZSql
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                If rstNumero.RecordCount > 0 Then
                    WCertificadoIva = rstNumero!Numero + 1
                    rstNumero.Close
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Numero SET "
                    ZSql = ZSql + " Numero = Numero + 1"
                    ZSql = ZSql + " Where Codigo = " + "'" + "93" + "'"
                    spNumero = ZSql
                    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
               Rem Open WEmpresa + "nro.txt" For Input As #10
                Rem Input #10, WCerti1
                Rem Input #10, WCerti2
                Rem Input #10, WCerti3
                Rem Close #10
    
                Rem WCertificadoIva = Val(WCerti3) + 1
                Rem Open WEmpresa + "nro.txt" For Output As #10
                Rem Print #10, WCerti1
                Rem Print #10, WCerti2
                Rem Print #10, WCertificadoIva
                Rem  Close #10
            End If
            
            Renglon = 0
            For iRow = 0 To 9
                WRow = iRow
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Then
                    
                    DbGrid1.Col = 3
                    XNumero1 = DbGrid1.Text
                    If XNumero1 = "99999999" Then
                        DbGrid1.Col = 0
                        WTipoDife = Left$(DbGrid1.Text, 2)
                        DbGrid1.Col = 1
                        WLetraDife = Left$(DbGrid1.Text, 1)
                        DbGrid1.Col = 2
                        WPuntoDife = Left$(DbGrid1.Text, 4)
                        Select Case WLetraDife
                            Case "A"
                                DbGrid1.Col = 4
                                WNetoDife = Val(DbGrid1.Text) / 1.21
                                Call Redondeo(WNetoDife)
                                WIvaDife = Val(DbGrid1.Text) - WNetoDife
                                Call Redondeo(WIvaDife)
                            Case Else
                                DbGrid1.Col = 4
                                WNetoDife = Val(DbGrid1.Text)
                                Call Redondeo(WNetoDife)
                                WIvaDife = 0
                        End Select
                        Call Alta_Dife
                        DbGrid1.Col = 3
                        DbGrid1.Text = WNumeroDife
                    End If
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    XOrden = Orden.Text
                    XRenglon = Auxi1
                    XProveedor = Proveedor.Text
                    XFecha = Fecha.Text
                    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XImporte = Str$(Debito)
                    XRetencion = Retencion.Text
                    XRetotra = RetIb.Text
                    XRetIva = RetIva.Text
                    XObservaciones = Observaciones.Text
                    XCuenta = ""
                    If Tipo1.Value = True Then
                        XTipoOrd = "1"
                    End If
                    If Tipo2.Value = True Then
                        XTipoOrd = "2"
                    End If
                    If Tipo3.Value = True Then
                        XTipoOrd = "3"
                        XCuenta = WCuenta(iRow, 1)
                    End If
                    If Tipo4.Value = True Then
                        XTipoOrd = "4"
                        spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
                        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                        If rstBanco.RecordCount > 0 Then
                            XCuenta = rstBanco!Cuenta
                            rstBanco.Close
                                Else
                            XCuenta = "999999"
                        End If
                    End If
                    If Tipo5.Value = True Then
                        XTipoOrd = "5"
                        XCuenta = "111"
                    End If
                    If Tipo6.Value = True Then
                        XTipoOrd = "6"
                    End If
                    
                    XTiporeg = "1"
                    DbGrid1.Col = 0
                    XTipo1 = Left$(DbGrid1.Text, 2)
                    DbGrid1.Col = 1
                    XLetra1 = Left$(DbGrid1.Text, 1)
                    DbGrid1.Col = 2
                    XPunto1 = Left$(DbGrid1.Text, 4)
                    DbGrid1.Col = 3
                    XNumero1 = Left$(DbGrid1.Text, 8)
                    DbGrid1.Col = 4
                    XImporte1 = DbGrid1.Text
                    DbGrid1.Col = 5
                    XObservaciones2 = Left$(DbGrid1.Text, 30)
                    XTipo2 = ""
                    XNumero2 = ""
                    XFecha2 = ""
                    XFechaOrd2 = ""
                    If Tipo4.Value = True Then
                        XBanco2 = Banco.Text
                            Else
                        XBanco2 = ""
                    End If
                    XImporte2 = ""
                    XEmpresa = "1"
                    XClave = XOrden + XRenglon
                    XRetganancias = ""
                    XConcepto = ""
                    XConcecionaria = ""
                    XImpolist = ""
                    
                    XParam = "'" + XClave + "','" _
                            + XOrden + "','" + XRenglon + "','" _
                            + XProveedor + "','" _
                            + XFecha + "','" + XFechaOrd + "','" _
                            + XTipoOrd + "','" _
                            + XRetganancias + "','" _
                            + XRetIva + "','" + XRetotra + "','" _
                            + XRetencion + "','" _
                            + XTiporeg + "','" _
                            + XTipo1 + "','" + XLetra1 + "','" _
                            + XPunto1 + "','" + XNumero1 + "','" _
                            + XImporte1 + "','" _
                            + XTipo2 + "','" + XNumero2 + "','" _
                            + XFecha2 + "','" + XBanco2 + "','" _
                            + XImporte2 + "','" + XObservaciones2 + "','" _
                            + XEmpresa + "','" + XConcepto + "','" _
                            + XObservaciones + "','" _
                            + XImporte + "','" + XFechaOrd2 + "','" _
                            + XConcesionaria + "','" _
                            + XImpolist + "','" _
                            + XCuenta + "'"
                
                    spPagos = "AltaPagos " + XParam
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                    WLetra = XLetra1
                    WTipo = XTipo1
                    WPunto = XPunto1
                    WNumero = XNumero1
                    WImporte = XImporte1
                    
                    If Tipo1.Value = True Then
                        ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
                        spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                        Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        If RstCtaPrv.RecordCount > 0 Then
                            XSaldo = Str$(RstCtaPrv!Saldo - Val(WImporte))
                            XParam = "'" + ClaveCtaprv + "','" _
                                         + XSaldo + "'"
                            spCtaprv = "ActualizaCtaprvSaldo " + XParam
                            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    End If
                    
                End If
                
                DbGrid1.Col = 6
                DbGrid1.Row = iRow
                XTipo2 = Left$(DbGrid1.Text, 2)
                DbGrid1.Col = 11
                DbGrid1.Row = iRow
                If Val(DbGrid1.Text) <> 0 Or XTipo2 <> "" Then
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    XOrden = Orden.Text
                    XRenglon = Auxi1
                    XProveedor = Proveedor.Text
                    XFecha = Fecha.Text
                    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XImporte = Str$(Debito)
                    XRetencion = Retencion.Text
                    XRetotra = RetIb.Text
                    XRetIva = RetIva.Text
                    XObservaciones = Observaciones.Text
                    If Tipo1.Value = True Then
                        XTipoOrd = "1"
                    End If
                    If Tipo2.Value = True Then
                        XTipoOrd = "2"
                    End If
                    If Tipo3.Value = True Then
                        XTipoOrd = "3"
                    End If
                    If Tipo4.Value = True Then
                        XTipoOrd = "4"
                    End If
                    If Tipo5.Value = True Then
                        XTipoOrd = "5"
                    End If
                    If Tipo6.Value = True Then
                        XTipoOrd = "6"
                    End If
                    XTiporeg = "2"
                    XTipo1 = ""
                    XLetra1 = ""
                    XPunto1 = ""
                    XNumero1 = ""
                    XImporte1 = ""
                    DbGrid1.Col = 6
                    XTipo2 = Left$(DbGrid1.Text, 2)
                    DbGrid1.Col = 7
                    XNumero2 = Left$(DbGrid1.Text, 8)
                    DbGrid1.Col = 8
                    XFecha2 = Left$(DbGrid1.Text, 10)
                    XFechaOrd2 = Right$(XFecha2, 4) + Mid$(XFecha2, 4, 2) + Left$(XFecha2, 2)
                    DbGrid1.Col = 9
                    XBanco2 = DbGrid1.Text
                    DbGrid1.Col = 10
                    XObservaciones2 = Left$(DbGrid1.Text, 20)
                    DbGrid1.Col = 11
                    XImporte2 = DbGrid1.Text
                    DbGrid1.Col = 12
                    ClaveRecibos = DbGrid1.Text
                    ClaveCtacte = DbGrid1.Text
                    XEmpresa = "1"
                    XClave = XOrden + XRenglon
                    XRetganancias = ""
                    XConcepto = ""
                    XConcecionaria = ""
                    XImpolist = ""
                    XCuenta = ""
                    If Val(XTipo2) = 6 Then
                        XCuenta = WCuenta(iRow, 2)
                    End If
                    
                    XParam = "'" + XClave + "','" _
                            + XOrden + "','" + XRenglon + "','" _
                            + XProveedor + "','" _
                            + XFecha + "','" + XFechaOrd + "','" _
                            + XTipoOrd + "','" _
                            + XRetganancias + "','" _
                            + XRetIva + "','" + XRetotra + "','" _
                            + XRetencion + "','" _
                            + XTiporeg + "','" _
                            + XTipo1 + "','" + XLetra1 + "','" _
                            + XPunto1 + "','" + XNumero1 + "','" _
                            + XImporte1 + "','" _
                            + XTipo2 + "','" + XNumero2 + "','" _
                            + XFecha2 + "','" + XBanco2 + "','" _
                            + XImporte2 + "','" + XObservaciones2 + "','" _
                            + XEmpresa + "','" + XConcepto + "','" _
                            + XObservaciones + "','" _
                            + XImporte + "','" + XFechaOrd2 + "','" _
                            + XConcesionaria + "','" _
                            + XImpolist + "','" _
                            + XCuenta + "'"
                
                    spPagos = "AltaPagos " + XParam
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If Val(XTipo2) = 3 Then
                    
                        Rem spRecibos = "ConsultaRecibosClave " + "'" + ClaveRecibos + "'"
                        Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstRecibos.RecordCount > 0 Then
                        Rem     XEstado2 = "X"
                        Rem     XDestino = ""
                        Rem     XParam = "'" + ClaveRecibos + "','" _
                        rem                  + XEstado2 + "','" _
                        rem                  + XDestino + "'"
                        Rem     spRecibos = "ActualizaRecibos " + XParam
                        Rem     Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        Rem End If
                        
                        XEstado2 = "X"
                        XDestino = ""
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Recibos SET "
                        ZSql = ZSql + " Estado2 = " + "'" + XEstado2 + "',"
                        ZSql = ZSql + " Destino = " + "'" + XDestino + "'"
                        ZSql = ZSql + " Where Numero2 = " + "'" + XNumero2 + "'"
                        ZSql = ZSql + " and Importe2 = " + "'" + XImporte2 + "'"
                        
                        spRecibos = ZSql
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    If Val(XTipo2) = 4 Then
                        spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtaCte.RecordCount > 0 Then
                            XSaldo = ""
                            XSaldoUs = ""
                            XEstado = "1"
                            XDate = Date$
                            rstCtaCte.Close
                            XParam = "'" + ClaveCtacte + "','" _
                                         + XSaldo + "','" _
                                         + XSaldoUs + "','" _
                                         + XEstado + "','" _
                                         + XDate + "'"
                            spCtaCte = "ActualizaCtacte " + XParam
                            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    End If
                    
                End If
                
            Next iRow
            
            XParam = "'" + Orden.Text + "','" _
                    + Paridad.Text + "'"
            spPagos = "ModificaPagosParidad " + XParam
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        
            With rstEmpresa
                .Index = "Empresa"
                Claveven$ = "1"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WCtaProveedor = !CtaProveedores
                    WCtaEfectivo = !CtaEfectivo
                    WCtaCheques = !CtaCheque
                End If
            End With
        
            If Tipo1.Value = True Then
        
                WLetra = "A"
                WTipo = "04"
                WPunto = "0000"
                WNumero = Orden.Text
                WProveedor = Proveedor.Text
        
                Call Ceros(WNumero, 8)
                Rem Call Ceros(WProveedor, 6)
        
                ClaveCtaprv = WProveedor + WLetra + WTipo + WPunto + WNumero
                spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                If RstCtaPrv.RecordCount > 0 Then
            
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = ""
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "OP"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    
                    XParam = "'" + XClave + "','" _
                        + XProveedor + "','" + XLetra + "','" _
                        + XTipo + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + XEstado + "','" _
                        + Xvencimiento + "','" + XVencimiento1 + "','" _
                        + XTotal + "','" _
                        + XSaldo + "','" _
                        + XOrdFecha + "','" + XOrdVencimiento + "','" _
                        + XImpre + "','" + XEmpresa + "','" _
                        + XSaldolist + "','" _
                        + XNroInterno + "','" + Xlista + "','" _
                        + XAcumulado + "'"
                    
                    spCtaprv = "ActualizaCtaCtePrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = ""
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "OP"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = ""
                    XPAgo = ""
                    
                    XParam = "'" + XClave + "','" _
                        + XProveedor + "','" + XLetra + "','" _
                        + XTipo + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + XEstado + "','" _
                        + Xvencimiento + "','" + XVencimiento1 + "','" _
                        + XTotal + "','" _
                        + XSaldo + "','" _
                        + XOrdFecha + "','" + XOrdVencimiento + "','" _
                        + XImpre + "','" + XEmpresa + "','" _
                        + XSaldolist + "','" _
                        + XNroInterno + "','" + Xlista + "','" _
                        + XAcumulado + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                    
                    spCtaprv = "AltaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
            End If
        
            If Tipo2.Value = True Then
        
                WLetra = "A"
                WTipo = "05"
                WPunto = "0000"
                WNumero = Orden.Text
                WProveedor = Proveedor.Text
        
                Call Ceros(WNumero, 8)
                Rem Call Ceros(WProveedor, 6)
            
                ClaveCtaprv = WProveedor + WLetra + WTipo + WPunto + WNumero
                spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                If RstCtaPrv.RecordCount > 0 Then
            
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = Str$(Debito * -1)
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "AN"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = ""
                    XPAgo = ""
                    
                    XParam = "'" + XClave + "','" _
                            + XProveedor + "','" + XLetra + "','" _
                            + XTipo + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + XEstado + "','" _
                            + Xvencimiento + "','" + XVencimiento1 + "','" _
                            + XTotal + "','" _
                            + XSaldo + "','" _
                            + XOrdFecha + "','" + XOrdVencimiento + "','" _
                            + XImpre + "','" + XEmpresa + "','" _
                            + XSaldolist + "','" _
                            + XNroInterno + "','" + Xlista + "','" _
                            + XAcumulado + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                        
                    spCtaprv = "ModificaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Else
                        
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = Str$(Debito * -1)
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "AN"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = ""
                    XPAgo = ""
                    
                    XParam = "'" + XClave + "','" _
                        + XProveedor + "','" + XLetra + "','" _
                        + XTipo + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + XEstado + "','" _
                        + Xvencimiento + "','" + XVencimiento1 + "','" _
                        + XTotal + "','" _
                        + XSaldo + "','" _
                        + XOrdFecha + "','" + XOrdVencimiento + "','" _
                        + XImpre + "','" + XEmpresa + "','" _
                        + XSaldolist + "','" _
                        + XNroInterno + "','" + Xlista + "','" _
                        + XAcumulado + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                    
                    spCtaprv = "AltaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
            End If
            
            If Tipo6.Value = True Then
            
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                XOrden = Orden.Text
                XRenglon = Auxi1
                XProveedor = Proveedor.Text
                XFecha = Fecha.Text
                XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                XImporte = ""
                XRetencion = Retencion.Text
                XRetotra = RetIb.Text
                XRetIva = RetIva.Text
                XObservaciones = Observaciones.Text
                XCuenta = ""
                XTipoOrd = "6"
                    
                XTiporeg = "1"
                XTipo1 = ""
                XLetra1 = ""
                XPunto1 = ""
                XNumero1 = ""
                XImporte1 = ""
                XObservaciones2 = "Aplicaicon de Pgos de Importacion"
                XTipo2 = ""
                XNumero2 = ""
                XFecha2 = ""
                XFechaOrd2 = ""
                XBanco2 = ""
                XImporte2 = ""
                XEmpresa = "1"
                XClave = XOrden + XRenglon
                XRetganancias = ""
                XConcepto = ""
                XConcecionaria = ""
                XImpolist = ""
                    
                XParam = "'" + XClave + "','" _
                            + XOrden + "','" + XRenglon + "','" _
                            + XProveedor + "','" _
                            + XFecha + "','" + XFechaOrd + "','" _
                            + XTipoOrd + "','" _
                            + XRetganancias + "','" _
                            + XRetIva + "','" + XRetotra + "','" _
                            + XRetencion + "','" _
                            + XTiporeg + "','" _
                            + XTipo1 + "','" + XLetra1 + "','" _
                            + XPunto1 + "','" + XNumero1 + "','" _
                            + XImporte1 + "','" _
                            + XTipo2 + "','" + XNumero2 + "','" _
                            + XFecha2 + "','" + XBanco2 + "','" _
                            + XImporte2 + "','" + XObservaciones2 + "','" _
                            + XEmpresa + "','" + XConcepto + "','" _
                            + XObservaciones + "','" _
                            + XImporte + "','" + XFechaOrd2 + "','" _
                            + XConcesionaria + "','" _
                            + XImpolist + "','" _
                            + XCuenta + "'"
                
                spPagos = "AltaPagos " + XParam
                Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)

            End If
        
            ClaveRetencion = WFecha + Proveedor.Text
            spRetencion = "ConsultaRetencion " + "'" + ClaveRetencion + "'"
            Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
            If rstRetencion.RecordCount > 0 Then
                XXNeto = Str$(rstRetencion!Neto + XNeto)
                XXRetenido = Str$(rstRetencion!Retenido + Val(Retencion.Text))
                XParam = "'" + ClaveRetencion + "','" + XXNeto + "','" _
                         + XXRetenido + "'"
                rstRetencion.Close
                spRetencion = "ActualizaRetencionPagos " + XParam
                Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
                        
            
            Rem XParam = "'" + Orden.Text + "','" _
            REM             + Carpeta.Text + "','" _
            REM             + ImpoCarpeta.Text + "','" _
            REM             + Carpeta1.Text + "','" _
            REM             + ImpoCarpeta1.Text + "','" _
            REM             + Carpeta2.Text + "','" _
            REM             + ImpoCarpeta2.Text + "','" _
            REM             + Carpeta3.Text + "','" _
            REM             + ImpoCarpeta3.Text + "','" _
            REM             + Carpeta4.Text + "','" _
            REM             + ImpoCarpeta4.Text + "'"
            Rem spPagos = "ModificaPagos " + XParam
            Rem Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Pagos SET "
            ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
            ZSql = ZSql + " Carpeta1 = " + "'" + Carpeta1.Text + "',"
            ZSql = ZSql + " Carpeta2 = " + "'" + Carpeta2.Text + "',"
            ZSql = ZSql + " Carpeta3 = " + "'" + Carpeta3.Text + "',"
            ZSql = ZSql + " Carpeta4 = " + "'" + Carpeta4.Text + "'"
            ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
            spPagos = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            
            
            Sql1 = "UPDATE Pagos SET "
            Sql2 = " CertificadoGan = " + "'" + Str$(WCertificadoGan) + "',"
            Sql3 = " CertificadoIb = " + "'" + Str$(WCertificadoIb) + "',"
            Sql4 = " CertificadoIva = " + "'" + Str$(WCertificadoIva) + "'"
            Sql5 = " Where Orden = " + "'" + Orden.Text + "'"
            spPagos = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        
            With rstEmpresa
                .Index = "Empresa"
                .Seek "=", Val(WEmpresa)
                If .NoMatch = False Then
                    WAuxiliar = !Nombre
                End If
            End With
        
            Call IMPREORDEN
            If Val(Retencion.Text) <> 0 Then
                Call Impreret
            End If
            If Val(RetIb.Text) <> 0 Then
                Call Impreretib
            End If
            If Val(RetIva.Text) <> 0 Then
                Call ImpreretIva
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
    Tipo6.Value = False
    Debitos.Caption = ""
    Creditos.Caption = ""
    Dife.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    RetIb.Text = ""
    RetIva.Text = ""
    WTipoprv = 0
    ParidadTotal = 0
    Existe = "N"
    
    Carpeta.Text = ""
    Carpeta1.Text = ""
    Carpeta2.Text = ""
    Carpeta3.Text = ""
    Carpeta4.Text = ""
    ImpoCarpeta.Text = ""
    ImpoCarpeta1.Text = ""
    ImpoCarpeta2.Text = ""
    ImpoCarpeta3.Text = ""
    ImpoCarpeta4.Text = ""
    
    spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambioAdm.RecordCount > 0 Then
        Paridad.Text = Str$(rstCambioAdm!Cambio)
        Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
        rstCambioAdm.Close
    End If
    
    Orden.SetFocus
    Orden.Text = ""
    
    Rem spPagos = "ListaPagosNumero"
    Rem Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPagos.RecordCount > 0 Then
    Rem     With rstPagos
    Rem         .MoveLast
    Rem         Orden.Text = rstPagos!Orden + 1
    Rem     End With
    Rem     rstPagos.Close
    Rem End If
    
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
    
    Orden.SetFocus
    Prgpago.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Command1_Click()
    Rem ClaveCtaprv = "0100000012"
    Rem spCtaprv = "BorrarCtaprv " + "'" + ClaveCtaprv + "'"
    Rem Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_ImprePago
    OPEN_FILE_ImpreRetIb
    OPEN_FILE_ImpreRetGan
End Sub

Private Sub Impresion_Click()

    Existe = "N"
    
    ClavePagos = Orden.Text + "01"
    spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        Existe = "S"
        rstPagos.Close
    End If
    
    If Existe = "S" Then

        WOrden = Orden.Text
        Call CmdLimpiar_Click
        Orden.Text = WOrden

        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        Existe = "N"
        
        ClavePagos = Orden.Text + "01"
        spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
        
            Existe = "S"
            Proveedor.Text = rstPagos!Proveedor
            Fecha.Text = rstPagos!Fecha
            Retencion.Text = rstPagos!Retencion
            Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            RetIb.Text = rstPagos!RetOtra
            RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
            RetIva.Text = rstPagos!RetIva
            RetIva.Text = Pusing("#,###,###.##", RetIva.Text)
            Tipo1.Value = False
            Tipo2.Value = False
            Tipo3.Value = False
            Tipo4.Value = False
            Tipo5.Value = False
            Tipo6.Value = False
            Select Case Val(rstPagos!TipoOrd)
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
                Case 6
                    Tipo6.Value = True
                Case Else
            End Select
            Observaciones.Text = rstPagos!Observaciones
            
            Carpeta.Text = IIf(IsNull(rstPagos!Carpeta), "", rstPagos!Carpeta)
            Carpeta1.Text = IIf(IsNull(rstPagos!Carpeta1), "", rstPagos!Carpeta1)
            Carpeta2.Text = IIf(IsNull(rstPagos!Carpeta2), "", rstPagos!Carpeta2)
            Carpeta3.Text = IIf(IsNull(rstPagos!Carpeta3), "", rstPagos!Carpeta3)
            Carpeta4.Text = IIf(IsNull(rstPagos!Carpeta4), "", rstPagos!Carpeta4)
            
            Rem ImpoCarpeta.Text = IIf(IsNull(rstPagos!ImpoCarpeta), "", rstPagos!ImpoCarpeta)
            Rem ImpoCarpeta1.Text = IIf(IsNull(rstPagos!ImpoCarpeta1), "", rstPagos!ImpoCarpeta1)
            Rem ImpoCarpeta2.Text = IIf(IsNull(rstPagos!ImpoCarpeta2), "", rstPagos!ImpoCarpeta2)
            Rem ImpoCarpeta3.Text = IIf(IsNull(rstPagos!ImpoCarpeta3), "", rstPagos!ImpoCarpeta3)
            Rem ImpoCarpeta4.Text = IIf(IsNull(rstPagos!ImpoCarpeta4), "", rstPagos!ImpoCarpeta4)
            
            Rem ImpoCarpeta.Text = Pusing("#,###,###.##", ImpoCarpeta.Text)
            Rem ImpoCarpeta1.Text = Pusing("#,###,###.##", ImpoCarpeta1.Text)
            Rem ImpoCarpeta2.Text = Pusing("#,###,###.##", ImpoCarpeta2.Text)
            Rem ImpoCarpeta3.Text = Pusing("#,###,###.##", ImpoCarpeta3.Text)
            Rem ImpoCarpeta4.Text = Pusing("#,###,###.##", ImpoCarpeta4.Text)
            
            WCertificadoGan = IIf(IsNull(rstPagos!CertificadoGan), "0", rstPagos!CertificadoGan)
            WCertificadoIb = IIf(IsNull(rstPagos!CertificadoIb), "0", rstPagos!CertificadoIb)
            WCertificadoIva = IIf(IsNull(rstPagos!CertificadoIva), "0", rstPagos!CertificadoIva)
            
            rstPagos.Close
                
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            
            Call IMPREORDEN
            
            If Val(Retencion.Text) <> 0 Then
                WRetencion = Val(Retencion.Text)
                Call Impreret
            End If
            
            If Val(RetIb.Text) <> 0 Then
                WRetIb = Val(RetIb.Text)
                Call Impreretib
            End If
            
            If Val(RetIva.Text) <> 0 Then
                WRetIva = Val(RetIva.Text)
                Call ImpreretIva
            End If
            
        End If
    End If
    
End Sub

Private Sub Limpia1_Click()
        A = DbGrid1.Row
        B = DbGrid1.Col
        If B <= 5 Then
            DbGrid1.Col = 0
            DbGrid1.Text = ""
            DbGrid1.Col = 1
            DbGrid1.Text = ""
            DbGrid1.Col = 2
            DbGrid1.Text = ""
            DbGrid1.Col = 3
            DbGrid1.Text = ""
            DbGrid1.Col = 4
            DbGrid1.Text = ""
            DbGrid1.Col = 5
            DbGrid1.Text = ""
                Else
            DbGrid1.Col = 6
            DbGrid1.Text = ""
            DbGrid1.Col = 7
            DbGrid1.Text = ""
            DbGrid1.Col = 8
            DbGrid1.Text = ""
            DbGrid1.Col = 9
            DbGrid1.Text = ""
            DbGrid1.Col = 10
            DbGrid1.Text = ""
            DbGrid1.Col = 11
            DbGrid1.Text = ""
        End If
        Call Suma_Datos
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        Existe = "N"
        
        ClavePagos = Orden.Text + "01"
        spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
        
            Existe = "S"
            Proveedor.Text = rstPagos!Proveedor
            Fecha.Text = rstPagos!Fecha
            Retencion.Text = rstPagos!Retencion
            Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            RetIb.Text = rstPagos!RetOtra
            RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
            RetIva.Text = rstPagos!RetIva
            RetIva.Text = Pusing("#,###,###.##", RetIva.Text)
            Tipo1.Value = False
            Tipo2.Value = False
            Tipo3.Value = False
            Tipo4.Value = False
            Tipo5.Value = False
            Tipo6.Value = False
            Select Case Val(rstPagos!TipoOrd)
                Case 1
                    Tipo1.Value = True
                Case 2
                    Tipo2.Value = True
                Case 3
                    Tipo3.Value = True
                Case 4
                    Banco.Text = rstPagos!Banco2
                    Tipo4.Value = True
                Case 5
                    Tipo5.Value = True
                Case 6
                    Tipo6.Value = True
                Case Else
            End Select
            Observaciones.Text = rstPagos!Observaciones
            
            Carpeta.Text = IIf(IsNull(rstPagos!Carpeta), "", rstPagos!Carpeta)
            Carpeta1.Text = IIf(IsNull(rstPagos!Carpeta1), "", rstPagos!Carpeta1)
            Carpeta2.Text = IIf(IsNull(rstPagos!Carpeta2), "", rstPagos!Carpeta2)
            Carpeta3.Text = IIf(IsNull(rstPagos!Carpeta3), "", rstPagos!Carpeta3)
            Carpeta4.Text = IIf(IsNull(rstPagos!Carpeta4), "", rstPagos!Carpeta4)
            
            Rem ImpoCarpeta.Text = IIf(IsNull(rstPagos!ImpoCarpeta), "", rstPagos!ImpoCarpeta)
            Rem ImpoCarpeta1.Text = IIf(IsNull(rstPagos!ImpoCarpeta1), "", rstPagos!ImpoCarpeta1)
            Rem ImpoCarpeta2.Text = IIf(IsNull(rstPagos!ImpoCarpeta2), "", rstPagos!ImpoCarpeta2)
            Rem ImpoCarpeta3.Text = IIf(IsNull(rstPagos!ImpoCarpeta3), "", rstPagos!ImpoCarpeta3)
            Rem ImpoCarpeta4.Text = IIf(IsNull(rstPagos!ImpoCarpeta4), "", rstPagos!ImpoCarpeta4)
            
            Rem ImpoCarpeta.Text = Pusing("#,###,###.##", ImpoCarpeta.Text)
            Rem ImpoCarpeta1.Text = Pusing("#,###,###.##", ImpoCarpeta1.Text)
            Rem ImpoCarpeta2.Text = Pusing("#,###,###.##", ImpoCarpeta2.Text)
            Rem ImpoCarpeta3.Text = Pusing("#,###,###.##", ImpoCarpeta3.Text)
            Rem ImpoCarpeta4.Text = Pusing("#,###,###.##", ImpoCarpeta4.Text)
            
            WCertificadoGan = IIf(IsNull(rstPagos!CertificadoGan), "0", rstPagos!CertificadoGan)
            WCertificadoIb = IIf(IsNull(rstPagos!CertificadoIb), "0", rstPagos!CertificadoIb)
            WCertificadoIva = IIf(IsNull(rstPagos!CertificadoIva), "0", rstPagos!CertificadoIva)
            
            Paridad.Text = IIf(IsNull(rstPagos!Paridad), "0", rstPagos!Paridad)
            Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
            
            rstPagos.Close
                
        End If
        
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
            spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
            Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambioAdm.RecordCount > 0 Then
                Paridad.Text = Str$(rstCambioAdm!Cambio)
                Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
                rstCambioAdm.Close
            End If
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
            spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Proveedor.Text = RstProveedor!Proveedor
                DesProveedor.Caption = RstProveedor!Nombre
                WPrvDireccion = RstProveedor!Direccion
                WPrvCuit = RstProveedor!Cuit
                WPrvIb = RstProveedor!NroIb
                WTipoprv = Val(RstProveedor!Tipo) + 1
                WTipoiva = Val(RstProveedor!Iva)
                WTipoIb = RstProveedor!CodIb
                RstProveedor.Close
                Observaciones.SetFocus
                    Else
                Proveedor.Text = Proveedor.Text
                Proveedor.SetFocus
            End If
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
            spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                Banco.Text = rstBanco!Banco
                DesBanco.Caption = rstBanco!Nombre
                WCtabanco = rstBanco!Cuenta
                DbGrid1.Col = 0
                DbGrid1.Row = 0
                DbGrid1.SetFocus
                rstBanco.Close
                    Else
                Banco.Text = Banco.Text
                Banco.SetFocus
            End If
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
            
            spProveedor = "ListaProveedoresordConsulta"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
            
                With RstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Auxi$ = Mascara("###########", Str$(RstProveedor!Proveedor))
                            Call Ceros(Auxi, 11)
                            IngresaItem = Auxi + "      " + RstProveedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = RstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                RstProveedor.Close
                
            End If
            
        Case 1
            Erase Deuda
            EntraDeuda = 0
            XParam = "'" + Proveedor.Text + "','" _
                        + Proveedor.Text + "'"
            spCtaprv = "ListaCtaPrvDesdeHasta " + XParam
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
            
                With RstCtaPrv
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Proveedor.Text = RstCtaPrv!Proveedor Then
                                WAuxi1 = RstCtaPrv!Saldo
                                Call Redondeo(WAuxi1)
                                If WAuxi1 <> 0 Then
                                    EntraDeuda = EntraDeuda + 1
                                    Deuda(EntraDeuda, 1) = !NroInterno
                                    Deuda(EntraDeuda, 2) = !Total
                                    Deuda(EntraDeuda, 3) = !Saldo
                                    Deuda(EntraDeuda, 4) = !Impre
                                    Deuda(EntraDeuda, 5) = !Letra
                                    Deuda(EntraDeuda, 6) = !Punto
                                    Deuda(EntraDeuda, 7) = !Numero
                                    Deuda(EntraDeuda, 8) = !Fecha
                                    Deuda(EntraDeuda, 9) = !Clave
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                RstCtaPrv.Close
                
            End If
            
            For Ciclo = 1 To EntraDeuda
            
                XNroInterno = Deuda(Ciclo, 1)
                XTotal = Deuda(Ciclo, 2)
                XSaldo = Deuda(Ciclo, 3)
                XImpre = Deuda(Ciclo, 4)
                XLetra = Deuda(Ciclo, 5)
                XPunto = Deuda(Ciclo, 6)
                XNumero = Deuda(Ciclo, 7)
                XFecha = Deuda(Ciclo, 8)
                XClave = Deuda(Ciclo, 9)

                spIvaComp = "Consultaivacomp " + "'" + XNroInterno + "'"
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    XParidad = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
                    XPAgo = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
                    rstIvaComp.Close
                End If
                                
                ParidadTotal = 0
                If XPAgo <> 2 Then
                    WSaldo = XSaldo
                    WSaldoUs = 0
                    Call Redondeo(WSaldo)
                        Else
                    spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
                    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCambioAdm.RecordCount > 0 Then
                        ParidadTotal = rstCambioAdm!Cambio
                        rstCambioAdm.Close
                    End If
                                    
                    WSaldo = XSaldo
                    WSaldoUs = (XSaldo / XParidad) * ParidadTotal
                    Call Redondeo(WSaldo)
                    Call Redondeo(WSaldoUs)
                End If
                
                Auxi$ = Str$(WSaldo)
                Auxi$ = Mascara("#,###,###.##", Auxi$)
                If WSaldoUs <> 0 Then
                    Auxi1$ = Str$(WSaldoUs)
                    Auxi1$ = Mascara("#,###,###.##", Auxi1$)
                        Else
                    Auxi1$ = ""
                End If
                IngresaItem = XImpre + " " + XLetra + " " + XPunto + " " + XNumero + " " + XFecha + " " + Auxi$ + " " + Auxi1$
                Pantalla.AddItem IngresaItem
                IngresaItem = XClave
                WIndice.AddItem IngresaItem
                
            Next Ciclo
            
        Case 2
            spRecibos = "ListaRecibosNroCheque"
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount Then
            
                With rstRecibos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(rstRecibos!Tiporeg) = 2 Then
                                If Val(rstRecibos!Tipo2) = 2 And rstRecibos!Estado2 <> "X" Then
                                    Auxi$ = Str$(rstRecibos!Importe2)
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    Numero = Str$(Val(rstRecibos!Numero2))
                                    Call Ceros(Numero, 6)
                                    IngresaItem = Numero + "  " + rstRecibos!Fecha2 + "  " + Auxi$ + "  " + rstRecibos!Banco2
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstRecibos!Clave
                                    WIndice.AddItem IngresaItem
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibos.Close
                
            End If
            
        Case 3
            spCtaCte = "ListaCtacte"
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount Then
            
                With rstCtaCte
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(rstCtaCte!Tipo) = 50 Then
                                WSaldo = rstCtaCte!Saldo
                                Call Redondeo(WSaldo)
                                If WSaldo <> 0 And rstCtaCte!Cliente <> Space$(6) Then
                                    Auxi$ = Str$(Abs(rstCtaCte!Saldo))
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    WNumero = rstCtaCte!Numero
                                    Call Ceros(WNumero, 6)
                                    IngresaItem = WNumero + "  " + rstCtaCte!Vencimiento1 + "  " + Auxi$ + "  " + rstCtaCte!Cliente
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstCtaCte!Clave
                                    WIndice.AddItem IngresaItem
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCtaCte.Close
                
            End If
            
        Case 4
            spCuenta = "ListaCuentas"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount Then
            
                With rstCuenta
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCuenta!Cuenta + "  " + rstCuenta!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCuenta!Cuenta
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCuenta.Close
                
            End If
                
     
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            Proveedor.Text = Claveven$
            spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                DesProveedor.Caption = RstProveedor!Nombre
                WPrvDireccion = RstProveedor!Direccion
                WPrvCuit = RstProveedor!Cuit
                WPrvIb = RstProveedor!NroIb
                WTipoprv = Val(RstProveedor!Tipo) + 1
                WTipoiva = Val(RstProveedor!Iva)
                WTipoIb = RstProveedor!CodIb
                RstProveedor.Close
            End If
            
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
            
            Indice = Pantalla.ListIndex
            ClaveCtaprv = WIndice.List(Indice)
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                XTipo = RstCtaPrv!Tipo
                XLetra = RstCtaPrv!Letra
                XPunto = RstCtaPrv!Punto
                XNumero = RstCtaPrv!Numero
                XSaldo = RstCtaPrv!Saldo
                XNroInterno = Str$(RstCtaPrv!NroInterno)
                RstCtaPrv.Close
            End If
            
            DbGrid1.Row = XRow
            DbGrid1.Col = 0
            DbGrid1.Text = XTipo
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 1
            DbGrid1.Text = XLetra
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 2
            DbGrid1.Text = XPunto
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 3
            DbGrid1.Text = XNumero
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 4
            DbGrid1.Text = XSaldo
            DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    
            Select Case Val(XTipo)
                Case 1
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 5
                    DbGrid1.Text = "Pago Factura nro. " + Str$(XNumero)
                Case 2
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 5
                    DbGrid1.Text = "Pago Nota de Debito nro. " + Str$(XNumero)
                Case 3
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 5
                    DbGrid1.Text = "Pago Nota de Credito nro. " + Str$(XNumero)
                Case 5
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 5
                    DbGrid1.Text = "Anticipo nro. " + Str$(XNumero)
                Case Else
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 5
                    DbGrid1.Text = ""
            End Select
                    
            DbGrid1.Row = XRow
            DbGrid1.Col = 4
            
            spIvaComp = "Consultaivacomp " + "'" + XNroInterno + "'"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                XParidad = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
                XPAgo = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
                rstIvaComp.Close
            End If
            
            ParidadTotal = 0
            If XPAgo = 2 Then
            
                spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
                Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
                If rstCambioAdm.RecordCount > 0 Then
                    ParidadTotal = rstCambioAdm!Cambio
                    rstCambioAdm.Close
                End If
                WSaldo = XSaldo
                WSaldoUs = (XSaldo / XParidad) * ParidadTotal
                WDife = WSaldoUs - WSaldo
                Call Redondeo(WDife)
                
                If WDife <> 0 Then
                    If WDife > 0 Then
                
                        XRow = XRow + 1
                        
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 0
                        DbGrid1.Text = "02"
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 1
                        DbGrid1.Text = XLetra
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 2
                        DbGrid1.Text = XPunto
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 3
                        DbGrid1.Text = "99999999"
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 4
                        DbGrid1.Text = Str$(WDife)
                        Rem DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 5
                        DbGrid1.Text = "N/D por Diferencia de Cambio "
                    
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 4
                        
                            Else
                
                        XRow = XRow + 1
                        
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 0
                        DbGrid1.Text = "03"
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 1
                        DbGrid1.Text = XLetra
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 2
                        DbGrid1.Text = XPunto
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 3
                        DbGrid1.Text = "99999999"
                
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 4
                        DbGrid1.Text = Str$(WDife)
                        Rem DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 5
                        DbGrid1.Text = "N/C por Diferencia de Cambio "
                    
                        DbGrid1.Row = XRow
                        DbGrid1.Col = 4
                        
                    End If
                End If
                
            End If
            
            If XRow < 9 Then
            
                XRow = XRow + 1
                DbGrid1.Row = XRow
            
                DbGrid1.Col = 0
                DbGrid1.Text = ""
                
                DbGrid1.Col = 1
                DbGrid1.Text = ""
                
                DbGrid1.Col = 2
                DbGrid1.Text = ""
                
                DbGrid1.Col = 3
                DbGrid1.Text = ""
                
                DbGrid1.Col = 4
                DbGrid1.Text = ""
                    
                DbGrid1.Col = 5
                DbGrid1.Text = ""
                    
                DbGrid1.Col = 0
                
            End If
                
            Call Suma_Datos

            End If
            
            End If
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 0
            DbGrid1.SetFocus
            
        Case 2
        
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
            DbGrid1.Text = ""
        
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
        
            Indice = Pantalla.ListIndex
            ClaveRecibos = WIndice.List(Indice)
            spRecibos = "ConsultaRecibosClave " + "'" + ClaveRecibos + "'"
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                
                    DbGrid1.Col = 6
                    If XIndice = 2 Then
                        DbGrid1.Text = "3"
                            Else
                        DbGrid1.Text = "4"
                    End If
                    
                    DbGrid1.Col = 7
                    DbGrid1.Text = rstRecibos!Numero2
                
                    DbGrid1.Col = 8
                    DbGrid1.Text = rstRecibos!Fecha2
                
                    DbGrid1.Col = 9
                    DbGrid1.Text = ""
                
                    DbGrid1.Col = 10
                    DbGrid1.Text = rstRecibos!Banco2
                
                    DbGrid1.Col = 11
                    DbGrid1.Text = rstRecibos!Importe2
                    DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    
                    DbGrid1.Col = 12
                    DbGrid1.Text = ClaveRecibos
                    
                    rstRecibos.Close
                    
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
            
            Indice = Pantalla.ListIndex
            ClaveCtacte = WIndice.List(Indice)
            spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                
                    DbGrid1.Col = 6
                    DbGrid1.Text = "4"
                    
                    DbGrid1.Col = 7
                    DbGrid1.Text = rstCtaCte!Numero
                
                    DbGrid1.Col = 8
                    DbGrid1.Text = rstCtaCte!Vencimiento1
                
                    DbGrid1.Col = 9
                    DbGrid1.Text = ""
                
                    DbGrid1.Col = 10
                    DbGrid1.Text = ""
                
                    DbGrid1.Col = 11
                    DbGrid1.Text = rstCtaCte!Saldo
                    DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    
                    DbGrid1.Col = 12
                    DbGrid1.Text = ClaveCtacte
                    
                    rstCtaCte.Close
                    
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
            
            End If
            
        Case 4
            Rem Indice = Pantalla.ListIndex
            Rem ClaveCuenta = WIndice.List(Indice)
            Rem spCuenta = "ConsultaCuentas " + "'" + ClaveCuenta + "'"
            Rem Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstCuenta.RecordCount > 0 Then
                
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
                        If DbGrid1.Text = "A" Or DbGrid1.Text = "C" Or DbGrid1.Text = "X" Or DbGrid1.Text = "E" Then
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
                    
                    ClaveCtaprv = Proveedor.Text
                    DbGrid1.Col = 1
                    ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                    DbGrid1.Col = 0
                    ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                    DbGrid1.Col = 2
                    ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                    DbGrid1.Col = 3
                    ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                    spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    If RstCtaPrv.RecordCount > 0 Then
                        DbGrid1.Col = 4
                        XRow = DbGrid1.Row
                        If Val(DbGrid1.Text) = 0 Then
                            DbGrid1.Text = RstCtaPrv!Saldo
                            RstCtaPrv.Close
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
                        Else
                    DbGrid1.Col = 4
                    KeyCode = 0
                End If
                
            Case 4
                Rem dada
                If KeyCode = 13 Then
                
                    If Tipo1.Value = True Then
                        ClaveCtaprv = Proveedor.Text
                        DbGrid1.Col = 1
                        ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                        DbGrid1.Col = 0
                        ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                        DbGrid1.Col = 2
                        ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                        DbGrid1.Col = 3
                        ClaveCtaprv = ClaveCtaprv + DbGrid1.Text
                        
                        spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                        Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        If RstCtaPrv.RecordCount > 0 Then
                            Saldo = RstCtaPrv!Saldo
                            XNroInterno = Str$(RstCtaPrv!NroInterno)
                            XLetra = RstCtaPrv!Letra
                            XPunto = RstCtaPrv!Punto
                            RstCtaPrv.Close
                                Else
                            Saldo = 0
                        End If
                
                        DbGrid1.Col = 4
                        XSaldo = Val(DbGrid1.Text)
                        If XSaldo > Saldo Then
                        
                            XSaldo = 0
                            DbGrid1.Text = ""
                            DbGrid1.Col = 4
                            KeyCode = 0
                            
                                Else
                                
                            DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                            
                            XRow = DbGrid1.Row
                            Call Suma_Datos
                            DbGrid1.Col = 5
                            KeyCode = 0
                        
                            spIvaComp = "Consultaivacomp " + "'" + XNroInterno + "'"
                            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                            If rstIvaComp.RecordCount > 0 Then
                                XParidad = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
                                XPAgo = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
                                rstIvaComp.Close
                            End If
            
                            ParidadTotal = 0
                            If XPAgo = 2 Then
                                spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
                                Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
                                If rstCambioAdm.RecordCount > 0 Then
                                    ParidadTotal = rstCambioAdm!Cambio
                                    rstCambioAdm.Close
                                End If
                                WSaldo = XSaldo
                                WSaldoUs = (XSaldo / XParidad) * ParidadTotal
                                WDife = WSaldoUs - WSaldo
                                Call Redondeo(WDife)
                
                                If WDife <> 0 Then
                                    If WDife > 0 Then
                
                                        XRow = XRow + 1
                                    
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 0
                                        DbGrid1.Text = "02"
                            
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 1
                                        DbGrid1.Text = XLetra
                            
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 2
                                        DbGrid1.Text = XPunto
                    
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 3
                                        DbGrid1.Text = "99999999"
                    
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 4
                                        DbGrid1.Text = WDife
                                        DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 5
                                        DbGrid1.Text = "N/D por Diferencia de Cambio "
                    
                                            Else
                
                                        XRow = XRow + 1
                        
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 0
                                        DbGrid1.Text = "03"
                
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 1
                                        DbGrid1.Text = XLetra
                
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 2
                                        DbGrid1.Text = XPunto
                
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 3
                                        DbGrid1.Text = "99999999"
                
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 4
                                        DbGrid1.Text = WDife
                                        DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    
                                        DbGrid1.Row = XRow
                                        DbGrid1.Col = 5
                                        DbGrid1.Text = "N/C por Diferencia de Cambio "
                    
                                    End If
                                End If
                            End If
                        End If
                            Else
                        columna = DbGrid1.Row
                        DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                        Call Suma_Datos
                        DbGrid1.Col = 5
                        DbGrid1.Row = columna
                        KeyCode = 0
                    End If
                End If
                
            Case 5
                If KeyCode = 13 Then
                    If Tipo3.Value = True Then
                        WProceso = 0
                        Cuenta.Text = WCuenta(DbGrid1.Row, 1)
                        IngreCuenta.Visible = True
                        Cuenta.SetFocus
                            Else
                        If Tipo4.Value = True Then
                            spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
                            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                            If rstBanco.RecordCount > 0 Then
                                WCuenta(DbGrid1.Row, 1) = rstBanco!Cuenta
                                rstBanco.Close
                                    Else
                                WCuenta(DbGrid1.Row, 1) = "999999"
                            End If
                        End If
                        If Tipo5.Value = True Then
                            WCuenta(DbGrid1.Row, 1) = "111"
                        End If
                        If DbGrid1.Row < 9 Then
                            DbGrid1.Row = DbGrid1.Row + 1
                            DbGrid1.Col = 0
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End If
                End If
                
            Case 6
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Or Val(DbGrid1.Text) = 4 Or Val(DbGrid1.Text) = 5 Or Val(DbGrid1.Text) = 6 Or Val(DbGrid1.Text) = 7 Or Val(DbGrid1.Text) = 8 Then
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
                                DbGrid1.Text = "Efectivo"
                                DbGrid1.Col = 11
                                DbGrid1.Columns(11).Locked = False
                                KeyCode = 0
                                
                            Case 5
                                DbGrid1.Col = 7
                                DbGrid1.Text = ""
                                DbGrid1.Col = 8
                                DbGrid1.Text = ""
                                DbGrid1.Col = 9
                                DbGrid1.Text = ""
                                DbGrid1.Col = 10
                                DbGrid1.Text = "U$S"
                                DbGrid1.Col = 11
                                DbGrid1.Columns(11).Locked = False
                                KeyCode = 0
                                
                            Case 3, 4
                                Call Consulta_Click
                                
                            Case 6
                                WProceso = 1
                                Cuenta.Text = WCuenta(DbGrid1.Row, 2)
                                IngreCuenta.Visible = True
                                Cuenta.SetFocus
                                
                            Case 7
                                DbGrid1.Col = 7
                                DbGrid1.Text = ""
                                DbGrid1.Col = 8
                                DbGrid1.Text = ""
                                DbGrid1.Col = 9
                                DbGrid1.Text = ""
                                DbGrid1.Col = 10
                                DbGrid1.Text = "Patacones"
                                DbGrid1.Col = 11
                                DbGrid1.Columns(11).Locked = False
                                KeyCode = 0
                                
                            Case 8
                                DbGrid1.Col = 7
                                DbGrid1.Text = ""
                                DbGrid1.Col = 8
                                DbGrid1.Text = ""
                                DbGrid1.Col = 9
                                DbGrid1.Text = ""
                                DbGrid1.Col = 10
                                DbGrid1.Text = "Lecop"
                                DbGrid1.Col = 11
                                DbGrid1.Columns(11).Locked = False
                                KeyCode = 0
                                
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
                    DbGrid1.Col = 9
                    ClaveBanco = DbGrid1.Text
                    spBanco = "ConsultaBanco " + "'" + ClaveBanco + "'"
                    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                    If rstBanco.RecordCount > 0 Then
                            DbGrid1.Col = 10
                            DbGrid1.Text = rstBanco!Nombre
                            DbGrid1.Col = 11
                            KeyCode = 0
                            DbGrid1.Columns(9).Locked = True
                            DbGrid1.Columns(11).Locked = False
                            rstBanco.Close
                                Else
                            DbGrid1.Col = 9
                            KeyCode = 0
                    End If
                End If

            Case 11
                
                If KeyCode = 13 Then
                    iRow = DbGrid1.Row
                    DbGrid1.Col = 11
                    DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    Call Suma_Datos
                    DbGrid1.Row = iRow
                    If DbGrid1.Row < 9 Then
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
            spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                If WProceso = 0 Then
                    WCuenta(DbGrid1.Row, 1) = Cuenta.Text
                    IngreCuenta.Visible = False
                    DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 0
                    KeyCode = 0
                    DbGrid1.SetFocus
                        Else
                    WCuenta(DbGrid1.Row, 2) = Cuenta.Text
                    IngreCuenta.Visible = False
                    DbGrid1.Col = 7
                    DbGrid1.Text = ""
                    DbGrid1.Col = 8
                    DbGrid1.Text = ""
                    DbGrid1.Col = 9
                    DbGrid1.Text = ""
                    DbGrid1.Col = 10
                    DbGrid1.Text = "Varios"
                    DbGrid1.Col = 11
                    DbGrid1.Columns(11).Locked = False
                    KeyCode = 0
                    DbGrid1.SetFocus
                End If
                WProceso = 0
                    Else
                Cuenta.SetFocus
            End If
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

ReDim UserData(0 To 12, 0 To 12)

mTotalRows& = 11

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
    Tipo6.Value = False
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    Tipo6.Value = False
    Debitos.Caption = ""
    Creditos.Caption = ""
    Dife.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    RetIb.Text = ""
    RetIva.Text = ""
    
    Carpeta.Text = ""
    Carpeta1.Text = ""
    Carpeta2.Text = ""
    Carpeta3.Text = ""
    Carpeta4.Text = ""
    Rem ImpoCarpeta.Text = ""
    Rem ImpoCarpeta1.Text = ""
    Rem ImpoCarpeta2.Text = ""
    Rem ImpoCarpeta3.Text = ""
    Rem ImpoCarpeta4.Text = ""
    
    WLeyenda(1) = "Compra de Bienes"
    WLeyenda(2) = "Ejericio Prof. Lib. c/Aj.Inf."
    WLeyenda(3) = "Alquileres y Arrendamientos"
    WLeyenda(6) = "Locacion de Obras y/o servicios"
    WLeyenda(7) = "Transporte de Carga"
    WLeyenda(8) = "Factura M"
    
    WParametro(0) = 0
    WParametro(1) = 2000
    WParametro(2) = 4000
    WParametro(3) = 8000
    WParametro(4) = 14000
    WParametro(5) = 24000
    WParametro(6) = 1000000
    
    WTasa1(1) = 0.1
    WTasa1(2) = 0.14
    WTasa1(3) = 0.18
    WTasa1(4) = 0.22
    WTasa1(5) = 0.26
    WTasa1(6) = 0.26
    
    spCambioAdm = "ConsultaCambioAdm " + "'" + Fecha.Text + "'"
    Set rstCambioAdm = db.OpenRecordset(spCambioAdm, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambioAdm.RecordCount > 0 Then
        Paridad.Text = Str$(rstCambioAdm!Cambio)
        Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
        rstCambioAdm.Close
    End If
    
    Orden.Text = ""
    Rem spPagos = "ListaPagosNumero"
    Rem Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPagos.RecordCount > 0 Then
    Rem     With rstPagos
    Rem         .MoveLast
    Rem         Orden.Text = rstPagos!Orden + 1
    Rem     End With
    Rem     rstPagos.Close
    Rem End If
    
End Sub


Private Sub IMPREORDEN()

    On Error GoTo WError
        
    da = 0
    With rstImprePago
        .Index = "Orden"
        .Seek ">=", 0
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

    With rstEmpresa
        .Index = "Empresa"
        Claveven$ = WEmpresa
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Impretit = !Nombre
                Else
            Impretit = ""
        End If
    End With
    
    Select Case Val(WEmpresa)
        Case 1, 10
            WEmpCuit = "30-54916508-3"
        Case Else
            WEmpCuit = "30-61052459-8"
    End Select
        
    WRenglon = 0
    Cantidad = 0
    Total = 0
    SubTotaL = 0
        
    Erase WImpresion, WDebito, WCredito, WImpre2
        
    For iRow = 0 To 10
        WRow = iRow
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Cantidad = Cantidad + 1
            DbGrid1.Col = 0
            Select Case Val(Left$(DbGrid1.Text, 2))
                Case 1
                    WImpresion(Cantidad, 2) = "Factura"
                Case 2
                    WImpresion(Cantidad, 2) = "N.Debito"
                Case 3
                    WImpresion(Cantidad, 2) = "N.Credito"
                Case 99
                    WImpresion(Cantidad, 2) = "Varios"
                Case Else
                    WImpresion(Cantidad, 2) = ""
            End Select
                            
            DbGrid1.Col = 3
            WImpresion(Cantidad, 3) = Left$(DbGrid1.Text, 8)
            DbGrid1.Col = 5
            WImpresion(Cantidad, 4) = DbGrid1.Text
            DbGrid1.Col = 4
            WImpresion(Cantidad, 5) = DbGrid1.Text
            If Val(WImpresion(Cantidad, 2)) = 3 Or Val(WImpresion(Cantidad, 2)) = 5 Then
                Total = Total - Val(WImpresion(Cantidad, 5))
                    Else
                Total = Total + Val(WImpresion(Cantidad, 5))
            End If
                    
            DbGrid1.Col = 0
            WTipo = DbGrid1.Text
            DbGrid1.Col = 1
            WLetra = DbGrid1.Text
            DbGrid1.Col = 2
            WPunto = DbGrid1.Text
            DbGrid1.Col = 3
            WNumero = DbGrid1.Text
                
            ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                WImpresion(Cantidad, 1) = RstCtaPrv!Fecha
                RstCtaPrv.Close
                    Else
                WImpresion(Cantidad, 1) = ""
            End If
                    
        End If
    Next iRow
        
    With rstEmpresa
        .Index = "Empresa"
        Claveven$ = "1"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            WCtaProveedor = !CtaProveedores
            WCtaEfectivo = !CtaEfectivo
            WCtaCheques = !CtaCheque
        End If
    End With
        
    If Tipo1.Value = True Or Tipo2.Value = True Then
        WDebito(1, 1) = WCtaProveedor
        WDebito(1, 2) = Total
            Else
        For iRow = 0 To 9
            WRow = iRow
            DbGrid1.Col = 4
            DbGrid1.Row = iRow
            If Val(DbGrid1.Text) <> 0 Then
                WDebito(iRow + 1, 1) = WCuenta(iRow, 1)
                WDebito(iRow + 1, 2) = Val(DbGrid1.Text)
            End If
        Next iRow
                    
    End If

    WCredito(1, 1) = WCtaProveedor
    If Retenido <> 0 Then
            WCredito(1, 2) = Retenido
    End If
        
    Lugar = 1
    Impre2 = 0
        
    For iRow = 0 To 9
        DbGrid1.Col = 11
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Lugar = Lugar + 1
            WCredito(Lugar, 4) = DbGrid1.Text
            DbGrid1.Col = 6
            Select Case Val(DbGrid1.Text)
                Case 2
                    WCredito(Lugar, 1) = "999999"
                    DbGrid1.Col = 9
                    ClaveBanco = DbGrid1.Text
                    spBanco = "ConsultaBanco " + "'" + ClaveBanco + "'"
                    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                    If rstBanco.RecordCount > 0 Then
                        WCredito(Lugar, 1) = rstBanco!Cuenta
                        rstBanco.Close
                    End If
                Case 3, 4
                    WCredito(Lugar, 1) = WCtaCheques
                Case Else
                    WCredito(Lugar, 1) = WCtaEfectivo
            End Select
                    
            Impre2 = Impre2 + 1
            DbGrid1.Col = 7
            WImpre2(Impre2, 1) = DbGrid1.Text
            DbGrid1.Col = 10
            WImpre2(Impre2, 2) = DbGrid1.Text
            DbGrid1.Col = 11
            WImpre2(Impre2, 3) = DbGrid1.Text
            DbGrid1.Col = 8
            WImpre2(Impre2, 4) = DbGrid1.Text
                    
            DbGrid1.Col = 10
            WCredito(Lugar, 2) = DbGrid1.Text
            DbGrid1.Col = 7
            WCredito(Lugar, 3) = DbGrid1.Text
            DbGrid1.Col = 11
            WCredito(Lugar, 4) = DbGrid1.Text
        End If
    Next iRow
        
    SubTotaL = Total - Retenido
    TotalDebito = Total
    TotalCredito = Total

    For WCiclo = 1 To 10
    
        WFecha1 = ""
        WNumero1 = ""
        WComprobante1 = ""
        WDescripcion1 = ""
        WImporte1 = 0
        WNumero2 = ""
        WBanco2 = ""
        WImporte2 = 0
        WFecha2 = ""
            
        If Val(WImpresion(WCiclo, 5)) <> 0 Then
            WFecha1 = WImpresion(WCiclo, 1)
            WNumero1 = WImpresion(WCiclo, 3)
            WComprobante1 = WImpresion(WCiclo, 2)
            WDescripcion1 = WImpresion(WCiclo, 4)
            WImporte1 = Val(WImpresion(WCiclo, 5))
        End If
                    
        If Val(WImpre2(WCiclo, 3)) <> 0 Then
            WNumero2 = WImpre2(WCiclo, 1)
            WBanco2 = WImpre2(WCiclo, 2)
            WImporte2 = Val(WImpre2(WCiclo, 3))
            WFecha2 = WImpre2(WCiclo, 4)
        End If
        
        WRenglon = WRenglon + 1
        With rstImprePago
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            !Clave = "1" + Auxi + Auxi1
            !Tipo = 1
            !Orden = Val(Orden.Text)
            !Renglon = WRenglon
            !Fecha = Fecha.Text
            !Proveedor = Proveedor.Text
            !Nombre = DesProveedor.Caption
            !Fecha1 = WFecha1
            !Numero1 = WNumero1
            !Comprobante1 = WComprobante1
            !Descripcion1 = WDescripcion1
            !Importe1 = WImporte1
            !Numero2 = WNumero2
            !Banco2 = WBanco2
            !Importe2 = WImporte2
            !Fecha2 = WFecha2
            !Neto = Total
            !Rete1 = Val(Retencion.Text)
            !Rete2 = Val(RetIb.Text)
            !Total = Val(RetIva.Text)
            !Observaciones = Observaciones.Text
            !Empresa = Impretit
            !Cuit = WEmpCuit
            !Paridad = ParidadTotal
            .Update
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            !Clave = "2" + Auxi + Auxi1
            !Tipo = 2
            !Orden = Val(Orden.Text)
            !Renglon = WRenglon
            !Fecha = Fecha.Text
            !Proveedor = Proveedor.Text
            !Nombre = DesProveedor.Caption
            !Fecha1 = WFecha1
            !Numero1 = WNumero1
            !Comprobante1 = WComprobante1
            !Descripcion1 = WDescripcion1
            !Importe1 = WImporte1
            !Numero2 = WNumero2
            !Banco2 = WBanco2
            !Importe2 = WImporte2
            !Fecha2 = WFecha2
            !Neto = Total
            !Rete1 = Val(Retencion.Text)
            !Rete2 = Val(RetIb.Text)
            !Total = Val(RetIva.Text)
            !Observaciones = Observaciones.Text
            !Empresa = Impretit
            !Cuit = WEmpCuit
            !Paridad = ParidadTotal
            .Update
        End With
        
    Next WCiclo

    LISTADO.ReportFileName = "Imprepago.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next
  

End Sub


Private Sub Impreret()

    On Error GoTo WError
        
    WRenglon = 0
    da = 0
    With rstImpreRetGan
        .Index = "Orden"
        .Seek ">=", 0
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
    
    Mes% = Val(Mid$(Fecha.Text, 3, 2))
    WCuatri = ""

    If Mes% <= 4 Then
        WCuatri = "Primer Cuatrimestre"
            Else
        If Mes% >= 5 And Mes% <= 8 Then
            WCuatri = "Segundo Cuatrimestre"
                Else
            If Mes% >= 9 Then
                WCuatri = "Tercer Cuatrimestre"
            End If
        End If
    End If

    Select Case Val(WEmpresa)
        Case 1
            WEmpNombre = "SURFACTAN S.A."
            WEmpDireccion = "Malvinas Argentinas 4589"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-54916508-3"
        Case Else
            WEmpNombre = "PELLITAL S.A."
            WEmpDireccion = "Uruguay 2671"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-61052459-8"
    End Select
    
    
    With rstImpreRetGan
        .AddNew
        Auxi = Orden.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        !Clave = "1" + Auxi + Auxi1
        !Tipo = 1
        !Orden = Val(Orden.Text)
        !Renglon = WRenglon
        !NroCertificado = WCertificadoGan
        !Empresa = WEmpNombre
        !Direccion = WEmpDireccion
        !Localidad = WEmpLocalidad
        !Fecha = Fecha.Text
        !Cuit = WEmpCuit
        !NombrePrv = DesProveedor.Caption
        !DireccionPrv = WPrvDireccion
        !CuitPrv = WPrvCuit
        !Concepto = WLeyenda$(Val(WTipoprv))
        !Pagado = Total - WRetencion
        !Retenido = WRetencion
        .Update
    End With
    
    With rstImpreRetGan
        .AddNew
        Auxi = Orden.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        !Clave = "2" + Auxi + Auxi1
        !Tipo = 2
        !Orden = Val(Orden.Text)
        !Renglon = WRenglon
        !NroCertificado = WCertificadoGan
        !Empresa = WEmpNombre
        !Direccion = WEmpDireccion
        !Localidad = WEmpLocalidad
        !Fecha = Fecha.Text
        !Cuit = WEmpCuit
        !NombrePrv = DesProveedor.Caption
        !DireccionPrv = WPrvDireccion
        !CuitPrv = WPrvCuit
        !Concepto = WLeyenda$(Val(WTipoprv))
        !Pagado = Total - WRetencion
        !Retenido = WRetencion
        .Update
    End With
        
    LISTADO.ReportFileName = "Impreretgan.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next
    

End Sub


Private Sub calcret_Click()

    WRetencion = 0
    Retencion.Text = ""
    RetIva.Text = ""
    
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WTipoprv = Val(RstProveedor!Tipo) + 1
        WTipoiva = Val(RstProveedor!Iva)
        RstProveedor.Close
    End If
    
    If Tipo1.Value = True Or Tipo2.Value = True Then
    
        If WTipoprv = 1 Or WTipoprv = 2 Or WTipoprv = 3 Or WTipoprv = 6 Or WTipoprv = 7 Then
        
            Rem VidPop "Fill Page 1"
            Rem VidPop "Admi540i"
            Rem
            Rem Call Ingreso(Iva$, "N", "", 5, 2, 11, 53, 7, 0, "###.##", A$)
            Rem
            Rem VidPop "Display Page 1"
                    
            Rem If Val(RecordPrv.pago) = 1 Then
            Rem     Concepto$ = "8"
            Rem             Else
            Rem     If Val(RecordPrv.pago) = 2 Then
            Rem             Concepto$ = "13"
            Rem                     Else
            Rem             If Val(RecordPrv.pago) = 3 Then
            Rem                     Concepto$ = "6"
            Rem                             Else
            Rem                     Return
            Rem             End If
            Rem     End If
            Rem End If

            XBruto = Val(Debitos.Caption)
            If WTipoiva = 2 Then
                XNeto = (XBruto / 1.21)
                    Else
                XNeto = XBruto
            End If
            XIva = XBruto - XNeto
            XTBase = XNeto
            
            WFecha = Right$(Fecha.Text, 2) + Mid$(Fecha.Text, 4, 2)
            
            ClaveRetencion = WFecha + Proveedor.Text
            spRetencion = "ConsultaRetencion " + "'" + ClaveRetencion + "'"
            Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
            If rstRetencion.RecordCount > 0 Then
                WNeto = rstRetencion!Neto
                WAnticipo = rstRetencion!Anticipo
                WBruto = rstRetencion!Bruto
                WIva = rstRetencion!Iva
                WRetenido = rstRetencion!Retenido
                rstRetencion.Close
                    Else
                XFecha = WFecha
                XProveedor = Proveedor.Text
                XXNeto = ""
                XXAnticipo = ""
                XXBruto = ""
                XXIva = ""
                XXRetenido = ""
                XClave = XFecha + XProveedor
                
                XParam = "'" + XClave + "','" _
                        + XFecha + "','" + XProveedor + "','" _
                        + XXNeto + "','" _
                        + XXRetenido + "','" + XXAnticipo + "','" _
                        + XXBruto + "','" _
                        + XXAcumulado + "'"
                    
                spRstRetencion = "AltaRetencion " + XParam
                Set RstRstRetencion = db.OpenRecordset(spRstRetencion, dbOpenSnapshot, dbSQLPassThrough)
                    
                WNeto = 0
                WAnticipo = 0
                WBruto = 0
                WIva = 0
                WRetenido = 0
            End If
            
            Select Case WTipoprv
                Case 1
                    WMinimo = 12000
                Case 2
                    WMinimo = 1200
                Case 3
                    WMinimo = 1200
                Case 6
                    WMinimo = 5000
                Case 7
                    WMinimo = 6500
                Case Else
            End Select

            WAcupag = WNeto + XTBase
            WAuxi = WAcupag - WMinimo

            If WAuxi <= 0 Then
                WAuxi = 0
                WRetencion = 0
            End If

            WTasa = 0.02
            If WTipoprv = 1 Then
                    WTasa = 0.02
            End If
            If WTipoprv = 3 Then
                    WTasa = 0.06
            End If
            If WTipoprv = 7 Then
                    WTasa = 0.0025
            End If

            Select Case WTipoprv
                Case 2
                    WRetencion = 0
                    WTope = 0
                    WTope1 = 0
                    
                    For da = 0 To 5
                        If WAuxi >= WParametro(da) And WAuxi < WParametro(da + 1) Then
                            WTope1 = WAuxi
                            WTope = WParametro(da)
                            WSum = WTope1 - WTope
                            WSum = WSum * WTasa1(da + 1)
                            WRetencion = WRetencion + WSum
                        End If
                        If WAuxi >= WParametro(da + 1) Then
                            WTope1 = WParametro(da + 1)
                            WTope = WParametro(da)
                            WSum = WTope1 - WTope
                            WSum = WSum * WTasa1(da + 1)
                            WRetencion = WRetencion + WSum
                        End If
                    Next da
                    
                Case Else
                    WRetencion = WAuxi * WTasa
                    
            End Select

            WRetencion = WRetencion - WRetenido

            If WRetencion < 20 Then
                WRetencion = 0
                        Else
                If WRetencion > XNeto Then
                        WRetencion = 0
                End If
            End If
                    
            Call Redondeo(WRetencion)
            Retencion.Text = WRetencion
            Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            
        End If
        
        
        Rem If WTipoprv = 8 Then
        
            WRete1 = 0
            WRete2 = 0
                
            For iRow = 0 To 10
            
                WRow = iRow
                DbGrid1.Row = WRow
                
                DbGrid1.Col = 0
                XTipo = Left$(DbGrid1.Text, 2)
                DbGrid1.Col = 1
                XLetra = Left$(DbGrid1.Text, 1)
                DbGrid1.Col = 2
                XPunto = Left$(DbGrid1.Text, 4)
                DbGrid1.Col = 3
                XNumero = Left$(DbGrid1.Text, 8)
                DbGrid1.Col = 4
                XImporte = DbGrid1.Text
                
                If Val(XImporte) <> 0 And XLetra = "M" Then
                
                    XBruto = Val(XImporte)
                    XNeto = (XBruto / 1.21)
                    XIva = XBruto - XNeto

                    If XNeto >= 1000 Then
            
                        WTasa = 0.03
                        WRete1 = WRete1 + (XNeto * WTasa)
                    
                        Sql1 = "Select *"
                        Sql2 = " FROM IvaComp"
                        Sql3 = " Where IvaComp.Proveedor = " + "'" + Proveedor.Text + "'"
                        Sql4 = " and IvaComp.Tipo = " + "'" + XTipo + "'"
                        Sql5 = " and IvaComp.Letra = " + "'" + XLetra + "'"
                        Sql6 = " and IvaComp.Punto = " + "'" + XPunto + "'"
                        Sql7 = " and IvaComp.Numero = " + "'" + XNumero + "'"
                        spIvaComp = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
                        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                        If rstIvaComp.RecordCount > 0 Then
                            WRete2 = WRete2 + rstIvaComp!Iva21
                            rstIvaComp.Close
                        End If
                
                    End If
                    
                End If
                
            Next iRow
            
            If Val(Retencion.Text) = 0 Then
                Call Redondeo(WRete1)
                WRetencion = WRete1
                Retencion.Text = Str$(WRete1)
                Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            End If
            
            Call Redondeo(WRete2)
            RetIva.Text = Str$(WRete2)
            RetIva.Text = Pusing("#,###,###.##", RetIva.Text)
            
        Rem End If
        
    End If

End Sub


Private Sub CalcRetIb()

    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WTipoIb = RstProveedor!CodIb
        WTipoiva = Val(RstProveedor!Iva)
        WTipoprv = Val(RstProveedor!Tipo) + 1
        RstProveedor.Close
    End If

    WRetIb = 0
    RetIb.Text = ""
    
    If Tipo1.Value = True Or Tipo2.Value = True Then
    
        If WTipoIb = 0 Or WTipoIb = 1 Then
        
                XBruto = Val(Debitos.Caption)
                If WTipoiva = 2 Then
                    XNeto = (XBruto / 1.21)
                        Else
                    XNeto = XBruto
                End If
                XIva = XBruto - XNeto
                XTBase = XNeto
                                
                If XTBase >= 400 Then
                
                    WImpoRetenido = 0
        
                    For iRow = 0 To 10
                        WRow = iRow
                        DbGrid1.Col = 4
                        DbGrid1.Row = iRow
                        If Val(DbGrid1.Text) <> 0 Then
                                
                            WImpre4 = Val(DbGrid1.Text)
                            If WTipoiva = 2 Then
                                WImpre4 = WImpre4 / 1.21
                            End If
                            Call Redondeo(WImpre4)
                            XImpre4 = Str$(WImpre4)
                    
                            Select Case WTipoIb
                                Case 0
                                    WRete = Val(XImpre4) * (0.75 / 100)
                                    Call Redondeo(WRete)
                                    WImpoRetenido = WImpoRetenido + WRete
                                
                                Case Else
                                    WRete = Val(XImpre4) * (1.75 / 100)
                                    Call Redondeo(WRete)
                                    WImpoRetenido = WImpoRetenido + WRete
                        
                            End Select
                        
                        End If
                    Next iRow
        
                    WRetIb = WImpoRetenido
                    
                End If
                
                Call Redondeo(WRetIb)
                RetIb.Text = WRetIb
                RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
                
                Rem XBruto = Val(Debitos.Caption)
                Rem If WTipoiva = 2 Then
                Rem     XNeto = (XBruto / 1.21)
                Rem         Else
                Rem     XNeto = XBruto
                Rem End If
                Rem XIva = XBruto - XNeto
                Rem XTBase = XNeto
                Rem
                Rem If XTBase >= 400 Then
                Rem     Select Case WTipoIb
                Rem         Case 0
                Rem             WRetIb = XTBase * (0.75 / 100)
                Rem         Case 1
                Rem             WRetIb = XTBase * (1.75 / 100)
                Rem         Case Else
                Rem             WRetIb = 0
                Rem     End Select
                Rem End If
                Rem
                Rem Call Redondeo(WRetIb)
                Rem RetIb.Text = WRetIb
                Rem RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
            
        End If
        
    End If

End Sub

Private Sub Impreretib()

    On Error GoTo WError
        
    WRenglon = 0
    da = 0
    With rstImpreRetIb
        .Index = "Orden"
        .Seek ">=", 0
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

    Select Case Val(WEmpresa)
        Case 1, 10
            WEmpNombre = "SURFACTAN S.A."
            WEmpDireccion = "Malvinas Argentinas 4589"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-54916508-3"
            WNroIb = "902-913585-2"
            WNroAgente = ""
        Case Else
            WEmpNombre = "PELLITAL S.A."
            WEmpDireccion = "Uruguay 2671"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-61052459-8"
            WNroIb = ""
            WNroAgente = ""
    End Select
    
    
    ImpreCopia(1) = "Original"
    ImpreCopia(2) = "Duplicado"
    
        
    WImpoRetenido = 0
        
    For iRow = 0 To 10
        WRow = iRow
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            DbGrid1.Col = 0
            Select Case Val(Left$(DbGrid1.Text, 2))
                Case 1
                    XImpre1 = "Factura"
                Case 2
                    XImpre1 = "N.Debito"
                Case 3
                    XImpre1 = "N.Credito"
                Case 5
                    XImpre1 = "Anticipo"
                Case 99
                    XImpre1 = "Varios"
                Case Else
                    XImpre1 = ""
            End Select
                                
            DbGrid1.Col = 3
            XImpre2 = Left$(DbGrid1.Text, 8)
                
            Rem spIvacomp = "ConsultaIvacomp " + "'" + XImpre2 + "'"
            Rem Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstIvacomp.RecordCount > 0 Then
            Rem     XImpre2 = rstIvacomp!Numero
            Rem     rstIvacomp.Close
            Rem End If
                    
            DbGrid1.Col = 0
            WTipo = DbGrid1.Text
            DbGrid1.Col = 1
            WLetra = DbGrid1.Text
            DbGrid1.Col = 2
            WPunto = DbGrid1.Text
            DbGrid1.Col = 3
            WNumero = DbGrid1.Text
                    
            ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                XImpre3 = RstCtaPrv!Fecha
                RstCtaPrv.Close
                    Else
                XImpre3 = ""
            End If
                        
            DbGrid1.Col = 4
            WImpre4 = Val(DbGrid1.Text)
            Rem If Val(WTipo) = 3 Or Val(WTipo) = 5 Then
            Rem    WImpre4 = WImpre4 * -1
            Rem End If
            If WTipoiva = 2 Then
                WImpre4 = WImpre4 / 1.21
            End If
            Call Redondeo(WImpre4)
            XImpre4 = Str$(WImpre4)
                    
            Select Case WTipoIb
                Case 0
                    WRete = Val(XImpre4) * (0.75 / 100)
                    Call Redondeo(WRete)
                    WImpoRetenido = WImpoRetenido + WRete
                            
                    WRenglon = WRenglon + 1
                    With rstImpreRetIb
                        .AddNew
                        Auxi = Orden.Text
                        Call Ceros(Auxi, 6)
                        Auxi1 = WRenglon
                        Call Ceros(Auxi1, 2)
                        !Clave = "1" + Auxi + Auxi1
                        !Tipo = 1
                        !Orden = Val(Orden.Text)
                        !Renglon = WRenglon
                        !Empresa = WEmpNombre
                        !Direccion = WEmpDireccion
                        !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                        !Localidad = WEmpLocalidad
                        !Fecha = Fecha.Text
                        !Cuit = WEmpCuit
                        !Copia = ImpreCopia(da)
                        !NroIb = WNroIb
                        !NroAgente = WNroAgente
                        !NombrePrv = DesProveedor.Caption
                        !DireccionPrv = WPrvDireccion
                        !CuitPrv = WPrvCuit
                        !NroIbPrv = WPrvIb
                        !Tipo1 = XImpre1
                        !Numero1 = XImpre2
                        !Fecha1 = XImpre3
                        !Categoria1 = "SUJETO A RETENCION 0.75%"
                        !Importe1 = Val(XImpre4)
                        !Porce1 = 0.75
                        !Retencion1 = WRete
                        .Update
                        .AddNew
                        Auxi = Orden.Text
                        Call Ceros(Auxi, 6)
                        Auxi1 = WRenglon
                        Call Ceros(Auxi1, 2)
                        !Clave = "2" + Auxi + Auxi1
                        !Tipo = 2
                        !Orden = Val(Orden.Text)
                        !Renglon = WRenglon
                        !Empresa = WEmpNombre
                        !Direccion = WEmpDireccion
                        !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                        !Localidad = WEmpLocalidad
                        !Fecha = Fecha.Text
                        !Cuit = WEmpCuit
                        !Copia = ImpreCopia(da)
                        !NroIb = WNroIb
                        !NroAgente = WNroAgente
                        !NombrePrv = DesProveedor.Caption
                        !DireccionPrv = WPrvDireccion
                        !CuitPrv = WPrvCuit
                        !NroIbPrv = WPrvIb
                        !Tipo1 = XImpre1
                        !Numero1 = XImpre2
                        !Fecha1 = XImpre3
                        !Categoria1 = "SUJETO A RETENCION 0.75%"
                        !Importe1 = Val(XImpre4)
                        !Porce1 = 0.75
                        !Retencion1 = WRete
                        .Update
                    End With
                            
                Case Else
                    WRete = Val(XImpre4) * (1.75 / 100)
                    Call Redondeo(WRete)
                    WImpoRetenido = WImpoRetenido + WRete
                    
                    WRenglon = WRenglon + 1
                    With rstImpreRetIb
                        .AddNew
                        Auxi = Orden.Text
                        Call Ceros(Auxi, 6)
                        Auxi1 = WRenglon
                        Call Ceros(Auxi1, 2)
                        !Clave = "1" + Auxi + Auxi1
                        !Tipo = 1
                        !Orden = Val(Orden.Text)
                        !Renglon = WRenglon
                        !Empresa = WEmpNombre
                        !Direccion = WEmpDireccion
                        !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                        !Localidad = WEmpLocalidad
                        !Fecha = Fecha.Text
                        !Cuit = WEmpCuit
                        !Copia = ImpreCopia(da)
                        !NroIb = WNroIb
                        !NroAgente = WNroAgente
                        !NombrePrv = DesProveedor.Caption
                        !DireccionPrv = WPrvDireccion
                        !CuitPrv = WPrvCuit
                        !NroIbPrv = WPrvIb
                        !Tipo1 = XImpre1
                        !Numero1 = XImpre2
                        !Fecha1 = XImpre3
                        !Categoria1 = "SUJETO A RETENCION 1.75%"
                        !Importe1 = Val(XImpre4)
                        !Porce1 = 1.75
                        !Retencion1 = WRete
                        .Update
                        .AddNew
                        Auxi = Orden.Text
                        Call Ceros(Auxi, 6)
                        Auxi1 = WRenglon
                        Call Ceros(Auxi1, 2)
                        !Clave = "2" + Auxi + Auxi1
                        !Tipo = 2
                        !Orden = Val(Orden.Text)
                        !Renglon = WRenglon
                        !Empresa = WEmpNombre
                        !Direccion = WEmpDireccion
                        !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                        !Localidad = WEmpLocalidad
                        !Fecha = Fecha.Text
                        !Cuit = WEmpCuit
                        !Copia = ImpreCopia(da)
                        !NroIb = WNroIb
                        !NroAgente = WNroAgente
                        !NombrePrv = DesProveedor.Caption
                        !DireccionPrv = WPrvDireccion
                        !CuitPrv = WPrvCuit
                        !NroIbPrv = WPrvIb
                        !Tipo1 = XImpre1
                        !Numero1 = XImpre2
                        !Fecha1 = XImpre3
                        !Categoria1 = "SUJETO A RETENCION 1.75%"
                        !Importe1 = Val(XImpre4)
                        !Porce1 = 1.75
                        !Retencion1 = WRete
                        .Update
                    End With
                    
            End Select
                    
        End If
    Next iRow
    
    For Ciclo = WRenglon + 1 To 10
        With rstImpreRetIb
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "1" + Auxi + Auxi1
            !Tipo = 1
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(da)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "2" + Auxi + Auxi1
            !Tipo = 2
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(da)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
        End With
    Next Ciclo
        
    LISTADO.ReportFileName = "Impreretib.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub

Private Sub ImpreretIva()

    On Error GoTo WError
        
    WRenglon = 0
    da = 0
    With rstImpreRetIb
        .Index = "Orden"
        .Seek ">=", 0
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

    Select Case Val(WEmpresa)
        Case 1, 10
            WEmpNombre = "SURFACTAN S.A."
            WEmpDireccion = "Malvinas Argentinas 4589"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-54916508-3"
            WNroIb = "902-913585-2"
            WNroAgente = ""
        Case Else
            WEmpNombre = "PELLITAL S.A."
            WEmpDireccion = "Uruguay 2671"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-61052459-8"
            WNroIb = ""
            WNroAgente = ""
    End Select
    
    ImpreCopia(1) = "Original"
    ImpreCopia(2) = "Duplicado"
    
    WReteIva = 0
                
    For iRow = 0 To 10
            
        WRow = iRow
        DbGrid1.Row = WRow
                
        DbGrid1.Col = 0
        XTipo = Left$(DbGrid1.Text, 2)
        Select Case Val(Left$(DbGrid1.Text, 2))
            Case 1
                XImpre1 = "Factura"
            Case 2
                XImpre1 = "N.Debito"
            Case 3
                XImpre1 = "N.Credito"
            Case 5
                XImpre1 = "Anticipo"
            Case 99
                XImpre1 = "Varios"
            Case Else
                XImpre1 = ""
        End Select
        DbGrid1.Col = 1
        XLetra = Left$(DbGrid1.Text, 1)
        DbGrid1.Col = 2
        XPunto = Left$(DbGrid1.Text, 4)
        DbGrid1.Col = 3
        XNumero = Left$(DbGrid1.Text, 8)
        XImpre2 = Left$(DbGrid1.Text, 8)
        DbGrid1.Col = 4
        XImporte = DbGrid1.Text
                
        If Val(XImporte) <> 0 Then
                
            XBruto = Val(XImporte)
            XNeto = (XBruto / 1.21)
            XIva = XBruto - XNeto

            If XNeto >= 1000 Then
            
                Sql1 = "Select *"
                Sql2 = " FROM IvaComp"
                Sql3 = " Where IvaComp.Proveedor = " + "'" + Proveedor.Text + "'"
                Sql4 = " and IvaComp.Tipo = " + "'" + XTipo + "'"
                Sql5 = " and IvaComp.Letra = " + "'" + XLetra + "'"
                Sql6 = " and IvaComp.Punto = " + "'" + XPunto + "'"
                Sql7 = " and IvaComp.Numero = " + "'" + XNumero + "'"
                spIvaComp = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    WReteIva = rstIvaComp!Iva21
                    rstIvaComp.Close
                End If
                
            End If
                    
            WTipo = XTipo
            WLetra = XLetra
            WPunto = XPunto
            WNumero = XNumero
                    
            ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                XImpre3 = RstCtaPrv!Fecha
                RstCtaPrv.Close
                    Else
                XImpre3 = ""
            End If
                        
            WImpre4 = Val(XImporte)
                            
            WRenglon = WRenglon + 1
            With rstImpreRetIb
                .AddNew
                Auxi = Orden.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                !Clave = "1" + Auxi + Auxi1
                !Tipo = 1
                !Orden = Val(Orden.Text)
                !Renglon = WRenglon
                !Empresa = WEmpNombre
                !Direccion = WEmpDireccion
                !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
                !Localidad = WEmpLocalidad
                !Fecha = Fecha.Text
                !Cuit = WEmpCuit
                !Copia = ImpreCopia(da)
                !NroIb = WNroIb
                !NroAgente = WNroAgente
                !NombrePrv = DesProveedor.Caption
                !DireccionPrv = WPrvDireccion
                !CuitPrv = WPrvCuit
                !NroIbPrv = WPrvIb
                !Tipo1 = XImpre1
                !Numero1 = XImpre2
                !Fecha1 = XImpre3
                !Categoria1 = "SUJETO A RETENCION 0.75%"
                !Importe1 = Val(XImpre4)
                !Porce1 = WReteIva
                !Retencion1 = WReteIva
                .Update
                .AddNew
                Auxi = Orden.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                !Clave = "2" + Auxi + Auxi1
                !Tipo = 2
                !Orden = Val(Orden.Text)
                !Renglon = WRenglon
                !Empresa = WEmpNombre
                !Direccion = WEmpDireccion
                !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
                !Localidad = WEmpLocalidad
                !Fecha = Fecha.Text
                !Cuit = WEmpCuit
                !Copia = ImpreCopia(da)
                !NroIb = WNroIb
                !NroAgente = WNroAgente
                !NombrePrv = DesProveedor.Caption
                !DireccionPrv = WPrvDireccion
                !CuitPrv = WPrvCuit
                !NroIbPrv = WPrvIb
                !Tipo1 = XImpre1
                !Numero1 = XImpre2
                !Fecha1 = XImpre3
                !Categoria1 = ""
                !Importe1 = Val(XImpre4)
                !Porce1 = WReteIva
                !Retencion1 = WReteIva
                .Update
            End With
                    
        End If
    Next iRow
    
    For Ciclo = WRenglon + 1 To 10
        With rstImpreRetIb
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "1" + Auxi + Auxi1
            !Tipo = 1
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(da)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "2" + Auxi + Auxi1
            !Tipo = 2
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(da)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
        End With
    Next Ciclo
        
    LISTADO.ReportFileName = "Impreretiva.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub



Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    If Ayuda.Text <> "" Then
        spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
            Else
        spProveedor = "ListaProveedoresOrdConsulta"
    End If
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
    
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Nombre), aa, WEspacios) Then
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
    
    RstProveedor.Close
    
    End If
    
    End If

End Sub

Private Sub CargaCarpeta_Click()

    IngreCarpeta.Height = 3375
    IngreCarpeta.Left = 4080
    IngreCarpeta.Top = 2160
    IngreCarpeta.Width = 3015
        
    IngreCarpeta.Visible = True
        
    Carpeta.SetFocus
    
End Sub

Private Sub GrabaCarpeta_Click()
    IngreCarpeta.Visible = False
End Sub

Private Sub Carpeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7
                    CargaEmpresa(1, 1) = "0001"
                    CargaEmpresa(1, 2) = "Empresa01"
                    CargaEmpresa(2, 1) = "0003"
                    CargaEmpresa(2, 2) = "Empresa03"
                    CargaEmpresa(3, 1) = "0005"
                    CargaEmpresa(3, 2) = "Empresa05"
                    CargaEmpresa(4, 1) = "0006"
                    CargaEmpresa(4, 2) = "Empresa06"
                    CargaEmpresa(5, 1) = "0007"
                    CargaEmpresa(5, 2) = "Empresa07"
                    ZHasta = 5
                    
                Case Else
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    ZHasta = 4
                    
            End Select
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta1.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta1.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7
                    CargaEmpresa(1, 1) = "0001"
                    CargaEmpresa(1, 2) = "Empresa01"
                    CargaEmpresa(2, 1) = "0003"
                    CargaEmpresa(2, 2) = "Empresa03"
                    CargaEmpresa(3, 1) = "0005"
                    CargaEmpresa(3, 2) = "Empresa05"
                    CargaEmpresa(4, 1) = "0006"
                    CargaEmpresa(4, 2) = "Empresa06"
                    CargaEmpresa(5, 1) = "0007"
                    CargaEmpresa(5, 2) = "Empresa07"
                    ZHasta = 5
                    
                Case Else
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    ZHasta = 4
                    
            End Select
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta1.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta2.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta2.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7
                    CargaEmpresa(1, 1) = "0001"
                    CargaEmpresa(1, 2) = "Empresa01"
                    CargaEmpresa(2, 1) = "0003"
                    CargaEmpresa(2, 2) = "Empresa03"
                    CargaEmpresa(3, 1) = "0005"
                    CargaEmpresa(3, 2) = "Empresa05"
                    CargaEmpresa(4, 1) = "0006"
                    CargaEmpresa(4, 2) = "Empresa06"
                    CargaEmpresa(5, 1) = "0007"
                    CargaEmpresa(5, 2) = "Empresa07"
                    ZHasta = 5
                    
                Case Else
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    ZHasta = 4
                    
            End Select
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta2.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta3.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta3.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7
                    CargaEmpresa(1, 1) = "0001"
                    CargaEmpresa(1, 2) = "Empresa01"
                    CargaEmpresa(2, 1) = "0003"
                    CargaEmpresa(2, 2) = "Empresa03"
                    CargaEmpresa(3, 1) = "0005"
                    CargaEmpresa(3, 2) = "Empresa05"
                    CargaEmpresa(4, 1) = "0006"
                    CargaEmpresa(4, 2) = "Empresa06"
                    CargaEmpresa(5, 1) = "0007"
                    CargaEmpresa(5, 2) = "Empresa07"
                    ZHasta = 5
                    
                Case Else
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    ZHasta = 4
                    
            End Select
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta3.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta4.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta4.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7
                    CargaEmpresa(1, 1) = "0001"
                    CargaEmpresa(1, 2) = "Empresa01"
                    CargaEmpresa(2, 1) = "0003"
                    CargaEmpresa(2, 2) = "Empresa03"
                    CargaEmpresa(3, 1) = "0005"
                    CargaEmpresa(3, 2) = "Empresa05"
                    CargaEmpresa(4, 1) = "0006"
                    CargaEmpresa(4, 2) = "Empresa06"
                    CargaEmpresa(5, 1) = "0007"
                    CargaEmpresa(5, 2) = "Empresa07"
                    ZHasta = 5
                    
                Case Else
                    CargaEmpresa(1, 1) = "0002"
                    CargaEmpresa(1, 2) = "Empresa02"
                    CargaEmpresa(2, 1) = "0004"
                    CargaEmpresa(2, 2) = "Empresa04"
                    CargaEmpresa(3, 1) = "0008"
                    CargaEmpresa(3, 2) = "Empresa08"
                    CargaEmpresa(4, 1) = "0009"
                    CargaEmpresa(4, 2) = "Empresa09"
                    ZHasta = 4
                    
            End Select
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta4.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Alta_Dife()

    XNroInterno = ""
    spIvaComp = "ListaIvacompNumero"
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveLast
            XNroInterno = Str$(rstIvaComp!NroInterno + 1)
        End With
        rstIvaComp.Close
    End If
    WNumeroDife = XNroInterno
    
    Call Ceros(XNroInterno, 6)
    Call Ceros(WTipoDife, 2)
    Call Ceros(WPuntoDife, 4)
    Call Ceros(WNumeroDife, 8)
    
    Rem graba el iva compras
    
    XProveedor = Proveedor.Text
    XTipo = WTipoDife
    XLetra = WLetraDife
    XPunto = WPuntoDife
    XNumero = WNumeroDife
    XFecha = Fecha.Text
    Xvencimiento = Fecha.Text
    XVencimiento1 = Fecha.Text
    XPeriodo = Fecha.Text
    XImpoNeto = Str$(WNetoDife)
    XIva21 = Str$(WIvaDife)
    XIva5 = ""
    XIva27 = ""
    XIb = ""
    XExento = ""
    Select Case Val(WTipoDife)
        Case 1
            XImpre = "FC"
        Case 2
            XImpre = "ND"
        Case 3
            XImpre = "NC"
        Case Else
            XImpre = "  "
    End Select
    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XContado = "2"
    XEmpresa = "1"
    XNetolist = ""
    XExentolist = ""
    XParidad = ""
    XPAgo = "1"
    
    XParam = "'" + XNroInterno + "','" _
                 + XProveedor + "','" + XTipo + "','" _
                 + XLetra + "','" _
                 + XPunto + "','" + XNumero + "','" _
                 + XFecha + "','" _
                 + Xvencimiento + "','" _
                 + XVencimiento1 + "','" + XPeriodo + "','" _
                 + XImpoNeto + "','" _
                 + XIva21 + "','" _
                 + XIva5 + "','" + XIva27 + "','" _
                 + XIb + "','" + XExento + "','" _
                 + XContado + "','" _
                 + XImpre + "','" + XOrdFecha + "','" _
                 + XEmpresa + "','" + XNetolist + "','" _
                 + XExentolist + "','" _
                 + XParidad + "','" _
                 + XPAgo + "'"
                
    spIvaComp = "AltaIvaCompras " + XParam
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                
                
    Rem graba las imputaciones contables
                                        
    RenglonDife = 0
    
    
    Rem renglon nro 1
                                        
    RenglonDife = RenglonDife + 1
    Auxi1 = Str$(RenglonDife)
    Call Ceros(Auxi1, 2)
    XRenglon = Auxi1
                        
    XTipomovi = "2"
    XTipocomp = WTipoDife
    XLetracomp = WLetraDife
    XPuntocomp = WPuntoDife
    XNrocomp = WNumeroDife
    XFecha = Fecha.Text
    XObservaciones = ""
    Select Case Val(WTipoDife)
        Case 2
            XCuenta = "6107"
            XDebito = Str$(Abs(WNetoDife))
            XCredito = ""
        Case Else
            XCuenta = "7308"
            XDebito = ""
            XCredito = Str$(Abs(WNetoDife))
    End Select
    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XTitulo = "Compras"
    XEmpresa = "1"
    XClave = XTipomovi + XNroInterno + XRenglon
    XDebitolist = ""
    XCreditolist = ""
    
    XParam = "'" + XClave + "','" _
                 + XTipomovi + "','" + XProveedor + "','" _
                 + XTipocomp + "','" _
                 + XLetracomp + "','" + XPuntocomp + "','" _
                 + XNrocomp + "','" _
                 + XRenglon + "','" _
                 + XFecha + "','" + XObservaciones + "','" _
                 + XCuenta + "','" _
                 + XDebito + "','" _
                 + XCredito + "','" + XFechaOrd + "','" _
                 + XTitulo + "','" + XEmpresa + "','" _
                 + XDebitolist + "','" _
                 + XCreditolist + "','" _
                 + XNroInterno + "'"
                                
    spImputac = "AltaImputacion " + XParam
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem renglon nro 2
                                        
    RenglonDife = RenglonDife + 1
    Auxi1 = Str$(RenglonDife)
    Call Ceros(Auxi1, 2)
    XRenglon = Auxi1
                        
    XTipomovi = "2"
    XTipocomp = WTipoDife
    XLetracomp = WLetraDife
    XPuntocomp = WPuntoDife
    XNrocomp = WNumeroDife
    XFecha = Fecha.Text
    XObservaciones = ""
    Select Case Val(WTipoDife)
        Case 2
            XCuenta = "151"
            XDebito = Str$(Abs(WIvaDife))
            XCredito = ""
        Case Else
            XCuenta = "151"
            XDebito = ""
            XCredito = Str$(Abs(WIvaDife))
    End Select
    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XTitulo = "Compras"
    XEmpresa = "1"
    XClave = XTipomovi + XNroInterno + XRenglon
    XDebitolist = ""
    XCreditolist = ""
    
    XParam = "'" + XClave + "','" _
                 + XTipomovi + "','" + XProveedor + "','" _
                 + XTipocomp + "','" _
                 + XLetracomp + "','" + XPuntocomp + "','" _
                 + XNrocomp + "','" _
                 + XRenglon + "','" _
                 + XFecha + "','" + XObservaciones + "','" _
                 + XCuenta + "','" _
                 + XDebito + "','" _
                 + XCredito + "','" + XFechaOrd + "','" _
                 + XTitulo + "','" + XEmpresa + "','" _
                 + XDebitolist + "','" _
                 + XCreditolist + "','" _
                 + XNroInterno + "'"
                                
    spImputac = "AltaImputacion " + XParam
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem renglon nro 3
                                        
    RenglonDife = RenglonDife + 1
    Auxi1 = Str$(RenglonDife)
    Call Ceros(Auxi1, 2)
    XRenglon = Auxi1
                        
    XTipomovi = "2"
    XTipocomp = WTipoDife
    XLetracomp = WLetraDife
    XPuntocomp = WPuntoDife
    XNrocomp = WNumeroDife
    XFecha = Fecha.Text
    XObservaciones = ""
    Select Case Val(WTipoDife)
        Case 2
            XCuenta = "2001"
            XDebito = ""
            XCredito = Str$(Abs(WNetoDife) + Abs(WIvaDife))
        Case Else
            XCuenta = "2001"
            XDebito = Str$(Abs(WNetoDife) + Abs(WIvaDife))
            XCredito = ""
    End Select
    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XTitulo = "Compras"
    XEmpresa = "1"
    XClave = XTipomovi + XNroInterno + XRenglon
    XDebitolist = ""
    XCreditolist = ""
    
    XParam = "'" + XClave + "','" _
                 + XTipomovi + "','" + XProveedor + "','" _
                 + XTipocomp + "','" _
                 + XLetracomp + "','" + XPuntocomp + "','" _
                 + XNrocomp + "','" _
                 + XRenglon + "','" _
                 + XFecha + "','" + XObservaciones + "','" _
                 + XCuenta + "','" _
                 + XDebito + "','" _
                 + XCredito + "','" + XFechaOrd + "','" _
                 + XTitulo + "','" + XEmpresa + "','" _
                 + XDebitolist + "','" _
                 + XCreditolist + "','" _
                 + XNroInterno + "'"
                                
    spImputac = "AltaImputacion " + XParam
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    Rem alta en la cuenta corriente de proveedores
                
    XProveedor = Proveedor.Text
    XLetra = WLetraDife
    XTipo = WTipoDife
    XPunto = WPuntoDife
    XNumero = WNumeroDife
    XFecha = Fecha.Text
    XEstado = "1"
    Xvencimiento = Fecha.Text
    XVencimiento1 = Fecha.Text
    XNroInterno = XNroInterno
    XTotal = Str$(WNetoDife + WIvaDife)
    XSaldo = Str$(WNetoDife + WIvaDife)
    XClave = Proveedor.Text + WLetraDife + WTipoDife + WPuntoDife + WNumeroDife
    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    Select Case Val(WTipoDife)
        Case 1
            XImpre = "FC"
        Case 2
            XImpre = "ND"
        Case 3
            XImpre = "NC"
        Case Else
            XImpre = ""
    End Select
    XEmpresa = "1"
    XSaldolist = ""
    Xlista = ""
    XAcumulado = ""
    XParidad = ""
    XPAgo = "1"
                    
    XParam = "'" + XClave + "','" _
                 + XProveedor + "','" + XLetra + "','" _
                 + XTipo + "','" _
                 + XPunto + "','" + XNumero + "','" _
                 + XFecha + "','" _
                 + XEstado + "','" _
                 + Xvencimiento + "','" + XVencimiento1 + "','" _
                 + XTotal + "','" _
                 + XSaldo + "','" _
                 + XOrdFecha + "','" + XOrdVencimiento + "','" _
                 + XImpre + "','" + XEmpresa + "','" _
                 + XSaldolist + "','" _
                 + XNroInterno + "','" + Xlista + "','" _
                 + XAcumulado + "','" _
                 + XParidad + "','" _
                 + XPAgo + "'"
                    
    spConsulta = "AltaCtaPrv " + XParam
    Set rstConsulta = db.OpenRecordset(spConsulta + cParam, dbOpenSnapshot, dbSQLPassThrough)
    
End Sub

