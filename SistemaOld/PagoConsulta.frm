VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgpagoConsulta 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Pagos a Proveedores"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   1155
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6855
   ScaleWidth      =   11880
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   42
      Top             =   4680
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   4200
      TabIndex        =   41
      Top             =   4320
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4920
      TabIndex        =   40
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   9
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   10
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   11
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   12
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   5280
      Width           =   375
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
      Left            =   3960
      TabIndex        =   35
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Carpeta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   33
      Text            =   " "
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox RetIb 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10560
      TabIndex        =   31
      Text            =   " "
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   10320
      TabIndex        =   30
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   7200
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox Retencion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7800
      TabIndex        =   28
      Text            =   " "
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Banco 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      Text            =   " "
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame IngreCuenta 
      Caption         =   "Cuenta Contable"
      Height          =   855
      Left            =   2880
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Cuenta 
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Text            =   " "
      Top             =   720
      Width           =   5415
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
      TabIndex        =   10
      Top             =   1440
      Width           =   3735
      Begin VB.OptionButton Tipo5 
         Caption         =   "Cheques Rechazados"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Tipo4 
         Caption         =   "Transferencias"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Tipo3 
         Caption         =   "Pagos Varios"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Pagos de Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Anticipos"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
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
      TabIndex        =   8
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
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2010
      ItemData        =   "PagoConsulta.frx":0000
      Left            =   7200
      List            =   "PagoConsulta.frx":0007
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3495
      Left            =   0
      TabIndex        =   51
      Top             =   3000
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   5400
      TabIndex        =   52
      Top             =   4320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label8 
      Caption         =   "Carpeta Importacion"
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ret.Ing.Brutos"
      Height          =   255
      Left            =   9120
      TabIndex        =   32
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retencion Ganan."
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label DesBanco 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2520
      TabIndex        =   26
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label bjm 
      Caption         =   "Banco"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   10320
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Debitos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label DesProveedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   " "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro. Orden de Pago"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "PrgpagoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Debito As Double
Private Credito As Double
Private WImpresion(20, 20) As String
Private WImpre2(20, 10) As String
Private WDebito(20, 2) As String
Private WCredito(20, 4) As String
Private WCuenta(20, 2) As String
Private WCuentaBco As String
Private Numero As String
Private WNumero As String
Private WSaldo As Double
Private WRetencion As Double
Private WRetIb As Double
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
Private WTipoIbCaba As Single
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
Private WParametro(0 To 20) As Double
Private WTasa1(20) As Double
Private WAuxi As Double
Private WAuxi1 As Double
Private Total As Double
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
Dim XParam As String
Dim WProceso As Integer
Dim WCerti As String
Dim WCerificado As Integer
Dim ImpreCopia(20) As String
Dim WRete As Double
Dim WImpoRetenido As Double
Dim XImpre1 As String
Dim XImpre2 As String
Dim XImpre3 As String
Dim XImpre4 As String
Dim WImpre4 As Double
Dim WOtro As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String
Dim WControlII As String

Private Sub Suma_Datos()
    Debitos.Caption = ""
    Creditos.Caption = ""
    
    For iRow = 1 To 12
        WTipo = WVector1.TextMatrix(iRow, 1)
        If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
            If Tipo1.Value = True Then
                If Val(WTipo) <> 0 Then
                    Debitos.Caption = Str$(Val(Debitos.Caption) + Val(WVector1.TextMatrix(iRow, 5)))
                End If
                    Else
                Debitos.Caption = Str$(Val(Debitos.Caption) + Val(WVector1.TextMatrix(iRow, 5)))
            End If
        End If
        If Val(WVector1.TextMatrix(iRow, 12)) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(WVector1.TextMatrix(iRow, 12)))
        End If
    Next iRow
    
    If Existe <> "S" Then
        Call calcret_Click
        Call CalcRetIb
    End If
    Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Retencion.Text) + Val(RetIb.Text))
    
    Debitos.Caption = Pusing("###,###.##", Debitos.Caption)
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
    
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
                    WVector1.TextMatrix(Debito, 1) = rstPagos!Tipo1
                    WVector1.TextMatrix(Debito, 2) = rstPagos!Letra1
                    WVector1.TextMatrix(Debito, 3) = rstPagos!Punto1
                    WVector1.TextMatrix(Debito, 4) = rstPagos!Numero1
                    WVector1.TextMatrix(Debito, 5) = Str$(rstPagos!Importe1)
                    WVector1.TextMatrix(Debito, 5) = Pusing("###,###.##", WVector1.TextMatrix(Debito, 5))
                    WVector1.TextMatrix(Debito, 6) = rstPagos!Observaciones2
                Case 2
                    Credito = Credito + 1
                    WVector1.TextMatrix(Credito, 7) = rstPagos!Tipo2
                    WVector1.TextMatrix(Credito, 8) = rstPagos!Numero2
                    WVector1.TextMatrix(Credito, 9) = rstPagos!Fecha2
                    WVector1.TextMatrix(Credito, 10) = rstPagos!Banco2
                    If rstPagos!Observaciones2 <> "" Then
                        WVector1.TextMatrix(Credito, 11) = rstPagos!Observaciones2
                    End If
                    WVector1.TextMatrix(Credito, 12) = Str$(rstPagos!Importe2)
                    WVector1.TextMatrix(Credito, 12) = Pusing("###,###.##", WVector1.TextMatrix(Credito, 12))
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
    Rem Retganancias.text = PUsing("###,###.##", Retganancias.text)
End Sub

Sub Imprime_Datos()
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
        WTipoIbCaba = RstProveedor!CodIbCaba
        RstProveedor.Close
        Call Format_datos
    End If
End Sub

Private Sub cmdClose_Click()
    
    With rstEmpresa
        .Close
    End With
    
    PrgpagoConsulta.Hide
    Unload Me
    PrgCcprv1.Show
    
End Sub


Private Sub Cerrar_Click()
    Rem Call CmdLimpiar_Click
    Rem Orden.SetFocus
    PrgpagoConsulta.Hide
    Unload Me
    PrgCcprv1.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Impresion_Click()
        Call IMPREORDEN
        If Val(Retencion.Text) <> 0 Then
            WRetencion = Val(Retencion.Text)
            Call Impreret
        End If
        If Val(RetIb.Text) <> 0 Then
            WOtro = "N"
            WRetIb = Val(RetIb.Text)
            Call Impreretib
        End If
End Sub

Private Sub Orden_GotFocus()

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
            Retencion.Text = Pusing("###,###.##", Retencion.Text)
            RetIb.Text = rstPagos!RetOtra
            RetIb.Text = Pusing("###,###.##", RetIb.Text)
            Tipo1.Value = False
            Tipo2.Value = False
            Tipo3.Value = False
            Tipo4.Value = False
            Tipo5.Value = False
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
                Case Else
            End Select
            Observaciones.Text = rstPagos!Observaciones
            Carpeta.Text = IIf(IsNull(rstPagos!Carpeta), "", rstPagos!Carpeta)
            rstPagos.Close
                
        End If
        
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If

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
            Retencion.Text = Pusing("###,###.##", Retencion.Text)
            RetIb.Text = rstPagos!RetOtra
            RetIb.Text = Pusing("###,###.##", RetIb.Text)
            Tipo1.Value = False
            Tipo2.Value = False
            Tipo3.Value = False
            Tipo4.Value = False
            Tipo5.Value = False
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
                Case Else
            End Select
            Observaciones.Text = rstPagos!Observaciones
            Carpeta.Text = IIf(IsNull(rstPagos!Carpeta), "", rstPagos!Carpeta)
            rstPagos.Close
                
        End If
        
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
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
                WTipoIbCaba = RstProveedor!CodIbCaba
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
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
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
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
                rstBanco.Close
                    Else
                Banco.Text = Banco.Text
                Banco.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Form_Load()

    Call Limpia_Vector

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
    RetIb.Text = ""
    Carpeta.Text = ""
    
    WLeyenda(1) = "Compra de Bienes"
    WLeyenda(2) = "Ejericio Prof. Lib. c/Aj.Inf."
    WLeyenda(3) = "Alquileres y Arrendamientos"
    
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
    
    Orden.Text = ""
    
    Orden.Text = WOPago
    Call Orden_KeyPress(13)
    Rem Orden.SetFocus
    
End Sub


Private Sub IMPREORDEN()

    Rem Open "da.txt" For Output As #1
    Open "lpt1" For Output As #1
    Rem Open "aa" For Output As #1
    
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70)
    
        Rem Print #1,Quality = -1

        Rem Printer.Font = "Times New Roman"
        Rem Printer.FontSize = "10"
        Rem Print #1, Tab(1); ""
        Rem Printer.FontSize = "9"

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

        Rem Retenido# = FNRedondeo#(Val(WRetencion.010$)/100)
        Rem Pagado#   = Val(Wpago.010$)/100

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

        For Ciclo% = 1 To 2

            Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72);
            Print #1, Tab(1); "ORDEN DE PAGO";
            Print #1, Tab(25); Impretit;
            Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
            
            Print #1, Tab(1); String$(127, "-")
            
            Print #1, Tab(1); "|";
            Print #1, Tab(10); "Proveedor";
            Print #1, Tab(60); "     Fecha    ";
            Print #1, Tab(115); "Numero";
            Print #1, Tab(127); "|"
            
            Print #1, Tab(1); "|";
            Print #1, Tab(2); "Sres.:"; DesProveedor.Caption;
            Print #1, Tab(63); Fecha.Text;
            Print #1, Tab(115); Orden.Text;
            Print #1, Tab(127); "|"
            
            Print #1, Tab(1); String$(127, "-")
            
            Print #1, Tab(1); "|";
            Print #1, Tab(14); "|";
            Print #1, Tab(29); "|";
            Print #1, Tab(39); "|";
            Print #1, Tab(76); "|";
            Print #1, Tab(87); "|";
            Print #1, Tab(88); "Valores o Docucmentos Entregados";
            Print #1, Tab(127); "|"
            
            Print #1, Tab(1); "|";
            Print #1, Tab(2); " Fecha";
            Print #1, Tab(14); "|";
            Print #1, Tab(15); "Numero";
            Print #1, Tab(29); "|";
            Print #1, Tab(30); "Comp.";
            Print #1, Tab(39); "|";
            Print #1, Tab(40); "Descripcion ";
            Print #1, Tab(76); "|";
            Print #1, Tab(78); "Importe";
            Print #1, Tab(87); "|";
            Print #1, Tab(88); "Nro.";
            Print #1, Tab(96); "|";
            Print #1, Tab(97); "Banco/Cliente";
            Print #1, Tab(116); "|";
            Print #1, Tab(117); " Importe";
            Print #1, Tab(127); "|"

            Print #1, Tab(1); String$(127, "-")
    
            For WCiclo = 1 To 10

                If Val(WImpresion(WCiclo, 5)) <> 0 Then
                
                    Print #1, Tab(1); "|";
                    Print #1, Tab(2); ""; WImpresion(WCiclo, 1);
                    Print #1, Tab(14); "|";
                    
                    spIvaComp = "ConsultaIvacomp " + "'" + WImpresion(WCiclo, 3) + "'"
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                           XImpre = rstIvaComp!Numero
                           rstIvaComp.Close
                               Else
                           XImpre = WImpresion(WCiclo, 3)
                    End If
                    
                    Rem Print #1, Tab(15); ""; XImpre;

                    Print #1, Tab(15); ""; WImpresion(WCiclo, 3);
                    Print #1, Tab(29); "|";
                    Print #1, Tab(30); ""; WImpresion(WCiclo, 2);
                    Print #1, Tab(39); "|";
                    Print #1, Tab(40); ""; WImpresion(WCiclo, 4);
                    Print #1, Tab(76); "|";
                    Print #1, Tab(77); ""; Alinea("###,###.##", WImpresion(WCiclo, 5));
                    Print #1, Tab(87); "|";
                    
                    If Val(WImpre2(WCiclo, 3)) <> 0 Then
                        DbGrid1.Col = 7
                        Print #1, Tab(88); WImpre2(WCiclo, 1);
                        DbGrid1.Col = 10
                        Print #1, Tab(96); "|";
                        Print #1, Tab(97); Left$(WImpre2(WCiclo, 2), 19);
                        DbGrid1.Col = 11
                        Print #1, Tab(116); "|";
                        Print #1, Tab(117); Alinea("###,###.##", WImpre2(WCiclo, 3));
                            Else
                        Print #1, Tab(92); "";
                        Print #1, Tab(96); "|";
                        Print #1, Tab(116); "|";
                    End If
                    Print #1, Tab(127); "|"
                    
                            Else
                            
                    Print #1, Tab(1); "|";
                    Print #1, Tab(14); "|";
                    Print #1, Tab(29); "|";
                    Print #1, Tab(39); "|";
                    Print #1, Tab(76); "|";
                    Print #1, Tab(87); "|";
                    If Val(WImpre2(WCiclo, 3)) <> 0 Then
                        DbGrid1.Col = 7
                        Print #1, Tab(88); WImpre2(WCiclo, 1);
                        DbGrid1.Col = 10
                        Print #1, Tab(96); "|";
                        Print #1, Tab(97); Left$(WImpre2(WCiclo, 2), 19);
                        DbGrid1.Col = 11
                        Print #1, Tab(116); "|";
                        Print #1, Tab(117); Alinea("###,###.##", WImpre2(WCiclo, 3));
                            Else
                        Print #1, Tab(92); "";
                        Print #1, Tab(96); "|";
                        Print #1, Tab(116); "|";
                    End If
                    Print #1, Tab(127); "|"

                End If

            Next WCiclo

            Print #1, Tab(1); String$(127, "-")
            
            Print #1, Tab(1); "|";
            Print #1, Tab(3); " Importe ";
            Print #1, Tab(10); ""; Alinea("###,###.##", Str$(Total));
            Print #1, Tab(20); " Ret. Ganancias";
            Print #1, Tab(40); ""; Alinea("###,###.##", Retencion.Text);
            Print #1, Tab(55); " Ret. Ing.Brutos";
            Print #1, Tab(75); ""; Alinea("###,###.##", RetIb.Text);
            Print #1, Tab(90); " Importe a Pagar";
            Print #1, Tab(110); ""; Alinea("###,###.##", Str$(Total - Val(Retencion.Text) - Val(RetIb.Text)));
            Print #1, Tab(127); "|"

            Print #1, Tab(1); String$(127, "-")

            Rem Print #1, Tab(1); "     Codigo";
            Rem Print #1, Tab(25); "     Importe";
            Rem Print #1, Tab(50); "  Codigo";
            Rem Print #1, Tab(75); "      Banco";
            Rem Print #1, Tab(115); "  Cheque";
            Rem Print #1, Tab(127); "  Importe";
            Rem Print #1, Tab(150); ""
            Rem
            Rem Print #1, Tab(1); String$(120, "_")
            Rem
            Rem For da = 1 To 10
            Rem         If Val(WDebito$(da, 2)) <> 0 Or Val(WCredito(da, 4)) <> 0 Then
            Rem             Print #1, Tab(1); ""; WDebito(da, 1);
            Rem             Print #1, Tab(25); "";
            Rem             If Val(WDebito(da, 2)) <> 0 Then
            Rem                 Print #1, Tab(30); Alinea("#,###,###.##", WDebito$(da, 2));
            Rem             End If
            Rem             Print #1, Tab(50); ""; WCredito(da, 1);
            Rem             Print #1, Tab(75); ""; WCredito(da, 2);
            Rem             Print #1, Tab(115); ""; WCredito(da, 3);
            Rem             Print #1, Tab(127); "";
            Rem             If Val(WCredito(da, 4)) <> 0 Then
            Rem                 Print #1, Tab(135); Alinea("#,###,###.##", WCredito$(da, 4));
            Rem             End If
            Rem             Print #1, Tab(150); ""
            Rem         End If
            Rem Next da
            Rem
            Rem Print #1, Tab(1); String$(120, "_")
            Rem
            Rem Print #1, Tab(1); " Total Debito";
            Rem Print #1, Tab(30); ""; Alinea("#,###,###.##", Str$(Total));
            Rem Print #1, Tab(50); "";
            Rem Print #1, Tab(75); "";
            Rem Print #1, Tab(115); " Total Credito";
            Rem Print #1, Tab(135); ""; Alinea("#,###,###.##", Str$(Total));
            Rem Print #1, Tab(150); ""
            Rem
            Rem
            Rem Print #1, Tab(1); String$(127, "_")
            
            Print #1, Tab(1); "|";
            Print #1, "OBSERVACIONES :"; Observaciones.Text;
            Print #1, Tab(127); "|"
            
            
            
            Print #1, Tab(1); String$(127, "-")

            Print #1, Tab(1); "|";
            Print #1, Tab(127); "|"
            Print #1, Tab(1); "|";
            Print #1, Tab(127); "|"
            Print #1, Tab(1); "|";
            Print #1, Tab(127); "|"
            
            Print #1, Tab(1); "|";
            Print #1, Tab(10); "   Confecciono";
            Print #1, Tab(35); "    Autorizo";
            Print #1, Tab(60); "   1ra Firma";
            Print #1, Tab(85); "   2da Firma";
            Print #1, Tab(105); "  Recibi Conforme";
            Print #1, Tab(127); "|"
            
            Print #1, Tab(1); String$(127, "-")
            Print #1, ""
            Print #1, ""
            Print #1, ""

  Next Ciclo%
  Print #1, Chr$(12)
  
  Close #1
  
  cc = 0

End Sub


Private Sub Impreret()

    Rem Open "dada.txt" For Output As #1
    Open "lpt1" For Output As #1

    Print #1, Chr$(15);

    Rem m# = Rete2784#
    
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

    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
    
    Open WEmpresa + "nro.txt" For Input As #10
    Input #10, WCerti
    Input #10, WCerti2
    Close #10
    
    WCertificado = Val(WCerti) + 1
    Open WEmpresa + "nro.txt" For Output As #10
    Print #10, WCertificado
    Print #10, WCerti2
    Close #10
    
    For da = 1 To 2
        Print #1, "Nro.Certificado  : "; WCertificado
        Print #1, "                                                     COMPROBANTE DE RETENCION"
        Print #1, "                                                IMPUESTO A LAS GANACIAS RG 2784                            "
        Print #1, WEmpNombre
        Print #1, WEmpDireccion
        Print #1, WEmpLocalidad
        Print #1, "Clave Unica de Identificacion Tributaria : ", WEmpCuit
        Print #1, "----------------------------------------------------------------------------------------------------------------------------------"
        Print #1, "SUJETO RETENIDO                                                                               |"
        Print #1, "                                                                                              |"
        Print #1, "Nombre/Razon Social : "; DesProveedor.Caption; Tab(95); "|"
        Print #1, "Domicilio           : "; WPrvDireccion; Tab(95); "|"
        Print #1, "Clave Unica de Identificacion Tributaria : "; WPrvCuit; Tab(95); "|"
        Print #1, "..............................................................................................|"
        Print #1, "                                                                                              |"
        Print #1, "DETALLE DE LA RETENCION                                                                       |"
        Print #1, "                                                                                              |"
        Print #1, "Concepto de la Retencion : "; WLeyenda$(Val(WTipoprv)); Tab(95); "|"
        Print #1, "Importe Pagado           : "; Alinea("###,###.##", Str$(Total - WRetencion)); Tab(95); "|"
        Print #1, "Importe Retenido         : "; Alinea("###,###.##", Str$(WRetencion)); Tab(95); "|"
        Print #1, "                                                                                              |----------------------------------"
        Print #1, ".................................................................................................................................."
        Print #1,
        Print #1, "La Presente Retencion efectuada el "; Fecha.Text; " se informara en la Declaracion Jurada del mes."
        Print #1, ""
        Print #1, "=================================================================================================================================="
    Next da
    Print #1, Chr$(12)
    
    Close #1

End Sub


Private Sub calcret_Click()

    WRetencion = 0
    
    If Tipo1.Value = True Or Tipo2.Value = True Then
    
        If WTipoprv = 1 Or WTipoprv = 2 Or WTipoprv = 3 Then
        
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

            If WTipoprv = 1 Then
                WMinimo = 12000
                                Else
                If WTipoprv = 2 Then
                        WMinimo = 1200
                                        Else
                        WMinimo = 1200
                End If
            End If

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
            Retencion.Text = Pusing("###,###.##", Retencion.Text)
            
        End If
        
    End If

End Sub


Private Sub CalcRetIb()

    WRetIb = 0
    
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
                Select Case WTipoIb
                    Case 0
                        WRetIb = XTBase * (0.75 / 100)
                    Case 1
                        WRetIb = XTBase * (1.75 / 100)
                    Case Else
                        WRetIb = 0
                End Select
            End If
                    
            Call Redondeo(WRetIb)
            RetIb.Text = WRetIb
            RetIb.Text = Pusing("###,###.##", RetIb.Text)
            
        End If
        
    End If

End Sub


Private Sub Impreretib()

    Rem Open "dada.txt" For Output As #1
    Open "lpt1" For Output As #1

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
    
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
    
    Open WEmpresa + "nro.txt" For Input As #10
    Input #10, WCerti1
    Input #10, WCerti2
    Close #10
    
    WCertificado = Val(WCerti2) + 1
    Open WEmpresa + "nro.txt" For Output As #10
    Print #10, WCerti1
    Print #10, WCertificado
    Close #10
    
    ImpreCopia(1) = "Original"
    ImpreCopia(2) = "Duplicado"
    
    For da = 1 To 2
        If WOtro = "S" Then
            Print #1, WEmpNombre;
            Print #1, Tab(50); "Nro.Certificado  : "; WCertificado
                Else
            Print #1, WEmpNombre;
            Print #1, Tab(50); "COPIA"
        End If
        Print #1, WEmpDireccion;
        Print #1, Tab(50); "Fecha Emision    : "; Fecha.Text
        Print #1, WEmpLocalidad;
        Print #1, Tab(50); ImpreCopia(da)
        Print #1, "C.U.I.T.    : ", WEmpCuit
        Print #1, "Ing. Brutos : ", WNroIb
        Print #1, "Nro. Agente : ", WNroAgente
        Print #1, "-----------------------------------------------------------------------------------------"
        Print #1, "SUJETO RETENIDO"
        Print #1, "Nombre/Razon Social  : "; DesProveedor.Caption
        Print #1, "Domicilio            : "; WPrvDireccion
        Print #1, "Nro. C.U.I.T.        : "; WPrvCuit
        Print #1, "Nro. Ingresos Brutos : "; WPrvIb
        Print #1, "Orden de Pago        : "; Orden.Text
        Print #1, "-----------------------------------------------------------------------------------------"
        Print #1, Tab(1); "Comprobante";
        Print #1, Tab(20); "Fecha";
        Print #1, Tab(35); "Categoria";
        Print #1, Tab(65); "Importe Pago";
        Print #1, Tab(85); "Porc.";
        Print #1, Tab(95); "Importe Retenido"
        Print #1, ""
        
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
                        Case 99
                            XImpre1 = "Varios"
                        Case Else
                            XImpre1 = ""
                    End Select
                                
                    DbGrid1.Col = 3
                    XImpre2 = Left$(DbGrid1.Text, 8)
                    
                    spIvaComp = "ConsultaIvacomp " + "'" + XImpre2 + "'"
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                           XImpre2 = rstIvaComp!Numero
                           rstIvaComp.Close
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
                        XImpre3 = RstCtaPrv!Fecha
                        RstCtaPrv.Close
                            Else
                        XImpre3 = ""
                    End If
                        
                    DbGrid1.Col = 4
                    WImpre4 = Val(DbGrid1.Text)
                    If Val(WTipo) = 3 Or Val(WTipo) = 5 Then
                       WImpre4 = WImpre4 * -1
                    End If
                    If WTipoiva = 2 Then
                        WImpre4 = WImpre4 / 1.21
                    End If
                    Call Redondeo(WImpre4)
                    XImpre4 = Str$(WImpre4)
                    
                    Select Case WTipoIb
                        Case 0
                            WRete = Val(XImpre4) * (0.75 / 100)
                            Call Redondeo(WRete)
                            Print #1, Tab(1); XImpre1 + " " + XImpre2;
                            Print #1, Tab(20); XImpre3;
                            Print #1, Tab(35); "SUJETO A RETENCION 0.75%";
                            Print #1, Tab(65); Alinea("###,###.##", XImpre4);
                            Print #1, Tab(85); "0.75";
                            Print #1, Tab(95); Alinea("###,###.##", Str$(WRete));
                            WImpoRetenido = WImpoRetenido + WRete
                            
                        Case Else
                            WRete = Val(XImpre4) * (1.75 / 100)
                            Call Redondeo(WRete)
                            Print #1, Tab(1); XImpre1 + " " + XImpre2;
                            Print #1, Tab(20); XImpre3;
                            Print #1, Tab(35); "SUJETO A RETENCION 1.75%";
                            Print #1, Tab(65); Alinea("###,###.##", XImpre4);
                            Print #1, Tab(85); "1.75";
                            Print #1, Tab(95); Alinea("###,###.##", Str$(WRete));
                            WImpoRetenido = WImpoRetenido + WRete
                            
                    End Select
                    
                End If
                Print #1, ""
        Next iRow
        
        Print #1, Tab(20); "---------------------------------";
        Print #1, Tab(70); "Total Retenido : ";
        Print #1, Tab(95); Alinea("###,###.##", Str$(WImpoRetenido))
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
    Next da
    Print #1, Chr$(12)
    
    Close #1

End Sub



Rem
Rem Controles de la grilla
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.Visible = True
            WTexto1.SetFocus
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.Visible = True
            Rem WTexto2.SetFocus
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            WTexto3.Visible = True
            WTexto3.SetFocus
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
            End If
        End If
        Rem Call Calcula_Click
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub


Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 6
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case 12
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 7
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Tipo1.Value = True Then
                If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                        Else
                    WControl = "N"
                End If
                    Else
                If Val(WVector1.Text) = 0 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                    WVector1.Col = 4
                        Else
                    WControl = "N"
                End If
            End If
            
        Case 2
            WVector1.Text = Left$(WVector1.Text, 1)
            If Tipo1.Value = True Then
                If WVector1.Text = "A" Or WVector1.Text = "C" Or WVector1.Text = "X" Or WVector1.Text = "E" Then
                    WControl = "S"
                        Else
                    WControl = "N"
                End If
            End If
                
        Case 3
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 4)
            WVector1.Text = Auxi$
            
        Case 4
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
                    
            ClaveCtaprv = Proveedor.Text
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 2)
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 1)
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 3)
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 4)
            
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                If Val(WVector1.TextMatrix(WVector1.Row, 5)) = 0 Then
                    WVector1.TextMatrix(WVector1.Row, 5) = RstCtaPrv!Saldo
                    RstCtaPrv.Close
                    Call Suma_Datos
                End If
                Rem WVector1.Col = 4
                    Else
                WControl = "N"
            End If
            
        
        Case 8
            WVector1.Col = 7
            If Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Then
                WVector1.Col = 8
                WControl = "N"
                    Else
                WVector1.Col = 8
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 8)
                WVector1.Text = Auxi$
            End If
                
        Case 9
            If Len(WVector1.Text) = 5 Then
                If Right$(WVector1.Text, 2) < 6 Then
                    WVector1.Text = WVector1.Text + "/2014"
                        Else
                    WVector1.Text = WVector1.Text + Right$(Fecha.Text, 5)
                End If
            End If
            Call Valida_fecha1(WVector1.Text, Auxi)
            If Auxi <> "S" Then
                WControl = "N"
                WControl = "N"
                    Else
                If Val(WVector1.TextMatrix(WVector1.Row, 10)) <> 0 Then
                    WVector1.Col = WVector1.Col + 2
                End If
            End If
                
        Case 10
            ClaveBanco = WVector1.Text
            spBanco = "ConsultaBanco " + "'" + ClaveBanco + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                WVector1.Col = 11
                WVector1.Text = rstBanco!Nombre
                rstBanco.Close
                    Else
                WControl = "N"
            End If

        Case 12
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
            End If
            Call Suma_Datos
            
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 15
    WVector1.FixedRows = 1
    WVector1.Rows = 14
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Letra"
                WVector1.ColWidth(Ciclo) = 550
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Punto"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 6
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1850
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 40
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 12
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 13
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 20
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 14
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 20
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 3
        WVector1.Col = Ciclo
        WTituloVector(Ciclo).Text = WVector1.Text
        WTituloVector(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTituloVector(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTituloVector(Ciclo).Width = WVector1.CellWidth
        WTituloVector(Ciclo).Height = WVector1.CellHeight
        WTituloVector(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = 11400
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
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
        Case Else
    End Select

End Sub



