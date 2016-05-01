VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRec2 
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
   Begin VB.Frame Ingrecuenta 
      Caption         =   "Ingreso de Cuenta Contable"
      Height          =   1095
      Left            =   3120
      TabIndex        =   34
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox Cuenta1 
         Height          =   285
         Left            =   480
         MaxLength       =   10
         TabIndex        =   35
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox Cuenta 
      Height          =   285
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   33
      Text            =   " "
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Impre 
      Caption         =   "Impresion"
      Height          =   300
      Left            =   9600
      TabIndex        =   30
      Top             =   120
      Width           =   975
   End
   Begin Crystal.CrystalReport listado 
      Left            =   10080
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "recibo.rpt"
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
      Left            =   6840
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
      TabIndex        =   20
      Top             =   1080
      Width           =   5295
      Begin VB.OptionButton Tipo3 
         Caption         =   "Varios"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Cobro de Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Anticipos"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
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
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   6840
      TabIndex        =   18
      Top             =   0
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
      Height          =   5415
      Left            =   0
      OleObjectBlob   =   "rec2.frx":0000
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8520
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "rec2.frx":09C2
      Left            =   5520
      List            =   "rec2.frx":09C9
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   9600
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   9600
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   9600
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   9600
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   9600
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Cuenta Contable"
      Height          =   255
      Left            =   5520
      TabIndex        =   32
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   8280
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Debitos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.   2) Ch.   3) Doc.  4) Varios"
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   7680
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Otra Retencion"
      Height          =   255
      Left            =   5520
      TabIndex        =   25
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Ret.Iva"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Rte.Ganancias"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label DesClientes 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   2520
      TabIndex        =   19
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod. Cilente"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero de Recibo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "PrgRec2"
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
Private WProv As String
Private WCuenta(20) As String
Private Debito As Double
Private Credito As Double

Private Sub Suma_Datos()
    Debitos.Caption = ""
    Creditos.Caption = ""
    
    Creditos.Caption = Str$(Val(Retganancias.Text) + Val(RetIva.Text) + Val(RetOtra.Text))
    For iRow = 0 To 19
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Debitos.Caption = Str$(Val(Debitos.Caption) + Val(DbGrid1.Text))
        End If
        DbGrid1.Col = 9
        DbGrid1.Row = iRow
        If Val(DbGrid1.Text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(DbGrid1.Text))
        End If
    Next iRow
    Debitos.Caption = Alinea("###,###.##", Debitos.Caption)
    Creditos.Caption = Alinea("###,###.##", Creditos.Caption)
    DbGrid1.Col = 0
    DbGrid1.Row = 0
End Sub

Private Sub Lee_Datos()

    For iRow = 0 To 19
        For iCol = 0 To 9
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iCol
    Next iRow

    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        With rstRecibos
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", Recibo.Text + Auxi1
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
                        DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                    Case 2
                        Credito = Credito + 1
                        DbGrid1.Row = Credito - 1
                        DbGrid1.Col = 5
                        DbGrid1.Text = !Tipo2
                        DbGrid1.Col = 6
                        DbGrid1.Text = !Numero2
                        DbGrid1.Col = 7
                        DbGrid1.Text = !Fecha2
                        DbGrid1.Col = 8
                        DbGrid1.Text = !Banco2
                        DbGrid1.Col = 9
                        DbGrid1.Text = !Importe2
                        DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                    Case Else
                End Select
                    Else
                Exit Do
            End If
        End With
    Loop
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
    Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
    RetIva.Text = Alinea("###,###.##", RetIva.Text)
    RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
End Sub

Sub Imprime_Datos()
    With rstClientes
        .Index = "Cliente"
        .Seek "=", Clientes.Text
        If .NoMatch = False Then
            Clientes.Text = !Cliente
            DesClientes.Caption = !Razon
            WRazon = !Razon
            WDireccion = !Direccion
            WLocalidad = !Localidad
            WPostal = !Postal
            WProvincia = Provincia(!Provincia)
            WProv = !Provincia
            Call Format_datos
        End If
    End With
End Sub

Private Sub cmdAdd_Click()

    If Recibo.Text <> "" And Fecha.Text <> "" Then
    
    Auxi1 = Recibo.Text
    Call Ceros(Auxi1, 6)
    Recibo.Text = Auxi1
        
    With rstRecibos
        Existe = "N"
        .Index = "Clave"
        Claveven$ = Recibo.Text + "01"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Existe = "S"
        End If
    End With
    
    If Existe <> "sS" Then
    
        Call Suma_Datos
        
        Debito = 0
        Credito = 0
        If Val(Debitos.Caption) <> 0 Then
            Debito = Val(Debitos.Caption)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = Val(Creditos.Caption)
        End If
        
        Call Redondeo(Debito)
        Call Redondeo(Credito)
        
        If Debito = Credito Or Tipo2.Value = True Or Tipo3.Value = True Then
    
        With rstRecibos
            Renglon = 0
            .Index = "Clave"
            For iRow = 0 To 19
        
                If Tipo1.Value = True Then
                    WRow = iRow
                    DbGrid1.Col = 4
                    DbGrid1.Row = iRow
                    If Val(DbGrid1.Text) <> 0 Then
                        DbGrid1.Col = 0
                        WTipo = DbGrid1.Text
                        DbGrid1.Col = 1
                        WLetra = DbGrid1.Text
                        DbGrid1.Col = 2
                        WPunto = DbGrid1.Text
                        DbGrid1.Col = 3
                        WNumero = DbGrid1.Text
                        DbGrid1.Col = 4
                        WImporte = Val(DbGrid1.Text)
                    
                        With rstCtaCte
                            .Index = "Clave"
                            Auxi$ = Clientes.Text
                            Call Ceros(Auxi$, 6)
                            Claveven$ = Auxi$
                            Claveven$ = WTipo + WNumero + "01"
                            .Seek "=", Claveven$
                            If .NoMatch = False Then
                                .Edit
                                WSaldo = !Saldo - WImporte
                                Call Redondeo(WSaldo)
                                !Saldo = WSaldo
                                WSaldo = !SaldoUs - WImporte
                                Call Redondeo(WSaldo)
                                !SaldoUs = WSaldo
                                !Wdate = Date$
                                .Update
                                .Bookmark = .LastModified
                            End If
                        End With
                        
                    End If
                End If
                
            Next iRow
        End With
        
        If Tipo1.Value = True Then
            With rstCtaCte
                WTipo = "06"
                WNumero = "00" + Recibo.Text
                .Index = "Clave"
                .Seek "=", WTipo + WNumero + "01"
                If .NoMatch Then
                    .AddNew
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Renglon = "01"
                    !Cliente = Clientes.Text
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = Fecha.Text
                    !Vencimiento1 = Fecha.Text
                    !Total = Credito * -1
                    !TotalUs = Credito * -1
                    !Saldo = 0
                    !SaldoUs = 0
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento1 = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Impre = "RC"
                    !Neto = 0
                    !Iva1 = 0
                    !Iva2 = 0
                    !Pedido = ""
                    !Remito = ""
                    !Orden = ""
                    !Paridad = 1
                    !Provincia = WProv
                    !Vendedor = WVendedor
                    !Rubro = WRubro
                    !Comprobante = ""
                    !Aceptada = ""
                    !Costo = 0
                    !Importe1 = 0
                    !Importe2 = 0
                    !Importe3 = 0
                    !Importe4 = 0
                    !Importe5 = 0
                    !Importe6 = 0
                    !Importe7 = 0
                    Auxi = WNumero
                    Call Ceros(Auxi, 8)
                    !Clave = WTipo + Auxi + "01"
                    !Wdate = Date$
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Renglon = "01"
                    !Cliente = Clientes.Text
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = Fecha.Text
                    !Vencimiento1 = Fecha.Text
                    !Total = Credito * -1
                    !TotalUs = Credito * -1
                    !Saldo = 0
                    !SaldoUs = 0
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento1 = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Impre = "RC"
                    !Neto = 0
                    !Iva1 = 0
                    !Iva2 = 0
                    !Pedido = ""
                    !Remito = ""
                    !Orden = ""
                    !Paridad = 1
                    !Provincia = WProv
                    !Vendedor = WVendedor
                    !Rubro = WRubro
                    !Comprobante = ""
                    !Aceptada = ""
                    !Costo = 0
                    !Importe1 = 0
                    !Importe2 = 0
                    !Importe3 = 0
                    !Importe4 = 0
                    !Importe5 = 0
                    !Importe6 = 0
                    !Importe7 = 0
                    Auxi = Recibo.Text
                    Call Ceros(Auxi, 8)
                    !Clave = WTipo + Auxi + "01"
                    !Wdate = Date$
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
        End If
        
        If Tipo2.Value = True Then
            With rstCtaCte
                WTipo = "07"
                WNumero = "00" + Recibo.Text
                .Index = "Clave"
                .Seek "=", WTipo + WNumero + "01"
                If .NoMatch Then
                    .AddNew
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Renglon = "01"
                    !Cliente = Clientes.Text
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = Fecha.Text
                    !Vencimiento1 = Fecha.Text
                    !Total = Credito * -1
                    !TotalUs = Credito * -1
                    !Saldo = Credito * -1
                    !SaldoUs = Credito * -1
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento1 = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Impre = "AN"
                    !Neto = 0
                    !Iva1 = 0
                    !Iva2 = 0
                    !Pedido = ""
                    !Remito = ""
                    !Orden = ""
                    !Paridad = 1
                    !Provincia = WProv
                    !Vendedor = WVendedor
                    !Rubro = WRubro
                    !Comprobante = ""
                    !Aceptada = ""
                    !Costo = 0
                    !Importe1 = 0
                    !Importe2 = 0
                    !Importe3 = 0
                    !Importe4 = 0
                    !Importe5 = 0
                    !Importe6 = 0
                    !Importe7 = 0
                    Auxi = Recibo.Text
                    Call Ceros(Auxi, 8)
                    !Clave = WTipo + Auxi + "01"
                    !Wdate = Date$
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Renglon = "01"
                    !Cliente = Clientes.Text
                    !Fecha = Fecha.Text
                    !Estado = "1"
                    !Vencimiento = Fecha.Text
                    !Vencimiento1 = Fecha.Text
                    !Total = Credito * -1
                    !TotalUs = Credito * -1
                    !Saldo = Credito * -1
                    !SaldoUs = Credito * -1
                    !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !OrdVencimiento1 = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Impre = "AN"
                    !Neto = 0
                    !Iva1 = 0
                    !Iva2 = 0
                    !Pedido = ""
                    !Remito = ""
                    !Orden = ""
                    !Paridad = 1
                    !Provincia = WProv
                    !Vendedor = WVendedor
                    !Rubro = WRubro
                    !Comprobante = ""
                    !Aceptada = ""
                    !Costo = 0
                    !Importe1 = 0
                    !Importe2 = 0
                    !Importe3 = 0
                    !Importe4 = 0
                    !Importe5 = 0
                    !Importe6 = 0
                    !Importe7 = 0
                    Auxi = Recibo.Text
                    Call Ceros(Auxi, 8)
                    !Clave = WTipo + Auxi + "01"
                    !Wdate = Date$
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
            
            With rstRecibos
                .AddNew
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                !Recibo = Recibo.Text
                !Renglon = Auxi1
                !Cliente = Clientes.Text
                !Fecha = Fecha.Text
                !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Tipo1.Value = True Then
                    !TipoRec = "1"
                End If
                If Tipo2.Value = True Then
                    !TipoRec = "2"
                End If
                !Retganancias = Val(Retganancias.Text)
                !RetIva = Val(RetIva.Text)
                !RetOtra = Val(RetOtra.Text)
                !Retencion = 0
                !Tiporeg = "1"
                DbGrid1.Col = 0
                !Tipo1 = "07"
                DbGrid1.Col = 1
                !Letra1 = ""
                DbGrid1.Col = 2
                !Punto1 = ""
                DbGrid1.Col = 3
                !Numero1 = Recibo.Text
                DbGrid1.Col = 4
                !Importe1 = Credito
                !Tipo2 = ""
                !Numero2 = ""
                !Fecha2 = ""
                !FechaOrd2 = ""
                !Banco2 = ""
                !Importe2 = 0
                !Estado2 = ""
                !Observaciones = Observaciones.Text
                !Empresa = 1
                !Clave = !Recibo + !Renglon
                !Importe = Credito
                !Cuenta = ""
                .Update
                .Bookmark = .LastModified
            End With
            
        End If
        
        If Tipo3.Value = True Then
            With rstRecibos
                .AddNew
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                !Recibo = Recibo.Text
                !Renglon = Auxi1
                !Cliente = Clientes.Text
                !Fecha = Fecha.Text
                !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Tipo1.Value = True Then
                    !TipoRec = "1"
                End If
                If Tipo2.Value = True Then
                    !TipoRec = "2"
                End If
                If Tipo3.Value = True Then
                    !TipoRec = "3"
                End If
                !Retganancias = Val(Retganancias.Text)
                !RetIva = Val(RetIva.Text)
                !RetOtra = Val(RetOtra.Text)
                !Retencion = 0
                !Tiporeg = "1"
                DbGrid1.Col = 0
                !Tipo1 = "99"
                DbGrid1.Col = 1
                !Letra1 = ""
                DbGrid1.Col = 2
                !Punto1 = ""
                DbGrid1.Col = 3
                !Numero1 = Recibo.Text
                DbGrid1.Col = 4
                !Importe1 = Credito
                !Tipo2 = ""
                !Numero2 = ""
                !Fecha2 = ""
                !FechaOrd2 = ""
                !Banco2 = ""
                !Importe2 = 0
                !Estado2 = ""
                !Observaciones = Observaciones.Text
                !Empresa = 1
                !Clave = !Recibo + !Renglon
                !Importe = Credito
                !Cuenta = Cuenta.Text
                .Update
                .Bookmark = .LastModified
            End With
        End If
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = "1"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                WCtaRetGan = !CtaRetGan
                WctaRetIva = !ctaRetIva
                WCtaretOtra = !CtaretOtro
                WCtaDeudores = !Ctadeudores
                WCtaEfectivo = !CtaEfectivo
                WCtaCheques = !CtaCheque
                WCtaDocumentos = !CtaDocumentos
                WctaTerceros = !CtaTerceros
            End If
        End With
        
        Rem With rstImputac
        Rem    Renglon = 0
        Rem    .Index = "Clave"
        Rem
        Rem    If Val(Retganancias.Text) <> 0 Then
        Rem        .AddNew
        Rem        !Tipomovi = "1"
        Rem        !Proveedor = "000000"
        Rem        !TipoComp = "01"
        Rem        !LetraComp = "A"
        Rem        !PuntoComp = "0000"
        Rem        !NroComp = Recibo.Text
        Rem        Renglon = Renglon + 1
        Rem        Auxi1 = Str$(Renglon)
        Rem        Call Ceros(Auxi1, 2)
        Rem        !Renglon = Auxi1$
        Rem        !Fecha = Fecha.Text
        Rem        !Observaciones = Left$(DesClientes.Caption, 30)
        Rem        !Cuenta = WCtaRetGan
        Rem        !Debito = Val(Retganancias.Text)
        Rem        !Credito = 0
        Rem        !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem        !Titulo = "Cobranzas"
        Rem        !Empresa = 1
        Rem        !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
        Rem        .Update
        Rem    End If
        Rem
        Rem    If Val(RetIva.Text) <> 0 Then
        Rem        .AddNew
        Rem        !Tipomovi = "1"
        Rem        !Proveedor = "000000"
        Rem        !TipoComp = "01"
        Rem        !LetraComp = "A"
        Rem        !PuntoComp = "0000"
        Rem        !NroComp = Recibo.Text
        Rem        Renglon = Renglon + 1
        Rem        Auxi1 = Str$(Renglon)
        Rem        Call Ceros(Auxi1, 2)
        Rem        !Renglon = Auxi1$
        Rem        !Fecha = Fecha.Text
        Rem        !Observaciones = Left$(DesClientes.Caption, 30)
        Rem        !Cuenta = WctaRetIva
        Rem        !Debito = Val(RetIva.Text)
        Rem        !Credito = 0
        Rem        !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem        !Titulo = "Cobranzas"
        Rem        !Empresa = 1
        Rem        !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
        Rem        .Update
        Rem    End If
        Rem
        Rem    If Val(RetOtra.Text) <> 0 Then
        Rem        .AddNew
        Rem        !Tipomovi = "1"
        Rem        !Proveedor = "000000"
        Rem        !TipoComp = "01"
        Rem        !LetraComp = "A"
        Rem        !PuntoComp = "0000"
        Rem        !NroComp = Recibo.Text
        Rem        Renglon = Renglon + 1
        Rem        Auxi1 = Str$(Renglon)
        Rem        Call Ceros(Auxi1, 2)
        Rem        !Renglon = Auxi1$
        Rem       !Fecha = Fecha.Text
        Rem        !Observaciones = Left$(DesClientes.Caption, 30)
        Rem        !Cuenta = WCtaretOtra
        Rem        !Debito = Val(RetOtra.Text)
        Rem        !Credito = 0
        Rem        !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem        !Titulo = "Cobranzas"
        Rem        !Empresa = 1
        Rem        !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
        Rem        .Update
        Rem    End If
        Rem
        Rem    If Tipo2.Value = True Then
        Rem        .AddNew
        Rem        !Tipomovi = "1"
        Rem        !Proveedor = "000000"
        Rem        !TipoComp = "01"
        Rem        !LetraComp = "A"
        Rem        !PuntoComp = "0000"
        Rem        !NroComp = Recibo.Text
        Rem        Renglon = Renglon + 1
        Rem        Auxi1 = Str$(Renglon)
        Rem        Call Ceros(Auxi1, 2)
        Rem        !Renglon = Auxi1$
        Rem        !Fecha = Fecha.Text
        Rem        !Observaciones = Left$(DesClientes.Caption, 30)
        Rem        !Cuenta = WctaTerceros
        Rem        !Debito = 0
        Rem        !Credito = Credito
        Rem        !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem        !Titulo = "Cobranzas"
        Rem        !Empresa = 1
        Rem        !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
        Rem        .Update
        Rem    End If

        Rem    For iRow = 0 To 19
        Rem        WRow = iRow
        Rem        DbGrid1.Col = 4
        Rem        DbGrid1.Row = iRow
        Rem        If Val(DbGrid1.Text) <> 0 Then
        Rem            .AddNew
        Rem            !Tipomovi = "1"
        Rem            !Proveedor = "000000"
        Rem            !TipoComp = "01"
        Rem            !LetraComp = "A"
        Rem            !PuntoComp = "0000"
        Rem            !NroComp = Recibo.Text
        Rem            Renglon = Renglon + 1
        Rem            Auxi1 = Str$(Renglon)
        Rem            Call Ceros(Auxi1, 2)
        Rem            !Renglon = Auxi1$
        Rem            !Fecha = Fecha.Text
        Rem            !Observaciones = Left$(DesClientes.Caption, 30)
        Rem            !Cuenta = WCtaDeudores
        Rem            !Debito = 0
        Rem            !Credito = Val(DbGrid1.Text)
        Rem            !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem            !Titulo = "Cobranzas"
        Rem            !Empresa = 1
        Rem            !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
        Rem            .Update
        Rem        End If
        Rem
        Rem        DbGrid1.Col = 9
        Rem        DbGrid1.Row = iRow
        Rem        If Val(DbGrid1.Text) <> 0 Then
        Rem            .AddNew
        Rem            !Tipomovi = "1"
        Rem            !Proveedor = "000000"
        Rem            !TipoComp = "01"
        Rem            !LetraComp = "A"
        Rem            !PuntoComp = "0000"
        Rem            !NroComp = Recibo.Text
        Rem            Renglon = Renglon + 1
        Rem            Auxi1 = Str$(Renglon)
        Rem            Call Ceros(Auxi1, 2)
        Rem            !Renglon = Auxi1$
        Rem            !Fecha = Fecha.Text
        Rem            !Observaciones = Left$(DesClientes.Caption, 30)
        Rem            !Debito = Val(DbGrid1.Text)
        Rem            !Credito = 0
        Rem            !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Rem            !Titulo = "Cobranzas"
        Rem            !Empresa = 1
        Rem            DbGrid1.Col = 5
        Rem            Select Case Val(DbGrid1.Text)
        Rem                Case 2
        Rem                    !Cuenta = WCtaCheques
        Rem                Case 3
        Rem                    !Cuenta = WCtaDocumentos
        Rem                Case Else
        Rem                    !Cuenta = WCtaEfectivo
        Rem            End Select
        Rem            !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
        Rem            .Update
        Rem        End If
        Rem
        Rem    Next iRow
        Rem End With
        
        Rem listado.GroupSelectionFormula = "{Recibos.recibo} in " + Chr$(34) + Recibo.Text + Chr$(34) + " to " + Chr$(34) + Recibo.Text + Chr$(34)
        Rem listado.Destination = 1
        Rem Listado.Action = 1
        
        Rem Call Impresion

        Call CmdLimpiar_Click
        Recibo.SetFocus
        
        End If
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Recibo.Text <> "" Then
                
            Rem Borro los datos anteriores
            
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Recibo.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow
    End If
    Clientes.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    For iRow = 0 To 19
        For iCol = 0 To 9
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
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
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Cuenta.Text = ""
    
    Ingrecuenta.Visible = False
    Erase WCuenta
    Pantalla.Visible = False
    Opcion.Visible = False
    
    With rstRecibos
        .Index = "Clave"
        .Seek "<", "99999999"
        If .NoMatch = False Then
            Recibo.Text = !Recibo + 1
                Else
            Recibo.Text = ""
        End If
    End With
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    Rem With rstImputac
    Rem    .Close
    Rem End With
    With rstClientes
        .Close
    End With
    With rstRecibos
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    DbsAdminis.Close
    DbsVentas.Close
    Recibo.SetFocus
    PrgRec2.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Impre_Click()
    Call Impresion
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        With rstRecibos
            Existe = "N"
            .Index = "Clave"
            Claveven$ = Recibo.Text + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Existe = "S"
                Clientes.Text = !Cliente
                Observaciones.Text = !Observaciones
                Fecha.Text = !Fecha
                Retganancias.Text = !Retganancias
                RetIva.Text = !RetIva
                RetOtra.Text = !RetOtra
                Tipo1.Value = True
                Tipo2.Value = False
                Select Case Val(!TipoRec)
                    Case 1
                        Tipo1.Value = True
                    Case 2
                        Tipo2.Value = True
                    Case Else
                End Select
                
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
            Clientes.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Clientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Clientes.Text <> "" Then
            With rstClientes
                .Index = "Cliente"
                Claveven$ = Clientes.Text
                .Seek "=", Clientes.Text
                If .NoMatch Then
                    Clientes.SetFocus
                        Else
                    Clientes.Text = !Cliente
                    DesClientes.Caption = !Razon
                    WRazon = !Razon
                    WDireccion = !Direccion
                    WLocalidad = !Localidad
                    WPostal = !Postal
                    WProvincia = Provincia(!Provincia)
                    WProv = !Provincia
                    Rem Call Imprime_Datos
                    Observaciones.SetFocus
                End If
            End With
        End If
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.SetFocus
    End If
End Sub

Private Sub Retganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
        Call Suma_Datos
        RetIva.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetIva.Text = Alinea("###,###.##", RetIva.Text)
        Call Suma_Datos
        RetOtra.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
        Call Suma_Datos
        DbGrid1.Col = 0
        DbGrid1.Row = 0
        DbGrid1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    XRow = DbGrid1.Row
    XCol = DbGrid1.Col

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Cuenta Corrientes"

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
            With rstClientes
                .Index = "Cliente"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Cliente + "     " + !Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case 1
            With rstCtaCte
    
                .Index = "Cliente"
                .Seek ">=", Clientes.Text
                If .NoMatch = False Then
                    Do
            
                        If Clientes.Text <> !Cliente Then
                            Exit Do
                        End If
                        
                        If !Saldo <> 0 Then
                            Auxi = Str$(!Saldo)
                            Auxi = Mascara("###,###.##", Auxi$)
                            Auxi1 = Str$(!Numero)
                            Call Ceros(Auxi1, 6)
                            IngresaItem = !Impre + " " + Auxi1 + " " + !Fecha + " " + Auxi
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Clave
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                        
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If Clientes.Text <> !Cliente Then
                            Exit Do
                        End If
                        
                    Loop
                End If
            End With
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            With rstClientes
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Clientes.Text = Claveven$
                .Index = "Cliente"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DesClientes.Caption = !Razon
                    WRazon = !Razon
                    WDireccion = !Direccion
                    WLocalidad = !Localidad
                    WPostal = !Postal
                    WProvincia = Provincia(!Provincia)
                    WProv = !Provincia
                                  Else
                    Clientes.Text = ""
                End If
            End With
                
            Clientes.SetFocus
            
        Case 1
        
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For iRow = 0 To 19
                DbGrid1.Row = iRow
                DbGrid1.Col = 0
                Compara2 = DbGrid1.Text
                DbGrid1.Col = 3
                Compara2 = Compara2 + DbGrid1.Text + "01"
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next iRow
            
            If Entra = "S" Then
            
            For iRow = 0 To 19
                DbGrid1.Row = iRow
                DbGrid1.Col = 0
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
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 0
                    Auxi = !Tipo
                    Call Ceros(Auxi, 2)
                    DbGrid1.Text = Auxi
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 1
                    DbGrid1.Text = ""
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 2
                    DbGrid1.Text = ""
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 3
                    Auxi = !Numero
                    Call Ceros(Auxi, 8)
                    DbGrid1.Text = Auxi
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    DbGrid1.Text = !Saldo
                    DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                End If
            End With
            
            End If
                
            DbGrid1.Row = XRow
            DbGrid1.Col = 0
            DbGrid1.SetFocus
                
        Case Else
    End Select
    
End Sub
Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
    
            Case 0
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Then
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
                
            Case 1
                Rem If KeyCode = 13 Then
                Rem     DBGrid1.Text = Left$(DBGrid1.Text, 1)
                Rem     If DBGrid1.Text = "A" Or DBGrid1.Text = "C" Then
                Rem         DBGrid1.Col = 2
                Rem         KeyCode = 0
                Rem         Rem no hago anda
                Rem             Else
                Rem         DBGrid1.Col = 1
                Rem         KeyCode = 0
                Rem     End If
                Rem End If
                
            Case 2
                Rem If KeyCode = 13 Then
                Rem     Auxi$ = Str$(Val(DBGrid1.Text))
                Rem     Call Ceros(Auxi$, 4)
                Rem     DBGrid1.Text = Auxi$
                Rem     DBGrid1.Col = 3
                Rem     KeyCode = 0
                Rem End If
                
            Case 3
                If KeyCode = 13 Then
                
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.Text = Auxi$
                
                    With rstCtaCte
                        .Index = "Clave"
                        Rem Auxi$ = Clientes.Text
                        Rem Call Ceros(Auxi$, 6)
                        Rem Claveven$ = Auxi$
                        Rem DBGrid1.Col = 1
                        Rem Claveven$ = Claveven$ + DBGrid1.Text
                        DbGrid1.Col = 0
                        Claveven$ = DbGrid1.Text
                        Rem DBGrid1.Col = 2
                        Rem Claveven$ = Claveven$ + DBGrid1.Text
                        DbGrid1.Col = 3
                        Claveven$ = Claveven$ + DbGrid1.Text + "01"
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
                End If
                
            Case 4
                If KeyCode = 13 Then
                    With rstCtaCte
                        .Index = "Clave"
                        Rem Auxi$ = Clientes.Text
                        Rem Call Ceros(Auxi$, 6)
                        Rem Claveven$ = Auxi$
                        Rem DBGrid1.Col = 1
                        Rem Claveven$ = Claveven$ + DBGrid1.Text
                        DbGrid1.Col = 0
                        Claveven$ = DbGrid1.Text
                        Rem DBGrid1.Col = 2
                        Rem Claveven$ = Claveven$ + DBGrid1.Text
                        DbGrid1.Col = 3
                        Claveven$ = Claveven$ + DbGrid1.Text + "01"
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Saldo = Alinea("###,###.##", Str$(!Saldo))
                                Else
                            Saldo = 0
                        End If
                    End With
                
                    DbGrid1.Col = 4
                    If Val(DbGrid1.Text) > Val(Saldo) Then
                        DbGrid1.Text = ""
                        DbGrid1.Col = 4
                        KeyCode = 0
                            Else
                        DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                        Call Suma_Datos
                        If DbGrid1.Row < 10 Then
                            DbGrid1.Row = DbGrid1.Row + 1
                            DbGrid1.Col = 0
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End If
                End If
                
            Case 5
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 2 Or Val(DbGrid1.Text) = 3 Or Val(DbGrid1.Text) = 4 Then
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.Text = Auxi$
                        Select Case Val(DbGrid1.Text)
                            Case 1, 4
                                DbGrid1.Col = 6
                                DbGrid1.Text = ""
                                DbGrid1.Col = 7
                                DbGrid1.Text = ""
                                DbGrid1.Col = 8
                                DbGrid1.Text = ""
                                DbGrid1.Col = 9
                                KeyCode = 0
                            Case Else
                                DbGrid1.Col = 6
                                KeyCode = 0
                        End Select
                            Else
                        DbGrid1.Col = 5
                        KeyCode = 0
                    End If
                End If
                
            Case 6
                If KeyCode = 13 Then
                    Auxi$ = Str$(Val(DbGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.Text = Auxi$
                    DbGrid1.Col = 7
                    KeyCode = 0
                
                End If
                
            Case 7
                If KeyCode = 13 Then
                    DbGrid1.Col = 7
                    Call Valida_fecha1(DbGrid1.Text, Auxi)
                    If Auxi <> "S" Then
                        DbGrid1.Col = 7
                        KeyCode = 0
                                Else
                        DbGrid1.Col = 8
                        KeyCode = 0
                    End If
                End If
                
            Case 8
                If KeyCode = 13 Then
                    DbGrid1.Col = 9
                    KeyCode = 0
                End If
                
            Case 9
                If KeyCode = 13 Then
                    iRow = DbGrid1.Row
                    DbGrid1.Col = 5
                    XTipo = DbGrid1.Text
                    DbGrid1.Col = 9
                    DbGrid1.Text = Alinea("###,###.##", DbGrid1.Text)
                    Call Suma_Datos
                    DbGrid1.Row = iRow
                    If Val(XTipo) = 4 Then
                        Cuenta1.Text = WCuenta(DbGrid1.Row)
                        Ingrecuenta.Visible = True
                        Cuenta1.SetFocus
                    End If
                    If DbGrid1.Row < 10 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 5
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 5
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

ReDim UserData(0 To 9, 0 To 19)

mTotalRows& = 20

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
For i = 0 To 9
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
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
             DbGrid1.Columns(newcnt).Locked = True
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 5
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 6
             DbGrid1.Columns(newcnt).Caption = "Numero/Cta"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 7
             DbGrid1.Columns(newcnt).Caption = "Fecha"
             DbGrid1.Columns(newcnt).Width = 1300
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 8
             DbGrid1.Columns(newcnt).Caption = "Banco"
             DbGrid1.Columns(newcnt).Width = 1500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 9
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
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

    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = "0"
    RetIva.Text = "0"
    RetOtra.Text = "0"
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Observaciones.Text = ""
    Cuenta.Text = ""
    
    With rstRecibos
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Recibo.Text = !Recibo + 1
                Else
            Recibo.Text = ""
        End If
    End With

End Sub


Sub Impresion()

        If Val(WEmpresa) = 2 Then
            Open "lpt2" For Output As #1
                Else
            Open "lpt1" For Output As #1
        End If

        For aa = 1 To 2

        Retencion = Val(Retganancias.Text) + Val(RetIva.Text) + Val(RetOtra.Text)

        Dolares = 0
        Pesos = 0
        Cheque = 0
        Documento = 0
        Total2 = 0
        
        Erase Vector
        
        For iRow = 0 To 19
        
                DbGrid1.Row = iRow
                
                DbGrid1.Col = 0
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 0) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 1
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 1) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 2
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 2) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 3
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 3) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 4
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 4) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 5
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 5) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 6
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 6) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 7
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 7) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 8
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 8) = DbGrid1.Text
                End If
                
                DbGrid1.Col = 9
                If DbGrid1.Text <> "" Then
                    Vector(iRow, 9) = DbGrid1.Text
                End If
                
                With rstCtaCte
                    .Index = "Clave"
                    WTipo = Vector(iRow, 0)
                    WNumero = Vector(iRow, 3)
                    Call Ceros(WTipo, 2)
                    Call Ceros(WNumero, 8)
                    Claveven$ = WTipo + WNumero + "01"
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        Vector(iRow, 10) = !Fecha
                    End If
                End With
                
        Next iRow

        For Ciclo = 0 To 19

                If Val(Vector(Ciclo, 9)) <> 0 Then
                        Select Case Val(Vector(Ciclo, 5))
                                Case 1, 4
                                    If Val(WCuenta(Ciclo)) <> 2 Then
                                        Pesos = Pesos + Val(Vector(Ciclo, 9))
                                            Else
                                        Dolares = Dolares + Val(Vector(Ciclo, 9))
                                    End If
                                Case 2
                                    Cheque = Cheque + Val(Vector(Ciclo, 9))
                                Case Else
                                    Documento = Documento + Val(Vector(Ciclo, 9))
                        End Select
                End If

                If Val(Vector(Ciclo, 4)) <> 0 Then
                        Total2 = Total2 + Val(Vector(Ciclo, 4))
                End If

        Next Ciclo

        Total1 = Pesos + Cheque + Documento + Retencion + Dolares

        Rem m = Total1
        Rem GoSub 4630

        Rem For Ciclo = 1 To 2
        Rem         Print #1, Chr$(18)
        Rem Next Ciclo

        Rem  Print #1, Chr$(27) + Chr$(69)
        
        
        If WEmpresa = 1 Then
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
                Else
            Print #1, ""
        End If
        
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Rem Print #1, ""
        Print #1, Tab(56); Alinea("######", Recibo.Text); "/"; Alinea("######", Str$(Wprovisorio))
        Print #1, ""
        Print #1, ""
        Print #1, Tab(55); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        If Tipo3.Value = True Then
            WRazon = Space$(30)
            WDireccion = Space$(30)
            WLocalidad = Space$(30)
            WProvincia = ""
            WPostal = ""
        End If

        Print #1, Tab(8); Left$(WRazon, 30);
        Print #1, Tab(40); "SE DEJA EXPRESA CONSTANCIA QUE LOS "
        Print #1, Tab(8); Left$(WDireccion, 30);
        Print #1, Tab(40); "VALORES QUE SE DETALLAN SOLO SERAN "
        Print #1, Tab(8); Left$(WLocalidad, 30);
        Print #1, Tab(40); "IMPUTADOS EN CONCEPTO DE CANCELACION"
        Print #1, Tab(8); WProvincia; " ("; WPostal; ")";
        Print #1, Tab(40); "DE DEUDA, UNA VEZ QUE SE HALLAN "
        Print #1, Tab(40); "EFECTIVIZADO LA TOTALIDAD DE LOS MISMOS"
        Print #1, ""
        Print #1, ""

        Print #1, Tab(1); m(1)
        Print #1, Tab(1); m(2)

        Print #1, ""
        Print #1, ""

        Print #1, Tab(2); "Efectivo ";
        Print #1, Tab(30); Alinea("###,###.##", Str$(Pesos))
        Print #1, ""
        Print #1, Tab(2); "Cheques ";
        Print #1, Tab(30); Alinea("###,###.##", Str$(Cheque))
        Print #1, ""
        Print #1, Tab(2); "Documentos ";
        Print #1, Tab(30); Alinea("###,###.##", Str$(Documento));

        For Ciclo = 0 To 17

                If Ciclo = 2 Then
                        Print #1, Tab(2); "Retencion Ganancias ";
                        Print #1, Tab(30); Alinea("###,###.##", Retganancias.Text);
                End If

                If Ciclo = 4 Then
                        Print #1, Tab(2); "Retencion Iva ";
                        Print #1, Tab(30); Alinea("###,###.##", RetIva.Text);
                End If

                If Ciclo = 6 Then
                        Print #1, Tab(2); "Retencion I.Brutos ";
                        Print #1, Tab(30); Alinea("###,###.##", RetOtra.Text);
                End If

                If Ciclo = 8 Then
                        Print #1, Tab(2); "Moneda Ext.";
                        Print #1, Tab(30); Alinea("###,###.##", Str$(Dolares));
                End If

                If Val(Vector(Ciclo, 4)) <> 0 Then
                        Print #1, Tab(42); Vector(Ciclo, 10);
                        Print #1, Tab(54); ImpreTipo(Val(Vector(Ciclo, 0)));
                        Print #1, Alinea("######", Vector(Ciclo, 3));
                        Print #1, Tab(65); " $ "; Alinea("###,###.##", Vector(Ciclo, 4))
                                                 Else
                        Print #1, ""
                End If
        Next Ciclo

        Print #1, Tab(30); Alinea("###,###.##", Str$(Total1));
        Print #1, Tab(65); " $ "; Alinea("###,###.##", Str$(Total2))

        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)

        dada = 0
        Erase Impre1
        Erase Impre2

        For Ciclo = 0 To 19

                If Val(Vector(Ciclo, 9)) <> 0 And Val(Vector(Ciclo, 5)) <> 1 Then

                        dada = dada + 1

                        If dada < 11 Then
                                Impre1(dada) = Ciclo + 1
                                        Else
                                Impre2(dada - 10) = Ciclo + 1
                        End If

                End If

        Next Ciclo

        For WCiclo = 1 To 10

                If Impre1(WCiclo) <> 0 Then
                        Ciclo = Impre1(WCiclo) - 1
                        Print #1, Tab(5); Alinea("######", Vector(Ciclo, 6));
                        Print #1, Tab(15); Vector(Ciclo, 7);
                        Print #1, Tab(30); Alinea("###,###.##", Vector(Ciclo, 9));
                        Print #1, Tab(43); Vector(Ciclo, 8);
                End If

                If Impre2(WCiclo) <> 0 Then
                        Ciclo = Impre2(WCiclo) - 1
                        Print #1, Tab(65); Alinea("######", Vector(Ciclo, 6));
                        Print #1, Tab(78); Vector(Ciclo, 7);
                        Print #1, Tab(92); Alinea("###,###.##", Vector(Ciclo, 9));
                        Print #1, Tab(107); Vector(Ciclo, 8);
                End If

                Print #1, ""

        Next WCiclo

        Print #1, Chr$(12)
        
        Next aa
        
        Close #1

End Sub

Private Sub Cuenta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta1.Text <> "" Then
            With rstCuenta
                .Index = "Cuenta"
                Claveven$ = Cuenta1.Text
                .Seek "=", Cuenta1.Text
                If .NoMatch Then
                    Cuenta.SetFocus
                        Else
                    WCuenta(DbGrid1.Row - 1) = Cuenta1.Text
                    Ingrecuenta.Visible = False
                    Rem DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 5
                    KeyCode = 0
                    DbGrid1.SetFocus
                End If
            End With
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

