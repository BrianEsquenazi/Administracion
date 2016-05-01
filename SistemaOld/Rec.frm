VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRecibos 
   Caption         =   "Ingreso de Recibos"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6735
   ScaleWidth      =   9480
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion"
      Height          =   300
      Left            =   6240
      TabIndex        =   30
      Top             =   2400
      Width           =   975
   End
   Begin Crystal.CrystalReport listado 
      Left            =   8520
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "recibo.rpt"
      CopiesToPrinter =   2
   End
   Begin VB.TextBox Observaciones 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox RetOtra 
      Height          =   285
      Left            =   4800
      TabIndex        =   24
      Text            =   " "
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox RetIva 
      Height          =   285
      Left            =   1680
      TabIndex        =   22
      Text            =   " "
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Retganancias 
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Text            =   " "
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Recibos"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   5295
      Begin VB.OptionButton Tipo1 
         Caption         =   "Cobro de Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Por Cta y Orden de Terceros"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.TextBox Consecionaria 
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   15
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox Opcion 
      Height          =   1425
      Left            =   6840
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   12
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
      Height          =   3015
      Left            =   0
      OleObjectBlob   =   "Recibos.frx":0000
      TabIndex        =   10
      Top             =   2760
      Width           =   9255
   End
   Begin VB.TextBox Recibo 
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
      Left            =   8400
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "Recibos.frx":09C2
      Left            =   5520
      List            =   "Recibos.frx":09C9
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7320
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   7320
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6240
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   300
      Left            =   7320
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   6240
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Creditos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   7800
      TabIndex        =   27
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Debitos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.   2) Ch.   3) Doc."
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Otra Retencion"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Ret.Iva"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Rte.Ganancias"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label DesConsecionaria 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod. Concesionaria"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero de Recibo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "PrgRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 10 ' Número máximo de campos del conjunto de registros.

Private Sub Suma_Datos()
    Debitos.Caption = ""
    Creditos.Caption = ""
    
    Creditos.Caption = Str$(Val(Retganancias.text) + Val(RetIva.text) + Val(RetOtra.text))
    For iRow = 0 To 9
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        If Val(DbGrid1.text) <> 0 Then
            Debitos.Caption = Str$(Val(Debitos.Caption) + Val(DbGrid1.text))
        End If
        DbGrid1.Col = 9
        DbGrid1.Row = iRow
        If Val(DbGrid1.text) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(DbGrid1.text))
        End If
    Next iRow
    Debitos.Caption = PUsing("###,###.##", Debitos.Caption)
    Creditos.Caption = PUsing("###,###.##", Creditos.Caption)
    DbGrid1.Col = 0
    DbGrid1.Row = 0
End Sub

Private Sub Lee_Datos()
    Renglon = 0
    DEbito = 0
    Credito = 0
    Do
        With rstRecibos
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", Recibo.text + Auxi1
            If .NoMatch = False Then
                Select Case Val(!Tiporeg)
                    Case 1
                        DEbito = DEbito + 1
                        DbGrid1.Row = DEbito - 1
                        DbGrid1.Col = 0
                        DbGrid1.text = !Tipo1
                        DbGrid1.Col = 1
                        DbGrid1.text = !Letra1
                        DbGrid1.Col = 2
                        DbGrid1.text = !Punto1
                        DbGrid1.Col = 3
                        DbGrid1.text = !Numero1
                        DbGrid1.Col = 4
                        DbGrid1.text = !Importe1
                        DbGrid1.text = PUsing("###,###.##", DbGrid1.text)
                    Case 2
                        Credito = Credito + 1
                        DbGrid1.Row = Credito - 1
                        DbGrid1.Col = 5
                        DbGrid1.text = !Tipo2
                        DbGrid1.Col = 6
                        DbGrid1.text = !Numero2
                        DbGrid1.Col = 7
                        DbGrid1.text = !Fecha2
                        DbGrid1.Col = 8
                        DbGrid1.text = !banco2
                        DbGrid1.Col = 9
                        DbGrid1.text = !Importe2
                        DbGrid1.text = PUsing("###,###.##", DbGrid1.text)
                    Case Else
                End Select
                    Else
                Exit Do
            End If
        End With
    Loop
End Sub
Sub Verifica_datos()
    If Val(Retganancias.text) = 0 Then
        Retganancias.text = "0"
    End If
    If Val(RetIva.text) = 0 Then
        RetIva.text = "0"
    End If
    If Val(RetOtra.text) = 0 Then
        RetOtra.text = "0"
    End If
End Sub
Sub Format_datos()
    Retganancias.text = PUsing("###,###.##", Retganancias.text)
    RetIva.text = PUsing("###,###.##", RetIva.text)
    RetOtra.text = PUsing("###,###.##", RetOtra.text)
End Sub

Sub Imprime_Datos()
    With rstConsecionaria
        .Index = "Consecionaria"
        .Seek "=", Consecionaria.text
        If .NoMatch = False Then
            Consecionaria.text = !Consecionaria
            DesConsecionaria.Caption = !Nombre
            Call Format_datos
        End If
    End With
End Sub

Private Sub cmdAdd_Click()

    If Recibo.text <> "" And Fecha.text <> "" Then
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        DEbito = 0
        Credito = 0
        If Val(Debitos.Caption) <> 0 Then
            DEbito = Val(Debitos.Caption)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = Val(Creditos.Caption)
        End If
        
        If DEbito = Credito Or Tipo2.Value = True Then
    
        With rstRecibos
            Renglon = 0
            .Index = "Clave"
            For iRow = 0 To 9
        
                If Tipo1.Value = True Then
                    WRow = iRow
                    DbGrid1.Col = 4
                    DbGrid1.Row = iRow
                    If Val(DbGrid1.text) <> 0 Then
                        .AddNew
                        Renglon = Renglon + 1
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        !Recibo = Recibo.text
                        !Renglon = Auxi1
                        !Consecionaria = Consecionaria.text
                        !Fecha = Fecha.text
                        !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                        If Tipo1.Value = True Then
                            !Tiporec = "1"
                        End If
                        If Tipo2.Value = True Then
                            !Tiporec = "2"
                        End If
                        !Retganancias = Val(Retganancias.text)
                        !RetIva = Val(RetIva.text)
                        !RetOtra = Val(RetOtra.text)
                        !Retencion = 0
                        !Tiporeg = "1"
                        DbGrid1.Col = 0
                        !Tipo1 = DbGrid1.text
                        DbGrid1.Col = 1
                        !Letra1 = DbGrid1.text
                        DbGrid1.Col = 2
                        !Punto1 = DbGrid1.text
                        DbGrid1.Col = 3
                        !Numero1 = DbGrid1.text
                        DbGrid1.Col = 4
                        !Importe1 = DbGrid1.text
                        !Tipo2 = ""
                        !Numero2 = ""
                        !Fecha2 = ""
                        !FechaOrd2 = ""
                        !banco2 = ""
                        !Importe2 = 0
                        !Estado2 = ""
                        !Observaciones = Observaciones.text
                        !Empresa = 1
                        !Clave = !Recibo + !Renglon
                        !Importe = Credito
                        .Update
                        .Bookmark = .LastModified
                    
                        WLetra = !Letra1
                        WTipo = !Tipo1
                        WPunto = !Punto1
                        WNumero = !Numero1
                        WImporte = !Importe1
                    
                        With rstCtaCte
                            .Index = "CtaCte"
                            Auxi$ = Consecionaria.text
                            Call Ceros(Auxi$, 6)
                            Claveven$ = Auxi$
                            Claveven$ = Claveven$ + WLetra + WTipo + WPunto + WNumero + "01"
                            .Seek "=", Claveven$
                            If .NoMatch = False Then
                                .Edit
                                !Saldo = !Saldo - WImporte
                                .Update
                                .Bookmark = .LastModified
                            End If
                        End With
                        
                    End If
                End If
                
                DbGrid1.Col = 9
                DbGrid1.Row = iRow
                If Val(DbGrid1.text) <> 0 Then
                    .AddNew
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Recibo = Recibo.text
                    !Renglon = Auxi1
                    !Consecionaria = Consecionaria.text
                    !Fecha = Fecha.text
                    !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    If Tipo1.Value = True Then
                        !Tiporec = "1"
                    End If
                    If Tipo2.Value = True Then
                        !Tiporec = "2"
                    End If
                    !Retganancias = Val(Retganancias.text)
                    !RetIva = Val(RetIva.text)
                    !RetOtra = Val(RetOtra.text)
                    !Retencion = 0
                    !Tiporeg = "2"
                    !Tipo1 = ""
                    !Letra1 = ""
                    !Punto1 = ""
                    !Numero1 = ""
                    !Importe1 = 0
                    DbGrid1.Col = 5
                    !Tipo2 = DbGrid1.text
                    DbGrid1.Col = 6
                    !Numero2 = DbGrid1.text
                    DbGrid1.Col = 7
                    !Fecha2 = DbGrid1.text
                    !FechaOrd2 = Right$(!Fecha2, 4) + Mid$(!Fecha2, 4, 2) + Left$(!Fecha2, 2)
                    DbGrid1.Col = 8
                    !banco2 = DbGrid1.text
                    DbGrid1.Col = 9
                    !Importe2 = DbGrid1.text
                    !Estado2 = "P"
                    !Observaciones = Observaciones.text
                    !Empresa = 1
                    !Clave = !Recibo + !Renglon
                    !Importe = Credito
                    .Update
                    .Bookmark = .LastModified
                End If
                
            Next iRow
        End With
        
        If Tipo1.Value = True Then
            With rstCtaCte
                Auxi = Consecionaria.text
                Call Ceros(Auxi, 6)
                WConsecionaria = Auxi
                WLetra = "A"
                WTipo = "04"
                WPunto = "0000"
                WNumero = "00" + Recibo.text
                .Index = "CtaCte"
                .Seek "=", WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                If .NoMatch Then
                    .AddNew
                    !Consecionaria = Consecionaria.text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Renglon = "01"
                    !Fecha = Fecha.text
                    !Estado = "1"
                    !Vencimiento = Fecha.text
                    !Total = Credito * -1
                    !Saldo = 0
                    !ClaveCtacte = WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                    !OrdFecha = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !OrdVencimiento = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !Impre = "RC"
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Consecionaria = Consecionaria.text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Renglon = "01"
                    !Fecha = Fecha.text
                    !Estado = "1"
                    !Vencimiento = Fecha.text
                    !Total = Credito * -1
                    !Saldo = 0
                    !ClaveCtacte = WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                    !OrdFecha = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !OrdVencimiento = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !Impre = "RC"
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
        End If
        
        If Tipo2.Value = True Then
            With rstCtaCte
                Auxi = Consecionaria.text
                Call Ceros(Auxi, 6)
                WConsecionaria = Auxi
                WLetra = "A"
                WTipo = "04"
                WPunto = "0000"
                WNumero = "00" + Recibo.text
                .Index = "CtaCte"
                .Seek "=", WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                If .NoMatch Then
                    .AddNew
                    !Consecionaria = Consecionaria.text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Renglon = "01"
                    !Fecha = Fecha.text
                    !Estado = "1"
                    !Vencimiento = Fecha.text
                    !Total = Credito * -1
                    !Saldo = Credito * -1
                    !ClaveCtacte = WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                    !OrdFecha = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !OrdVencimiento = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !Impre = "RC"
                    Rem !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Consecionaria = Consecionaria.text
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Renglon = "01"
                    !Fecha = Fecha.text
                    !Estado = "1"
                    !Vencimiento = Fecha.text
                    !Total = Credito * -1
                    !Saldo = Credito * -1
                    !ClaveCtacte = WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                    !OrdFecha = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !OrdVencimiento = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !Impre = "RC"
                    !Empresa = 1
                    .Update
                    .Bookmark = .LastModified
                End If
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
        
        With rstImputac
            Renglon = 0
            .Index = "Clave"
            
            If Val(Retganancias.text) <> 0 Then
                .AddNew
                !Tipomovi = "1"
                !Proveedor = "000000"
                !TipoComp = "01"
                !LetraComp = "A"
                !PuntoComp = "0000"
                !NroComp = Recibo.text
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                !Renglon = Auxi1$
                !Fecha = Fecha.text
                !Observaciones = DesConsecionaria.Caption
                !Cuenta = WCtaRetGan
                !DEbito = Val(Retganancias.text)
                !Credito = 0
                !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                !Titulo = "Cobranzas"
                !Empresa = 1
                !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                .Update
            End If
            
            If Val(RetIva.text) <> 0 Then
                .AddNew
                !Tipomovi = "1"
                !Proveedor = "000000"
                !TipoComp = "01"
                !LetraComp = "A"
                !PuntoComp = "0000"
                !NroComp = Recibo.text
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                !Renglon = Auxi1$
                !Fecha = Fecha.text
                !Observaciones = DesConsecionaria.Caption
                !Cuenta = WctaRetIva
                !DEbito = Val(RetIva.text)
                !Credito = 0
                !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                !Titulo = "Cobranzas"
                !Empresa = 1
                !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                .Update
            End If
            
            If Val(RetOtra.text) <> 0 Then
                .AddNew
                !Tipomovi = "1"
                !Proveedor = "000000"
                !TipoComp = "01"
                !LetraComp = "A"
                !PuntoComp = "0000"
                !NroComp = Recibo.text
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                !Renglon = Auxi1$
                !Fecha = Fecha.text
                !Observaciones = DesConsecionaria.Caption
                !Cuenta = WCtaretOtra
                !DEbito = Val(RetOtra.text)
                !Credito = 0
                !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                !Titulo = "Cobranzas"
                !Empresa = 1
                !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                .Update
            End If
            
            If Tipo2.Value = True Then
                .AddNew
                !Tipomovi = "1"
                !Proveedor = "000000"
                !TipoComp = "01"
                !LetraComp = "A"
                !PuntoComp = "0000"
                !NroComp = Recibo.text
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                !Renglon = Auxi1$
                !Fecha = Fecha.text
                !Observaciones = DesConsecionaria.Caption
                !Cuenta = WctaTerceros
                !DEbito = 0
                !Credito = Credito
                !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                !Titulo = "Cobranzas"
                !Empresa = 1
                !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                .Update
            End If

            For iRow = 0 To 9
                WRow = iRow
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                If Val(DbGrid1.text) <> 0 Then
                    .AddNew
                    !Tipomovi = "1"
                    !Proveedor = "000000"
                    !TipoComp = "01"
                    !LetraComp = "A"
                    !PuntoComp = "0000"
                    !NroComp = Recibo.text
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Renglon = Auxi1$
                    !Fecha = Fecha.text
                    !Observaciones = DesConsecionaria.Caption
                    !Cuenta = WCtaDeudores
                    !DEbito = 0
                    !Credito = Val(DbGrid1.text)
                    !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !Titulo = "Cobranzas"
                    !Empresa = 1
                    !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                    .Update
                End If
                
                DbGrid1.Col = 9
                DbGrid1.Row = iRow
                If Val(DbGrid1.text) <> 0 Then
                    .AddNew
                    !Tipomovi = "1"
                    !Proveedor = "000000"
                    !TipoComp = "01"
                    !LetraComp = "A"
                    !PuntoComp = "0000"
                    !NroComp = Recibo.text
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Renglon = Auxi1$
                    !Fecha = Fecha.text
                    !Observaciones = DesConsecionaria.Caption
                    !DEbito = Val(DbGrid1.text)
                    !Credito = 0
                    !FechaOrd = Right$(Fecha.text, 4) + Mid$(Fecha.text, 4, 2) + Left$(Fecha.text, 2)
                    !Titulo = "Cobranzas"
                    !Empresa = 1
                    DbGrid1.Col = 5
                    Select Case Val(DbGrid1.text)
                        Case 2
                            !Cuenta = WCtaCheques
                        Case 3
                            !Cuenta = WCtaDocumentos
                        Case Else
                            !Cuenta = WCtaEfectivo
                    End Select
                    !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                    .Update
                End If
                
            Next iRow
        End With
        
        Listado.GroupSelectionFormula = "{Recibos.recibo} in " + Chr$(34) + Recibo.text + Chr$(34) + " to " + Chr$(34) + Recibo.text + Chr$(34)
        Listado.Destination = 1
        Listado.Action = 1

        Call CmdLimpiar_Click
        Recibo.SetFocus
        
        End If
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Existe = "S" Then
        T$ = "Borrar Recibo"
        M$ = "Desea Borrar el Recibo "
        Respuesta% = MsgBox(M$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            With rstRecibos
                Renglon = 0
                .Index = "Clave"
                For iRow = 0 To 50
                    Renglon = iRow + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    Claveven$ = Recibo.text + Auxi1$
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                    
                        .Edit
                        
                        If !Tiporeg = "1" Then
                        
                            WLetra = !Letra1
                            WTipo = !Tipo1
                            WPunto = !Punto1
                            WNumero = !Numero1
                            WImporte = !Importe1
                    
                            With rstCtaCte
                                .Index = "CtaCte"
                                Auxi$ = Consecionaria.text
                                Call Ceros(Auxi$, 6)
                                Claveven$ = Claveven$ + WLetra + WTipo + WPunto + WNumero + "01"
                                .Seek "=", Claveven$
                                If .NoMatch = False Then
                                    .Edit
                                    !Saldo = !Saldo + WImporte
                                    .Update
                                    .Bookmark = .LastModified
                                End If
                            End With
                        End If
                        
                        .Delete
                        
                                Else
                                
                        Exit For
                        
                    End If
                
                Next iRow
            End With
        
            If Tipo1.Value = True Then
                With rstCtaCte
                    Auxi = Consecionaria.text
                    Call Ceros(Auxi, 6)
                    WConsecionaria = Auxi
                    WLetra = "A"
                    WTipo = "04"
                    WPunto = "0000"
                    WNumero = "00" + Recibo.text
                    .Index = "CtaCte"
                    .Seek "=", WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                    If .NoMatch = False Then
                        .Edit
                        .Delete
                        .Bookmark = .LastModified
                    End If
                End With
            End If
        
            If Tipo2.Value = True Then
                With rstCtaCte
                    Auxi = Consecionaria.text
                    Call Ceros(Auxi, 6)
                    WConsecionaria = Auxi
                    WLetra = "A"
                    WTipo = "04"
                    WPunto = "0000"
                    WNumero = "00" + Recibo.text
                    .Index = "CtaCte"
                    .Seek "=", WConsecionaria + WLetra + WTipo + WPunto + WNumero + "01"
                    If .NoMatch = False Then
                        .Edit
                        .Delete
                        .Bookmark = .LastModified
                    End If
                End With
            End If
        
            With rstImputac

                For iRow = 1 To 50
                
                    Renglon = iRow
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    Wrenglon = Auxi1
                    WTipomovi = "1"
                    WTipoComp = "01"
                    WLetraComp = "A"
                    WPuntoComp = "0000"
                    WNroComp = Recibo.text
                    Claveven = WTipomovi + WTipoComp + WLetraComp + WPuntoComp + WNroComp + Wrenglon
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        .Edit
                        .Delete
                    End If
                Next iRow
            End With

            Call CmdLimpiar_Click
        End If
    
    End If
    Recibo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    For iCol = 0 To 9
        For iRow = 0 To 9
        
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.text = ""
        Next iRow
    Next iCol
    Recibo.text = ""
    Consecionaria.text = ""
    DesConsecionaria.Caption = ""
    Observaciones.text = ""
    Fecha.text = "  /  /    "
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.text = "0"
    RetIva.text = "0"
    RetOtra.text = "0"
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    cmdDelete.Enabled = False
    
    With rstRecibos
        .Index = "Clave"
        .Seek "<", "99999999"
        If .NoMatch = False Then
            Recibo.text = !Recibo + 1
                Else
            Recibo.text = ""
        End If
    End With
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstImputac
        .Close
    End With
    With rstConsecionaria
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
    PrgRecibos.Hide
    Menu.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Impresion_Click()
        Listado.GroupSelectionFormula = "{Recibos.recibo} in " + Chr$(34) + Recibo.text + Chr$(34) + " to " + Chr$(34) + Recibo.text + Chr$(34)
        Listado.Destination = 1
        Listado.Action = 1
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Recibo.text
        Call Ceros(Auxi1, 6)
        Recibo.text = Auxi1
        
        With rstRecibos
            cmdDelete.Enabled = False
            Existe = "N"
            .Index = "Clave"
            Claveven$ = Recibo.text + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                cmdDelete.Enabled = True
                Existe = "S"
                Consecionaria.text = !Consecionaria
                Observaciones.text = !Observaciones
                Fecha.text = !Fecha
                Retganancias.text = !Retganancias
                RetIva.text = !RetIva
                RetOtra.text = !RetOtra
                Tipo1.Value = True
                Tipo2.Value = False
                Select Case Val(!Tiporec)
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
        Call Valida_fecha1(Fecha.text, Auxi)
        If Auxi = "S" Then
            Consecionaria.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Consecionaria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Consecionaria.text) <> 0 Then
            With rstConsecionaria
                .Index = "Consecionaria"
                Claveven$ = Consecionaria.text
                .Seek "=", Consecionaria.text
                If .NoMatch Then
                    Consecionaria.text = Claveven$
                        Else
                    Consecionaria.text = !Consecionaria
                    DesConsecionaria.Caption = !Nombre
                    Rem Call Imprime_Datos
                End If
            End With
        End If
        Observaciones.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.SetFocus
    End If
End Sub

Private Sub Retganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.text = PUsing("###,###.##", Retganancias.text)
        Call Suma_Datos
        RetIva.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetIva.text = PUsing("###,###.##", RetIva.text)
        Call Suma_Datos
        RetOtra.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtra.text = PUsing("###,###.##", RetOtra.text)
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

     Opcion.AddItem "Consecionarias"
     Opcion.AddItem "Cuenta Corriebtes"

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
            With rstConsecionaria
                .Index = "Consecionaria"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Consecionaria + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Consecionaria
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case 1
            With rstCtaCte
                .Index = "ClaveImpre"
                .Seek ">", Consecionaria.text + Space$(100)
                If .NoMatch = False Then
                Do
                    If .EOF = False Then
                        If Val(Consecionaria.text) = Val(!Consecionaria) Then
                            If !Saldo <> 0 Then
                                Auxi$ = Str$(!Saldo)
                                Auxi$ = PUsing("###,###.##", Auxi$)
                                IngresaItem = !Impre + " " + !Letra + " " + !Punto + " " + !Numero + " " + !Fecha + " " + Auxi$
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !ClaveCtacte
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
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            With rstConsecionaria
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Consecionaria.text = Claveven$
                .Index = "Consecionaria"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DesConsecionaria.Caption = !Nombre
                            Else
                    Consecionaria.text = ""
                End If
            End With
                
            Consecionaria.SetFocus
            
        Case 1
            With rstCtaCte

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "CtaCte"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 0
                    DbGrid1.text = !Tipo
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 1
                    DbGrid1.text = !Letra
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 2
                    DbGrid1.text = !Punto
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 3
                    DbGrid1.text = !Numero
                
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    DbGrid1.text = !Saldo
                    DbGrid1.text = PUsing("###,###.##", DbGrid1.text)
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 4
                    
                End If
            End With
                
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
                    If Val(DbGrid1.text) = 1 Or Val(DbGrid1.text) = 2 Or Val(DbGrid1.text) = 3 Then
                        Auxi$ = Str$(Val(DbGrid1.text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.text = Auxi$
                        DbGrid1.Col = 1
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If
                
            Case 1
                If KeyCode = 13 Then
                    DbGrid1.text = Left$(DbGrid1.text, 1)
                    If DbGrid1.text = "A" Or DbGrid1.text = "C" Then
                        DbGrid1.Col = 2
                        KeyCode = 0
                        Rem no hago anda
                            Else
                        DbGrid1.Col = 1
                        KeyCode = 0
                    End If
                End If
                
            Case 2
                If KeyCode = 13 Then
                    Auxi$ = Str$(Val(DbGrid1.text))
                    Call Ceros(Auxi$, 4)
                    DbGrid1.text = Auxi$
                    DbGrid1.Col = 3
                    KeyCode = 0
                End If
                
            Case 3
                If KeyCode = 13 Then
                
                    Auxi$ = Str$(Val(DbGrid1.text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.text = Auxi$
                
                    With rstCtaCte
                        .Index = "CtaCte"
                        Auxi$ = Consecionaria.text
                        Call Ceros(Auxi$, 6)
                        Claveven$ = Auxi$
                        DbGrid1.Col = 1
                        Claveven$ = Claveven$ + DbGrid1.text
                        DbGrid1.Col = 0
                        Claveven$ = Claveven$ + DbGrid1.text
                        DbGrid1.Col = 2
                        Claveven$ = Claveven$ + DbGrid1.text
                        DbGrid1.Col = 3
                        Claveven$ = Claveven$ + DbGrid1.text + "01"
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            DbGrid1.Col = 4
                            XRow = DbGrid1.Row
                            If Val(DbGrid1.text) = 0 Then
                                DbGrid1.text = !Saldo
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
                        .Index = "CtaCte"
                        Auxi$ = Consecionaria.text
                        Call Ceros(Auxi$, 6)
                        Claveven$ = Auxi$
                        DbGrid1.Col = 1
                        Claveven$ = Claveven$ + DbGrid1.text
                        DbGrid1.Col = 0
                        Claveven$ = Claveven$ + DbGrid1.text
                        DbGrid1.Col = 2
                        Claveven$ = Claveven$ + DbGrid1.text
                        DbGrid1.Col = 3
                        Claveven$ = Claveven$ + DbGrid1.text + "01"
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            Saldo = !Saldo
                                Else
                            Saldo = 0
                        End If
                    End With
                
                    DbGrid1.Col = 4
                    If Val(DbGrid1.text) > Saldo Then
                        DbGrid1.text = ""
                        DbGrid1.Col = 4
                        KeyCode = 0
                            Else
                        DbGrid1.text = PUsing("###,###.##", DbGrid1.text)
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
                    If Val(DbGrid1.text) = 1 Or Val(DbGrid1.text) = 2 Or Val(DbGrid1.text) = 3 Then
                        Auxi$ = Str$(Val(DbGrid1.text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.text = Auxi$
                        If Val(DbGrid1.text) = 1 Then
                            DbGrid1.Col = 6
                            DbGrid1.text = ""
                            DbGrid1.Col = 7
                            DbGrid1.text = ""
                            DbGrid1.Col = 8
                            DbGrid1.text = ""
                            DbGrid1.Col = 9
                            KeyCode = 0
                                Else
                            DbGrid1.Col = 6
                            KeyCode = 0
                        End If
                            Else
                        DbGrid1.Col = 5
                        KeyCode = 0
                    End If
                End If
                
            Case 6
                If KeyCode = 13 Then
                
                    Auxi$ = Str$(Val(DbGrid1.text))
                    Call Ceros(Auxi$, 8)
                    DbGrid1.text = Auxi$
                    DbGrid1.Col = 7
                    KeyCode = 0
                
                End If
                
            Case 7
                If KeyCode = 13 Then
                    DbGrid1.Col = 7
                    
                    Call Valida_fecha1(DbGrid1.text, Auxi)
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
                    DbGrid1.Col = 9
                    DbGrid1.text = PUsing("###,###.##", DbGrid1.text)
                    Call Suma_Datos
                    DbGrid1.Row = iRow
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

ReDim UserData(0 To 9, 0 To 9)

mTotalRows& = 10

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
             DbGrid1.Columns(newcnt).Locked = False
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Letra"
             DbGrid1.Columns(newcnt).Width = 450
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Punto"
             DbGrid1.Columns(newcnt).Width = 600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 5
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 6
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 7
             DbGrid1.Columns(newcnt).Caption = "Fecha"
             DbGrid1.Columns(newcnt).Width = 1150
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
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
     
    Tipo1.Value = True
    Tipo2.Value = False
    
    Retganancias.text = "0"
    RetIva.text = "0"
    RetOtra.text = "0"

    Recibo.text = ""
    Consecionaria.text = ""
    DesConsecionaria.Caption = ""
    Fecha.text = "  /  /    "
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.text = "0"
    RetIva.text = "0"
    RetOtra.text = "0"
    Recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Observaciones.text = ""
    cmdDelete.Enabled = False
    
    With rstRecibos
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Recibo.text = !Recibo + 1
                Else
            Recibo.text = ""
        End If
    End With

End Sub
