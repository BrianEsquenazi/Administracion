VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCtaCte 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Clientes"
   ClientHeight    =   6840
   ClientLeft      =   1485
   ClientTop       =   615
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   9015
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   4095
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox TipoCliente 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ComboBox TipoFecha 
         Height          =   315
         Left            =   3720
         TabIndex        =   24
         Top             =   960
         Width           =   2535
      End
      Begin VB.Frame Frame4 
         Caption         =   "Moneda"
         Height          =   855
         Left            =   3480
         TabIndex        =   17
         Top             =   1440
         Width           =   2775
         Begin VB.OptionButton Dolares 
            Caption         =   "Dolares"
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Pesos 
            Caption         =   "Pesos"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Comprobantes"
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   6135
         Begin VB.OptionButton Diferencia 
            Caption         =   "N/D por Dif.Cambio"
            Height          =   495
            Left            =   3000
            TabIndex        =   30
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton Total 
            Caption         =   "Total"
            Height          =   255
            Left            =   5160
            TabIndex        =   22
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Documentos 
            Caption         =   "Documentos"
            Height          =   495
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton CtaCte 
            Caption         =   "Cta. Cte."
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Listado"
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3135
         Begin VB.OptionButton Tipo2 
            Caption         =   "Completo"
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Tipo1 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox Hasta 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   27
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   285
         Left            =   4920
         TabIndex        =   28
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   3720
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   3720
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7440
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ctacte.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Clientes"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2010
      ItemData        =   "ctacte.frx":0000
      Left            =   120
      List            =   "ctacte.frx":0007
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7680
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WNume As String
Private WPasa As String
Private WTitulo As String
Private Importe3 As Double
Private Acumula As Double
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String

Private Sub Acepta_Click()

    On Error GoTo WError
    
    WTitulo = ""
    
    If TipoCliente.ListIndex = 1 Then
        Pesos.Value = False
        Dolares.Value = True
    End If
    
    
    If CtaCte.Value = True Then
        WTitulo = "Cuenta Corriente - "
    End If
    If Documentos.Value = True Then
        WTitulo = "Documentos - "
    End If
    If Diferencia.Value = True Then
        WTitulo = "Diferencia Cambio - "
    End If
    If Total.Value = True Then
        WTitulo = "Total - "
    End If
    
    If TipoCliente.ListIndex = 1 Then
        WTitulo = "Exterior - "
    End If
    
    If Pesos.Value = True Then
        WTitulo = WTitulo + "Pesos"
    End If
    If Dolares.Value = True Then
        WTitulo = WTitulo + "Dolares"
    End If
    

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesdeFecha = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHastaFecha = WAno + WMes + WDia

    Rem With rstCtacte
    Rem         .MoveFirst
    Rem         Do
    Rem             If rstCtacte!Cliente >= Desde.Text And rstCtacte!Cliente <= Hasta.Text Then
    Rem                 WPasa = "N"
    Rem                 If CtaCte.Value = True Then
    Rem                     If rstCtacte!Tipo < 50 Then
    Rem                         WPasa = "S"
    Rem                     End If
    Rem                 End If
    Rem
    Rem                 If Documentos.Value = True Then
    Rem                     If rstCtacte!Tipo >= 50 Then
    Rem                         WPasa = "S"
    Rem                     End If
    Rem                 End If
    Rem
    Rem                 If Total.Value = True Then
    Rem                     WPasa = "S"
    Rem                 End If
    Rem
    Rem                 If WPasa = "S" Then
    Rem                     If Pesos.Value = True Then
    Rem                         If rstCtacte!Total > 0 Then
    Rem                             WImporte1 = rstCtacte!Total
    Rem                             WImporte2 = 0
    Rem                                 Else
    Rem                             WImporte1 = 0
    Rem                             WImporte2 = rstCtacte!Total
    Rem                         End If
    Rem                         WImporte3 = !Saldo
    Rem                             Else
    Rem                         If rstCtacte!TotalUs > 0 Then
    Rem                             WImporte1 = rstCtacte!TotalUs
    Rem                             WImporte2 = 0
    Rem                                 Else
    Rem                             WImporte1 = 0
    Rem                             WImporte2 = rstCtacte!TotalUs
    Rem                         End If
    Rem                         WImporte3 = rstCtacte!SaldoUS
    Rem                     End If
    Rem                 End If
    Rem
    Rem             End If
    Rem
    Rem             WClave = rstCtacte!Clave
    Rem             WImporte1 = Str$(WImporte1)
    Rem             WImporte2 = Str$(WImporte2)
    Rem             WImporte3 = Str$(WImporte3)
    Rem
    Rem             rstCtacte.Close
    Rem
    Rem             XParam = "'" + WClave + "','" _
    Rem                     + WImporte1 + "','" _
    Rem                     + WImporte2 + "','" _
    Rem                     + WImporte3 + "'"
    Rem
    Rem             spCtacte = "ModificaCtacteImporte " + XParam
    Rem             Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem XParam = "'" + Desde.Text + "','" _
    Rem              + Hasta.Text + "'"
    Rem spCtacte = "ListaCtacteDesdeHasta" + XParam
    Rem Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstCtacte.RecordCount > 0 Then
    Rem
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    Rem End If

    spCtacte = "ModificaCtacteTipo1"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    spCtacte = "ModificaCtacteTipo2"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    spCtacte = "ModificaCtacteImporte0"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

    If CtaCte.Value = True Or Diferencia.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte1 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte2 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte3 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte4 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Documentos.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte5 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte6 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte7 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte8 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Total.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte9 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte10 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
                spCtacte = "ModificaCtacte11 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + Desde.Text + "','" _
                        + Hasta.Text + "'"
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
    
    ZZSaldoIni = 0

    ZSql = ""
    ZSql = ZSql + "Select ctacte.clave, ctacte.cliente, ctacte.tipo, ctacte.impre, ctacte.numero, ctacte.renglon, ctacte.Fecha, ctacte.vencimiento, ctacte.vencimiento1, ctacte.total, ctacte.saldo, ctacte.totalus, ctacte.saldous, ctacte.ordfecha, ctacte.ordvencimiento, ctacte.ordvencimiento1, ctacte.Importe1, ctacte.importe2, ctacte.importe3, ctacte.importe4 "
    ZSql = ZSql + " FROM Ctacte"
    ZSql = ZSql + " Where Ctacte.Cliente >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Ctacte.Cliente <= " + "'" + Hasta.Text + "'"
    spCtacte = ZSql
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
                    
                    If Diferencia.Value = True Then
                        If !Tipo = 4 And !Impre = "ND" Then
                            WPasa = "S"
                        End If
                    End If
                    
                    If Total.Value = True Then
                        WPasa = "S"
                    End If
                        
                    If WPasa = "S" Then
                
                        If Tipo2.Value = True Or !Importe3 <> 0 Then
                        
                            If TipoFecha.ListIndex = 0 Or (!OrdFecha >= WDesdeFecha And !OrdFecha <= WHastaFecha) Then
                        
                                WTipo = !Tipo
                                WImpre = !Impre
                                WNumero = !Numero
                                WRenglon = !Renglon
                                WCliente = !Cliente
                                WFecha = !Fecha
                                Rem WEstado = !Estado
                                Wvencimiento = !Vencimiento
                                WVencimiento1 = !Vencimiento1
                                WTotal = !Total
                                WTotalUs = !Totalus
                                WSaldo = !Saldo
                                WSaldoUs = !Saldous
                                Rem  WNeto = !Neto
                                Rem WIva1 = !Iva1
                                Rem WWIva2 = !Iva2
                                WOrdFecha = !OrdFecha
                                WOrdVencimiento = !OrdVencimiento
                                WOrdVencimiento1 = !OrdVencimiento1
                                Rem  WPedido = !Pedido
                                Rem WRemito = !Remito
                                Rem WOrden = !Orden
                                Rem WParidad = !Paridad
                                Rem WProvincia = !Provincia
                                Rem WVendedor = !Vendedor
                                Rem WRubro = !Rubro
                                Rem WCcomprobante = !Comprobante
                                Rem WAceptada = !Aceptada
                                Rem WCosto = !Costo
                                WImporte1 = !Importe1
                                WImporte2 = !Importe2
                                If TipoFecha.ListIndex = 1 Then
                                    WImporte3 = !Total
                                        Else
                                    WImporte3 = !Importe3
                                End If
                                WImporte4 = !Importe4
                                Rem WImporte5 = !Importe5
                                Rem WImporte6 = !Importe6
                                Rem WImporte7 = !Importe7
                                
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
                            
                            If TipoFecha.ListIndex = 1 And !OrdFecha < WDesdeFecha Then
                                ZZSaldoIni = ZZSaldoIni + !Total
                            End If
                        
                        End If
                    
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
        End With
        rstCtacte.Close
    
    End If

    If TipoFecha.ListIndex = 1 Then
    
        
        With rstImpCtaCte

            .Index = "Clave"
                                    
            .AddNew
            
            !Tipo = "01"
            !Impre = "SI"
            !Numero = "00000000"
            !Renglon = "01"
            !Cliente = WCliente
            !Fecha = "00/00/0000"
            !Estado = ""
            !Vencimiento = ""
            !Vencimiento1 = ""
            !Total = ZZSaldoIni
            !Totalus = ZZSaldoIni
            !Saldo = ZZSaldoIni
            !Saldous = ZZSaldoIni
            !Neto = 0
            !Iva1 = 0
            !Iva2 = 0
            !OrdFecha = "00000000"
            !OrdVencimiento = ""
            !OrdVencimiento1 = ""
            !Pedido = ""
            !Remito = ""
            !Orden = ""
            !Paridad = 0
            !Provincia = 0
            !vendedor = 0
            !Rubro = 0
            !Comprobante = 0
            !Aceptada = ""
            !Costo = 0
            !Importe1 = 0
            !Importe2 = 0
            !Importe3 = ZZSaldoIni
            !Importe4 = 0
            !Importe5 = 0
            !Importe6 = 0
            !Importe7 = 0
            !Clave = !Tipo + !Numero + !Renglon
            WNume = Str$(!Numero)
            Call Ceros(WNume, 8)
            !ClaveImpre = !Cliente + !OrdFecha + !Tipo + WNume
            !Empresa = Val(WEmpresa)

            .Update
            
        End With
    
    End If
    
    
    With rstImpCtaCte
            .Index = "ClaveImpre"
            .MoveFirst
            Do
            
                WRazon = ""
                spCliente = "ConsultaCliente " + !Cliente
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WRazon = rstCliente!Razon
                    WProvincia = rstCliente!Provincia
                    rstCliente.Close
                End If
            
                If TipoCliente.ListIndex = 1 And WProvincia <> 24 Then
                    .Delete
                        Else
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
                    !Titulo = WTitulo
                    .Update
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
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
    If Diferencia.Value = True Then
        WTitulo = "Diferencia Cambio - "
    End If
    
    If TipoCliente.ListIndex = 1 Then
        WTitulo = "Exterior - "
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
    
    Listado.GroupSelectionFormula = "{impCtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.ReportFileName = "wimpctacte.rpt"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Rem Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Rem Listado.Connect = Connect()
    
    Listado.Action = 1
    
    Exit Sub

WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgCtaCte.Hide
    Unload Me
    Menu.Show
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
    End If
            
    Pantalla.Visible = True

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

Private Sub pantalla_Click()
    Pantalla.Visible = False
       
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Desde.Text = rstCliente!Cliente
        Hasta.Text = rstCliente!Cliente
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
                Else
            DesdeFecha.SetFocus
        End If
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            DesdeFecha.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
End Sub

Sub Form_Load()

    TipoFecha.Clear
    
    TipoFecha.AddItem "Toda la informacion"
    TipoFecha.AddItem "Entre fechas"
    
    TipoFecha.ListIndex = 0
    
    TipoCliente.Clear
    
    TipoCliente.AddItem "Todos los Clientes"
    TipoCliente.AddItem "Exterior"
    
    TipoCliente.ListIndex = 0

    Desde.Text = ""
    Hasta.Text = ""
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Tipo1.Value = True
    Tipo2.Value = False
    Pesos.Value = True
    Dolares.Value = False
    CtaCte.Value = True
    Diferencia.Value = False
    Documentos.Value = False
    Total.Value = False
    Frame2.Visible = True
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    With rstClientes
        .Index = "Razon"
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
    
    End If

End Sub



