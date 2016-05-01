VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCtaCteFec 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Clientes"
   ClientHeight    =   7425
   ClientLeft      =   1605
   ClientTop       =   585
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   9015
   Begin VB.TextBox Ayuda 
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
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
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
         Left            =   5760
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         TabIndex        =   17
         Top             =   2280
         Width           =   2775
         Begin VB.OptionButton Dolares 
            Caption         =   "Dolares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Pesos 
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Comprobantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   720
         TabIndex        =   16
         Top             =   3360
         Width           =   4815
         Begin VB.OptionButton Total 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3480
            TabIndex        =   22
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Documentos 
            Caption         =   "Documentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   495
            Left            =   1800
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton CtaCte 
            Caption         =   "Cta. Cte."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   3135
         Begin VB.OptionButton Tipo2 
            Caption         =   "Completo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Tipo1 
            Caption         =   "Pendiente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox Hasta 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Desde 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Left            =   4440
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         Left            =   4440
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7920
      Top             =   1200
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
      Left            =   7680
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "ctactefec.frx":0000
      Left            =   120
      List            =   "ctactefec.frx":0007
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7680
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCtaCteFec"
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
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim XParam As String
Dim WRecibo(15000, 6) As String

Private Sub Acepta_Click()

    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFecha = WAno + WMes + WDia

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    spCtacte = "ModificaCtacteTipo1"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    spCtacte = "ModificaCtacteTipo2"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
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
    
    
    
    XParam = "'" + Desde.Text + "','" _
            + Hasta.Text + "'"
    spCtacte = "ListaCtacteDesdeHasta " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then

    With rstCtacte
            .MoveFirst
            Do
            
                If !OrdFecha <= WFecha Then
                
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
                
                    If Total.Value = True Then
                        WPasa = "S"
                    End If
                    
                    If WPasa = "S" Then
            
                        WTipo = !Tipo
                        WImpre = !Impre
                        WNumero = !Numero
                        WRenglon = !Renglon
                        WCliente = !Cliente
                        XFecha = !Fecha
                        WEstado = !Estado
                        Wvencimiento = !Vencimiento
                        WVencimiento1 = !Vencimiento1
                        WTotal = !Total
                        WTotalUs = !Totalus
                        WSaldo = !Saldo
                        WSaldoUs = !Saldous
                        WNeto = !Neto
                        WIva1 = !Iva1
                        WWIva2 = !Iva2
                        WOrdFecha = !OrdFecha
                        WOrdVencimiento = !OrdVencimiento
                        WOrdVencimiento1 = !OrdVencimiento1
                        WPedido = !Pedido
                        WRemito = !Remito
                        WOrden = !Orden
                        WParidad = !Paridad
                        WProvincia = !Provincia
                        WVendedor = !vendedor
                        WRubro = !Rubro
                        WCcomprobante = !Comprobante
                        WAceptada = !Aceptada
                        WCosto = !Costo
                        WImporte1 = !Importe1
                        WImporte2 = !Importe2
                        WImporte3 = !Importe3
                        WImporte4 = !Importe4
                        WImporte5 = !Importe5
                        WImporte6 = !Importe6
                        WImporte7 = !Importe7
                        WClave = !Clave
                
                        With rstImpCtaCte
        
                            .Index = "Clave"
                                            
                            .AddNew
                    
                            !Tipo = WTipo
                            !Impre = WImpre
                            !Numero = WNumero
                            !Renglon = WRenglon
                            !Cliente = WCliente
                            !Fecha = XFecha
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
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            !Importe7 = 0
                            !Clave = WClave
                            WNume = Str$(!Numero)
                            Call Ceros(WNume, 8)
                            !ClaveImpre = !Cliente + !OrdFecha + !Tipo + WNume

                            If !Total > 0 Then
                                !Importe1 = !Total
                                !Importe2 = 0
                                    Else
                                !Importe1 = 0
                                !Importe2 = !Total
                            End If
                            !Importe3 = !Saldo

                            .Update
                    
                        End With
                
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
    
    
    Erase WRecibo
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
            + Hasta.Text + "'"
    spRecibos = "ListaRecibosCliente " + XParam
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
            
        With rstRecibos
            .MoveFirst
            Do
                        
                If WFecha < !FechaOrd Then
                
                    If !Importe1 <> 0 Then
                    
                        Renglon = Renglon + 1
                
                        WRecibo(Renglon, 1) = !Tipo1
                        WRecibo(Renglon, 2) = !Numero1
                        WRecibo(Renglon, 3) = Str$(!Importe1)
                        WRecibo(Renglon, 4) = !Clave
                        WRecibo(Renglon, 5) = !Recibo
                        WRecibo(Renglon, 6) = !Cliente
                            
                    End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        
        rstRecibos.Close
        
    End If
    
    
    For Ciclo = 1 To Renglon
    
        WTipo = WRecibo(Ciclo, 1)
        WNumero = WRecibo(Ciclo, 2)
        WImporte = Val(WRecibo(Ciclo, 3))
        XClave = WRecibo(Ciclo, 4)
        XRecibo = WRecibo(Ciclo, 5)
        WCliente = WRecibo(Ciclo, 6)
                    
        Call Ceros(WTipo, 2)
        Call Ceros(WNumero, 8)
                    
        WClave = WTipo + WNumero + "01"
        
        WProv = 0
        WParidad = 0
        
        spClientes = "ConsultaClientes " + "'" + WCliente + "'"
        Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
        If rstClientes.RecordCount > 0 Then
            WProv = rstClientes!Provincia
            rstClientes.Close
        End If
        
        If WProv = 24 Then
            Auxi1 = XRecibo
            Call Ceros(Auxi1, 8)
            ClaveCtacte = "06" + Auxi1 + "01"
            spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
            Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtacte.RecordCount > 0 Then
                WParidad = Str$(rstCtacte!Paridad)
                rstCtacte.Close
                    Else
                ClaveCtacte = "07" + Auxi1 + "01"
                spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtacte.RecordCount > 0 Then
                    WParidad = Str$(rstCtacte!Paridad)
                    rstCtacte.Close
                End If
            End If
        End If
        
        With rstImpCtaCte
            .Index = "Clave"
            .Seek "=", WClave
            If .NoMatch = False Then
                .Edit
                If WProv = 24 And WParidad <> 0 Then
                    !Importe3 = !Importe3 + (WImporte / Val(WParidad))
                        Else
                    !Importe3 = !Importe3 + WImporte
                End If
                .Update
            End If
        End With
    
    Next Ciclo
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    XClave = WAno + WMes + WDia
    WParidad = 0

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            WParidad = Str$(rstCambios!Cambio)
        End With
        rstCambios.Close
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
                    WProv = rstCliente!Provincia
                    rstCliente.Close
                End If
            
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
                If Pesos.Value = True Then
                    If WProv = 24 And WParidad <> 0 Then
                        !Importe3 = !Importe3 * Val(WParidad)
                    End If
                End If
                Acumula = Acumula + !Importe3
                Call Redondeo(Acumula)
                !Importe4 = Acumula
                !Razon = WRazon
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    If Dolares.Value = True Then
        With rstImpCtaCte
            .Index = "ClaveImpre"
            .MoveFirst
            Do
                .Edit
                If rstImpCtaCte!Totalus <> 0 Then
                    Pari = rstImpCtaCte!Total / rstImpCtaCte!Totalus
                    !Importe1 = !Importe1 / Pari
                    !Importe2 = !Importe2 / Pari
                    !Importe3 = !Importe3 / Pari
                End If
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    End If

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
    
    If Pesos.Value = True Then
        WTitulo = WTitulo + "Pesos"
    End If
    If Dolares.Value = True Then
        WTitulo = WTitulo + "Dolares"
    End If
    
    WTitulo = WTitulo + " al " + Fecha.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Cuenta Corriente a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    If Tipo1.Value = True Then
        Listado.GroupSelectionFormula = "{impCtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34) + " and {impCtaCte.Importe3} <> 0.00"
            Else
        Listado.GroupSelectionFormula = "{impCtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34) + " and {IMPCtaCte.Importe3} <> 999999.99"
    End If
    
    
    Listado.ReportFileName = "wctactefec.rpt"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Rem Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    Desde.SetFocus
    PrgCtaCteFec.Hide
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
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

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
    Ayuda.Visible = False
       
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

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Fecha.SetFocus
    End If
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Tipo1.Value = True
    Tipo2.Value = False
    Pesos.Value = True
    Dolares.Value = False
    CtaCte.Value = True
    Documentos.Value = False
    Total.Value = False
    Frame2.Visible = True
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstCliente!Cliente
                            IngresaItem = Auxi + "    " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
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
        rstCliente.Close
    End If
    End If

End Sub



