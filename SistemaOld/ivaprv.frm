VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgIvaprv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Iva por Provincias"
   ClientHeight    =   3825
   ClientLeft      =   3315
   ClientTop       =   2175
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   282
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
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
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Davod.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "ivaprv.frx":0000
      Left            =   840
      List            =   "ivaprv.frx":0007
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgIvaprv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstProvincia As Recordset
Dim spProvincia As String
Dim rstClientes As Recordset
Dim spClientes As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim XParam As String

Private Sub Acepta_Click()

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    WTitulo = "del " + Desde.Text + " al " + Hasta.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Ventas por provincias"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    da = ""
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

    spCtacte = "ListaCtacte"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
    With rstCtacte
            .MoveFirst
            Do
                If Val(!Tipo) >= 1 And Val(!Tipo) <= 5 Then
                
                    WCliente = !Cliente
                    WProvincia = ""
                
                    If !OrdFecha >= WDesde And !OrdFecha <= WHasta Then
                
                            WTipo = !Tipo
                            WImpre = !Impre
                            WNumero = !Numero
                            WRenglon = !Renglon
                            WCliente = !Cliente
                            WFecha = !Fecha
                            WEstado = !Estado
                            Wvencimiento = !Vencimiento
                            WVencimiento1 = !Vencimiento1
                            WTotal = !Total
                            WTotalUs = !Totalus
                            WSaldo = !Saldo
                            WSaldoUs = !Saldous
                            WNeto = !Neto
                            WIva1 = !Iva1
                            WIva2 = !Iva2
                            WImpoIb = !ImpoIb
                            WOrdFecha = !OrdFecha
                            WOrdVencimiento = !OrdVencimiento
                            WOrdVencimiento1 = !OrdVencimiento1
                            WPedido = !Pedido
                            WRemito = !Remito
                            WOrden = !Orden
                            WParidad = !Paridad
                            WVendedor = !Vendedor
                            WRubro = !Rubro
                            WComprobante = !Comprobante
                            WAceptada = !Aceptada
                            WCosto = !Costo
                            Rem WImporte1 = !Importe1
                            Rem WImporte2 = !Importe2
                            Rem WImporte3 = !Importe3
                            WImporte4 = !Importe4
                            WImporte5 = !Importe5
                            WImporte6 = !Importe6
                            WImporte7 = !Importe7
                            WImporte8 = !Importe8
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
                                !Vendedor = WVendedor
                                !Rubro = WRubro
                                !Comprobante = WComprobante
                                !Aceptada = WAceptada
                                !Costo = WCosto
                                !Importe1 = WImpoIb
                                !Importe2 = 0
                                !Importe3 = 0
                                !Importe4 = 0
                                !Importe5 = 0
                                !Importe6 = 0
                                !Importe7 = 0
                                !Clave = WClave
                        
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
    End If
    
    
    With rstImpCtaCte
            .Index = "Clave"
            .MoveFirst
            Do
                .Edit
                !Importe3 = 0
                !Importe4 = 0
                !Importe5 = 0
                !Importe6 = 0
                !Importe7 = 0
                DSA = !Numero
                
                
                If !Tipo >= 1 And !Tipo <= 5 Then
                
                    WCliente = !Cliente
                    WProvincia = ""
                    WRazon = ""
                                
                    spClientes = "ConsultaCliente " + "'" + WCliente + "'"
                    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
                    If rstClientes.RecordCount > 0 Then
                        WProvincia = rstClientes!Provincia
                        rstClientes.Close
                        spProvincia = "ConsultaProvincia " + "'" + WProvincia + "'"
                        Set rstProvincia = db.OpenRecordset(spProvincia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstProvincia.RecordCount > 0 Then
                            WRazon = rstProvincia!Nombre
                            rstProvincia.Close
                        End If
                    End If
                
                    !Provincia = WProvincia
                    !Razon = WRazon
                    
                    If !OrdFecha >= WDesde And !OrdFecha <= WHasta Then
                        If !Iva1 <> 0 Then
                            !Importe3 = !Importe1
                            !Importe4 = !Neto
                            !Importe5 = !Iva1
                            !Importe6 = !Iva2
                            !Importe7 = 0
                                Else
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            If Val(WEmpresa) = 4 Or Val(WEmpresa) = 8 Then
                                !Importe7 = !Total * !Paridad
                                    Else
                                !Importe7 = !Total
                            End If
                        End If
                    End If
                
                End If
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Listado.GroupSelectionFormula = "{ImpCtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()
    With rstImpCtaCte
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgIvaven.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_ImpCtacte
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub
Sub Form_load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

