VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCtaCte2 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Clientes"
   ClientHeight    =   7320
   ClientLeft      =   570
   ClientTop       =   1155
   ClientWidth     =   10995
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   10995
   Begin VB.CommandButton reclamo 
      Caption         =   "reclamos"
      Height          =   495
      Left            =   10440
      TabIndex        =   24
      Top             =   1560
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   4935
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8705
      _Version        =   327680
      Rows            =   4000
      Cols            =   10
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   3255
   End
   Begin VB.CommandButton datoscli 
      Caption         =   "Datos del Cliente"
      Height          =   735
      Left            =   9960
      TabIndex        =   21
      Top             =   720
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.CommandButton Listar 
      Caption         =   "Listar"
      Height          =   300
      Left            =   3480
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Datos"
      Height          =   1335
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Todos 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Pendiente 
         Caption         =   "Pendiente"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Listado"
      Height          =   1335
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Total 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Documentos 
         Caption         =   "Documentos"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton CtaCte 
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   1335
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Dolares 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Pesos 
         Caption         =   "Pesos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Cliente 
      Height          =   375
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      Height          =   300
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1230
      ItemData        =   "ctacte2.frx":0000
      Left            =   120
      List            =   "ctacte2.frx":0007
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCtaCte2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Importe1 As Double
Private Importe2 As Double
Private Importe3 As Double
Private WTipo As Integer
Private WSaldo As Double
Private Acumula As Double
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstCtacte As Recordset
Dim spCtecte As String
Dim XParam As String
Private WNume As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgCtaCte2.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()
    Sql1 = "UPDATE CtaCte SET "
    Sql2 = " Saldo  = 0" + ","
    Sql3 = " SaldoUs = 0"
    Sql4 = " Where Cliente = " + "'" + Cliente.Text + "'"
                     
    spCtacte = Sql1 + Sql2 + Sql3 + Sql4
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaClienteConsulta1"
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
    
    Rem Pantalla.Visible = True

End Sub

Private Sub datoscli_Click()
    PCliente = Cliente.Text
    prgcli.Show
End Sub

Private Sub pantalla_Click()
    Rem Pantalla.Visible = False
       
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    Cliente.Text = Claveven$
    Rem by nan
    cliente2 = Cliente.Text
    spCliente = "ConsultaCliente " + "'" + Claveven$ + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Cliente.Text = rstCliente!Cliente
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
        Call Proceso_Click
            Else
        Cliente.Text = Claveven$
    End If
    Cliente.SetFocus
    descliente2 = DesCliente
End Sub

Private Sub Form_Load()


    Call Limpia_Vector
    
    Muestra.ColWidth(0) = 150
    Muestra.ColWidth(1) = 500
    Muestra.ColWidth(2) = 1000
    Muestra.ColWidth(3) = 1300
    Muestra.ColWidth(4) = 1000
    Muestra.ColWidth(5) = 1100
    Muestra.ColWidth(6) = 1000
    Muestra.ColWidth(7) = 1300
    Muestra.ColWidth(8) = 1300
    Muestra.ColWidth(9) = 1200
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Numero"
    
    Muestra.Col = 3
    Muestra.Text = "Fecha"
    
    Muestra.Col = 4
    Muestra.Text = "Debito"
    
    Muestra.Col = 5
    Muestra.Text = "Credito"
    
    Muestra.Col = 6
    Muestra.Text = "Saldo"
    
    Muestra.Col = 7
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 8
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 9
    Muestra.Text = "Acumulado"
 
    Cliente.Text = ""
    DesCliente.Caption = ""

    Pesos.Value = True
    CtaCte.Value = True
    Pendiente.Value = True
    
    Call Consulta_Click
    
    Cliente.Text = PCliente
    
End Sub

Private Sub Proceso_Click()

    Cliente.Text = UCase(Cliente.Text)
    
    WSalida = "N"
        
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Numero"
    
    Muestra.Col = 3
    Muestra.Text = "Fecha"
    
    Muestra.Col = 4
    Muestra.Text = "Debito"
    
    Muestra.Col = 5
    Muestra.Text = "Credito"
    
    Muestra.Col = 6
    Muestra.Text = "Saldo"
    
    Muestra.Col = 7
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 8
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 9
    Muestra.Text = "Acumulado"

    Renglon = 0
    WSaldo = 0
    
    XParam = "'" + Cliente.Text + "'"
    spCtacte = "ListaCtacteCliente " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
    With rstCtacte
    
        .MoveFirst
        If .NoMatch = False Then
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
                
                If Total.Value = True Then
                    WPasa = "S"
                End If
                    
                If WPasa = "S" Then
                    If Pesos.Value = True Then
                        If !Total > 0 Then
                            Importe1 = !Total
                            Importe2 = 0
                                Else
                            Importe1 = 0
                            Importe2 = !Total
                        End If
                        Importe3 = !Saldo
                            Else
                        If !Totalus > 0 Then
                            Importe1 = !Totalus
                            Importe2 = 0
                                Else
                            Importe1 = 0
                            Importe2 = !Totalus
                        End If
                        Importe3 = !Saldous
                    End If
                    
                    Call Redondeo(Importe3)
                
                    If Importe3 <> 0 Or Todos.Value = True Then
                
                        Renglon = Renglon + 1
            
                        Muestra.Row = Renglon
                
                        Select Case !Tipo
                            Case 1
                                Muestra.Col = 1
                                Muestra.Text = "Fac"
                            Case 2
                                Muestra.Col = 1
                                Muestra.Text = "Dev"
                            Case 3
                                Muestra.Col = 1
                                Muestra.Text = "Fac"
                            Case 4
                                Muestra.Col = 1
                                Select Case Left$(!Impre, 2)
                                    Case "DC"
                                        Muestra.Text = "D.C"
                                    Case "CH"
                                        Muestra.Text = "CHR"
                                    Case Else
                                        Muestra.Text = "N/D"
                                End Select
                            Case 5
                                Muestra.Col = 1
                                Select Case Left$(!Impre, 2)
                                    Case "DC"
                                        Muestra.Text = "D.C"
                                    Case "CH"
                                        Muestra.Text = "CHR"
                                    Case Else
                                        Muestra.Text = "N/C"
                                End Select
                            Case 6
                                Muestra.Col = 1
                                Muestra.Text = "Rec"
                            Case 7
                                Muestra.Col = 1
                                Muestra.Text = "Ant"
                            Case 10
                                Muestra.Col = 1
                                Muestra.Text = "FCR"
                            Case 50
                                Muestra.Col = 1
                                Muestra.Text = "Doc"
                            Case Else
                        End Select
                        
                        Muestra.Col = 2
                        Muestra.Text = Pusing("######", Str$(!Numero))
                
                        Muestra.Col = 3
                        Muestra.Text = !Fecha
                
                        If Importe1 <> 0 Then
                            Muestra.Col = 4
                            Muestra.Text = Pusing("###,###,###.##", Str$(Importe1))
                                Else
                            Muestra.Col = 4
                            Muestra.Text = ""
                        End If
                
                        If Importe2 <> 0 Then
                            Muestra.Col = 5
                            Muestra.Text = Pusing("###,###,###.##", Str$(Importe2))
                                Else
                            Muestra.Col = 5
                            Muestra.Text = ""
                        End If
                
                        If Importe3 <> 0 Then
                            Muestra.Col = 6
                            Muestra.Text = Pusing("###,###,###.##", Str$(Importe3))
                                Else
                            Muestra.Col = 6
                            Muestra.Text = ""
                        End If
                        
                        WSaldo = WSaldo + Importe3
                
                        Muestra.Col = 7
                        Muestra.Text = !Vencimiento
                        
                        Muestra.Col = 8
                        Muestra.Text = !Vencimiento1
                        
                        Muestra.Col = 9
                        Muestra.Text = Pusing("###,###,###.##", Str$(WSaldo))
                    
                    End If
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    End If
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Saldo.Caption = Pusing("###,###,###.##", Str$(WSaldo))
    
    Muestra.TopRow = 1
    
    Muestra.SetFocus

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        cliente2 = Cliente.Text
        Cliente.Text = WCliente
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                rstCliente.Close
                Call Proceso_Click
                Muestra.SetFocus
                    Else
                Cliente.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Numero"
    
    Muestra.Col = 3
    Muestra.Text = "Fecha"
    
    Muestra.Col = 4
    Muestra.Text = "Debito"
    
    Muestra.Col = 5
    Muestra.Text = "Credito"
    
    Muestra.Col = 6
    Muestra.Text = "Saldo"
    
    Muestra.Col = 7
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 8
    Muestra.Text = "Vencimiento"
    
    Muestra.Col = 9
    Muestra.Text = "Acumulado"
    
End Sub

Private Sub Listar_Click()

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


    WDesde = UCase(Cliente.Text)
    WHasta = UCase(Cliente.Text)
    
    spCtacte = "ModificaCtacteTipo1"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    spCtacte = "ModificaCtacteTipo2"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    spCtacte = "ModificaCtacteImporte0"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)

    If CtaCte.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte1 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte2 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte3 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte4 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Documentos.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte5 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte6 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte7 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte8 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            End If
    End If
                
    If Total.Value = True Then
            If Pesos.Value = True Then
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte9 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte10 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
                spCtacte = "ModificaCtacte11 " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                XParam = "'" + WDesde + "','" _
                        + WHasta + "'"
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
    
    XParam = "'" + WDesde + "','" _
            + WHasta + "'"
    spCtacte = "ListaCtacteDesdeHasta " + XParam
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
                
                If Total.Value = True Then
                    WPasa = "S"
                End If
                    
                If WPasa = "S" Then
            
                If Todos.Value = True Or !Importe3 <> 0 Then
            
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
    
    With rstImpCtaCte
            .Index = "ClaveImpre"
            .MoveFirst
            Do
            
                WRazon = ""
                spCliente = "ConsultaCliente " + !Cliente
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WRazon = rstCliente!Razon
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
                Acumula = Acumula + !Importe3
                Call Redondeo(Acumula)
                !Importe4 = Acumula
                !Razon = WRazon
                !Titulo = WTitulo
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
    
    Listado.GroupSelectionFormula = "{impCtaCte.Cliente} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.ReportFileName = "wimpctacte.rpt"
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Rem Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Rem Listado.Connect = Connect()
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.Action = 1
    
End Sub


Private Sub Muestra_DblClick()

    Muestra.Col = 1
    Tipo = Muestra.Text
    
    If Tipo = "Rec" Then
        Muestra.Col = 2
        WRecibo = Muestra.Text
        PrgRec.Show
    End If

    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta1"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
    With rstCliente
        .MoveFirst
        Do
            If .EOF = False Then
            
                DA = Len(!Razon) - WEspacios
                
                For aa = 1 To DA
                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Razon), aa, WEspacios) Then
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
    
    End If

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

Private Sub reclamo_Click()
Prgreclamo.Show
End Sub
