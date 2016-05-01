VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaComando 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Comando"
   ClientHeight    =   4080
   ClientLeft      =   2175
   ClientTop       =   945
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   8145
   Begin VB.TextBox MiraII 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox mira 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   2040
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
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   2040
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
         Height          =   495
         Left            =   3480
         TabIndex        =   3
         Top             =   600
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
         Height          =   495
         Left            =   3480
         TabIndex        =   2
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
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
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6840
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEsta7.rpt"
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
End
Attribute VB_Name = "PrgListaComando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Uno As String
Private Dos As String
Private Costo As Double
Private Producto As String
Private Auxiliar(1000, 7) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTermnado As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstLinea As Recordset
Dim spLinea As String
Dim rstComando As Recordset
Dim spComando As String
Dim rstStockHistorico As Recordset
Dim spStockHistorico As String
Dim XParam As String
Private Vecosto(5000, 2) As String
Dim Posi As Integer
Dim ImpreFecha(12, 2) As String
Dim ZMes As String
Dim ZAno As String
Dim PVenta(12) As Double
Dim PKilos(12) As Double
Dim PCosto(12) As Double
Dim PStock(12) As Double
Dim PPedidos(12) As Double
Dim PAtraso(12) As Double
Dim PFactor(12) As Double
Dim PPrecio(12) As Double
Dim PPorceVenta(12) As Double
Dim PPorceAtraso(12) As Double
Dim PRotacion(12) As Double
Dim WVectorPedido(20000, 11) As String
Dim LugarPedido As Integer
Dim WEntrada(50000) As String
Dim ZStockHistorico(20, 10) As String

Private Sub Acepta_Click()

    Rem On Error GoTo WError
    
    MesIni = Val(Mid$(DesdeFec.Text, 4, 2))
    AnoIni = Val(Right$(DesdeFec.Text, 4))
    
    ZMes = MesIni
    ZAno = AnoIni
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    
    ZInicioPeriodo = ZAno + ZMes
    
    MesFin = Val(Mid$(HastaFec.Text, 4, 2))
    AnoFin = Val(Right$(HastaFec.Text, 4))
    
    ZMes = MesFin
    ZAno = AnoFin
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    
    ZFinPeriodo = ZAno + ZMes
    
    MesCicla = MesIni
    AnoCicla = AnoIni
    Erase ImpreFecha
    
    For Ciclo = 1 To 12
    
        ImpreFecha(Ciclo, 1) = Str$(MesCicla)
        ImpreFecha(Ciclo, 2) = Str$(AnoCicla)
        
        MesCicla = MesCicla + 1
        If MesCicla > 12 Then
            MesCicla = 1
            AnoCicla = AnoCicla + 1
        End If
        
        ZMes = MesCicla
        ZAno = AnoCicla
        
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        
        WComparaPeriodo = ZAno + ZMes
        
        If WComparaPeriodo > ZFinPeriodo Then
            Exit For
        End If
        
    Next Ciclo
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            Nomempresa = !Nombre
        End If
    End With
    
    With rstEstaComando
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
    
    WTitulo1 = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    WTitulo3 = Nomempresa
    
    Rem inicializa los DATOS
    
    Sql1 = "UPDATE Comando SET "
    Sql2 = "Venta1 = 0,"
    Sql3 = "Venta2 = 0,"
    Sql4 = "Venta3 = 0,"
    Sql5 = "Venta4 = 0,"
    Sql6 = "Venta5 = 0,"
    Sql7 = "Venta6 = 0,"
    Sql8 = "Venta7 = 0,"
    Sql9 = "Venta8 = 0,"
    Sql10 = "Venta9 = 0,"
    Sql11 = "Venta10 = 0,"
    Sql12 = "Venta11 = 0,"
    Sql13 = "Venta12 = 0,"
    Sql14 = "Kilos1 = 0,"
    Sql15 = "Kilos2 = 0,"
    Sql16 = "Kilos3 = 0,"
    Sql17 = "Kilos4 = 0,"
    Sql18 = "Kilos5 = 0,"
    Sql19 = "Kilos6 = 0,"
    Sql20 = "Kilos7 = 0,"
    Sql21 = "Kilos8 = 0,"
    Sql22 = "Kilos9 = 0,"
    Sql23 = "Kilos10 = 0,"
    Sql24 = "Kilos11 = 0,"
    Sql25 = "Kilos12 = 0,"
    Sql26 = "Costo1 = 0,"
    Sql27 = "Costo2 = 0,"
    Sql28 = "Costo3 = 0,"
    Sql29 = "Costo4 = 0,"
    Sql30 = "Costo5 = 0,"
    Sql31 = "Costo6 = 0,"
    Sql32 = "Costo7 = 0,"
    Sql33 = "Costo8 = 0,"
    Sql34 = "Costo9 = 0,"
    Sql35 = "Costo10 = 0,"
    Sql36 = "Costo11 = 0,"
    Sql37 = "Costo12 = 0,"
    Sql38 = "Stock1 = 0,"
    Sql39 = "Stock2 = 0,"
    Sql40 = "Stock3 = 0,"
    Sql41 = "Stock4 = 0,"
    Sql42 = "Stock5 = 0,"
    Sql43 = "Stock6 = 0,"
    Sql44 = "Stock7 = 0,"
    Sql45 = "Stock8 = 0,"
    Sql46 = "Stock9 = 0,"
    Sql47 = "Stock10 = 0,"
    Sql48 = "Stock11 = 0,"
    Sql49 = "Stock12 = 0,"
    Sql50 = "Pedidos1 = 0,"
    Sql51 = "Pedidos2 = 0,"
    Sql52 = "Pedidos3 = 0,"
    Sql53 = "Pedidos4 = 0,"
    Sql54 = "Pedidos5 = 0,"
    Sql55 = "Pedidos6 = 0,"
    Sql56 = "Pedidos7 = 0,"
    Sql57 = "Pedidos8 = 0,"
    Sql58 = "Pedidos9 = 0,"
    Sql59 = "Pedidos10 = 0,"
    Sql60 = "Pedidos11 = 0,"
    Sql61 = "Pedidos12 = 0,"
    Sql62 = "Atraso1 = 0,"
    Sql63 = "Atraso2 = 0,"
    Sql64 = "Atraso3 = 0,"
    Sql65 = "Atraso4 = 0,"
    Sql66 = "Atraso5 = 0,"
    Sql67 = "Atraso6 = 0,"
    Sql68 = "Atraso7 = 0,"
    Sql69 = "Atraso8 = 0,"
    Sql70 = "Atraso9 = 0,"
    Sql71 = "Atraso10 = 0,"
    Sql72 = "Atraso11 = 0,"
    Sql73 = "Atraso12 = 0,"
    Sql74 = "Proye1 = 0,"
    Sql75 = "Proye2 = 0,"
    Sql76 = "Proye3 = 0,"
    Sql77 = "Proye4 = 0,"
    Sql78 = "Proye5 = 0,"
    Sql79 = "Proye6 = 0,"
    Sql80 = "Proye7 = 0,"
    Sql81 = "Proye8 = 0,"
    Sql82 = "Proye9 = 0,"
    Sql83 = "Proye10 = 0,"
    Sql84 = "Proye11 = 0,"
    Sql85 = "Proye12 = 0"
                     
    spComando = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                Sql41 + Sql42 + Sql43 + Sql44 + Sql45 + Sql46 + Sql47 + Sql48 + Sql49 + Sql50 + _
                Sql51 + Sql52 + Sql53 + Sql54 + Sql55 + Sql56 + Sql57 + Sql58 + Sql59 + Sql60 + _
                Sql61 + Sql62 + Sql63 + Sql64 + Sql65 + Sql66 + Sql67 + Sql68 + Sql69 + Sql70 + _
                Sql71 + Sql72 + Sql73 + Sql74 + Sql75 + Sql76 + Sql77 + Sql78 + Sql79 + Sql80 + _
                Sql81 + Sql82 + Sql83 + Sql84 + Sql85
    Set rstComando = db.OpenRecordset(spComando, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem XParam = "'" + WDesde + "','" _
    rem              + WHasta + "'"
    Rem spEstadistica = "ListaEstadisticaFecha" + XParam
    Rem Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstEstadistica.RecordCount > 0 Then
    
    
    Rem dadadada
    
    ZZLugar = 0
    
    Sql1 = "Select estadistica.tipo, estadistica.Articulo, estadistica.cantidad, estadistica.Precio, estadistica.PrecioUs, estadistica.Importe, estadistica.ImporteUs, estadistica.cliente, estadistica.LInea, estadistica.Fecha, estadistica.OrdFecha, estadistica.WArticulo  "
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.Ordfecha >= " + "'" + WDesde + "'"
    Sql4 = " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        
        With rstEstadistica
    
            .MoveFirst
            
            Do
                
                ZZLugar = ZZLugar + 1
                mira.Text = ZZLugar
                DoEvents
                
                WTipo = rstEstadistica!Tipo
                Rem WNumero = rstEstadistica!NUMERO
                Rem WRenglon = rstEstadistica!Renglon
                WArticulo = rstEstadistica!Articulo
                WCantidad = rstEstadistica!Cantidad
                WPrecio = rstEstadistica!Precio
                WPrecioUs = rstEstadistica!PrecioUs
                WImporte = rstEstadistica!Importe
                WimporteUs = rstEstadistica!ImporteUs
                WCliente = rstEstadistica!Cliente
                Rem WParidad = rstEstadistica!Paridad
                Rem WVendedor = rstEstadistica!Vendedor
                Rem WRubro = rstEstadistica!Rubro
                WLinea = rstEstadistica!Linea
                Rem WCosto1 = rstEstadistica!Costo1
                Rem WCosto2 = rstEstadistica!Costo2
                Rem WCoeficiente = rstEstadistica!Coeficiente
                Rem WPedido = rstEstadistica!Pedido
                WFecha = rstEstadistica!Fecha
                Rem WImporte1 = rstEstadistica!Importe1
                Rem WImporte2 = rstEstadistica!Importe2
                Rem WImporte3 = rstEstadistica!Importe3
                Rem WImporte4 = rstEstadistica!Importe4
                WOrdFecha = rstEstadistica!OrdFecha
                WWArticulo = rstEstadistica!WArticulo
                Rem WRemito = rstEstadistica!Remito
                Rem WClave = rstEstadistica!Clave
                
                Impo1 = 0
                Impo2 = 0
                Impo3 = 0
                Impo4 = 0
                Impo5 = 0
                Impo6 = 0
                Impo7 = 0
                Impo8 = 0
                Impo9 = 0
                Impo10 = 0
                Impo11 = 0
                Impo12 = 0
                    
                Pesos1 = 0
                Pesos2 = 0
                Pesos3 = 0
                Pesos4 = 0
                Pesos5 = 0
                Pesos6 = 0
                Pesos7 = 0
                Pesos8 = 0
                Pesos9 = 0
                Pesos10 = 0
                Pesos11 = 0
                Pesos12 = 0
                    
                Canti1 = 0
                Canti2 = 0
                Canti3 = 0
                Canti4 = 0
                Canti5 = 0
                Canti6 = 0
                Canti7 = 0
                Canti8 = 0
                Canti9 = 0
                Canti10 = 0
                Canti11 = 0
                Canti12 = 0
                
                MesCompara = Val(Mid$(WFecha, 4, 2))
                AnoCompara = Val(Right$(WFecha, 4))
                    
                If !Tipo = 2 Then
                    WCantidad = Abs(WCantidad) * -1
                    WImporte = Abs(WImporte) * -1
                    WimporteUs = Abs(WimporteUs) * -1
                End If
                        
                If MesCompara = Val(ImpreFecha(1, 1)) And AnoCompara = Val(ImpreFecha(1, 2)) Then
                    Impo1 = WimporteUs
                    Pesos1 = WImporte
                    Canti1 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(2, 1)) And AnoCompara = Val(ImpreFecha(2, 2)) Then
                    Impo2 = WimporteUs
                    Pesos2 = WImporte
                    Canti2 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(3, 1)) And AnoCompara = Val(ImpreFecha(3, 2)) Then
                    Impo3 = WimporteUs
                    Pesos3 = WImporte
                    Canti3 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(4, 1)) And AnoCompara = Val(ImpreFecha(4, 2)) Then
                    Impo4 = WimporteUs
                    Pesos4 = WImporte
                    Canti4 = WCantidad
                End If
                
                If MesCompara = Val(ImpreFecha(5, 1)) And AnoCompara = Val(ImpreFecha(5, 2)) Then
                    Impo5 = WimporteUs
                    Pesos5 = WImporte
                    Canti5 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(6, 1)) And AnoCompara = Val(ImpreFecha(6, 2)) Then
                    Impo6 = WimporteUs
                    Pesos6 = WImporte
                    Canti6 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(7, 1)) And AnoCompara = Val(ImpreFecha(7, 2)) Then
                    Impo7 = WimporteUs
                    Pesos7 = WImporte
                    Canti7 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(8, 1)) And AnoCompara = Val(ImpreFecha(8, 2)) Then
                    Impo8 = WimporteUs
                    Pesos8 = WImporte
                    Canti8 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(9, 1)) And AnoCompara = Val(ImpreFecha(9, 2)) Then
                    Impo9 = WimporteUs
                    Pesos9 = WImporte
                    Canti9 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(10, 1)) And AnoCompara = Val(ImpreFecha(10, 2)) Then
                    Impo10 = WimporteUs
                    Pesos10 = WImporte
                    Canti10 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(11, 1)) And AnoCompara = Val(ImpreFecha(11, 2)) Then
                    Impo11 = WimporteUs
                    Pesos11 = WImporte
                    Canti11 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(12, 1)) And AnoCompara = Val(ImpreFecha(12, 2)) Then
                    Impo12 = WimporteUs
                    Pesos12 = WImporte
                    Canti12 = WCantidad
                End If
                
                ZLinea = 0
                If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Then
                    ZLinea = 6
                        Else
                    If WArticulo = "PT-99999-999" Then
                        ZLinea = 99
                            Else
                        ZLinea = WLinea
                    End If
                End If
                Select Case ZLinea
                    Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16, 17, 19, 99
                        WCliente = "Z99999"
                    Case Else
                End Select
                WClave = WCliente + WArticulo
                
                With rstEstaComando
                    .Index = "Clave"
                    .Seek "=", WClave
                    If .NoMatch = True Then
                        .AddNew
                        !Clave = WClave
                        !Cliente = WCliente
                        !Codigo = WArticulo
                        !Linea = ZLinea
                        !Impo1 = Impo1
                        !Impo2 = Impo2
                        !Impo3 = Impo3
                        !Impo4 = Impo4
                        !Impo5 = Impo5
                        !Impo6 = Impo6
                        !Impo7 = Impo7
                        !Impo8 = Impo8
                        !Impo9 = Impo9
                        !Impo10 = Impo10
                        !Impo11 = Impo11
                        !Impo12 = Impo12
                        !Titulo1 = WTitulo1
                        !Titulo2 = WTitulo2
                        !Titulo3 = WTitulo3
                        !Pesos1 = Pesos1
                        !Pesos2 = Pesos2
                        !Pesos3 = Pesos3
                        !Pesos4 = Pesos4
                        !Pesos5 = Pesos5
                        !Pesos6 = Pesos6
                        !Pesos7 = Pesos7
                        !Pesos8 = Pesos8
                        !Pesos9 = Pesos9
                        !Pesos10 = Pesos10
                        !Pesos11 = Pesos11
                        !Pesos12 = Pesos12
                        !Canti1 = Canti1
                        !Canti2 = Canti2
                        !Canti3 = Canti3
                        !Canti4 = Canti4
                        !Canti5 = Canti5
                        !Canti6 = Canti6
                        !Canti7 = Canti7
                        !Canti8 = Canti8
                        !Canti9 = Canti9
                        !Canti10 = Canti10
                        !Canti11 = Canti11
                        !Canti12 = Canti12
                        .Update
                            Else
                        .Edit
                        !Impo1 = !Impo1 + Impo1
                        !Impo2 = !Impo2 + Impo2
                        !Impo3 = !Impo3 + Impo3
                        !Impo4 = !Impo4 + Impo4
                        !Impo5 = !Impo5 + Impo5
                        !Impo6 = !Impo6 + Impo6
                        !Impo7 = !Impo7 + Impo7
                        !Impo8 = !Impo8 + Impo8
                        !Impo9 = !Impo9 + Impo9
                        !Impo10 = !Impo10 + Impo10
                        !Impo11 = !Impo11 + Impo11
                        !Impo12 = !Impo12 + Impo12
                        !Pesos1 = !Pesos1 + Pesos1
                        !Pesos2 = !Pesos2 + Pesos2
                        !Pesos3 = !Pesos3 + Pesos3
                        !Pesos4 = !Pesos4 + Pesos4
                        !Pesos5 = !Pesos5 + Pesos5
                        !Pesos6 = !Pesos6 + Pesos6
                        !Pesos7 = !Pesos7 + Pesos7
                        !Pesos8 = !Pesos8 + Pesos8
                        !Pesos9 = !Pesos9 + Pesos9
                        !Pesos10 = !Pesos10 + Pesos10
                        !Pesos11 = !Pesos11 + Pesos11
                        !Pesos12 = !Pesos12 + Pesos12
                        !Canti1 = !Canti1 + Canti1
                        !Canti2 = !Canti2 + Canti2
                        !Canti3 = !Canti3 + Canti3
                        !Canti4 = !Canti4 + Canti4
                        !Canti5 = !Canti5 + Canti5
                        !Canti6 = !Canti6 + Canti6
                        !Canti7 = !Canti7 + Canti7
                        !Canti8 = !Canti8 + Canti8
                        !Canti9 = !Canti9 + Canti9
                        !Canti10 = !Canti10 + Canti10
                        !Canti11 = !Canti11 + Canti11
                        !Canti12 = !Canti12 + Canti12
                        .Update
                    End If
                End With
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        rstEstadistica.Close
    End If
    
    
    
    
    
    
    ZZLugar = 0
    
    With rstEsta8
        .Index = "Clave"
        .MoveFirst
        Do
            If WDesde <= !OrdFecha And !OrdFecha <= WHasta Then
            
                ZZLugar = ZZLugar + 1
                mira.Text = ZZLugar
                DoEvents
            
                WTipo = !Tipo
                WNumero = !NUMERO
                WRenglon = !Renglon
                WArticulo = !Articulo
                WCantidad = !Cantidad
                WPrecio = !Precio
                WPrecioUs = !PrecioUs
                WImporte = !Importe
                WimporteUs = !ImporteUs
                WCliente = !Cliente
                WParidad = !Paridad
                WRubro = !Rubro
                WCosto1 = !Costo1
                WCosto2 = !Costo2
                WCoeficiente = !Coeficiente
                WPedido = !Pedido
                WFecha = !Fecha
                WImporte1 = !Importe1
                WImporte2 = !Importe2
                WImporte3 = !Importe3
                WImporte4 = !Importe4
                WOrdFecha = !OrdFecha
                WWArticulo = !WArticulo
                WRemito = !Remito
                WClave = !Clave
                        
                If WPedido = 0 Then
                    WCantidad = 0
                End If
                        
                Impo1 = 0
                Impo2 = 0
                Impo3 = 0
                Impo4 = 0
                Impo5 = 0
                Impo6 = 0
                Impo7 = 0
                Impo8 = 0
                Impo9 = 0
                Impo10 = 0
                Impo11 = 0
                Impo12 = 0
                    
                Pesos1 = 0
                Pesos2 = 0
                Pesos3 = 0
                Pesos4 = 0
                Pesos5 = 0
                Pesos6 = 0
                Pesos7 = 0
                Pesos8 = 0
                Pesos9 = 0
                Pesos10 = 0
                Pesos11 = 0
                Pesos12 = 0
                    
                Canti1 = 0
                Canti2 = 0
                Canti3 = 0
                Canti4 = 0
                Canti5 = 0
                Canti6 = 0
                Canti7 = 0
                Canti8 = 0
                Canti9 = 0
                Canti10 = 0
                Canti11 = 0
                Canti12 = 0
                
                MesCompara = Val(Mid$(WFecha, 4, 2))
                AnoCompara = Val(Right$(WFecha, 4))
                    
                If !Tipo = 2 Then
                    WCantidad = Abs(WCantidad) * -1
                    WImporte = Abs(WImporte) * -1
                    WimporteUs = Abs(WimporteUs) * -1
                End If
                        
                If MesCompara = Val(ImpreFecha(1, 1)) And AnoCompara = Val(ImpreFecha(1, 2)) Then
                    Impo1 = WimporteUs
                    Pesos1 = WImporte
                    Canti1 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(2, 1)) And AnoCompara = Val(ImpreFecha(2, 2)) Then
                    Impo2 = WimporteUs
                    Pesos2 = WImporte
                    Canti2 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(3, 1)) And AnoCompara = Val(ImpreFecha(3, 2)) Then
                    Impo3 = WimporteUs
                    Pesos3 = WImporte
                    Canti3 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(4, 1)) And AnoCompara = Val(ImpreFecha(4, 2)) Then
                    Impo4 = WimporteUs
                    Pesos4 = WImporte
                    Canti4 = WCantidad
                End If
                
                If MesCompara = Val(ImpreFecha(5, 1)) And AnoCompara = Val(ImpreFecha(5, 2)) Then
                    Impo5 = WimporteUs
                    Pesos5 = WImporte
                    Canti5 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(6, 1)) And AnoCompara = Val(ImpreFecha(6, 2)) Then
                    Impo6 = WimporteUs
                    Pesos6 = WImporte
                    Canti6 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(7, 1)) And AnoCompara = Val(ImpreFecha(7, 2)) Then
                    Impo7 = WimporteUs
                    Pesos7 = WImporte
                    Canti7 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(8, 1)) And AnoCompara = Val(ImpreFecha(8, 2)) Then
                    Impo8 = WimporteUs
                    Pesos8 = WImporte
                    Canti8 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(9, 1)) And AnoCompara = Val(ImpreFecha(9, 2)) Then
                    Impo9 = WimporteUs
                    Pesos9 = WImporte
                    Canti9 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(10, 1)) And AnoCompara = Val(ImpreFecha(10, 2)) Then
                    Impo10 = WimporteUs
                    Pesos10 = WImporte
                    Canti10 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(11, 1)) And AnoCompara = Val(ImpreFecha(11, 2)) Then
                    Impo11 = WimporteUs
                    Pesos11 = WImporte
                    Canti11 = WCantidad
                End If
                    
                If MesCompara = Val(ImpreFecha(12, 1)) And AnoCompara = Val(ImpreFecha(12, 2)) Then
                    Impo12 = WimporteUs
                    Pesos12 = WImporte
                    Canti12 = WCantidad
                End If
                
                ZLinea = 0
                If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Then
                    ZLinea = 6
                        Else
                    If WArticulo = "PT-99999-999" Then
                        ZLinea = 99
                            Else
                        Sql1 = "Select *"
                        Sql2 = " FROM Terminado"
                        Sql3 = " Where Terminado.Codigo = " + "'" + WArticulo + "'"
                        spTerminado = Sql1 + Sql2 + Sql3
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                    End If
                End If
                Select Case ZLinea
                    Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16, 17, 19, 99
                        WCliente = "Z99999"
                    Case Else
                End Select
                WClave = WCliente + WArticulo
                
                With rstEstaComando
                    .Index = "Clave"
                    .Seek "=", WClave
                    If .NoMatch = True Then
                        .AddNew
                        !Clave = WClave
                        !Cliente = WCliente
                        !Codigo = WArticulo
                        !Linea = ZLinea
                        !Impo1 = Impo1
                        !Impo2 = Impo2
                        !Impo3 = Impo3
                        !Impo4 = Impo4
                        !Impo5 = Impo5
                        !Impo6 = Impo6
                        !Impo7 = Impo7
                        !Impo8 = Impo8
                        !Impo9 = Impo9
                        !Impo10 = Impo10
                        !Impo11 = Impo11
                        !Impo12 = Impo12
                        !Titulo1 = WTitulo1
                        !Titulo2 = WTitulo2
                        !Titulo3 = WTitulo3
                        !Pesos1 = Pesos1
                        !Pesos2 = Pesos2
                        !Pesos3 = Pesos3
                        !Pesos4 = Pesos4
                        !Pesos5 = Pesos5
                        !Pesos6 = Pesos6
                        !Pesos7 = Pesos7
                        !Pesos8 = Pesos8
                        !Pesos9 = Pesos9
                        !Pesos10 = Pesos10
                        !Pesos11 = Pesos11
                        !Pesos12 = Pesos12
                        !Canti1 = Canti1
                        !Canti2 = Canti2
                        !Canti3 = Canti3
                        !Canti4 = Canti4
                        !Canti5 = Canti5
                        !Canti6 = Canti6
                        !Canti7 = Canti7
                        !Canti8 = Canti8
                        !Canti9 = Canti9
                        !Canti10 = Canti10
                        !Canti11 = Canti11
                        !Canti12 = Canti12
                        .Update
                            Else
                        .Edit
                        !Impo1 = !Impo1 + Impo1
                        !Impo2 = !Impo2 + Impo2
                        !Impo3 = !Impo3 + Impo3
                        !Impo4 = !Impo4 + Impo4
                        !Impo5 = !Impo5 + Impo5
                        !Impo6 = !Impo6 + Impo6
                        !Impo7 = !Impo7 + Impo7
                        !Impo8 = !Impo8 + Impo8
                        !Impo9 = !Impo9 + Impo9
                        !Impo10 = !Impo10 + Impo10
                        !Impo11 = !Impo11 + Impo11
                        !Impo12 = !Impo12 + Impo12
                        !Pesos1 = !Pesos1 + Pesos1
                        !Pesos2 = !Pesos2 + Pesos2
                        !Pesos3 = !Pesos3 + Pesos3
                        !Pesos4 = !Pesos4 + Pesos4
                        !Pesos5 = !Pesos5 + Pesos5
                        !Pesos6 = !Pesos6 + Pesos6
                        !Pesos7 = !Pesos7 + Pesos7
                        !Pesos8 = !Pesos8 + Pesos8
                        !Pesos9 = !Pesos9 + Pesos9
                        !Pesos10 = !Pesos10 + Pesos10
                        !Pesos11 = !Pesos11 + Pesos11
                        !Pesos12 = !Pesos12 + Pesos12
                        !Canti1 = !Canti1 + Canti1
                        !Canti2 = !Canti2 + Canti2
                        !Canti3 = !Canti3 + Canti3
                        !Canti4 = !Canti4 + Canti4
                        !Canti5 = !Canti5 + Canti5
                        !Canti6 = !Canti6 + Canti6
                        !Canti7 = !Canti7 + Canti7
                        !Canti8 = !Canti8 + Canti8
                        !Canti9 = !Canti9 + Canti9
                        !Canti10 = !Canti10 + Canti10
                        !Canti11 = !Canti11 + Canti11
                        !Canti12 = !Canti12 + Canti12
                        .Update
                    End If
                End With
                
            End If
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    
    Erase WEntrada
    LugarEntrada = 0
    ZZLugar = 0
    
    spArticulo = "ListaArticuloStock"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
    
            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                
                If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DW" Then
                
                    ZZLugar = ZZLugar + 1
                    mira.Text = ZZLugar
                    DoEvents
                    
                    LugarEntrada = LugarEntrada + 1
                    WEntrada(LugarEntrada) = Left$(rstArticulo!Codigo, 3) + "00" + Right$(rstArticulo!Codigo, 7)
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        rstArticulo.Close
    End If
    
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        With rstTerminado
            .MoveFirst
                
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
               
                    If Left$(rstTerminado!Codigo, 2) = "PT" Or Left$(rstTerminado!Codigo, 2) = "PE" Or Left$(rstTerminado!Codigo, 2) = "SU" Or Left$(rstTerminado!Codigo, 2) = "SE" Then
                    
                        ZZLugar = ZZLugar + 1
                        mira.Text = ZZLugar
                        DoEvents
                        
                        LugarEntrada = LugarEntrada + 1
                        WEntrada(LugarEntrada) = rstTerminado!Codigo
                        
                    End If
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            
        End With
        rstTerminado.Close
    
    End If
    
    MiraII.Text = LugarEntrada
    
    For Ciclo = 1 To LugarEntrada
    
        mira.Text = Ciclo
        DoEvents
    
        WArticulo = WEntrada(Ciclo)
        WCliente = "Z99999"
        WClave = WCliente + WArticulo
        
        With rstEstaComando
            .Index = "Codigo"
            .Seek "=", WArticulo
            If .NoMatch = True Then
                .AddNew
                
                ZLinea = 0
                If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Then
                    ZLinea = 6
                        Else
                    If WArticulo = "PT-99999-999" Then
                        ZLinea = 99
                            Else
                        Sql1 = "Select *"
                        Sql2 = " FROM Terminado"
                        Sql3 = " Where Terminado.Codigo = " + "'" + WArticulo + "'"
                        spTerminado = Sql1 + Sql2 + Sql3
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                    End If
                End If
                
                !Clave = WClave
                !Cliente = WCliente
                !Codigo = WArticulo
                
                !Linea = ZLinea
                
                !Impo1 = 0
                !Impo2 = 0
                !Impo3 = 0
                !Impo4 = 0
                !Impo5 = 0
                !Impo6 = 0
                !Impo7 = 0
                !Impo8 = 0
                !Impo9 = 0
                !Impo10 = 0
                !Impo11 = 0
                !Impo12 = 0
                
                !Titulo1 = WTitulo1
                !Titulo2 = WTitulo2
                !Titulo3 = WTitulo3
                
                !Pesos1 = 0
                !Pesos2 = 0
                !Pesos3 = 0
                !Pesos4 = 0
                !Pesos5 = 0
                !Pesos6 = 0
                !Pesos7 = 0
                !Pesos8 = 0
                !Pesos9 = 0
                !Pesos10 = 0
                !Pesos11 = 0
                !Pesos12 = 0
                
                !Canti1 = 0
                !Canti2 = 0
                !Canti3 = 0
                !Canti4 = 0
                !Canti5 = 0
                !Canti6 = 0
                !Canti7 = 0
                !Canti8 = 0
                !Canti9 = 0
                !Canti10 = 0
                !Canti11 = 0
                !Canti12 = 0
                
                .Update
            End If
        End With
    
    Next Ciclo
    
    
    
    zcicla = 0
    MiraII.Text = 0
    
    With rstEstaComando
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                zcicla = zcicla + 1
                mira.Text = zcicla
                DoEvents
                
                WTerminado = !Codigo
                WCliente = !Cliente
                ZLinea = !Linea
                ZTipo = "0"
                
                ZVenta1 = !Impo1
                ZVenta2 = !Impo2
                ZVenta3 = !Impo3
                ZVenta4 = !Impo4
                ZVenta5 = !Impo5
                ZVenta6 = !Impo6
                ZVenta7 = !Impo7
                ZVenta8 = !Impo8
                ZVenta9 = !Impo9
                ZVenta10 = !Impo10
                ZVenta11 = !Impo11
                ZVenta12 = !Impo12
                
                ZCanti1 = !Canti1
                ZCanti2 = !Canti2
                ZCanti3 = !Canti3
                ZCanti4 = !Canti4
                ZCanti5 = !Canti5
                ZCanti6 = !Canti6
                ZCanti7 = !Canti7
                ZCanti8 = !Canti8
                ZCanti9 = !Canti9
                ZCanti10 = !Canti10
                ZCanti11 = !Canti11
                ZCanti12 = !Canti12
                
                Select Case ZLinea
                    Case 3, 4, 5, 7, 8, 9, 11, 12, 14, 19
                        ZTipo = "1"
                    Case 6, 16, 17
                        ZTipo = "2"
                    Case 10, 22, 24, 25, 26, 27, 28, 29, 30
                        ZTipo = "3"
                    Case 20
                        ZTipo = "5"
                    Case 21
                        ZTipo = "6"
                    Case 99
                        ZTipo = "7"
                    Case Else
                        aa = WTerminado
                        WRubro = 0
                        spCliente = "ConsultaCliente " + WCliente
                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCliente.RecordCount > 0 Then
                            WRubro = rstCliente!Rubro
                            rstCliente.Close
                        End If
                        If WCliente = "P00005" Then
                            ZTipo = "4"
                                Else
                            If WRubro = 10 Then
                                ZTipo = "5"
                                    Else
                                ZTipo = "6"
                            End If
                        End If
                End Select
                
                If WCliente = "Z99999" Then
                
                    ZStock1 = 0
                    ZCosto1 = 0
                    ZStock2 = 0
                    ZCosto2 = 0
                    ZStock3 = 0
                    ZCosto3 = 0
                    ZStock4 = 0
                    ZCosto4 = 0
                    ZStock5 = 0
                    ZCosto5 = 0
                    ZStock6 = 0
                    ZCosto6 = 0
                    ZStock7 = 0
                    ZCosto7 = 0
                    ZStock8 = 0
                    ZCosto8 = 0
                    ZStock9 = 0
                    ZCosto9 = 0
                    ZStock10 = 0
                    ZCosto10 = 0
                    ZStock11 = 0
                    ZCosto11 = 0
                    ZStock12 = 0
                    ZCosto12 = 0
                
                    For Cicla = 1 To 12
                    
                        ZMes = ImpreFecha(Cicla, 1)
                        ZAno = ImpreFecha(Cicla, 2)
                        Call Ceros(ZMes, 2)
                        Call Ceros(ZAno, 4)
                        ZFecha = ZAno + ZMes
                        ZClave = WTerminado + ZFecha
                    
                        Sql1 = "Select *"
                        Sql2 = " FROM StockHistorico"
                        Sql3 = " Where Clave = " + "'" + ZClave + "'"
                        spStockHistorico = Sql1 + Sql2 + Sql3
                        Set rstStockHistorico = db.OpenRecordset(spStockHistorico, dbOpenSnapshot, dbSQLPassThrough)
                        If rstStockHistorico.RecordCount > 0 Then
                        
                            ZPlanta1 = rstStockHistorico!Planta1
                            ZPlanta2 = rstStockHistorico!Planta2
                            ZPlanta3 = rstStockHistorico!Planta3
                            ZPlanta4 = rstStockHistorico!Planta4
                            ZPlanta5 = rstStockHistorico!Planta5
                            ZCosto = rstStockHistorico!Costo
                            rstStockHistorico.Close
                        
                            Select Case Cicla
                                Case 1
                                    ZStock1 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto1 = ZCosto * ZCanti1
                                Case 2
                                    ZStock2 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto2 = ZCosto * ZCanti2
                                Case 3
                                    ZStock3 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto3 = ZCosto * ZCanti3
                                Case 4
                                    ZStock4 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto4 = ZCosto * ZCanti4
                                Case 5
                                    ZStock5 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto5 = ZCosto * ZCanti5
                                Case 6
                                    ZStock6 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto6 = ZCosto * ZCanti6
                                Case 7
                                    ZStock7 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto7 = ZCosto * ZCanti7
                                Case 8
                                    ZStock8 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto8 = ZCosto * ZCanti8
                                Case 9
                                    ZStock9 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto9 = ZCosto * ZCanti9
                                Case 10
                                    ZStock10 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto10 = ZCosto * ZCanti10
                                Case 11
                                    ZStock11 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto11 = ZCosto * ZCanti11
                                Case 12
                                    ZStock12 = (ZPlanta1 + ZPlanta2 + ZPlanta3 + ZPlanta4 + ZPlanta5) * ZCosto
                                    ZCosto12 = ZCosto * ZCanti12
                                Case Else
                            End Select
                        End If
    
                    Next Cicla
                
                            Else
                            
                    ZCosto1 = 0
                    ZCosto2 = 0
                    ZCosto3 = 0
                    ZCosto4 = 0
                    ZCosto5 = 0
                    ZCosto6 = 0
                    ZCosto7 = 0
                    ZCosto8 = 0
                    ZCosto9 = 0
                    ZCosto10 = 0
                    ZCosto11 = 0
                    ZCosto12 = 0
                    
                    ZStock1 = 0
                    ZStock2 = 0
                    ZStock3 = 0
                    ZStock4 = 0
                    ZStock5 = 0
                    ZStock6 = 0
                    ZStock7 = 0
                    ZStock8 = 0
                    ZStock9 = 0
                    ZStock10 = 0
                    ZStock11 = 0
                    ZStock12 = 0
                    
                End If
                
                Sql1 = "UPDATE Comando SET "
                Sql2 = "Venta1 = Venta1 + " + Str$(ZVenta1) + ","
                Sql3 = "Venta2 = Venta2 + " + Str$(ZVenta2) + ","
                Sql4 = "Venta3 = Venta3 + " + Str$(ZVenta3) + ","
                Sql5 = "Venta4 = Venta4 + " + Str$(ZVenta4) + ","
                Sql6 = "Venta5 = Venta5 + " + Str$(ZVenta5) + ","
                Sql7 = "Venta6 = Venta6 + " + Str$(ZVenta6) + ","
                Sql8 = "Venta7 = Venta7 + " + Str$(ZVenta7) + ","
                Sql9 = "Venta8 = Venta8 + " + Str$(ZVenta8) + ","
                Sql10 = "Venta9 = Venta9 + " + Str$(ZVenta9) + ","
                Sql11 = "Venta10 = Venta10 + " + Str$(ZVenta10) + ","
                Sql12 = "Venta11 = Venta11 + " + Str$(ZVenta11) + ","
                Sql13 = "Venta12 = Venta12 + " + Str$(ZVenta12) + ","
                Sql14 = "Kilos1 = Kilos1 + " + Str$(ZCanti1) + ","
                Sql15 = "Kilos2 = Kilos2 + " + Str$(ZCanti2) + ","
                Sql16 = "Kilos3 = Kilos3 + " + Str$(ZCanti3) + ","
                Sql17 = "Kilos4 = Kilos4 + " + Str$(ZCanti4) + ","
                Sql18 = "Kilos5 = Kilos5 + " + Str$(ZCanti5) + ","
                Sql19 = "Kilos6 = Kilos6 + " + Str$(ZCanti6) + ","
                Sql20 = "Kilos7 = Kilos7 + " + Str$(ZCanti7) + ","
                Sql21 = "Kilos8 = Kilos8 + " + Str$(ZCanti8) + ","
                Sql22 = "Kilos9 = Kilos9 + " + Str$(ZCanti9) + ","
                Sql23 = "Kilos10 = Kilos10 + " + Str$(ZCanti10) + ","
                Sql24 = "Kilos11 = Kilos11 + " + Str$(ZCanti11) + ","
                Sql25 = "Kilos12 = Kilos12 + " + Str$(ZCanti12) + ","
                Sql26 = "Costo1 = Costo1 + " + Str$(ZCosto1) + ","
                Sql27 = "Costo2 = Costo2 + " + Str$(ZCosto2) + ","
                Sql28 = "Costo3 = Costo3 + " + Str$(ZCosto3) + ","
                Sql29 = "Costo4 = Costo4 + " + Str$(ZCosto4) + ","
                Sql30 = "Costo5 = Costo5 + " + Str$(ZCosto5) + ","
                Sql31 = "Costo6 = Costo6 + " + Str$(ZCosto6) + ","
                Sql32 = "Costo7 = Costo7 + " + Str$(ZCosto7) + ","
                Sql33 = "Costo8 = Costo8 + " + Str$(ZCosto8) + ","
                Sql34 = "Costo9 = Costo9 + " + Str$(ZCosto9) + ","
                Sql35 = "Costo10 = Costo10 + " + Str$(ZCosto10) + ","
                Sql36 = "Costo11 = Costo11 + " + Str$(ZCosto11) + ","
                Sql37 = "Costo12 = Costo12 + " + Str$(ZCosto12) + ","
                Sql38 = "Stock1 = Stock1 + " + Str$(ZStock1) + ","
                Sql39 = "Stock2 = Stock2 + " + Str$(ZStock2) + ","
                Sql40 = "Stock3 = Stock3 + " + Str$(ZStock3) + ","
                Sql41 = "Stock4 = Stock4 + " + Str$(ZStock4) + ","
                Sql42 = "Stock5 = Stock5 + " + Str$(ZStock5) + ","
                Sql43 = "Stock6 = Stock6 + " + Str$(ZStock6) + ","
                Sql44 = "Stock7 = Stock7 + " + Str$(ZStock7) + ","
                Sql45 = "Stock8 = Stock8 + " + Str$(ZStock8) + ","
                Sql46 = "Stock9 = Stock9 + " + Str$(ZStock9) + ","
                Sql47 = "Stock10 = Stock10 + " + Str$(ZStock10) + ","
                Sql48 = "Stock11 = Stock11 + " + Str$(ZStock11) + ","
                Sql49 = "Stock12 = Stock12 + " + Str$(ZStock12)
                Sql50 = " Where Tipo = " + "'" + ZTipo + "'"
                
                spComando = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                            Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                            Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                            Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                            Sql41 + Sql42 + Sql43 + Sql44 + Sql45 + Sql46 + Sql47 + Sql48 + Sql49 + Sql50
                Set rstComando = db.OpenRecordset(spComando, dbOpenSnapshot, dbSQLPassThrough)

                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem
    Rem procesa los atrasos de pedidos
    Rem
    
    Sql1 = "UPDATE Pedido SET "
    Sql2 = " Suma1 = 0,"
    Sql3 = " Suma2 = 0,"
    Sql4 = " Dias = 0"
    spPedido = Sql1 + Sql2 + Sql3 + Sql4
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Erase WVectorPedido
    LugarPedido = 0
    
    Sql1 = "Select Pedido, Fecha, FechaOrd, Terminado, Clave, FecEntrega, OrdFecEntrega, Cliente, FechaInicial, OrdFechaInicial, TipoPed, FechaActualizacion, OrdFechaActualizacion"
    Sql2 = " FROM Pedido"
    Sql3 = " Where Pedido.FechaOrd >= " + "'" + WDesde + "'"
    Sql4 = " and Pedido.FechaOrd <= " + "'" + WHasta + "'"
    Sql5 = " and Pedido.TipoPed <> 5"
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Rem If rstPedido!Pedido = 329005 Then Stop
                    Rem If rstPedido!Pedido = 329057 Then Stop
                    Rem If rstPedido!Pedido = 329272 Then Stop
                    Rem If rstPedido!Pedido = 328596 Then Stop
                    Rem If rstPedido!Pedido = 328630 Then Stop
                    Rem If rstPedido!Pedido = 328632 Then Stop
                
                    LugarPedido = LugarPedido + 1
                    
                    mira.Text = LugarPedido
                    DoEvents
                    
                    
                    WVectorPedido(LugarPedido, 1) = rstPedido!Pedido
                    WVectorPedido(LugarPedido, 2) = rstPedido!Fecha
                    WVectorPedido(LugarPedido, 4) = rstPedido!Terminado
                    WVectorPedido(LugarPedido, 5) = rstPedido!Clave
                    WVectorPedido(LugarPedido, 6) = rstPedido!FechaOrd
                    
                    XFechaInicial = IIf(IsNull(rstPedido!FechaInicial), "", rstPedido!FechaInicial)
                    XOrdFechaInicial = IIf(IsNull(rstPedido!OrdFechaInicial), "", rstPedido!OrdFechaInicial)
                    If XFechaInicial <> "" Then
                        WVectorPedido(LugarPedido, 3) = rstPedido!FechaInicial
                        WVectorPedido(LugarPedido, 7) = rstPedido!OrdFechaInicial
                        
                            Else
                        WVectorPedido(LugarPedido, 3) = rstPedido!FecEntrega
                        WVectorPedido(LugarPedido, 7) = rstPedido!OrdFecEntrega
                    End If
                    
                    WVectorPedido(LugarPedido, 8) = rstPedido!Cliente
                    
                    WVectorPedido(Renglon, 9) = rstPedido!Tipoped
                    XFechaActualizacion = IIf(IsNull(rstPedido!FechaActualizacion), "", rstPedido!FechaActualizacion)
                    XOrdFechaActualizacion = IIf(IsNull(rstPedido!OrdFechaActualizacion), "", rstPedido!OrdFechaActualizacion)
                    WVectorPedido(Renglon, 10) = XFechaActualizacion
                    WVectorPedido(Renglon, 11) = XOrdFechaActualizacion
                    
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    MiraII.Text = LugarPedido
    
    For Ciclo = 1 To LugarPedido
    
    
        mira.Text = Ciclo
        DoEvents
    
        WPedido = WVectorPedido(Ciclo, 1)
        WFecha = WVectorPedido(Ciclo, 2)
        WFechaEntrega = WVectorPedido(Ciclo, 3)
        WTerminado = WVectorPedido(Ciclo, 4)
        WClave = WVectorPedido(Ciclo, 5)
        WFechaord = WVectorPedido(Ciclo, 6)
        WOrdFechaEntrega = WVectorPedido(Ciclo, 7)
        WCliente = WVectorPedido(Ciclo, 8)
        WTipoPedido = WVectorPedido(Ciclo, 9)
        WFechaActualizacion = WVectorPedido(Ciclo, 10)
        WOrdFechaActualizacion = WVectorPedido(Ciclo, 11)
        
        Rem If WPedido = 329005 Then Stop
        Rem If WPedido = 329057 Then Stop
        Rem If WPedido = 329272 Then Stop
        Rem If WPedido = 328596 Then Stop
        Rem If WPedido = 328630 Then Stop
        Rem If WPedido = 328632 Then Stop
        
        If Val(WTipoPedido) = 4 Then
        
            If WFechaActualizacion <> "" Then
                Entra = "S"
                FechaFactu = WFechaActualizacion
                FechaFactuOrd = WOrdFechaActualizacion
                    Else
                Entra = "N"
                FechaFactu = ""
                FechaFactuOrd = ""
            End If
        
                Else
        
            Entra = "N"
            FechaFactu = ""
            FechaFactuOrd = ""
        
            Sql1 = "Select Fecha, OrdFecha, Articulo, Pedido "
            Sql2 = " FROM Estadistica"
            Sql3 = " Where Estadistica.Pedido = " + "'" + WPedido + "'"
            Sql4 = " and Estadistica.Articulo = " + "'" + WTerminado + "'"
            Sql5 = " order by OrdFecha"
            spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Entra = "S"
                            FechaFactu = rstEstadistica!Fecha
                            FechaFactuOrd = rstEstadistica!OrdFecha
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEstadistica.Close
            End If
        
        End If
        
        If Entra = "S" Then
        
            If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DW" Then
                WLinea = "6"
                    Else
                WLinea = ""
                Sql1 = "Select *"
                Sql2 = " FROM Terminado"
                Sql3 = " Where Terminado.Codigo = " + "'" + WTerminado + "'"
                spTerminado = Sql1 + Sql2 + Sql3
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WLinea = Str$(rstTerminado!Linea)
                    rstTerminado.Close
                End If
            End If
        
            WSuma1 = 1
            WSuma2 = 0
            WFechaHastaOrd = FechaFactuOrd
            WFechaDesdeOrd = WOrdFechaEntrega
            WFechaHasta = FechaFactu
            WFechaDesde = WFechaEntrega
            
            If WFechaHastaOrd > WFechaDesdeOrd Then
                WSuma2 = 1
                AAAA = WTerminado
            End If
            
            Select Case Val(WLinea)
                Case 3, 4, 5, 7, 8, 9, 11, 12, 14, 19
                    WSumaLinea = "1"
                Case 6, 16, 17
                    WSumaLinea = "2"
                Case 10, 22, 24, 25, 26, 27, 28, 29, 30
                    WSumaLinea = "3"
                Case 20
                    WSumaLinea = "5"
                Case 21
                    WSumaLinea = "6"
                Case Else
                    WRubro = 0
                    spCliente = "ConsultaCliente " + WCliente
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WRubro = rstCliente!Rubro
                        rstCliente.Close
                    End If
                    If WCliente = "P00005" Then
                        WSumaLinea = "4"
                            Else
                        If WRubro = 10 Then
                            WSumaLinea = "5"
                                Else
                            WSumaLinea = "6"
                        End If
                    End If
            End Select
            
            Pedido1 = 0
            Pedido2 = 0
            Pedido3 = 0
            Pedido4 = 0
            Pedido5 = 0
            Pedido6 = 0
            Pedido7 = 0
            Pedido8 = 0
            Pedido9 = 0
            Pedido10 = 0
            Pedido11 = 0
            Pedido12 = 0
            
            Atraso1 = 0
            Atraso2 = 0
            Atraso3 = 0
            Atraso4 = 0
            Atraso5 = 0
            Atraso6 = 0
            Atraso7 = 0
            Atraso8 = 0
            Atraso9 = 0
            Atraso10 = 0
            Atraso11 = 0
            Atraso12 = 0
            
            MesCompara = Val(Mid$(WFecha, 4, 2))
            AnoCompara = Val(Right$(WFecha, 4))
            
            If MesCompara = Val(ImpreFecha(1, 1)) And AnoCompara = Val(ImpreFecha(1, 2)) Then
                Pedido1 = WSuma1
                Atraso1 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(2, 1)) And AnoCompara = Val(ImpreFecha(2, 2)) Then
                Pedido2 = WSuma1
                Atraso2 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(3, 1)) And AnoCompara = Val(ImpreFecha(3, 2)) Then
                Pedido3 = WSuma1
                Atraso3 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(4, 1)) And AnoCompara = Val(ImpreFecha(4, 2)) Then
                Pedido4 = WSuma1
                Atraso4 = WSuma2
            End If
                
            If MesCompara = Val(ImpreFecha(5, 1)) And AnoCompara = Val(ImpreFecha(5, 2)) Then
                Pedido5 = WSuma1
                Atraso5 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(6, 1)) And AnoCompara = Val(ImpreFecha(6, 2)) Then
                Pedido6 = WSuma1
                Atraso6 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(7, 1)) And AnoCompara = Val(ImpreFecha(7, 2)) Then
                Pedido7 = WSuma1
                Atraso7 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(8, 1)) And AnoCompara = Val(ImpreFecha(8, 2)) Then
                Pedido8 = WSuma1
                Atraso8 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(9, 1)) And AnoCompara = Val(ImpreFecha(9, 2)) Then
                Pedido9 = WSuma1
                Atraso9 = WSuma2
            End If
            
            If MesCompara = Val(ImpreFecha(10, 1)) And AnoCompara = Val(ImpreFecha(10, 2)) Then
                Pedido10 = WSuma1
                Atraso10 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(11, 1)) And AnoCompara = Val(ImpreFecha(11, 2)) Then
                Pedido11 = WSuma1
                Atraso11 = WSuma2
            End If
                    
            If MesCompara = Val(ImpreFecha(12, 1)) And AnoCompara = Val(ImpreFecha(12, 2)) Then
                Pedido12 = WSuma1
                Atraso12 = WSuma2
            End If
            
            Sql1 = "UPDATE Comando SET "
            Sql2 = "Pedidos1 = Pedidos1 + " + Str$(Pedido1) + ","
            Sql3 = "Pedidos2 = Pedidos2 + " + Str$(Pedido2) + ","
            Sql4 = "Pedidos3 = Pedidos3 + " + Str$(Pedido3) + ","
            Sql5 = "Pedidos4 = Pedidos4 + " + Str$(Pedido4) + ","
            Sql6 = "Pedidos5 = Pedidos5 + " + Str$(Pedido5) + ","
            Sql7 = "Pedidos6 = Pedidos6 + " + Str$(Pedido6) + ","
            Sql8 = "Pedidos7 = Pedidos7 + " + Str$(Pedido7) + ","
            Sql9 = "Pedidos8 = Pedidos8 + " + Str$(Pedido8) + ","
            Sql10 = "Pedidos9 = Pedidos9 + " + Str$(Pedido9) + ","
            Sql11 = "Pedidos10 = Pedidos10 + " + Str$(Pedido10) + ","
            Sql12 = "Pedidos11 = Pedidos11 + " + Str$(Pedido11) + ","
            Sql13 = "Pedidos12 = Pedidos12 + " + Str$(Pedido12) + ","
            Sql14 = "Atraso1 = Atraso1 + " + Str$(Atraso1) + ","
            Sql15 = "Atraso2 = Atraso2 + " + Str$(Atraso2) + ","
            Sql16 = "Atraso3 = Atraso3 + " + Str$(Atraso3) + ","
            Sql17 = "Atraso4 = Atraso4 + " + Str$(Atraso4) + ","
            Sql18 = "Atraso5 = Atraso5 + " + Str$(Atraso5) + ","
            Sql19 = "Atraso6 = Atraso6 + " + Str$(Atraso6) + ","
            Sql20 = "Atraso7 = Atraso7 + " + Str$(Atraso7) + ","
            Sql21 = "Atraso8 = Atraso8 + " + Str$(Atraso8) + ","
            Sql22 = "Atraso9 = Atraso9 + " + Str$(Atraso9) + ","
            Sql23 = "Atraso10 = Atraso10 + " + Str$(Atraso10) + ","
            Sql24 = "Atraso11 = Atraso11 + " + Str$(Atraso11) + ","
            Sql25 = "Atraso12 = Atraso12 + " + Str$(Atraso12) + ","
            Sql26 = "Impre1 = " + "'" + ZDescri1 + "',"
            Sql27 = "Impre2 = " + "'" + ZDescri2 + "',"
            Sql28 = "Impre3 = " + "'" + ZDescri3 + "',"
            Sql29 = "Impre4 = " + "'" + ZDescri4 + "',"
            Sql30 = "Impre5 = " + "'" + ZDescri5 + "',"
            Sql31 = "Impre6 = " + "'" + ZDescri6 + "',"
            Sql32 = "Impre7 = " + "'" + ZDescri7 + "',"
            Sql33 = "Impre8 = " + "'" + ZDescri8 + "',"
            Sql34 = "Impre9 = " + "'" + ZDescri9 + "',"
            Sql35 = "Impre10 = " + "'" + ZDescri10 + "',"
            Sql36 = "Impre11 = " + "'" + ZDescri11 + "',"
            Sql37 = "Impre12 = " + "'" + ZDescri12 + "'"
            Sql38 = " Where Tipo = " + "'" + WSumaLinea + "'"
                
            spComando = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                        Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                        Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                        Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38
                            
            Set rstComando = db.OpenRecordset(spComando, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    For Ciclo = 1 To 7
    
        Sql1 = "Select *"
        Sql2 = " FROM Comando"
        Sql3 = " Where Tipo = " + "'" + Str$(Ciclo) + "'"
        spComando = Sql1 + Sql2 + Sql3
        Set rstComando = db.OpenRecordset(spComando, dbOpenSnapshot, dbSQLPassThrough)
        If rstComando.RecordCount > 0 Then
            
            PVenta(1) = rstComando!Venta1
            PVenta(2) = rstComando!Venta2
            PVenta(3) = rstComando!Venta3
            PVenta(4) = rstComando!Venta4
            PVenta(5) = rstComando!Venta5
            PVenta(6) = rstComando!Venta6
            PVenta(7) = rstComando!Venta7
            PVenta(8) = rstComando!Venta8
            PVenta(9) = rstComando!Venta9
            PVenta(10) = rstComando!Venta10
            PVenta(11) = rstComando!Venta11
            PVenta(12) = rstComando!Venta12
            
            PKilos(1) = rstComando!Kilos1
            PKilos(2) = rstComando!Kilos2
            PKilos(3) = rstComando!Kilos3
            PKilos(4) = rstComando!Kilos4
            PKilos(5) = rstComando!Kilos5
            PKilos(6) = rstComando!Kilos6
            PKilos(7) = rstComando!Kilos7
            PKilos(8) = rstComando!Kilos8
            PKilos(9) = rstComando!Kilos9
            PKilos(10) = rstComando!Kilos10
            PKilos(11) = rstComando!Kilos11
            PKilos(12) = rstComando!Kilos12
    
            PCosto(1) = rstComando!Costo1
            PCosto(2) = rstComando!Costo2
            PCosto(3) = rstComando!Costo3
            PCosto(4) = rstComando!Costo4
            PCosto(5) = rstComando!Costo5
            PCosto(6) = rstComando!Costo6
            PCosto(7) = rstComando!Costo7
            PCosto(8) = rstComando!Costo8
            PCosto(9) = rstComando!Costo9
            PCosto(10) = rstComando!Costo10
            PCosto(11) = rstComando!Costo11
            PCosto(12) = rstComando!Costo12
    
            PStock(1) = rstComando!Stock1
            PStock(2) = rstComando!Stock2
            PStock(3) = rstComando!Stock3
            PStock(4) = rstComando!Stock4
            PStock(5) = rstComando!Stock5
            PStock(6) = rstComando!Stock6
            PStock(7) = rstComando!Stock7
            PStock(8) = rstComando!Stock8
            PStock(9) = rstComando!Stock9
            PStock(10) = rstComando!Stock10
            PStock(11) = rstComando!Stock11
            PStock(12) = rstComando!Stock12
    
            PPedidos(1) = rstComando!Pedidos1
            PPedidos(2) = rstComando!Pedidos2
            PPedidos(3) = rstComando!Pedidos3
            PPedidos(4) = rstComando!Pedidos4
            PPedidos(5) = rstComando!Pedidos5
            PPedidos(6) = rstComando!Pedidos6
            PPedidos(7) = rstComando!Pedidos7
            PPedidos(8) = rstComando!Pedidos8
            PPedidos(9) = rstComando!Pedidos9
            PPedidos(10) = rstComando!Pedidos10
            PPedidos(11) = rstComando!Pedidos11
            PPedidos(12) = rstComando!Pedidos12

            PAtraso(1) = rstComando!Atraso1
            PAtraso(2) = rstComando!Atraso2
            PAtraso(3) = rstComando!Atraso3
            PAtraso(4) = rstComando!Atraso4
            PAtraso(5) = rstComando!Atraso5
            PAtraso(6) = rstComando!Atraso6
            PAtraso(7) = rstComando!Atraso7
            PAtraso(8) = rstComando!Atraso8
            PAtraso(9) = rstComando!Atraso9
            PAtraso(10) = rstComando!Atraso10
            PAtraso(11) = rstComando!Atraso11
            PAtraso(12) = rstComando!Atraso12
            
            MesesVenta = 0
            VentaTotal = 0
                
            If ImpreFecha(1, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(1)
            End If
                
            If ImpreFecha(2, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(2)
            End If
                
            If ImpreFecha(3, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(3)
            End If
                
            If ImpreFecha(4, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(4)
            End If
                
            If ImpreFecha(5, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(5)
            End If
                
            If ImpreFecha(6, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(6)
            End If
            
            If ImpreFecha(7, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(7)
            End If
                
            If ImpreFecha(8, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(8)
            End If
                
            If ImpreFecha(9, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(9)
            End If
            
            If ImpreFecha(10, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(10)
            End If
                
            If ImpreFecha(11, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(11)
            End If
                
            If ImpreFecha(12, 1) <> "" Then
                MesesVenta = MesesVenta + 1
                VentaTotal = VentaTotal + PVenta(12)
            End If
                
            If MesesVenta <> 0 Then
                ZPromedio = VentaTotal / MesesVenta
                    Else
                ZPromedio = 0
            End If
            
            Erase PFactor
            Erase PPrecio
            Erase PPorceVenta
            Erase PPorceAtraso
            Erase PRotacion
            
            For da = 1 To 12
            
                If PCosto(da) <> 0 Then
                    PFactor(da) = PVenta(da) / PCosto(da)
                End If
                
                If PKilos(da) <> 0 Then
                    PPrecio(da) = PVenta(da) / PKilos(da)
                End If
                
                If ZPromedio <> 0 And PVenta(da) <> 0 Then
                    Rem PDife = PVenta(da) - ZPromedio
                    PPorceVenta(da) = ((PVenta(da) / ZPromedio) - 1) * 100
                End If
                
                If PCosto(da) <> 0 Then
                    PRotacion(da) = PStock(da) / PCosto(da)
                End If
                
                If PPedidos(da) <> 0 Then
                    PPorceAtraso(da) = PAtraso(da) / (PPedidos(da) / 100)
                End If
            
            Next da
            
            ZDescri1 = ImpreFecha(1, 1) + "/" + ImpreFecha(1, 2)
            ZDescri2 = ImpreFecha(2, 1) + "/" + ImpreFecha(2, 2)
            ZDescri3 = ImpreFecha(3, 1) + "/" + ImpreFecha(3, 2)
            ZDescri4 = ImpreFecha(4, 1) + "/" + ImpreFecha(4, 2)
            ZDescri5 = ImpreFecha(5, 1) + "/" + ImpreFecha(5, 2)
            ZDescri6 = ImpreFecha(6, 1) + "/" + ImpreFecha(6, 2)
            ZDescri7 = ImpreFecha(7, 1) + "/" + ImpreFecha(7, 2)
            ZDescri8 = ImpreFecha(8, 1) + "/" + ImpreFecha(8, 2)
            ZDescri9 = ImpreFecha(9, 1) + "/" + ImpreFecha(9, 2)
            ZDescri10 = ImpreFecha(10, 1) + "/" + ImpreFecha(10, 2)
            ZDescri11 = ImpreFecha(11, 1) + "/" + ImpreFecha(11, 2)
            ZDescri12 = ImpreFecha(12, 1) + "/" + ImpreFecha(12, 2)
            
            rstComando.Close
            
            Sql1 = "UPDATE Comando SET "
            Sql2 = "Factor1 = " + "'" + Str$(PFactor(1)) + "',"
            Sql3 = "Factor2 = " + "'" + Str$(PFactor(2)) + "',"
            Sql4 = "Factor3 = " + "'" + Str$(PFactor(3)) + "',"
            Sql5 = "Factor4 = " + "'" + Str$(PFactor(4)) + "',"
            Sql6 = "Factor5 = " + "'" + Str$(PFactor(5)) + "',"
            Sql7 = "Factor6 = " + "'" + Str$(PFactor(6)) + "',"
            Sql8 = "Factor7 = " + "'" + Str$(PFactor(7)) + "',"
            Sql9 = "Factor8 = " + "'" + Str$(PFactor(8)) + "',"
            Sql10 = "Factor9 = " + "'" + Str$(PFactor(9)) + "',"
            Sql11 = "Factor10 = " + "'" + Str$(PFactor(10)) + "',"
            Sql12 = "Factor11 = " + "'" + Str$(PFactor(11)) + "',"
            Sql13 = "Factor12 = " + "'" + Str$(PFactor(12)) + "',"
            Sql14 = "Precio1 = " + "'" + Str$(PPrecio(1)) + "',"
            Sql15 = "Precio2 = " + "'" + Str$(PPrecio(2)) + "',"
            Sql16 = "Precio3 = " + "'" + Str$(PPrecio(3)) + "',"
            Sql17 = "Precio4 = " + "'" + Str$(PPrecio(4)) + "',"
            Sql18 = "Precio5 = " + "'" + Str$(PPrecio(5)) + "',"
            Sql19 = "Precio6 = " + "'" + Str$(PPrecio(6)) + "',"
            Sql20 = "Precio7 = " + "'" + Str$(PPrecio(7)) + "',"
            Sql21 = "Precio8 = " + "'" + Str$(PPrecio(8)) + "',"
            Sql22 = "Precio9 = " + "'" + Str$(PPrecio(9)) + "',"
            Sql23 = "Precio10 = " + "'" + Str$(PPrecio(10)) + "',"
            Sql24 = "Precio11 = " + "'" + Str$(PPrecio(11)) + "',"
            Sql25 = "Precio12 = " + "'" + Str$(PPrecio(12)) + "',"
            Sql26 = "PorceVenta1 = " + "'" + Str$(PPorceVenta(1)) + "',"
            Sql27 = "PorceVenta2 = " + "'" + Str$(PPorceVenta(2)) + "',"
            Sql28 = "PorceVenta3 = " + "'" + Str$(PPorceVenta(3)) + "',"
            Sql29 = "PorceVenta4 = " + "'" + Str$(PPorceVenta(4)) + "',"
            Sql30 = "PorceVenta5 = " + "'" + Str$(PPorceVenta(5)) + "',"
            Sql31 = "PorceVenta6 = " + "'" + Str$(PPorceVenta(6)) + "',"
            Sql32 = "PorceVenta7 = " + "'" + Str$(PPorceVenta(7)) + "',"
            Sql33 = "PorceVenta8 = " + "'" + Str$(PPorceVenta(8)) + "',"
            Sql34 = "PorceVenta9 = " + "'" + Str$(PPorceVenta(9)) + "',"
            Sql35 = "PorceVenta10 = " + "'" + Str$(PPorceVenta(10)) + "',"
            Sql36 = "PorceVenta11 = " + "'" + Str$(PPorceVenta(11)) + "',"
            Sql37 = "PorceVenta12 = " + "'" + Str$(PPorceVenta(12)) + "',"
            Sql38 = "PorceAtraso1 = " + "'" + Str$(PPorceAtraso(1)) + "',"
            Sql39 = "PorceAtraso2 = " + "'" + Str$(PPorceAtraso(2)) + "',"
            Sql40 = "PorceAtraso3 = " + "'" + Str$(PPorceAtraso(3)) + "',"
            Sql41 = "PorceAtraso4 = " + "'" + Str$(PPorceAtraso(4)) + "',"
            Sql42 = "PorceAtraso5 = " + "'" + Str$(PPorceAtraso(5)) + "',"
            Sql43 = "PorceAtraso6 = " + "'" + Str$(PPorceAtraso(6)) + "',"
            Sql44 = "PorceAtraso7 = " + "'" + Str$(PPorceAtraso(7)) + "',"
            Sql45 = "PorceAtraso8 = " + "'" + Str$(PPorceAtraso(8)) + "',"
            Sql46 = "PorceAtraso9 = " + "'" + Str$(PPorceAtraso(9)) + "',"
            Sql47 = "PorceAtraso10 = " + "'" + Str$(PPorceAtraso(10)) + "',"
            Sql48 = "PorceAtraso11 = " + "'" + Str$(PPorceAtraso(11)) + "',"
            Sql49 = "PorceAtraso12 = " + "'" + Str$(PPorceAtraso(12)) + "',"
            Sql50 = "Rotacion1 = " + "'" + Str$(PRotacion(1)) + "',"
            Sql51 = "Rotacion2 = " + "'" + Str$(PRotacion(2)) + "',"
            Sql52 = "Rotacion3 = " + "'" + Str$(PRotacion(3)) + "',"
            Sql53 = "Rotacion4 = " + "'" + Str$(PRotacion(4)) + "',"
            Sql54 = "Rotacion5 = " + "'" + Str$(PRotacion(5)) + "',"
            Sql55 = "Rotacion6 = " + "'" + Str$(PRotacion(6)) + "',"
            Sql56 = "Rotacion7 = " + "'" + Str$(PRotacion(7)) + "',"
            Sql57 = "Rotacion8 = " + "'" + Str$(PRotacion(8)) + "',"
            Sql58 = "Rotacion9 = " + "'" + Str$(PRotacion(9)) + "',"
            Sql59 = "Rotacion10 = " + "'" + Str$(PRotacion(10)) + "',"
            Sql60 = "Rotacion11 = " + "'" + Str$(PRotacion(11)) + "',"
            Sql61 = "Rotacion12 = " + "'" + Str$(PRotacion(12)) + "',"
            Sql62 = "Promedio = " + "'" + Str$(ZPromedio) + "',"
            Sql63 = "Impre1 = " + "'" + ZDescri1 + "',"
            Sql64 = "Impre2 = " + "'" + ZDescri2 + "',"
            Sql65 = "Impre3 = " + "'" + ZDescri3 + "',"
            Sql66 = "Impre4 = " + "'" + ZDescri4 + "',"
            Sql67 = "Impre5 = " + "'" + ZDescri5 + "',"
            Sql68 = "Impre6 = " + "'" + ZDescri6 + "',"
            Sql69 = "Impre7 = " + "'" + ZDescri7 + "',"
            Sql70 = "Impre8 = " + "'" + ZDescri8 + "',"
            Sql71 = "Impre9 = " + "'" + ZDescri9 + "',"
            Sql72 = "Impre10 = " + "'" + ZDescri10 + "',"
            Sql73 = "Impre11 = " + "'" + ZDescri11 + "',"
            Sql74 = "Impre12 = " + "'" + ZDescri12 + "'"
            Sql75 = " Where Tipo = " + "'" + Str$(Ciclo) + "'"
                
            spComando = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                        Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                        Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                        Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                        Sql41 + Sql42 + Sql43 + Sql44 + Sql45 + Sql46 + Sql47 + Sql48 + Sql49 + Sql50 + _
                        Sql51 + Sql52 + Sql53 + Sql54 + Sql55 + Sql56 + Sql57 + Sql58 + Sql59 + Sql60 + _
                        Sql61 + Sql62 + Sql63 + Sql64 + Sql65 + Sql66 + Sql67 + Sql68 + Sql69 + Sql70 + _
                        Sql71 + Sql72 + Sql73 + Sql74 + Sql75

            Set rstComando = db.OpenRecordset(spComando, dbOpenSnapshot, dbSQLPassThrough)
            
        End If

    Next Ciclo
    
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Tablero de Comando"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    Rem Dos = " and {Estadistica.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Rem Listado.GroupSelectionFormula = Uno + Dos
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT  Comando.Tipo, Comando.Venta1, Comando.Venta2, Comando.Venta3, Comando.Venta4, Comando.Venta5, Comando.Venta6, Comando.Venta7, Comando.Venta8, Comando.Venta9, Comando.Venta10, Comando.Venta11, Comando.Venta12, Comando.Kilos1, Comando.Kilos2, Comando.Kilos3, Comando.Kilos4, Comando.Kilos5, Comando.Kilos6, Comando.Kilos7, Comando.Kilos8, Comando.Kilos9, Comando.Kilos10, Comando.Kilos11, Comando.Kilos12, Comando.Stock1, Comando.Stock2, Comando.Stock3, Comando.Stock4, Comando.Stock5, Comando.Stock6, Comando.Stock7, Comando.Stock8, Comando.Stock9, Comando.Stock10, Comando.Stock11, Comando.Stock12, Comando.Pedidos1, Comando.Pedidos2, Comando.Pedidos3, Comando.Pedidos4, Comando.Pedidos5, Comando.Pedidos6, Comando.Pedidos7, Comando.Pedidos8, Comando.Pedidos9, Comando.Pedidos10, Comando.Pedidos11, Comando.Pedidos12 , " _
                     + "Comando.Atraso1, Comando.Atraso2, Comando.Atraso3, Comando.Atraso4, Comando.Atraso5, Comando.Atraso6, Comando.Atraso7, Comando.Atraso8, Comando.Atraso9, Comando.Atraso10, Comando.Atraso11, Comando.Atraso12, Comando.Impre1, Comando.Impre2, Comando.Impre3, Comando.Impre4, Comando.Impre5, Comando.Impre6, Comando.Impre7, Comando.Impre8, Comando.Impre9, Comando.Impre10, Comando.Impre11, Comando.Impre12, Comando.Descripcion, Comando.Promedio, Comando.Factor1, Comando.Factor2, Comando.Factor3, Comando.Factor4, Comando.Factor5, Comando.Factor6, Comando.Factor7, Comando.Factor8, Comando.Factor9, Comando.Factor10, Comando.Factor11, Comando.Factor12, Comando.Precio1, Comando.Precio2, Comando.Precio3, Comando.Precio4, Comando.Precio5, Comando.Precio6, Comando.Precio7, Comando.Precio8, Comando.Precio9, Comando.Precio10, Comando.Precio11, Comando.Precio12, " _
                     + "Comando.PorceVenta1, Comando.PorceVenta2, Comando.PorceVenta3, Comando.PorceVenta4, Comando.PorceVenta5, Comando.PorceVenta6, Comando.PorceVenta7, Comando.PorceVenta8, Comando.PorceVenta9, Comando.PorceVenta10, Comando.PorceVenta11, Comando.PorceVenta12, Comando.PorceAtraso1, Comando.Porcndo.Rotacion5, Comando.Rotacion6, Comando.Rotacion7, Comando.Rotacion8, Comando.Rotacion9, Comando.Rotacion10, Comando.Rotacion11, Comando.Rotacion12 " _
                     + "From " _
                     + DSQ + ".dbo.Comando Comando  " _
                     + "Where Comando.Tipo >= 0 AND " _
                     + "Comando.Tipo <= 99"
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Rem Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "ListaComando.rpt"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEstaComando
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    PrgListaComando.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub DesdeFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
End Sub

Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            DesdeFec.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Posi = 0
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_EstaComando
    OPEN_FILE_Esta8
End Sub






