VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactux 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturacion de Pedidos"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partida"
      Height          =   1935
      Left            =   5400
      TabIndex        =   56
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox WCanti3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   64
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WCanti2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   63
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox WCanti1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   62
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Wlote3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   61
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox WLote2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox WLote1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   255
         Left            =   1440
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Partida"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox Tipoventa 
      Height          =   315
      Left            =   3240
      TabIndex        =   55
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton ReImpre 
      Caption         =   "ReImpresion"
      Height          =   495
      Left            =   10200
      TabIndex        =   54
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Canti5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   48
      Text            =   " "
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Canti4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   47
      Text            =   " "
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox Canti3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   46
      Text            =   " "
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Canti2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   45
      Text            =   " "
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox Canti1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2160
      TabIndex        =   44
      Text            =   " "
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Envase5 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   43
      Text            =   " "
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Envase4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   42
      Text            =   " "
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Envase3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   41
      Text            =   " "
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Envase2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Text            =   " "
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Envase1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   39
      Text            =   " "
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox Paridad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   34
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
      Height          =   495
      Left            =   9120
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   8760
      TabIndex        =   23
      Top             =   5760
      Width           =   2535
      Begin VB.Label Label16 
         Caption         =   "Interes"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Dto."
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Dto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Interes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Iva2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Iva1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Iva 10.5%"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Iva 21%"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Neto"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   22
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   6120
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Orden 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   18
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Remito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   16
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Cliente 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   6360
      TabIndex        =   9
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Numero 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   7
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   450
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      Height          =   450
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      Height          =   450
      Left            =   10200
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "prgfactux.frx":0000
      Left            =   6480
      List            =   "prgfactux.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "prgfactux.frx":0015
      TabIndex        =   2
      Top             =   1560
      Width           =   11415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.Label Descri5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1200
      TabIndex        =   53
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Descri4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1200
      TabIndex        =   52
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Descri3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1200
      TabIndex        =   51
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Descri2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1200
      TabIndex        =   50
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Descri1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1200
      TabIndex        =   49
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Paridad"
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Pedido"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Orden de compra"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Remito"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Vencimiento"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Factura"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgFactux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' N�mero m�ximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WPlazo1 As Integer
Private WPlazo2 As Integer
Private WDias1 As Integer
Private WDias2 As Integer
Private WFecha As String
Private Wvencimiento As String
Private WVencimiento1 As String
Private WPago1 As Integer
Private WPago2 As Integer
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WImporte As Double
Private WCodIva As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private Precio As String
Private dada As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private WAceptada As String
Private Stk(19, 4) As String
Private Envase(5, 2) As String
Private parcial As String
Private Auxiliar(100, 10) As String
Private BajaLote(3, 2) As String
Private XLote(100, 7) As String
Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovenv As Recordset
Dim spMovenv As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer

Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias1 = rstPago!Dias
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        WDescuento = rstPago!Descuento
        WPago = rstPago!Nombre
        rstPago.Close
    End If
    
    WFecha = Fecha.Text
    Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
    
    spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias2 = rstPago!Dias
        WPlazo2 = rstPago!Plazo
        rstPago.Close
   End If
    
    Call Calcula_vencimiento(WFecha, WDias2, WVencimiento1)

End Sub

Private Sub Borra_Click()

    Rem DBGrid1.Col = 0
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 1
    Rem DBGrid1.Text = ""

    Rem DBGrid1.Col = 2
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 3
    Rem DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = "S"
    
    XLote(WRow, 1) = ""
    XLote(WRow, 2) = ""
    XLote(WRow, 3) = ""
    XLote(WRow, 4) = ""
    XLote(WRow, 5) = ""
    XLote(WRow, 6) = ""
    
End Sub

Private Sub Calcula_Click()

    WNeto = 0
    
    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 3
            Precio = DBGrid1.Text
            
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            If Val(Cantidad) <> 0 Then
                WNeto = WNeto + (Val(Cantidad) * Val(Precio))
            End If
                    
        Next iRow
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 4
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    WImpoInteres = 0

    If Val(Paridad.Text) <> 0 Then
        WNeto = WNeto * Val(Paridad.Text)
    End If
    
    XNeto = WNeto
    
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    If WTasa <> 0 Then
        WImpoInteres = (WNeto * WPlazo1 * WTasa) / 36000
        Call Redondeo(WImpoInteres)
        WNeto = WNeto + WImpoInteres
    End If
    
    WIva1 = 0
    WIva2 = 0
    
    
    Select Case Val(WCodIva)
        Case 2
            WIva1 = WNeto * 0.21
            WIva2 = WNeto * 0.105
            Call Redondeo(WIva1)
            Call Redondeo(WIva2)
        Case 4
            WIva1 = 0
            WIva2 = 0
        Case Else
            WIva1 = WNeto * 0.21
            Call Redondeo(WIva1)
    End Select
    
    If WNeto <> 0 Then
        Call Convierte1_datos(Str$(WNeto), Auxi)
        Neto.Caption = Pusing("###,###.##", Auxi)
            Else
        Neto.Caption = "0.00"
    End If
    
    If WImpoDto <> 0 Then
        Call Convierte1_datos(Str$(WImpoDto), Auxi)
        Dto.Caption = Pusing("###,###.##", Auxi)
            Else
        Dto.Caption = "0.00"
    End If
    
    If WImpoInteres <> 0 Then
        Call Convierte1_datos(Str$(WImpoInteres), Auxi)
        Interes.Caption = Pusing("###,###.##", Auxi)
            Else
        Interes.Caption = "0.00"
    End If
    
    If WIva1 <> 0 Then
        Call Convierte1_datos(Str$(WIva1), Auxi)
        Iva1.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva1.Caption = "0.00"
    End If
    
    If WIva2 <> 0 Then
        Call Convierte1_datos(Str$(WIva2), Auxi)
        Iva2.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva2.Caption = "0.00"
    End If
    
    WTotal = WNeto + WIva1 + WIva2
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    
    PrgFactu.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    
        Call Calcula_Click
        
        Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
        Rem     WImporte = WNeto
        Rem     WNeto = WNeto / 1.21
        Rem     Call Redondeo(WNeto)
        Rem     WIva1 = WImporte - WNeto
        Rem     WIva2 = 0
        Rem End If
        
        WTipo = "01"
        WNumero = Numero.Text
        WRenglon = "01"
        WCliente = Cliente.Text
        WFecha = Fecha.Text
        WEstado = "0"
        Rem Wvencimiento = Wvencimiento
        Rem WVencimiento1 = WVencimiento1
        Call Convierte_datos(Str$(Total), Auxi)
        XTotal = Str$(WTotal)
        XTotalUs = Str$(WTotal * Val(Paridad.Text))
        XSaldo = Str$(WTotal)
        XSaldoUs = Str$(WTotal * Val(Paridad.Text))
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
        WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
        WImpre = "FC"
        XNet = Str$(WNeto)
        XIva1 = Str$(WIva1)
        XIva2 = Str$(WIva2)
        WPedido = Pedido.Text
        WRemito = Remito.Text
        WOrden = Orden.Text
        WParidad = Paridad.Text
        WProvincia = WProv
        XVendedor = Str$(WVendedor)
        XRubro = Str$(WRubro)
        WComprobante = ""
        WAceptada = Str$(Tipoventa.ListIndex)
        Call Ceros(WAceptada, 1)
        WCosto = ""
        WImporte1 = ""
        WImporte2 = ""
        WImporte3 = ""
        WImporte4 = ""
        WImporte5 = ""
        WImporte6 = ""
        WImporte7 = ""
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        WClave = "01" + Auxi + "01"
        XEmpresa = "1"
        WDate = Date$
        
        XParam = "'" + WClave + "','" _
                    + WTipo + "','" + WNumero + "','" _
                    + WRenglon + "','" + WCliente + "','" _
                    + WFecha + "','" + WEstado + "','" _
                    + Wvencimiento + "','" + WVencimiento1 + "','" _
                    + XTotal + "','" + XTotalUs + "','" _
                    + XSaldo + "','" + XSaldoUs + "','" _
                    + WOrdFecha + "','" + WOrdVencimiento + "','" _
                    + WOrdVencimiento1 + "','" + WImpre + "','" _
                    + WEmpresa + "','" _
                    + XNet + "','" + XIva1 + "','" _
                    + XIva2 + "','" + WPedido + "','" _
                    + WRemito + "','" + WOrden + "','" _
                    + WParidad + "','" + WProvincia + "','" _
                    + XVendedor + "','" + XRubro + "','" _
                    + WComprobante + "','" + WAceptada + "','" _
                    + WCosto + "','" _
                    + WImporte1 + "','" _
                    + WImporte2 + "','" _
                    + WImporte3 + "','" _
                    + WImporte4 + "','" _
                    + WImporte5 + "','" _
                    + WImporte6 + "','" _
                    + WImporte7 + "','" _
                    + WDate + "'"
                        
        spCtacte = "AltaCtacte " + XParam
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
        Erase Auxiliar
        Auxi = 0
        
        Suma = 0
        Renglon = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                Suma = Suma + 1
                WRenglon = WRenglon + 1
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text
                WBase = Val(Right$(Articulo, 3))
                If WBase <= 5 Then
                    Articulo = Left$(Articulo, 7) + "100"
                End If
                    
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
                    
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
                        
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                            
                    Auxi1 = Str$(Numero.Text)
                    Call Ceros(Auxi1, 8)
                    WTipo = "01"
                    WNumero = Numero.Text
                    XRenglon = Str$(Renglon)
                    WArticulo = Articulo
                    XXCantidad = Str$(Cantidad)
                    XPrecio = Str$(Precio)
                    XPrecioUs = Str$(Precio * Val(Paridad.Text))
                    XImporte = Str$(Precio * Cantidad)
                    XImporteUs = Str$(Precio * Val(Paridad.Text) * Cantidad)
                    WCliente = Cliente.Text
                    WParidad = Paridad.Text
                    XVendedor = Str$(WVendedor)
                    XRubro = Str$(WRubro)
                    XLinea = Str$(WLinea)
                    XCosto2 = ""
                    XCosto1 = ""
                    WCoeficiente = ""
                    WPedido = Pedido.Text
                    WFecha = Fecha.Text
                    WImporte1 = ""
                    WImporte2 = ""
                    WImporte3 = ""
                    WImporte4 = ""
                    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XArticulo = Left$(Articulo, 8)
                    If Tipoventa.ListIndex = 1 Then
                        WRemito = "C" + Remito.Text
                            Else
                        WRemito = Remito.Text
                    End If
                    WClave = "01" + Auxi1 + Auxi
                    WDate = Date$
                    XCanti = ""
                    XImpo = ""
                    XImpoUs = ""
                    XMarca = "X"
                    WLote1 = ""
                    WLote2 = ""
                    Wlote3 = ""
                    WCanti1 = ""
                    WCanti2 = ""
                    WCanti3 = ""
                        
                    XParam = "'" + WClave + "','" _
                                + WTipo + "','" + WNumero + "','" _
                                + XRenglon + "','" + WArticulo + "','" _
                                + XXCantidad + "','" + XPrecio + "','" _
                                + XPrecioUs + "','" + XImporte + "','" _
                                + XImporteUs + "','" + WCliente + "','" _
                                + WParidad + "','" + XVendedor + "','" _
                                + XRubro + "','" + XLinea + "','" _
                                + XCosto1 + "','" + XCosto2 + "','" _
                                + WCoeficiente + "','" + WPedido + "','" _
                                + WFecha + "','" + WImporte1 + "','" _
                                + WImporte2 + "','" + WImporte3 + "','" _
                                + WImporte4 + "','" + WOrdFecha + "','" _
                                + XArticulo + "','" + WRemito + "','" _
                                + WDate + "','" + XCanti + "','" _
                                + XImpo + "','" _
                                + XImpoUs + "','" _
                                + XMarca + "','" _
                                + WLote1 + "','" _
                                + WCanti1 + "','" _
                                + WLote2 + "','" _
                                + WCanti2 + "','" _
                                + Wlote3 + "','" _
                                + WCanti3 + "'"
                    
                    spEstadistica = "AltaEstadistica " + XParam
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Auxiliar(Renglon, 1) = Articulo
                    Auxiliar(Renglon, 2) = Cantidad
                    Auxiliar(Renglon, 3) = Precio
                    Auxiliar(Renglon, 4) = WRenglon
                    Auxiliar(Renglon, 5) = WLote1
                    Auxiliar(Renglon, 6) = WCanti1
                    Auxiliar(Renglon, 7) = WLote2
                    Auxiliar(Renglon, 8) = WCanti2
                    Auxiliar(Renglon, 9) = Wlote3
                    Auxiliar(Renglon, 10) = WCanti3
                        
                End If
                                        
            Next iRow
            
        Next a
        
        For da = 1 To Renglon
        
            Articulo = Auxiliar(da, 1)
            Cantidad = Auxiliar(da, 2)
            Precio = Auxiliar(da, 3)
            WRenglon = Auxiliar(da, 4)
            WLote1 = Auxiliar(da, 5)
            WCanti1 = Auxiliar(da, 6)
            WLote2 = Auxiliar(da, 7)
            WCanti2 = Auxiliar(da, 8)
            Wlote3 = Auxiliar(da, 9)
            WCanti3 = Auxiliar(da, 10)
            
            Auxi = Pedido.Text
            Call Ceros(Auxi, 6)
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            ClavePedido = Auxi + Auxi1
            
            XParam = "'" + Pedido.Text + "','" _
                        + Articulo + "'"
                                
            spPedido = "ConsultaPedidoFactura " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                WFacturado = Str$(rstPedido!Facturado + Cantidad)
                If Val(WFacturado) > rstPedido!Cantidad Then
                    WFacturado = Str$(rstPedido!Cantidad)
                End If
                ClavePedido = rstPedido!Clave
                rstPedido.Close
                XParam = "'" + ClavePedido + "','" _
                            + WFacturado + "'"
                                           
                spPedido = "ModificaPedidoFacturas " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            End If
                
            ClavePrecio = Cliente.Text + Articulo
            
            spPrecios = "ConsultaPrecios " + "'" + ClavePrecio + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
            
                WFecha1 = ""
                WFactura1 = ""
                WPrecio1 = ""
                WCantidad1 = ""
                
                WFecha2 = ""
                WFactura2 = ""
                WPrecio2 = ""
                WCantidad2 = ""
                
                WFecha3 = ""
                WFactura3 = ""
                WPrecio3 = ""
                WCantidad3 = ""
                
                WFecha4 = ""
                WFactura4 = ""
                WPrecio4 = ""
                WCantidad4 = ""
                
                WFecha5 = ""
                WFactura5 = ""
                WPrecio5 = ""
                WCantidad5 = ""
                
                If rstPrecios!Cantidad2 <> O Then
                    WFecha1 = rstPrecios!Fecha2
                    WFactura1 = rstPrecios!Factura2
                    WPrecio1 = Str$(rstPrecios!Precio2)
                    WCantidad1 = Str$(rstPrecios!Cantidad2)
                End If
                                
                If rstPrecios!Cantidad2 <> O Then
                    WFecha2 = rstPrecios!Fecha3
                    WFactura2 = rstPrecios!Factura3
                    WPrecio2 = Str$(rstPrecios!Precio3)
                    WCantidad2 = Str$(rstPrecios!Cantidad3)
                End If
                                
                If rstPrecios!Cantidad2 <> O Then
                    WFecha3 = rstPrecios!Fecha4
                    WFactura3 = rstPrecios!Factura4
                    WPrecio3 = Str$(rstPrecios!Precio4)
                    WCantidad3 = Str$(rstPrecios!Cantidad4)
                End If
                                
                If rstPrecios!Cantidad2 <> O Then
                    WFecha4 = rstPrecios!Fecha5
                    WFactura4 = rstPrecios!Factura5
                    WPrecio4 = Str$(rstPrecios!Precio5)
                    WCantidad4 = Str$(rstPrecios!Cantidad5)
                End If
                                
                WFecha5 = Fecha.Text
                WFactura5 = Numero.Text
                WPrecio5 = Str$(Precio)
                WCantidad5 = Str$(Cantidad)
                                
                WDate = Date$
                
                rstPrecios.Close
                
                XParam = "'" + ClavePrecio + "','" _
                            + WFecha1 + "','" _
                            + WFactura1 + "','" _
                            + WPrecio1 + "','" _
                            + WCantidad1 + "','" _
                            + WFecha2 + "','" _
                            + WFactura2 + "','" _
                            + WPrecio2 + "','" _
                            + WCantidad2 + "','" _
                            + WFecha3 + "','" _
                            + WFactura3 + "','" _
                            + WPrecio3 + "','" _
                            + WCantidad3 + "','" _
                            + WFecha4 + "','" _
                            + WFactura4 + "','" _
                            + WPrecio4 + "','" _
                            + WCantidad4 + "','" _
                            + WFecha5 + "','" _
                            + WFactura5 + "','" _
                            + WPrecio5 + "','" _
                            + WCantidad5 + "','" _
                            + WDate + "'"
                                           
                spPrecios = "ModificaPreciosFactura " + XParam
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            End If
        Next da
        
        spNumero = "ConsultaNumero " + "'" + "01" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            WCodigo = "01"
            WNumero = Numero.Text
            rstNumero.Close
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
        Rem Listado.GroupSelectionFormula = "{Pedido.Pedido} in " + Pedido.Text + " to " + Pedido.Text
        Rem Listado.Destination = 1
        Rem Listado.Action = 1
        
        Call Calcula_Saldo
        
        Erase Envase
        Envase(1, 1) = Envase1.Text
        Envase(2, 1) = Envase2.Text
        Envase(3, 1) = Envase3.Text
        Envase(4, 1) = Envase4.Text
        Envase(5, 1) = Envase5.Text
        
        Envase(1, 2) = Canti1.Text
        Envase(2, 2) = Canti2.Text
        Envase(3, 2) = Canti3.Text
        Envase(4, 2) = Canti4.Text
        Envase(5, 2) = Canti5.Text
        
        For XDa = 1 To 5
            For da = 1 To 9
                If Val(Envase(XDa, 1)) = Val(Stk(da, 1)) Then
                    Stk(da, 3) = Canti1.Text
                End If
            Next da
        Next XDa
        
        For da = 1 To 9
            Stk(da, 4) = Str$(Val(Stk(da, 2)) + Val(Stk(da, 3)))
        Next da
        
        Renglon = 0
        
        For da = 1 To 5
        
            If Val(Envase(da, 2)) <> 0 Then
            
                Renglon = Renglon + 1
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Val(Remito.Text))
                Call Ceros(Auxi1, 6)
                    
                WTipo = "1"
                WCodigo = Remito.Text
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WEnvase = Envase(da, 1)
                WCantidad = Envase(da, 2)
                WMovimiento = "S"
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WClave = Auxi1 + Auxi
                
                XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WCodigo + "','" _
                        + WRenglon + "','" _
                        + WFecha + "','" _
                        + WFechaord + "','" _
                        + WCliente + "','" _
                        + WEnvase + "','" _
                        + WMovimiento + "','" _
                        + WCantidad + "'"
                    
                spMovenv = "AltaMovenv " + XParam
                Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next da
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
    Exit Sub

WError:
     Resume Next
        
End Sub


Private Sub Limpia_Click()

    CargaLote.Visible = False
    Erase XLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    Wlote3.Text = ""

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    Envase4.Text = ""
    Envase5.Text = ""
    
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    Canti4.Text = ""
    Canti5.Text = ""
    
    Numero.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4
                Select Case KeyCode
                    Case 13
                        DBGrid1.Col = 4
                        DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = 0
                        XTerminado = DBGrid1.Text
                        DBGrid1.Col = 4
                        XCantidad = Val(DBGrid1.Text)
                        WRow = DBGrid1.Row
                        
                        If DBGrid1.Row < 40 Then
                           DBGrid1.Row = DBGrid1.Row + 1
                           WRow = DBGrid1.Row
                           DBGrid1.Col = 4
                           KeyCode = 0
                        End If
                        Call Calcula_Click
                        DBGrid1.Row = WRow
                        
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(WRow, 1) = WLote1.Text
                    XLote(WRow, 2) = WCanti1.Text
                    XLote(WRow, 3) = WLote2.Text
                    XLote(WRow, 4) = WCanti2.Text
                    XLote(WRow, 5) = Wlote3.Text
                    XLote(WRow, 6) = WCanti3.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 4
                       KeyCode = 0
                    End If
                    Call Calcula_Click
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 4
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Emision de facturas")
            End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WLote2.SetFocus
                Else
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo1
            G% = MsgBox(m$, 0, "Emiison de facturas")
            WLote1.SetFocus
        End If
        Rem WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
        Rem WLote2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(WRow, 1) = WLote1.Text
                    XLote(WRow, 2) = WCanti1.Text
                    XLote(WRow, 3) = WLote2.Text
                    XLote(WRow, 4) = WCanti2.Text
                    XLote(WRow, 5) = Wlote3.Text
                    XLote(WRow, 6) = WCanti3.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 4
                       KeyCode = 0
                    End If
                    Call Calcula_Click
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 4
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Emision de Facturas")
            End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            Wlote3.SetFocus
                Else
            XSaldo2 = WSaldo2
            XSaldo2 = Pusing("###,###.##", XSaldo2)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo2
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote2.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + Wlote3.Text + "','" _
                        + XTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + XTerminado + "','" _
                            + Wlote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(Wlote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    XLote(WRow, 1) = WLote1.Text
                    XLote(WRow, 2) = WCanti1.Text
                    XLote(WRow, 3) = WLote2.Text
                    XLote(WRow, 4) = WCanti2.Text
                    XLote(WRow, 5) = Wlote3.Text
                    XLote(WRow, 6) = WCanti3.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 4
                       KeyCode = 0
                    End If
                    Call Calcula_Click
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 4
                    KeyCode = 0
                    DBGrid1.SetFocus
                    Exit Sub
                        Else
                    Wlote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
            Call Verifica_Lote
            If WEstado = "S" Then
                XLote(WRow, 1) = WLote1.Text
                XLote(WRow, 2) = WCanti1.Text
                XLote(WRow, 3) = WLote2.Text
                XLote(WRow, 4) = WCanti2.Text
                XLote(WRow, 5) = Wlote3.Text
                XLote(WRow, 6) = WCanti3.Text
                CargaLote.Visible = False
                DBGrid1.Col = 5
                DBGrid1.Text = "S"
                If DBGrid1.Row < 40 Then
                    DBGrid1.Row = DBGrid1.Row + 1
                    WRow = DBGrid1.Row
                    XRow = DBGrid1.Row
                    DBGrid1.Col = 4
                    KeyCode = 0
                End If
                Call Calcula_Click
                DBGrid1.Row = XRow
                DBGrid1.Col = 4
                KeyCode = 0
                Exit Sub
            End If
                Else
            XSaldo3 = WSaldo3
            XSaldo3 = Pusing("###,###.##", XSaldo3)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo3
            G% = MsgBox(m$, 0, "Emision de facturas")
            Wlote3.SetFocus
        End If
        
        Rem WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
        Rem Call Verifica_Lote
        Rem If WEstado = "S" Then
        Rem     XLote(WRow, 1) = WLote1.Text
        Rem     XLote(WRow, 2) = WCanti1.Text
        Rem     XLote(WRow, 3) = WLote2.Text
        Rem     XLote(WRow, 4) = WCanti2.Text
        Rem     XLote(WRow, 5) = Wlote3.Text
        Rem     XLote(WRow, 6) = WCanti3.Text
        Rem     CargaLote.Visible = False
        Rem     DBGrid1.Col = 5
        Rem     DBGrid1.Text = "S"
        Rem     If DBGrid1.Row < 40 Then
        Rem         DBGrid1.Row = DBGrid1.Row + 1
        Rem         WRow = DBGrid1.Row
        Rem         XRow = DBGrid1.Row
        Rem         DBGrid1.Col = 4
        Rem         KeyCode = 0
        Rem     End If
        Rem     Call Calcula_Click
        Rem     DBGrid1.Row = XRow
        Rem     DBGrid1.Col = 4
        Rem     KeyCode = 0
        Rem     Exit Sub
        Rem End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la �ltima fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ning�n valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila bas�ndose en su marcador.
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
' DBGrid est� solicitando filas, as� que se las damos

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
    ' Busca la posici�n para empezar a leer, bas�ndose en el marcador
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
    ' Establece el marcador mediante CurRow&, que es tambi�n
    ' nuestro �ndice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz despu�s de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se est�n actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()


    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
    
    Tipoventa.Clear
    
    Tipoventa.AddItem "Venta Normal"
    Tipoventa.AddItem "Mercaderia en Consignacion"
    
    Tipoventa.ListIndex = 0

    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 5, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 5
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Parcial a Entregar"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 5
             DBGrid1.Columns(newcnt).Caption = "OK"
             DBGrid1.Columns(newcnt).Width = 300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

    Erase XLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    Wlote3.Text = ""

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Numero.SetFocus
     
End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    WNeto = 0
    
    Erase Auxiliar
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Canti = !Cantidad - !Facturado
                
                    Renglon = Renglon + 1
                
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = !Terminado
                    Auxi1 = !Terminado
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad))
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(!Precio))
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", Str$(Canti))
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = Canti
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For da = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(da, 1)
        Canti = Auxiliar(da, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
        
            DBGrid1.Col = 1
            DBGrid1.Text = rstPrecios!Descripcion
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
            Precio = rstPrecios!Precio
            rstPrecios.Close
        End If

        If Val(Canti) <> 0 Then
            WNeto = WNeto + (Val(Canti) * Precio)
        End If
        
    Next da
    
    Call Calcula_Click

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Graba.Enabled = True
    Borra.Enabled = True

End Sub

Private Sub Proceso1_Click()

    WNeto = 0

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    
    
    XParam = "'" + "01" + "','" _
                + Numero.Text + "'"
    
    spEstadistica = "ConsultaEstadistica1 " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstEstadistica!Articulo
                    Auxi1 = rstEstadistica!Articulo
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Precio)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Paridad)
                    Paridad.Text = Pusing("###,###.##", dada)
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!Precio)
                    End If
                    
                    Auxiliar(Renglon, 1) = Auxi1
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    XRenglon = Renglon
    Renglon = 0
    
    For da = 1 To XRenglon
    
        Auxi1 = Auxiliar(da, 1)
                    
        ClavePrecios = Cliente.Text + Auxi1
        
        spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                    
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
                    
            DBGrid1.Col = 1
            DBGrid1.Text = rstPrecios!Descripcion
            rstPrecios.Close
        End If
    Next da
    
    Renglon = Renglon + 1
            
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
                
    DBGrid1.Col = 0
    DBGrid1.Text = ""
                
    DBGrid1.Col = 2
    DBGrid1.Text = ""
                
    DBGrid1.Col = 3
    DBGrid1.Text = ""
                
    DBGrid1.Col = 4
    DBGrid1.Text = ""
                            
    DBGrid1.Col = 1
    DBGrid1.Text = ""
    
    Call Calcula_FechaVto
    Call Calcula_Click

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.FirstRow = 0
    DBGrid1.Row = 0
    DBGrid1.Col = 0
    
    DBGrid1.SetFocus
    
    Graba.Enabled = False
    Borra.Enabled = False

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "01" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Pedido.Text = rstCtacte!Pedido
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            Vencimiento.Text = rstCtacte!Vencimiento
            Remito.Text = rstCtacte!Remito
            Orden.Text = rstCtacte!Orden
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!Vendedor
                WProv = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
            End If
            Call Proceso1_Click
                    Else
            Rem .Index = "Numero"
            Rem .Seek "=", Val(Numero.Text)
            Rem If .NoMatch = False Then
            Rem     m$ = "Comprobante ya existente"
            Rem     A% = MsgBox(m$, 0, "Ingreso de Facturas")
            Rem     Numero.SetFocus
            Rem        Else
            Rem     WNumero = Numero.Text
            Rem    Rem Call Limpia_Click
            Rem    Numero.Text = WNumero
            Rem    Pedido.SetFocus
            Rem End If
            WNumero = Numero.Text
            Rem Call Limpia_Click
            Numero.Text = WNumero
            Pedido.SetFocus
                
        End If
    End If
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Cliente.Text = rstPedido!Cliente
            rstPedido.Close
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!Vendedor
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WProv = rstCliente!Provincia
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
            End If
            Call Calcula_FechaVto
            Call Proceso_Click
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            spCambios = "ConsultaCambio  " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = "1.00"
            End If
            If Val(Paridad.Text) <> 0 Then
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
                Remito.SetFocus
                    Else
                m$ = "No exsite paridad cargada para esta fecha"
                a% = MsgBox(m$, 0, "Emision de facturas")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de facturas")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub ReImpre_Click()
    Call Impresion
    Call Impresion_Remito
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Remito.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Orden.SetFocus
    End If
End Sub

Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        DBGrid1.SetFocus
    End If
End Sub

Sub Impresion()

    If Val(WEmpresa) = 1 Then
        Open "lpt1" For Output As #1
        Rem Open "DADA.TXT" For Output As #1
            Else
        Open "lpt1" For Output As #1
        Rem Open "DADA.TXT" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
    End If
    
    Rem Width #1, 255

    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72);
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)

    For XX% = 1 To 2
    
        If XX% = 1 Then
            Print #1, ""
                Else
            Print #1, ""
        End If

        If Val(WEmpresa) = 1 Then
            Print #1, ""
        End If
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(WEmpresa) = 1 Then
            Print #1, Tab(59); Fecha.Text
                Else
            Print #1, Tab(57); Fecha.Text
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(8); WRazon
        Print #1, Tab(8); WDireccion
        Print #1, Tab(8); Left$(WLocalidad, 33);
        Print #1, Tab(55); Cliente.Text;
        Print #1, Tab(69); Orden.Text
        Print #1, Tab(8); Provincia(Val(WProv)); " ("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(8); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, ""
        Print #1, Tab(5); Left$(WPago, 40); " ";
        Print #1, Vencimiento.Text;
        Print #1, Tab(60); Remito.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(Alinea("##,###.##", DBGrid1.Text))
            
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    Print #1, Tab(1); Alinea("#####.##", Str$(Cantidad));
                    Print #1, " Kg";
                    Print #1, Tab(15); Left$(Descri, 40);
                    parcial = Str$(Precio * Cantidad)
                    
                    Rem If Val(WCodIva) = 1 Or Val(WCodIva) = 2 Then
                    Rem     Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem     Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem             Else
                    Rem     Precio = Str$(Val(Precio) * 1.21)
                    Rem     Parcial = Str$(Val(Parcial) * 1.21)
                    Rem     Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem     Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem End If
                    
                    Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    
                    Print #1, Tab(68); Alinea("###,###.##", parcial)
                    
                    Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a

        For aa = Impre To 22
                Print #1, ""
        Next aa

        Rem M# = Total# / 100
        Rem GoSub 4630

        Print #1, Tab(1); "EL IMPORTE DE ESTA FACTURA ESTA EXPRESADO EN DOLARES."
        Print #1, Tab(1); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
        Paridad = Val(Paridad.Text)
        ImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption)) * Paridad
        Impotot = Val(Total.Caption) * Paridad
        Print #1, Tab(1); "TIPO DE CAMBIO:";
        Print #1, Alinea("##.##", Str$(Paridad));
        Print #1, " I.V.A.:";
        Print #1, Alinea("#,###.##", Str$(ImpoIva));
        Print #1, " TOTAL:";
        Print #1, Alinea("###,###.##", Str$(Impotot))
        Print #1, Tab(1); "CONDICIONES : SI POR FUERZA MAYOR NO FUESE  POSIBLE EL"
        If Val(WEmpresa) = 1 Then
            Print #1, Tab(1); "PAGO EN DOLARES BILLETE; SURFACTAN S.A. PODRA OPTAR EN"
                Else
            Print #1, Tab(1); "PAGO EN DOLARES BILLETE; PELLITAL S.A. PODRA OPTAR EN "
        End If
        Print #1, Tab(1); "RECIBIR PESOS BONEX/89 COTIZACION MERCADO NVA.YORK, EN"
        Print #1, Tab(1); "CANTIDAD SUFICIENTE PARA  ADQUIRIR EL  EQUIVALENTE  AL"
        Print #1, Tab(1); "PRECIO EN DOLARES. SI EL IMPORTE NO SE CANCELARA EN EL"
        Print #1, Tab(1); "PLAZO ESTIPULADO A PARTIR DE SU VENCIMIENTO Y HASTA LA"
        Print #1, Tab(1); "FECHA EFVO. PAGO SE APLICARA UNA TASA DEL "
        Print #1, Alinea("##.##", Str$(WTasa));
        Print #1, " %MENSUAL"
        
        Print #1, Tab(68); Alinea("###,###.##", Str$(XNeto))

        If Val(Dto.Caption) <> 0 Then
                Print #1, Tab(56); "Dto."; Alinea("###.##", Str$(WDescuento));
                Print #1, Tab(68); Alinea("###,###.##", Dto.Caption)
                        Else
                Print #1, ""
        End If

        If Val(Interes.Caption) <> 0 Then
                Print #1, Tab(56); "Interes";
                Print #1, Tab(68); Alinea("###,###.##", Interes.Caption)
                                                  Else
                Print #1, ""
        End If

        Print #1, Tab(3); M1;
        Print #1, Tab(68); Alinea("###,###.##", Neto.Caption)
        Print #1, Tab(3); M2;
        If Val(Iva1.Caption) <> 0 Then
                Print #1, Tab(61); "21";
                Print #1, Tab(68); Alinea("###,###.##", Iva1.Caption)
                        Else
                Print #1, ""
        End If

        Select Case XX
                Case 1
                        Print #1, Tab(10); "ORIGINAL";
                Case 2
                        Print #1, Tab(10); "DUPLICADO";
                Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                Case Else
        End Select

        If Val(Iva2.Caption) <> 0 Then
                Print #1, Tab(61); "10.5";
                Print #1, Tab(68); Alinea("###,###.##", Iva2.Caption)
                        Else
                Print #1, ""
        End If

        Print #1, Tab(68); Alinea("###,###.##", Total.Caption); Chr$(12)

        Next XX%

        Close #1
        
End Sub

Sub Impresion_Remito()

        If Val(WEmpresa) = 1 Then
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
                Else
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        End If
  
        Rem  #1, 255

        For FF = 1 To 2

        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(7); WRazon
        Print #1, Tab(7); Left$(WDireccion, 33)
        Print #1, Tab(7); Left$(WLocalidad, 33);
        Print #1, Tab(44); Pedido.Text;
        Print #1, Tab(57); Cliente.Text;
        Print #1, Tab(68); Orden.Text
        Print #1, Tab(7); Provincia(Val(WProv)); "("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(7); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, Tab(30); WDirentrega;
        Print #1, ""
        If FF = 1 Then
            Print #1, Tab(60); "ORIGINAL"
                Else
            Print #1, Tab(60); "DUPLICADO"
        End If
        Print #1, ""
        
        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
            
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                        
                        Print #1, Tab(14); Left$(Descri, 40);
                        Print #1, Tab(58); Alinea("#####.##", Str$(Cantidad));
                        Print #1, " Kg";
                        Print #1, Tab(71); "Netos"
                        Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a
        
        For aa = Impre To 22
                Print #1, ""
        Next aa
        
        Print #1, ""
        Print #1, Tab(10); "Lugar de Pago : Ayacucho 1231 5to Piso Dto. 'A' Capital Federal"
        Print #1, ""

        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)

        For XDa = 1 To 1
                For da = 1 To 9
                        If Val(Stk(da, 4)) <> 0 Then
                                        
                                Select Case da
                                        Case 1
                                                Lugar = 22
                                        Case 2
                                                Lugar = 33
                                        Case 3
                                                Lugar = 44
                                        Case 4
                                                Lugar = 55
                                        Case 5
                                                Lugar = 66
                                        Case 6
                                                Lugar = 77
                                        Case 7
                                                Lugar = 89
                                        Case 8
                                                Lugar = 101
                                        Case 9
                                                Lugar = 113
                                        Case Else
                                End Select
                                                         
                                If da = 9 Then
                                    Digi = 10
                                            Else
                                    Digi = 10
                                End If
                                
                                spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(da, XDa))) + "'"
                                Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                If rstEnvases.RecordCount > 0 Then
                                    Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                    rstEnvases.Close
                                            Else
                                    Print #1, Tab(Lugar); Stk(da, XDa);
                                End If
                            End If
        
                Next da
                Print #1, ""
        
        Next XDa
        
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
        For XDa = 2 To 4
                For da = 1 To 9
        
                        If Val(Stk(da, 4)) <> 0 Then
        
                                Select Case da
                                        Case 1
                                                Lugar = 14
                                        Case 2
                                                Lugar = 21
                                        Case 3
                                                Lugar = 29
                                        Case 4
                                                Lugar = 36
                                        Case 5
                                                Lugar = 43
                                        Case 6
                                                Lugar = 50
                                        Case 7
                                                Lugar = 57
                                        Case 8
                                                Lugar = 64
                                        Case 9
                                                Lugar = 71
                                        Case Else
                                End Select
        
                                If Val(Stk(da, XDa)) <> 0 Then
                                        Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(da, XDa))));
                                End If
        
                         End If
                Next da
        
                Print #1, ""
                Print #1, ""
        
        Next XDa
        
        Print #1, ""
        Select Case XX
                Case 1
                        Print #1, Tab(10); "ORIGINAL";
                Case 2
                        Print #1, Tab(10); "DUPLICADO";
                Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                Case Else
        End Select
        Print #1, Tab(10); "Nro. Control : "; Remito.Text
        Print #1, Chr$(12)

        Next FF

        Close #1


End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti1.SetFocus
                Else
            Envase1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri2.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti2.SetFocus
                Else
            Envase2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri3.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti3.SetFocus
                Else
            Envase3.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri4.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti4.SetFocus
                Else
            Envase4.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri5.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti5.SetFocus
                Else
            Envase5.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Calcula_Saldo()

    Rem On Error GoTo Error_saldo


    Erase Stk

    Stk(1, 1) = "020"
    Stk(2, 1) = "021"
    Stk(3, 1) = "022"
    Stk(4, 1) = "023"
    Stk(5, 1) = "024"
    Stk(6, 1) = "025"
    Stk(7, 1) = "026"
    Stk(8, 1) = "030"
    Stk(9, 1) = "028"

    XParam = "'" + Cliente.Text + "','" _
                + Cliente.Text + "'"

    spMovenv = "ListaMovenvDesdeHastaCliente " + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
            .MoveFirst
            Do
                If .EOF = False Then

                    For da = 1 To 9
                        If Val(Stk(da, 1)) = !Envase Then
                            If !Movimiento = "S" Then
                                Stk(da, 2) = Str$(Val(Stk(da, 2)) + !Cantidad)
                                    Else
                                Stk(da, 2) = Str$(Val(Stk(da, 2)) - !Cantidad)
                            End If
                        End If
                    
                    Next da
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovenv.Close
    End If

End Sub



Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    
    If Val(WLote1.Text) <> 0 Then
        Suma = Suma + Val(WCanti1.Text)
    End If
    If Val(WLote2.Text) <> 0 Then
        Suma = Suma + Val(WCanti2.Text)
    End If
    If Val(Wlote3.Text) <> 0 Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    
    If Suma = XCantidad Then
        WEstado = "S"
    End If
    
End Sub





