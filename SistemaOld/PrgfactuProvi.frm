VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactuProvi 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturacion de Pedidos (Provisorio)"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.CommandButton Graba1 
      Caption         =   "Fc. Exportacion"
      Height          =   495
      Left            =   10200
      TabIndex        =   55
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Marca 
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   54
      Top             =   5760
      Width           =   3015
   End
   Begin VB.TextBox Cip 
      Height          =   285
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   53
      Text            =   " "
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox Consignatario 
      Height          =   285
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   52
      Text            =   " "
      Top             =   7920
      Width           =   2535
   End
   Begin MSMask.MaskEdBox fecorden 
      Height          =   255
      Left            =   4200
      TabIndex        =   51
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox NroOrden 
      Height          =   285
      Left            =   960
      MaxLength       =   10
      TabIndex        =   50
      Text            =   " "
      Top             =   7560
      Width           =   1815
   End
   Begin VB.TextBox Pago2 
      Height          =   285
      Left            =   960
      MaxLength       =   50
      TabIndex        =   49
      Text            =   " "
      Top             =   7200
      Width           =   5055
   End
   Begin VB.TextBox Pago1 
      Height          =   285
      Left            =   960
      MaxLength       =   50
      TabIndex        =   48
      Text            =   " "
      Top             =   6840
      Width           =   5055
   End
   Begin VB.TextBox Envio2 
      Height          =   285
      Left            =   960
      MaxLength       =   50
      TabIndex        =   47
      Text            =   " "
      Top             =   6480
      Width           =   5055
   End
   Begin VB.TextBox Envio1 
      Height          =   285
      Left            =   960
      MaxLength       =   50
      TabIndex        =   46
      Text            =   " "
      Top             =   6120
      Width           =   5055
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
      Height          =   2415
      Left            =   8760
      TabIndex        =   23
      Top             =   5760
      Width           =   2535
      Begin VB.Label Flete 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   59
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Seguro 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   58
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Flete"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Seguro"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   975
      End
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
         Top             =   2040
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
         Top             =   2040
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
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   1200
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
      ItemData        =   "PrgfactuProvi.frx":0000
      Left            =   6480
      List            =   "PrgfactuProvi.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "PrgfactuProvi.frx":0015
      TabIndex        =   2
      Top             =   1560
      Width           =   11415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8520
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.Label Label20 
      Caption         =   "Cip"
      Height          =   375
      Left            =   4200
      TabIndex        =   45
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label Label19 
      Caption         =   "Consignatario"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label18 
      Caption         =   "Fecha Orden"
      Height          =   375
      Left            =   2880
      TabIndex        =   43
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "Nro orden"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Label Label14 
      Caption         =   "Pago"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Label Label13 
      Caption         =   "Envio"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label rrr 
      Caption         =   "Marca"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5760
      Width           =   1215
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
Attribute VB_Name = "PrgFactuProvi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 20 ' Número máximo de campos del conjunto de registros.
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
Private WImpoIb As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private parcial As String
Private WSeguro As Double
Private WFlete As Double
Private WTexto1 As String
Private WTexto2 As String
Private Auxiliar(100, 15) As String
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
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim XParam As String
Dim WLote(5, 2) As String
Dim WImpresion(100, 10) As String
Dim XEnvase(100, 6) As String
Dim XCanti As String


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
    
End Sub

Private Sub Calcula_Click()

    WNeto = 0
    
    For a = 0 To 5
        
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

    Rem If Val(Paridad.Text) <> 0 Then
    Rem     WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    
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
    WImpoIb = 0
    
    Rem Select Case Val(WCodIva)
    Rem     Case 2
    Rem         WIva1 = WNeto * 0.21
    Rem         WIva2 = WNeto * 0.105
    Rem         Call Redondeo(WIva1)
    Rem         Call Redondeo(WIva2)
    Rem     Case 4
    Rem         WIva1 = 0
    Rem         WIva2 = 0
    Rem     Case Else
    Rem         WIva1 = WNeto * 0.21
    Rem         Call Redondeo(WIva1)
    Rem End Select
    
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
    
    Seguro.Caption = Alinea("###,###.##", Str$(WSeguro))
    Flete.Caption = Alinea("###,###.##", Str$(WFlete))
    WTotal = WNeto + WIva1 + WIva2 + WSeguro + WFlete
    
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    
    PrgFactuProvi.Hide
    Unload Me
    Menu.Show
    
End Sub



Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Graba_Click()

        Renglon = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 5
        
                Suma = a * 10
                DBGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRenglon = WRenglon + 1
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Articulo = DBGrid1.Text
                    
                    DBGrid1.Col = 4
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 19
                    Marca = DBGrid1.Text
                    
                    DBGrid1.Col = 0
                    aa = DBGrid1.Text
                    DBGrid1.Col = 1
                    aa = DBGrid1.Text
                    DBGrid1.Col = 2
                    aa = DBGrid1.Text
                    DBGrid1.Col = 3
                    aa = DBGrid1.Text
                    DBGrid1.Col = 4
                    aa = DBGrid1.Text
                    DBGrid1.Col = 5
                    aa = DBGrid1.Text
                    DBGrid1.Col = 6
                    aa = DBGrid1.Text
                    DBGrid1.Col = 7
                    aa = DBGrid1.Text
                    DBGrid1.Col = 8
                    aa = DBGrid1.Text
                    DBGrid1.Col = 9
                    aa = DBGrid1.Text
                    
                    If Cantidad <> 0 And Marca <> "X" Then
                        m$ = Articulo + " Verifique la discrminacion de lotes"
                        G% = MsgBox(m$, 0, "Emision de facturas")
                        Exit Sub
                    End If
                Next iRow
        Next a

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
        XTotalUs = Str$(WTotal)
        XSaldo = Str$(WTotal)
        XSaldoUs = Str$(WTotal)
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
        WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
        WImpre = "FC"
        XNet = Str$(WNeto * Val(Paridad.Text))
        XIva1 = Str$(WIva1 * Val(Paridad.Text))
        XIva2 = Str$(WIva2 * Val(Paridad.Text))
        XImpoIb = Str$(WImpoIb * Val(Paridad.Text))
        XSeguro = Str$(WSeguro)
        XFlete = Str$(WFlete)
        WPedido = Pedido.Text
        WRemito = Remito.Text
        WOrden = Orden.Text
        WParidad = Paridad.Text
        WProvincia = WProv
        XVendedor = Str$(WVendedor)
        XRubro = Str$(WRubro)
        WComprobante = ""
        WAceptada = ""
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
                    + WImporte1 + "','" + WImporte2 + "','" _
                    + WImporte3 + "','" + WImporte4 + "','" _
                    + WImporte5 + "','" + WImporte6 + "','" _
                    + WImporte7 + "','" + WDate + "','" _
                    + XSeguro + "','" + XFlete + "','" _
                    + XImpoIb + "'"
                        
        spCtacte = "AltaCtacte " + XParam
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
        Renglon = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For a = 0 To 5
        
                Suma = a * 10
                DBGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRenglon = WRenglon + 1
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Articulo = DBGrid1.Text
                    
                    DBGrid1.Col = 3
                    Precio = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 4
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 9
                    lote1 = Val(DBGrid1.Text)
                    DBGrid1.Col = 10
                    Canti1 = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 11
                    lote2 = Val(DBGrid1.Text)
                    DBGrid1.Col = 12
                    Canti2 = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 13
                    lote3 = Val(DBGrid1.Text)
                    DBGrid1.Col = 14
                    Canti3 = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 15
                    lote4 = Val(DBGrid1.Text)
                    DBGrid1.Col = 16
                    Canti4 = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 17
                    lote5 = Val(DBGrid1.Text)
                    DBGrid1.Col = 18
                    Canti5 = Val(DBGrid1.Text)
                    
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
                        XCantidad = Str$(Cantidad)
                        XPrecio = Str$(Precio * Val(Paridad.Text))
                        XPrecioUs = Str$(Precio)
                        XImporte = Str$(Precio * Cantidad * Val(Paridad.Text))
                        XImporteUs = Str$(Precio * Cantidad)
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
                        WRemito = Remito.Text
                        WClave = "01" + Auxi1 + Auxi
                        WDate = Date$
                        XCanti = ""
                        XImpo = ""
                        XImpoUs = ""
                        XMarca = ""
                        WLote1 = Str$(lote1)
                        WCanti1 = Str$(Canti1)
                        WLote2 = Str$(lote2)
                        WCanti2 = Str$(Canti2)
                        Wlote3 = Str$(lote3)
                        WCanti3 = Str$(Canti3)
                        WLote4 = Str$(lote4)
                        WCanti4 = Str$(Canti4)
                        WLote5 = Str$(lote5)
                        WCanti5 = Str$(Canti5)
                        XTipoproDy = "T"
                        XArticuloDy = "  -   -   "
                    
                        XParam = "'" + WClave + "','" _
                            + WTipo + "','" + WNumero + "','" _
                            + XRenglon + "','" + WArticulo + "','" _
                            + XCantidad + "','" + XPrecio + "','" _
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
                            + XImpo + "','" + XImpoUs + "','" _
                            + XMarca + "','" _
                            + WLote1 + "','" + WCanti1 + "','" _
                            + WLote2 + "','" + WCanti2 + "','" _
                            + Wlote3 + "','" + WCanti3 + "','" _
                            + WLote4 + "','" + WCanti4 + "','" _
                            + WLote5 + "','" + WCanti5 + "','" _
                            + XTipoproDy + "','" + XArticuloDy + "'"
                    
                        spEstadistica = "AltaEstadistica " + XParam
                        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Auxiliar(Renglon, 1) = Articulo
                        Auxiliar(Renglon, 2) = Cantidad
                        Auxiliar(Renglon, 3) = Precio
                        Auxiliar(Renglon, 4) = WRenglon
                        Auxiliar(Renglon, 5) = lote1
                        Auxiliar(Renglon, 6) = Canti1
                        Auxiliar(Renglon, 7) = lote2
                        Auxiliar(Renglon, 8) = Canti2
                        Auxiliar(Renglon, 9) = lote3
                        Auxiliar(Renglon, 10) = Canti3
                        Auxiliar(Renglon, 11) = lote4
                        Auxiliar(Renglon, 12) = Canti4
                        Auxiliar(Renglon, 13) = lote5
                        Auxiliar(Renglon, 14) = Canti5

                    End If
                                        
                Next iRow
            
        Next a
        
        For DA = 1 To Renglon
        
            Articulo = Auxiliar(DA, 1)
            Cantidad = Auxiliar(DA, 2)
            Precio = Auxiliar(DA, 3)
            WRenglon = Auxiliar(DA, 4)
            lote1 = Auxiliar(DA, 5)
            Cantidad1 = Auxiliar(DA, 6)
            lote2 = Auxiliar(DA, 7)
            Cantidad2 = Auxiliar(DA, 8)
            lote3 = Auxiliar(DA, 9)
            Cantidad3 = Auxiliar(DA, 10)
            lote4 = Auxiliar(DA, 11)
            Cantidad4 = Auxiliar(DA, 12)
            lote5 = Auxiliar(DA, 13)
            Cantidad5 = Auxiliar(DA, 14)
            
            WLote(1, 1) = lote1
            WLote(1, 2) = Cantidad1
            WLote(2, 1) = lote2
            WLote(2, 2) = Cantidad2
            WLote(3, 1) = lote3
            WLote(3, 2) = Cantidad3
            WLote(4, 1) = lote4
            WLote(4, 2) = Cantidad4
            WLote(5, 1) = lote5
            WLote(5, 2) = Cantidad5
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                WCodigo = Articulo
                WPedido = Str$(rstTerminado!Pedido - Cantidad)
                WSalidas = Str$(rstTerminado!Salidas + Cantidad)
                WDate = Date$
                
                WLinea = rstTerminado!Linea
                rstTerminado.Close
                
                
                XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                           
                spTerminado = "ModificaTerminadoFacturas " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            For Da2 = 1 To 5
            
                If WLote(Da2, 1) <> 0 Then
                    Lote = WLote(Da2, 1)
                    Cantilote = WLote(Da2, 2)
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                + Articulo + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            
                            WClave = rstHoja!Clave
                            WSaldo = Str$(rstHoja!Saldo - Cantilote)
                            WDate = Date$
                            rstHoja.Close
                            
                            XParam = "'" + WClave + "','" _
                                + WDate + "','" _
                                + WSaldo + "'"
                            spHoja = "ModificaHojaSaldo " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            XParam = "'" + Articulo + "','" _
                                        + Lote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                WSaldo = Str$(rstMovguia!Saldo - Cantilote)
                                WDate = Date$
                                rstMovguia.Close
                        
                                XParam = "'" + WClave + "','" _
                                        + WDate + "','" _
                                        + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                End If
            
            Next Da2

            Auxi = Pedido.Text
            Call Ceros(Auxi, 6)
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            ClavePedido = Auxi + Auxi1
            
            XParam = "'" + Left$(ClavePedido, 6) + "','" _
                        + Right$(ClavePedido, 2) + "'"
            spPedido = "ConsultaPedido2 " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                WFacturado = Str$(rstPedido!Facturado + Cantidad)
                If Val(WFacturado) > rstPedido!Cantidad Then
                    WFacturado = Str$(rstPedido!Cantidad)
                End If
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
                    WFecha1 = rstPrecios!fecha2
                    WFactura1 = rstPrecios!Factura2
                    WPrecio1 = Str$(rstPrecios!Precio2)
                    WCantidad1 = Str$(rstPrecios!Cantidad2)
                End If
                                
                If rstPrecios!Cantidad3 <> O Then
                    WFecha2 = rstPrecios!Fecha3
                    WFactura2 = rstPrecios!Factura3
                    WPrecio2 = Str$(rstPrecios!Precio3)
                    WCantidad2 = Str$(rstPrecios!Cantidad3)
                End If
                                
                If rstPrecios!Cantidad4 <> O Then
                    WFecha3 = rstPrecios!Fecha4
                    WFactura3 = rstPrecios!Factura4
                    WPrecio3 = Str$(rstPrecios!Precio4)
                    WCantidad3 = Str$(rstPrecios!Cantidad4)
                End If
                                
                If rstPrecios!Cantidad5 <> O Then
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
        Next DA
                    
        spNumero = "ConsultaNumero " + "'" + "02" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            WCodigo = "02"
            WNumero = Numero.Text
            rstNumero.Close
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call Impresion_Remito
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
End Sub


Private Sub Graba1_Click()

        Call Calcula_Click
        
        Call Impresion
        
        Numero.SetFocus
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    
    For a = 0 To 5
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 19
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
    Seguro.Caption = ""
    Flete.Caption = ""
    
    Marca.Text = ""
    Envio1.Text = ""
    Envio2.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    NroOrden.Text = ""
    fecorden.Text = "  /  /    "
    Consignatario.Text = ""
    Cip.Text = ""
    
    spNumero = "ConsultaNumero " + "'" + "02" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
    
    Numero.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 39 Then
        KeyCode = 13
    End If

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                Select Case KeyCode
                    Case 13
                        WAuxi = DBGrid1.Col
                        DBGrid1.Col = 4
                        DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = 5
                        DBGrid1.Text = Pusing("####", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = 8
                        DBGrid1.Text = Pusing("#####.##", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = WAuxi
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                End Select
                        
            Case 9, 11, 13, 15, 17
                Select Case KeyCode
                    Case 13
                        WAuxi = DBGrid1.Col
                        
                        WEntra = "N"
                        
                        DBGrid1.Col = 0
                        XTerminado = DBGrid1.Text
                        DBGrid1.Col = WAuxi
                        XLote = DBGrid1.Text
                        
                        If Val(XLote) <> 0 Then
                        
                        WControla = 0
                        spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                            rstTerminado.Close
                        End If
            
                        If WControla = 0 Then
                            XParam = "'" + XLote + "','" _
                                    + XTerminado + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                XSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                WEntra = "S"
                                rstHoja.Close
                            End If
                
                            If WEntra = "N" Then
                                XParam = "'" + XTerminado + "','" _
                                        + XLote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    WEntra = "S"
                                    rstMovguia.Close
                                End If
                            End If
                
                                Else
                    
                            WEntra = "S"
                            
                        End If
                        
                        If WEntra = "S" Then
                        
                            DBGrid1.Col = WAuxi + 1
                            KeyCode = 0
                            
                                Else
                                
                            Select Case WAuxi
                                Case 11
                                    DBGrid1.Col = 9
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 10
                                    a = DBGrid1.Text
                                Case 13
                                    DBGrid1.Col = 9
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 10
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 11
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 12
                                    a = DBGrid1.Text
                                Case 15
                                    DBGrid1.Col = 9
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 10
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 11
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 12
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 13
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 14
                                    a = DBGrid1.Text
                                Case 17
                                    DBGrid1.Col = 9
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 10
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 11
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 12
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 13
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 14
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 15
                                    a = DBGrid1.Text
                                    DBGrid1.Col = 16
                                    a = DBGrid1.Text
                                Case Else
                            End Select
                                
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + XLote + " inexistente"
                            G% = MsgBox(m$, 0, "Emision de facturas")
                            DBGrid1.Col = WAuxi
                            KeyCode = 0
                            
                        End If
                        
                            Else
                            
                        DBGrid1.Col = 4
                        Canti1 = Val(DBGrid1.Text)
                        Canti2 = 0
                        DBGrid1.Col = 10
                        Canti2 = Canti2 + Val(DBGrid1.Text)
                        DBGrid1.Col = 12
                        Canti2 = Canti2 + Val(DBGrid1.Text)
                        DBGrid1.Col = 14
                        Canti2 = Canti2 + Val(DBGrid1.Text)
                        DBGrid1.Col = 16
                        Canti2 = Canti2 + Val(DBGrid1.Text)
                        DBGrid1.Col = 18
                        Canti2 = Canti2 + Val(DBGrid1.Text)
                        
                        If Canti1 = Canti2 Then
                            If DBGrid1.Row < 40 Then
                                DBGrid1.Col = 19
                                DBGrid1.Text = "X"
                                DBGrid1.Row = DBGrid1.Row + 1
                                WRow = DBGrid1.Row
                                DBGrid1.Col = 0
                                KeyCode = 0
                                DBGrid1.Col = 4
                            End If
                                Else
                            DBGrid1.Col = WAuxi
                            KeyCode = 0
                        End If
                        
                        End If
                        
                    Case Else
                        Rem If KeyCode <> 0 Then stop
                End Select
                
            Case 10, 12, 14, 16, 18
                Select Case KeyCode
                    Case 13
                        WAuxi = DBGrid1.Col
                        XCantidad = Val(DBGrid1.Text)
                        
                        WEntra = "N"
                        
                        DBGrid1.Col = 0
                        XTerminado = DBGrid1.Text
                        DBGrid1.Col = WAuxi - 1
                        XLote = DBGrid1.Text
            
                        WControla = 0
                        spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                            rstTerminado.Close
                        End If
            
                        If WControla = 0 Then
                            XParam = "'" + XLote + "','" _
                                    + XTerminado + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                XSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                WEntra = "S"
                                rstHoja.Close
                            End If
                
                            If WEntra = "N" Then
                                XParam = "'" + XTerminado + "','" _
                                        + XLote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    WEntra = "S"
                                    rstMovguia.Close
                                End If
                            End If
                
                                Else
                    
                            WEntra = "S"
                            
                        End If
                        
                        If WEntra = "S" Then
                            
                            If XCantidad > XSaldo Then
                                m$ = XTerminado + " Cantidad Insuficiente Stock : " + Str$(XSaldo)
                                G% = MsgBox(m$, 0, "Emision de facturas")
                                    Else
                                If WAuxi = 18 Then
                                    DBGrid1.Col = 4
                                    Canti1 = Val(DBGrid1.Text)
                                    Canti2 = 0
                                    DBGrid1.Col = 10
                                    Canti2 = Canti2 + Val(DBGrid1.Text)
                                    DBGrid1.Col = 12
                                    Canti2 = Canti2 + Val(DBGrid1.Text)
                                    DBGrid1.Col = 14
                                    Canti2 = Canti2 + Val(DBGrid1.Text)
                                    DBGrid1.Col = 16
                                    Canti2 = Canti2 + Val(DBGrid1.Text)
                                    DBGrid1.Col = 18
                                    Canti2 = Canti2 + Val(DBGrid1.Text)
                            
                                    If Canti1 = Canti2 Then
                                        If DBGrid1.Row < 40 Then
                                            DBGrid1.Col = 19
                                            DBGrid1.Text = "X"
                                            DBGrid1.Row = DBGrid1.Row + 1
                                            WRow = DBGrid1.Row
                                            DBGrid1.Col = 0
                                            KeyCode = 0
                                            DBGrid1.Col = 4
                                        End If
                                            Else
                                        DBGrid1.Col = WAuxi
                                        KeyCode = 0
                                    End If
                                        Else
                                    DBGrid1.Col = WAuxi + 1
                                    KeyCode = 0
                                End If
                            End If
                            
                                Else
                                
                            m$ = XTerminado + " Cantidad Insuficiente Stock : " + Str$(XSaldo)
                            G% = MsgBox(m$, 0, "Emiison de facturas")
                                
                        End If
                        
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                End Select
            
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
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
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

    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 19, 0 To 80)

mTotalRows& = 80

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
For i = 0 To 19
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Parcial"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 780
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 7
             DBGrid1.Columns(newcnt).Caption = "Numero"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 8
             DBGrid1.Columns(newcnt).Caption = "Bruto"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 9
             DBGrid1.Columns(newcnt).Caption = "Lote 1"
             DBGrid1.Columns(newcnt).Width = 800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 10
             DBGrid1.Columns(newcnt).Caption = "Cantidad 1"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 11
             DBGrid1.Columns(newcnt).Caption = "Lote 2"
             DBGrid1.Columns(newcnt).Width = 800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 12
             DBGrid1.Columns(newcnt).Caption = "Cantidad 2"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 13
             DBGrid1.Columns(newcnt).Caption = "Lote 3"
             DBGrid1.Columns(newcnt).Width = 800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 14
             DBGrid1.Columns(newcnt).Caption = "Cantidad 3"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 15
             DBGrid1.Columns(newcnt).Caption = "Lote 4"
             DBGrid1.Columns(newcnt).Width = 800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 16
             DBGrid1.Columns(newcnt).Caption = "Cantidad 4"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 17
             DBGrid1.Columns(newcnt).Caption = "Lote 5"
             DBGrid1.Columns(newcnt).Width = 800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 18
             DBGrid1.Columns(newcnt).Caption = "Cantidad 5"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 19
             DBGrid1.Columns(newcnt).Caption = ""
             DBGrid1.Columns(newcnt).Width = 250
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

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
    Seguro.Caption = ""
    Flete.Caption = ""
    
    Marca.Text = ""
    Envio1.Text = ""
    Envio2.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    NroOrden.Text = ""
    fecorden.Text = "  /  /    "
    Consignatario.Text = ""
    Cip.Text = ""
    
    spNumero = "ConsultaNumero " + "'" + "02" + "'"
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

    For a = 0 To 5
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 19
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    WNeto = 0
    WSeguro = 0
    WFlete = 0
    
    Erase Auxiliar
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Canti = !Cantidad
                
                    Select Case Mid$(!Terminado, 1, 2)
                        Case "Z2"
                            WSeguro = WSeguro + (!Precio * !Cantidad)
                                                            
                        Case "Z1"
                            WFlete = WFlete + (!Precio * !Cantidad)
                    
                        Case Else
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
                            
                            DBGrid1.Col = 9
                            DBGrid1.Text = Str$(!lote1)
                            
                            DBGrid1.Col = 10
                            DBGrid1.Text = Str$(!CantiLote1)
                            
                            DBGrid1.Col = 11
                            DBGrid1.Text = Str$(!lote2)
                            
                            DBGrid1.Col = 12
                            DBGrid1.Text = Str$(!CantiLote2)
                            
                            DBGrid1.Col = 13
                            DBGrid1.Text = Str$(!lote3)
                            
                            DBGrid1.Col = 14
                            DBGrid1.Text = Str$(!CantiLote3)
                            
                            DBGrid1.Col = 15
                            DBGrid1.Text = Str$(!lote4)
                            
                            DBGrid1.Col = 16
                            DBGrid1.Text = Str$(!CantiLote4)
                            
                            DBGrid1.Col = 17
                            DBGrid1.Text = Str$(!lote5)
                            
                            DBGrid1.Col = 18
                            DBGrid1.Text = Str$(!CantiLote5)
                            
                            Auxiliar(Renglon, 1) = Auxi1
                            Auxiliar(Renglon, 2) = Canti
                            
                            XEnvase(Renglon, 1) = rstPedido!Envase1
                            XEnvase(Renglon, 2) = rstPedido!Canti1
                            XEnvase(Renglon, 3) = rstPedido!Envase2
                            XEnvase(Renglon, 4) = rstPedido!Canti2
                            XEnvase(Renglon, 5) = rstPedido!Envase3
                            XEnvase(Renglon, 6) = rstPedido!Canti3
                            
                            
                    End Select
                                
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
    
    For DA = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(DA, 1)
        Canti = Auxiliar(DA, 2)
        
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
        
    Next DA
    
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

    For a = 0 To 5
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 19
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
                
                    dada = Str$(rstEstadistica!PrecioUs)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Paridad)
                    Paridad.Text = Pusing("###,###.##", dada)
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!PrecioUs)
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
    
    For DA = 1 To XRenglon
    
        Auxi1 = Auxiliar(DA, 1)
                    
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
    Next DA
    
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
            WSeguro = IIf(IsNull(rstCtacte!Seguro), "0", rstCtacte!Seguro)
            WFlete = IIf(IsNull(rstCtacte!Flete), "0", rstCtacte!Flete)
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
                rstCambios.Close
                         Else
                 Paridad.Text = ""
            End If
            Rem Paridad.Text = "1"
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

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        DBGrid1.SetFocus
    End If
End Sub

Sub Impresion()

        Open "LPT1" For Output As #99

        Print #99, Chr$(27) + Chr$(40) + "19U";
        Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
        Print #99, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72)
        Print #99, Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72)

        For XX = 1 To 1

        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(18); Left$(fecorden.Text, 2);
        Print #99, Tab(21); Mid$(fecorden.Text, 4, 2);
        Print #99, Tab(24); Right$(fecorden.Text, 2);
        Print #99, Tab(27); Left$(NroOrden.Text, 6);
        Print #99, Tab(37); Consignatario.Text;
        Print #99, Tab(68); Left$(Fecha.Text, 2);
        Print #99, Tab(71); Mid$(Fecha.Text, 4, 2);
        Print #99, Tab(74); Right$(Fecha.Text, 2)
        Print #99, ""

        Print #99, Tab(45); Envio1.Text
        Print #99, Tab(3); Left$(WRazon, 40);
        Print #99, Tab(45); Envio2.Text

        Print #99, ""
        Print #99, Tab(3); Left$(WDireccion, 40);
        Print #99, Tab(45); Pago1.Text
        Print #99, Tab(3); Left$(WLocalidad, 40);
        Print #99, Tab(45); Pago2.Text
        Print #99, ""
        Print #99, ""
        Print #99, Tab(85); "USD"
        Print #99, ""
        Print #99, ""

        Suma1 = 0
        Suma2 = 0
        Suma3 = 0
        Erase WImpresion
        WRenglon = 0
        
        Impre = 0
        
        For a = 0 To 5
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                XProducto = DBGrid1.Text
                
                DBGrid1.Col = 1
                XDescri = DBGrid1.Text
                
                DBGrid1.Col = 3
                XPrecio = DBGrid1.Text
            
                DBGrid1.Col = 4
                XCantidad = DBGrid1.Text
                
                DBGrid1.Col = 5
                XCantidad1 = DBGrid1.Text
                
                DBGrid1.Col = 6
                XTipo = DBGrid1.Text
                
                DBGrid1.Col = 7
                XNumero = DBGrid1.Text
                
                DBGrid1.Col = 8
                XBruto = DBGrid1.Text
                
                WRenglon = WRenglon + 1
                
                WImpresion(WRenglon, 1) = XProducto
                WImpresion(WRenglon, 2) = XDescri
                WImpresion(WRenglon, 3) = ""
                WImpresion(WRenglon, 4) = XPrecio
                WImpresion(WRenglon, 5) = XCantidad
                WImpresion(WRenglon, 6) = XCantidad1
                WImpresion(WRenglon, 7) = XTipo
                WImpresion(WRenglon, 8) = XNumero
                WImpresion(WRenglon, 9) = XBruto
                    
            Next iRow
            
        Next a
        
        XPasa = 0
        
        
        For a = 1 To 99
        
                Producto = WImpresion(a, 1)
                Descri = WImpresion(a, 2)
                Precio = Val(Alinea("##,###.##", WImpresion(a, 4)))
                Cantidad = Val(WImpresion(a, 5))
                Cantidad1 = Val(WImpresion(a, 6))
                WTipo = WImpresion(a, 7)
                WNumero = WImpresion(a, 8)
                Bruto = Val(WImpresion(a, 9))
                    
                If Cantidad <> 0 Then
                
                        If WNumero = WImpresion(a + 1, 8) And XPasa = 0 Then
                        
                            Print #99, Tab(2); Alinea("###", Str$(1));
                            Print #99, Tab(8); WTipo;
                            Print #99, Tab(12); WNumero;
                            Print #99, Tab(22); "Palet conteniendo : ";
                            Print #99, Tab(60); Alinea("#####.##", Str$(24));
                            Suma2 = Suma2 + 24
                            
                            XPasa = 1
                            
                            XCanti = XEnvase(a, 2)
                            Call Ceros(XCanti, 2)
                            spEnvase = "ConsultaEnvases " + "'" + XEnvase(a, 1) + "'"
                            Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                            If rstEnvase.RecordCount > 0 Then
                                WEnva = "(" + XCanti + "X" + Left$(rstEnvase!Descripcion, 10) + ")"
                                rstEnvase.Close
                                    Else
                                WEnva = ""
                            End If
                            
                            Print #99, Tab(22); Left$(Descri, 21); " "; WEnva;
                            Print #99, Tab(60); Alinea("#####.##", Str$(Bruto));
                            Print #99, Tab(68); Alinea("#####.#", Str$(Cantidad));
                            Print #99, Tab(75); Alinea("###.##", Str$(Precio));
                            Print #99, Tab(83); Alinea("###,###.##", Str$(Cantidad * Precio))
                            
                            Suma1 = Suma1 + Cantidad1
                            Suma2 = Suma2 + Bruto
                            Suma3 = Suma3 + Cantidad
                            
                                    Else
                                    
                            If WNumero <> WImpresion(a + 1, 8) Then
                                XPasa = 0
                            End If
                                
                            If WNumero = WImpresion(a - 1, 8) Then
                            
                                XCanti = XEnvase(a, 2)
                                Call Ceros(XCanti, 2)
                                spEnvase = "ConsultaEnvases " + "'" + XEnvase(a, 1) + "'"
                                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                If rstEnvase.RecordCount > 0 Then
                                    WEnva = "(" + XCanti + "X" + Left$(rstEnvase!Descripcion, 10) + ")"
                                    rstEnvase.Close
                                        Else
                                    WEnva = ""
                                End If
                            
                                Print #99, Tab(22); Left$(Descri, 21); " "; WEnva;
                                Print #99, Tab(60); Alinea("#####.##", Str$(Bruto));
                                Print #99, Tab(68); Alinea("#####.#", Str$(Cantidad));
                                Print #99, Tab(75); Alinea("###.##", Str$(Precio));
                                Print #99, Tab(83); Alinea("###,###.##", Str$(Cantidad * Precio))
                                    Else
                                Print #99, Tab(2); Alinea("###", Str$(Cantidad1));
                                Print #99, Tab(8); WTipo;
                                Print #99, Tab(12); WNumero;
                                Print #99, Tab(22); Left$(Descri, 37);
                                Print #99, Tab(60); Alinea("#####.##", Str$(Bruto));
                                Print #99, Tab(68); Alinea("#####.#", Str$(Cantidad));
                                Print #99, Tab(75); Alinea("###.##", Str$(Precio));
                                Print #99, Tab(83); Alinea("###,###.##", Str$(Cantidad * Precio))
                            End If
                                
                            Suma1 = Suma1 + Cantidad1
                            Suma2 = Suma2 + Bruto
                            Suma3 = Suma3 + Cantidad
                            
                        End If
                
                        Impre = Impre + 1
                    
                End If
            
        Next a
        
        
        
        For DA = Impre To 21
            Print #99, ""
        Next DA

        Print #99, Tab(5); "Todas las disputas que puedan surgir en el presente contrato seran finalmente arregladas"
        Print #99, Tab(5); "de acuerdo a las Reglas de Conciliacion y Arbitraje  de  la  Camara  Internacional   de"
        Print #99, Tab(5); "Comercio por uno o mas arbitros de acuerdo de dichas reglas"
        Print #99, Tab(5); "INCOTERMS 1990";
        
        Call Numtolet
        
        WTexto1 = UCase(WTexto1)
        WTexto2 = UCase(WTexto2)

        Print #99, Tab(22); "Son Dolares estadounidenses"
        Print #99, Tab(20); WTexto1

        Print #99, Tab(2); Alinea("###", Str$(Suma1));
        Print #99, Tab(20); WTexto2;
        Print #99, Tab(60); Alinea("#####.#", Str$(Suma2));
        Print #99, Tab(68); Alinea("#####", Str$(Suma3));
        Print #99, Tab(83); Alinea("###,###.##", Neto.Caption)
        Print #99, ""
        Print #99, Tab(83); Alinea("###,###.##", Neto.Caption)
        Print #99, ""
        If WSeguro <> 0 Then
                Print #99, Tab(83); Alinea("###,###.##", Str$(WSeguro))
                        Else
                Print #99, ""
        End If
        Print #99, ""
        If WFlete <> 0 Then
                Print #99, Tab(83); Alinea("###,###.##", Str$(WFlete))
                        Else
                Print #99, ""
        End If
        Print #99, Tab(6); Marca.Text
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(60); Cip.Text;
        Print #99, Tab(83); Alinea("###,###.##", Total.Caption)

        Next XX
        
        Close #99
End Sub
        

Sub Impresion_Remito()

        Open "lpt1" For Output As #1

        For FF = 1 To 2


        Print #1, Chr$(27) + Chr$(40) + "19U";
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
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

        For a = 0 To 5
        
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
        Print #1, Tab(10); "Lugar de Pago : Paraguay 1359 Piso 2 Capital Federal"
        Print #1, ""

        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
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

Private Sub Marca_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envio1.SetFocus
    End If
End Sub
Private Sub Envio1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envio2.SetFocus
    End If
End Sub
Private Sub Envio2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pago1.SetFocus
    End If
End Sub
Private Sub Pago1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Pago2.SetFocus
    End If
End Sub
Private Sub Pago2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroOrden.SetFocus
    End If
End Sub
Private Sub NroOrden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fecorden.SetFocus
    End If
End Sub
Private Sub Fecorden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Consignatario.SetFocus
    End If
End Sub
Private Sub Consignatario_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cip.SetFocus
    End If
End Sub
Private Sub Cip_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Marca.SetFocus
    End If
End Sub

Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = "dolares"
    sCentimos = "centavos"
    
    Numero = CStr(Val(Total.Caption))
    
    WTexto1 = Numero2Letra(Numero, , sMoneda & " ", sCentimos & " ")
    WTexto1 = WTexto1 + Space$(50)
    
    Pasa = 0
    
    For DA = 40 To 1 Step -1
        If Mid$(WTexto1, DA, 1) = Space$(1) Then
            Pasa = 1
        End If
        If Pasa = 1 Then
            If Mid$(WTexto1, DA, 1) <> Space$(1) Then
                Exit For
            End If
        End If
    Next DA
    
    WTexto2 = Mid$(WTexto1, DA + 2, 35)
    WTexto1 = Left$(WTexto1, DA)
    
End Sub

