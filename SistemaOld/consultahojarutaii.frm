VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaHojaRutaII 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Hoja de Ruta"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   Begin VB.ListBox WPantalla 
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
      Height          =   2595
      ItemData        =   "consultahojarutaii.frx":0000
      Left            =   2880
      List            =   "consultahojarutaii.frx":0007
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   10186
      _Version        =   327680
      Rows            =   1000
      Cols            =   10
      BackColor       =   16777152
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   5400
      MouseIcon       =   "consultahojarutaii.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "consultahojarutaii.frx":031F
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgConsultaHojaRutaII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstHojaRuta As Recordset
Dim spHojaRuta As String
Dim XParam As String

Dim ZVector(100, 10) As String
Dim ZAyuda(100) As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub cmdClose1_Click()
    PrgConsultaHojaRutaII.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 900
    Muestra.ColWidth(2) = 900
    Muestra.ColWidth(3) = 900
    Muestra.ColWidth(4) = 900
    Muestra.ColWidth(5) = 2400
    Muestra.ColWidth(6) = 1400
    Muestra.ColWidth(7) = 2400
    Muestra.ColWidth(8) = 900
    Muestra.ColWidth(9) = 10
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Factura"
    
    Muestra.Col = 3
    Muestra.Text = "Remito"
    
    Muestra.Col = 4
    Muestra.Text = "Pedido"
    
    Muestra.Col = 5
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 6
    Muestra.Text = "Producto"
    
    Muestra.Col = 7
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 8
    Muestra.Text = "Kilos"
    
    Muestra.Col = 9
    Muestra.Text = ""
    
    Fecha.Text = "  /  /    "
    
End Sub

Private Sub Proceso_Click()
    
    Call Limpia_Vector

    WRenglon = 0
    WRenglonII = 0
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFecha = WAno + WMes + WDia
    
    Sql1 = "Select *"
    Sql2 = " FROM HojaRuta"
    Sql3 = " Where HojaRuta.Fecha = " + "'" + Fecha.Text + "'"
    Sql4 = " Order by HojaRuta.Clave"
    rsHojaRuta = Sql1 + Sql2 + Sql3 + Sql4
    Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    If rstHojaRuta.RecordCount > 0 Then
        With rstHojaRuta
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    ZVector(WRenglon, 1) = rstHojaRuta!Pedido
                    ZVector(WRenglon, 2) = rstHojaRuta!Cliente
                    ZVector(WRenglon, 3) = rstHojaRuta!Razon
                    ZVector(WRenglon, 4) = rstHojaRuta!Remito
                    ZVector(WRenglon, 5) = rstHojaRuta!Hoja
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHojaRuta.Close
    End If
    
    For Ciclo = 1 To WRenglon
    
        ZZPedido = ZVector(Ciclo, 1)
        ZZCliente = ZVector(Ciclo, 2)
        ZZRazon = ZVector(Ciclo, 3)
        ZZRemito = ZVector(Ciclo, 4)
        ZZHoja = ZVector(Ciclo, 5)
        ZZFactura = 0
        
        ZSql = ""
        ZSql = ZSql + "Select Remito, TipoPedido"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZZPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZZRemito = IIf(IsNull(rstPedido!Remito), "", rstPedido!Remito)
            ZZTipoPedido = IIf(IsNull(rstPedido!TipoPedido), "0", rstPedido!TipoPedido)
            rstPedido.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select Pedido, Remito, Numero"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Pedido = " + "'" + Trim(ZZPedido) + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            With rstCtacte
                .MoveFirst
                Do
                    If .EOF = False Then
                                
                        If Val(ZZRemito) = Val(rstCtacte!Remito) Then
                            ZZFactura = rstCtacte!Numero
                        End If
                                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCtacte.Close
        End If
        
        If ZZFactura <> 0 Then
    
            ZSql = ""
            ZSql = ZSql + "Select Numero, Cliente, Articulo, Cantidad"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Numero = " + "'" + Str$(ZZFactura) + "'"
            ZSql = ZSql + " and Estadistica.Cliente = " + "'" + ZZCliente + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            WRenglonII = WRenglonII + 1
                        
                            Muestra.TextMatrix(WRenglonII, 1) = ZZHoja
                            Muestra.TextMatrix(WRenglonII, 2) = ZZFactura
                            Muestra.TextMatrix(WRenglonII, 3) = ZZRemito
                            Muestra.TextMatrix(WRenglonII, 4) = ZZPedido
                            Muestra.TextMatrix(WRenglonII, 5) = ZZRazon
                            Muestra.TextMatrix(WRenglonII, 6) = rstEstadistica!Articulo
                            Muestra.TextMatrix(WRenglonII, 7) = ""
                            Muestra.TextMatrix(WRenglonII, 8) = Str$(rstEstadistica!Cantidad)
                            Muestra.TextMatrix(WRenglonII, 8) = Pusing("###,###", Muestra.TextMatrix(WRenglonII, 8))
                            Muestra.TextMatrix(WRenglonII, 9) = ZZCliente
                                        
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEstadistica.Close
            End If
            
                Else
    
            ZSql = ""
            ZSql = ZSql + "Select Pedido, CantiLote1, CantiLote2, CantiLote3, CantiLote4, CantiLote5, Cantidad, Terminado, CantidadFac"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZZPedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            WRenglonII = WRenglonII + 1
                            
                            ZCantidad1 = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                            ZCantidad2 = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                            ZCantidad3 = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                            ZCantidad4 = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                            ZCantidad5 = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                            ZCantidadFac = IIf(IsNull(rstPedido!CantidadFac), "0", rstPedido!CantidadFac)
                            ZSumaCantidad = ZCantidad1 + ZCantidad2 + ZCantidad3 + ZCantidad4 + ZCantidad5
                                    
                            If ZSumaCantidad = 0 Then
                                ZSumaCantidad = ZCantidadFac
                            End If
                                    
                            If ZSumaCantidad <> 0 Then
                                ZKilos = ZSumaCantidad
                                    Else
                                ZKilos = rstPedido!Cantidad
                            End If
                        
                            Muestra.TextMatrix(WRenglonII, 1) = ZZHoja
                            Muestra.TextMatrix(WRenglonII, 2) = ""
                            Muestra.TextMatrix(WRenglonII, 3) = ZZRemito
                            Muestra.TextMatrix(WRenglonII, 4) = ZZPedido
                            Muestra.TextMatrix(WRenglonII, 5) = ZZRazon
                            Muestra.TextMatrix(WRenglonII, 6) = rstPedido!Terminado
                            Muestra.TextMatrix(WRenglonII, 7) = ""
                            Muestra.TextMatrix(WRenglonII, 8) = Str$(ZKilos)
                            Muestra.TextMatrix(WRenglonII, 8) = Pusing("###,###", Muestra.TextMatrix(WRenglonII, 8))
                            Muestra.TextMatrix(WRenglonII, 9) = ZZCliente
                                        
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPedido.Close
            End If
            
        End If
    
    Next Ciclo
        
        
    
    For Ciclo = 1 To WRenglonII
    
        Cliente = Muestra.TextMatrix(Ciclo, 9)
        Terminado = Muestra.TextMatrix(Ciclo, 6)
        If Left$(Terminado, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Muestra.TextMatrix(Ciclo, 7) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Muestra.TextMatrix(Ciclo, 7) = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                
        End Select
        
    Next Ciclo
    
End Sub

Private Sub Limpia_Vector()

    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Factura"
    
    Muestra.Col = 3
    Muestra.Text = "Remito"
    
    Muestra.Col = 4
    Muestra.Text = "Pedido"
    
    Muestra.Col = 5
    Muestra.Text = "Razon Social"
    
    Muestra.Col = 6
    Muestra.Text = "Producto"
    
    Muestra.Col = 7
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 8
    Muestra.Text = "Kilos"
    
    Muestra.Col = 9
    Muestra.Text = ""
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Call Proceso_Click
        End If
    End If
End Sub

Private Sub Muestra_Click()

    If Muestra.Col = 5 Then

    WPantalla.Clear
    WPantalla.AddItem ""
    ZLugar = 0
    ZAyuda(ZLugar) = ""
    
    Sql1 = "Select DISTINCT Razon, Cliente"
    Sql2 = " FROM HojaRuta"
    Sql3 = " Where HojaRuta.Fecha = " + "'" + Fecha.Text + "'"
    Sql4 = " Order by HojaRuta.Razon"
    spHojaRuta = Sql1 + Sql2 + Sql3 + Sql4
    Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    With rstHojaRuta
        .MoveFirst
        Do
            If .EOF = False Then
                WPantalla.AddItem rstHojaRuta!Razon
                ZLugar = ZLugar + 1
                ZAyuda(ZLugar) = rstHojaRuta!Cliente
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstHojaRuta.Close
    
    WPantalla.Visible = True
    
    End If



End Sub

Private Sub WPantalla_Click()
    
    WPantalla.Visible = False

    ClienteSeleccion = ZAyuda(WPantalla.ListIndex)
    
    Call Limpia_Vector

    WRenglon = 0
    WRenglonII = 0
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFecha = WAno + WMes + WDia
    
    Sql1 = "Select *"
    Sql2 = " FROM HojaRuta"
    Sql3 = " Where HojaRuta.Fecha = " + "'" + Fecha.Text + "'"
    Sql4 = " Order by HojaRuta.Clave"
    rsHojaRuta = Sql1 + Sql2 + Sql3 + Sql4
    Set rstHojaRuta = db.OpenRecordset(rsHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
    If rstHojaRuta.RecordCount > 0 Then
        With rstHojaRuta
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ClienteSeleccion = "" Or ClienteSeleccion = rstHojaRuta!Cliente Then
                
                        WRenglon = WRenglon + 1
                        ZVector(WRenglon, 1) = rstHojaRuta!Pedido
                        ZVector(WRenglon, 2) = rstHojaRuta!Cliente
                        ZVector(WRenglon, 3) = rstHojaRuta!Razon
                        ZVector(WRenglon, 4) = rstHojaRuta!Remito
                        ZVector(WRenglon, 5) = rstHojaRuta!Hoja
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHojaRuta.Close
    End If
    
    For Ciclo = 1 To WRenglon
    
        ZZPedido = ZVector(Ciclo, 1)
        ZZCliente = ZVector(Ciclo, 2)
        ZZRazon = ZVector(Ciclo, 3)
        ZZRemito = ZVector(Ciclo, 4)
        ZZHoja = ZVector(Ciclo, 5)
        ZZFactura = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZZPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZZRemito = IIf(IsNull(rstPedido!Remito), "", rstPedido!Remito)
            rstPedido.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Pedido = " + "'" + Trim(ZZPedido) + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            With rstCtacte
                .MoveFirst
                Do
                    If .EOF = False Then
                                
                        If Val(ZZRemito) = Val(rstCtacte!Remito) Then
                            ZZFactura = rstCtacte!Numero
                        End If
                                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCtacte.Close
        End If
    
        If ZZFactura <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Numero = " + "'" + Str$(ZZFactura) + "'"
            ZSql = ZSql + " and Estadistica.Cliente = " + "'" + ZZCliente + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            WRenglonII = WRenglonII + 1
                        
                            Muestra.TextMatrix(WRenglonII, 1) = ZZHoja
                            Muestra.TextMatrix(WRenglonII, 2) = ZZFactura
                            Muestra.TextMatrix(WRenglonII, 3) = ZZRemito
                            Muestra.TextMatrix(WRenglonII, 4) = ZZPedido
                            Muestra.TextMatrix(WRenglonII, 5) = ZZRazon
                            Muestra.TextMatrix(WRenglonII, 6) = rstEstadistica!Articulo
                            Muestra.TextMatrix(WRenglonII, 7) = ""
                            Muestra.TextMatrix(WRenglonII, 8) = Str$(rstEstadistica!Cantidad)
                            Muestra.TextMatrix(WRenglonII, 8) = Pusing("###,###", Muestra.TextMatrix(WRenglonII, 8))
                            Muestra.TextMatrix(WRenglonII, 9) = ZZCliente
                                        
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEstadistica.Close
            End If
            
            
                Else
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Pedido = " + "'" + ZZPedido + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            WRenglonII = WRenglonII + 1
                            
                            ZCantidad1 = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                            ZCantidad2 = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                            ZCantidad3 = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                            ZCantidad4 = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                            ZCantidad5 = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                            ZCantidadFac = IIf(IsNull(rstPedido!CantidadFac), "0", rstPedido!CantidadFac)
                            ZSumaCantidad = ZCantidad1 + ZCantidad2 + ZCantidad3 + ZCantidad4 + ZCantidad5
                                    
                            If ZSumaCantidad = 0 Then
                                ZSumaCantidad = ZCantidadFac
                            End If
                            
                            If ZSumaCantidad <> 0 Then
                                ZKilos = ZSumaCantidad
                                    Else
                                ZKilos = rstPedido!Cantidad
                            End If
                        
                            Muestra.TextMatrix(WRenglonII, 1) = ZZHoja
                            Muestra.TextMatrix(WRenglonII, 2) = ""
                            Muestra.TextMatrix(WRenglonII, 3) = ZZRemito
                            Muestra.TextMatrix(WRenglonII, 4) = ZZPedido
                            Muestra.TextMatrix(WRenglonII, 5) = ZZRazon
                            Muestra.TextMatrix(WRenglonII, 6) = rstPedido!Terminado
                            Muestra.TextMatrix(WRenglonII, 7) = ""
                            Muestra.TextMatrix(WRenglonII, 8) = Str$(ZKilos)
                            Muestra.TextMatrix(WRenglonII, 8) = Pusing("###,###", Muestra.TextMatrix(WRenglonII, 8))
                            Muestra.TextMatrix(WRenglonII, 9) = ZZCliente
                                        
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPedido.Close
            End If
            
        End If
    
    Next Ciclo
    
    For Ciclo = 1 To WRenglonII
    
        Cliente = Muestra.TextMatrix(Ciclo, 9)
        Terminado = Muestra.TextMatrix(Ciclo, 6)
        If Left$(Terminado, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Muestra.TextMatrix(Ciclo, 7) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Muestra.TextMatrix(Ciclo, 7) = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                
        End Select
        
    Next Ciclo
    
    Muestra.Col = 1
    Muestra.Row = 1
    Muestra.TopRow = 1
    
End Sub
