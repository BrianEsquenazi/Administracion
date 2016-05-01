VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgTrazabilidad 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trazabilidad de Factura de Exportacion"
   ClientHeight    =   8310
   ClientLeft      =   1125
   ClientTop       =   420
   ClientWidth     =   10020
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   10020
   Visible         =   0   'False
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   8280
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Procesar 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   8280
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   8280
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Cliente 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   7
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3720
      TabIndex        =   5
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
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      Height          =   570
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7680
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreRemito.rpt"
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6855
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12091
      _Version        =   327680
      Rows            =   1000
      Cols            =   7
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Factura"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "PrgTrazabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private WGastos As Double
Private WTexto1 As String
Private WTexto2 As String
Private Auxiliar(100, 50) As String

Dim ZZControlLote(100, 60) As String
Dim ControlLote(12, 2) As String
Dim ControlEnvase(12, 2) As String

Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
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
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstImpreRemito As Recordset
Dim spImpreRemio As String

Dim XParam As String
Dim WLote(12, 2) As String
Dim WImpresion(100, 10) As String
Dim XEnvase(100, 6) As String
Dim XCanti As String
Private WTipoPedido As String

Dim VectorCosto(100, 3) As String
Dim ZZZProducto As String
Dim ZZZCosto As Double

Dim ZZClave As String
Dim ZZNumero As String
Dim ZZRenglon As String
Dim ZZFecha As String
Dim ZZNombre As String
Dim ZZDireccion As String
Dim ZZLocalidad As String
Dim ZZPedido As String
Dim ZZCliente As String
Dim ZZOrden As String
Dim ZZDescripcion As String
Dim ZZCantidad As String
Dim ZZRemito As String

Dim ZZVector(100, 10) As String
Dim ZZImpre(100, 10) As String
Dim ZZCampo1 As String
Dim ZZCampo2 As String
Dim ZLote6 As Double
Dim ZLote7 As Double
Dim ZLote8 As Double
Dim ZLote9 As Double
Dim ZLote10 As Double
Dim ZLote11 As Double
Dim ZLote12 As Double

Dim ZZComprobante As Integer
Dim ZZCuit As String
Dim ZZPais As String
Dim ZZCuitII As String
Dim ZZRazon As String
Dim ZZDomicilio As String

Dim ZZLote(100, 2) As String
Dim Empe(12, 10) As String
Dim ZZHoja(1000, 25) As String
Dim ZZTrabajo(1000, 10) As String
Dim ZZLugar As Integer
Dim ZZLugarII As Integer
Dim ZZPorce As Double
Dim ZZCanti As Double

Dim ZZGrabaFactura As String

Private Sub cmdClose_Click()
    PrgTrazabilidad.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Procesar_Click()

    Erase ZZTrabajo
    Erase ZZHoja
    ZZLugar = 0
    ZZLugarII = 0
    
    OPEN_FILE_Trazabilidad
        

        
    For Ciclo = 1 To 99
        
        XXClave = WVector1.TextMatrix(Ciclo, 4)
        XXCodigo = WVector1.TextMatrix(Ciclo, 1)
        XXCantidadTotal = Val(WVector1.TextMatrix(Ciclo, 3))
        
        If XXCantidadTotal <> 0 Then
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Clave = " + "'" + XXClave + "'"
            ZSql = ZSql + " Order by Estadistica.Clave"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
            
                Erase ZZLote
            
                ZZLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                ZZLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                
                ZZLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                ZZLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "0", rstEstadistica!lote2)
                
                ZZLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                ZZLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "0", rstEstadistica!lote3)
                
                ZZLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                ZZLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "0", rstEstadistica!lote4)
                
                ZZLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                ZZLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "0", rstEstadistica!lote5)
                        
                WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                
                If Len(Trim(WLoteAdicional)) = 98 Then
                
                    ZZLote(6, 2) = Val(Mid$(WLoteAdicional, 1, 8))
                    ZZLote(6, 1) = Val(Mid$(WLoteAdicional, 9, 6))
                    
                    ZZLote(7, 2) = Val(Mid$(WLoteAdicional, 15, 8))
                    ZZLote(7, 1) = Val(Mid$(WLoteAdicional, 23, 6))
                    
                    ZZLote(8, 2) = Val(Mid$(WLoteAdicional, 29, 8))
                    ZZLote(8, 1) = Val(Mid$(WLoteAdicional, 37, 6))
                    
                    ZZLote(9, 2) = Val(Mid$(WLoteAdicional, 43, 8))
                    ZZLote(9, 1) = Val(Mid$(WLoteAdicional, 51, 6))
                    
                    ZZLote(10, 2) = Val(Mid$(WLoteAdicional, 57, 8))
                    ZZLote(10, 1) = Val(Mid$(WLoteAdicional, 65, 6))
                    
                    ZZLote(11, 1) = Val(Mid$(WLoteAdicional, 71, 8))
                    ZZLote(11, 2) = Val(Mid$(WLoteAdicional, 79, 6))
                    
                    ZZLote(12, 1) = Val(Mid$(WLoteAdicional, 85, 8))
                    ZZLote(12, 2) = Val(Mid$(WLoteAdicional, 93, 6))
                    
                End If
                
                For ZZCiclo = 1 To 12
                    If ZZLote(ZZCiclo, 2) <> 0 Then
                        ZZLugar = ZZLugar + 1
                        ZZTrabajo(ZZLugar, 1) = XXCodigo
                        ZZTrabajo(ZZLugar, 2) = ZZLote(ZZCiclo, 1)
                        ZZTrabajo(ZZLugar, 3) = ZZLote(ZZCiclo, 2)
                        ZZTrabajo(ZZLugar, 4) = XXCodigo
                        ZZTrabajo(ZZLugar, 5) = Str$(rstEstadistica!Numero) + "-" + Str$(rstEstadistica!Renglon)
                        ZZTrabajo(ZZLugar, 6) = ZZLote(ZZCiclo, 1)
                        ZZTrabajo(ZZLugar, 7) = "100"
                        ZZTrabajo(ZZLugar, 8) = ZZLote(ZZCiclo, 2)
                        ZZTrabajo(ZZLugar, 9) = XXCantidadTotal
                        ZZTrabajo(ZZLugar, 10) = XyCodigo
                    End If
                Next ZZCiclo
                
                rstEstadistica.Close
                
            End If
       End If
       
    Next Ciclo
    

    
    If Val(WEmpresa) = 1 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        Hasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        Hasta = 4
    End If
    
    
    XEmpresa = WEmpresa
    
    For Ciclo = 1 To 1000
       
        XXCodigo = ZZTrabajo(Ciclo, 1)
        XXLote = ZZTrabajo(Ciclo, 2)
        XXCantidad = ZZTrabajo(Ciclo, 3)
        XXCodigoOriginal = ZZTrabajo(Ciclo, 4)
        XXClaveOriginal = ZZTrabajo(Ciclo, 5)
        XXLoteOriginal = ZZTrabajo(Ciclo, 6)
        XXPorce = ZZTrabajo(Ciclo, 7)
        XXCantidadOriginal = ZZTrabajo(Ciclo, 8)
        XXCantidadTotal = ZZTrabajo(Ciclo, 9)
        
        If Val(XXCantidad) <> 0 Then
        
            For CiclaEmpresa = 1 To Hasta
        
                WEmpresa = Empe(CiclaEmpresa, 1)
                txtOdbc = Empe(CiclaEmpresa, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Hoja"
                ZSql = ZSql + " Where Hoja.Hoja = " + "'" + XXLote + "'"
                ZSql = ZSql + " and Hoja.Producto = " + "'" + XXCodigo + "'"
                ZSql = ZSql + " Order by clave"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    
                    With rstHoja
                    
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                                
                                If rstHoja!Canti1 <> 0 Then
                                    If rstHoja!realant <> 0 Then
                                        ZZReal = rstHoja!realant
                                            Else
                                        ZZReal = rstHoja!Real
                                    End If
                                    ZZPorce = (Val(XXCantidad) / ZZReal)
                                    ZZCanti = rstHoja!Canti1 * ZZPorce
                                    Call Redondeo(ZZCanti)
                                    If rstHoja!Tipo = "M" Then
                                        ZZLugarII = ZZLugarII + 1
                                        ZZHoja(ZZLugarII, 1) = rstHoja!Articulo
                                        ZZHoja(ZZLugarII, 2) = Str$(rstHoja!Cantidad)
                                        ZZHoja(ZZLugarII, 3) = Str$(rstHoja!lote1)
                                        ZZHoja(ZZLugarII, 4) = Str$(ZZCanti)
                                        ZZHoja(ZZLugarII, 5) = XXCodigo
                                        ZZHoja(ZZLugarII, 6) = XXLote
                                        ZZHoja(ZZLugarII, 7) = XXCodigoOriginal
                                        ZZHoja(ZZLugarII, 8) = XXClaveOriginal
                                        ZZHoja(ZZLugarII, 9) = XXLoteOriginal
                                        ZZHoja(ZZLugarII, 20) = XXCantidadOriginal
                                        ZZHoja(ZZLugarII, 22) = XXCantidadTotal
                                            Else
                                        ZZLugar = ZZLugar + 1
                                        ZZTrabajo(ZZLugar, 1) = rstHoja!Terminado
                                        ZZTrabajo(ZZLugar, 2) = Str$(rstHoja!lote1)
                                        ZZTrabajo(ZZLugar, 3) = Str$(ZZCanti)
                                        ZZTrabajo(ZZLugar, 4) = XXCodigoOriginal
                                        ZZTrabajo(ZZLugar, 5) = XXClaveOriginal
                                        ZZTrabajo(ZZLugar, 6) = XXLoteOriginal
                                        ZZTrabajo(ZZLugar, 8) = XXCantidadOriginal
                                        ZZTrabajo(ZZLugar, 9) = XXCantidadTotal
                                    End If
                                End If
                                
                                If rstHoja!Canti2 <> 0 Then
                                    If rstHoja!realant <> 0 Then
                                        ZZReal = rstHoja!realant
                                            Else
                                        ZZReal = rstHoja!Real
                                    End If
                                    ZZPorce = (Val(XXCantidad) / ZZReal)
                                    ZZCanti = rstHoja!Canti2 * ZZPorce
                                    Call Redondeo(ZZCanti)
                                    If rstHoja!Tipo = "M" Then
                                        ZZLugarII = ZZLugarII + 1
                                        ZZHoja(ZZLugarII, 1) = rstHoja!Articulo
                                        ZZHoja(ZZLugarII, 2) = Str$(rstHoja!Cantidad)
                                        ZZHoja(ZZLugarII, 3) = Str$(rstHoja!lote2)
                                        ZZHoja(ZZLugarII, 4) = Str$(ZZCanti)
                                        ZZHoja(ZZLugarII, 5) = XXCodigo
                                        ZZHoja(ZZLugarII, 6) = XXLote
                                        ZZHoja(ZZLugarII, 7) = XXCodigoOriginal
                                        ZZHoja(ZZLugarII, 8) = XXClaveOriginal
                                        ZZHoja(ZZLugarII, 9) = XXLoteOriginal
                                        ZZHoja(ZZLugarII, 20) = XXCantidadOriginal
                                        ZZHoja(ZZLugarII, 22) = XXCantidadTotal
                                            Else
                                        ZZLugar = ZZLugar + 1
                                        ZZTrabajo(ZZLugar, 1) = rstHoja!Terminado
                                        ZZTrabajo(ZZLugar, 2) = Str$(rstHoja!lote2)
                                        ZZTrabajo(ZZLugar, 3) = Str$(ZZCanti)
                                        ZZTrabajo(ZZLugar, 4) = XXCodigoOriginal
                                        ZZTrabajo(ZZLugar, 5) = XXClaveOriginal
                                        ZZTrabajo(ZZLugar, 6) = XXLoteOriginal
                                        ZZTrabajo(ZZLugar, 8) = XXCantidadOriginal
                                        ZZTrabajo(ZZLugar, 9) = XXCantidadTotal
                                    End If
                                End If
                                
                                If rstHoja!Canti3 <> 0 Then
                                    If rstHoja!realant <> 0 Then
                                        ZZReal = rstHoja!realant
                                            Else
                                        ZZReal = rstHoja!Real
                                    End If
                                    ZZPorce = (Val(XXCantidad) / ZZReal)
                                    ZZCanti = rstHoja!Canti3 * ZZPorce
                                    Call Redondeo(ZZCanti)
                                    If rstHoja!Tipo = "M" Then
                                        ZZLugarII = ZZLugarII + 1
                                        ZZHoja(ZZLugarII, 1) = rstHoja!Articulo
                                        ZZHoja(ZZLugarII, 2) = Str$(rstHoja!Cantidad)
                                        ZZHoja(ZZLugarII, 3) = Str$(rstHoja!lote3)
                                        ZZHoja(ZZLugarII, 4) = Str$(ZZCanti)
                                        ZZHoja(ZZLugarII, 5) = XXCodigo
                                        ZZHoja(ZZLugarII, 6) = XXLote
                                        ZZHoja(ZZLugarII, 7) = XXCodigoOriginal
                                        ZZHoja(ZZLugarII, 8) = XXClaveOriginal
                                        ZZHoja(ZZLugarII, 9) = XXLoteOriginal
                                        ZZHoja(ZZLugarII, 20) = XXCantidadOriginal
                                        ZZHoja(ZZLugarII, 22) = XXCantidadTotal
                                            Else
                                        ZZLugar = ZZLugar + 1
                                        ZZTrabajo(ZZLugar, 1) = rstHoja!Terminado
                                        ZZTrabajo(ZZLugar, 2) = Str$(rstHoja!lote3)
                                        ZZTrabajo(ZZLugar, 3) = Str$(ZCanti)
                                        ZZTrabajo(ZZLugar, 4) = XXCodigoOriginal
                                        ZZTrabajo(ZZLugar, 5) = XXClaveOriginal
                                        ZZTrabajo(ZZLugar, 6) = XXLoteOriginal
                                        ZZTrabajo(ZZLugar, 8) = XXCantidadOriginal
                                        ZZTrabajo(ZZLugar, 9) = XXCantidadTotal
                                    End If
                                End If
                                
                                .MoveNext
                                
                                If .EOF = True Then
                                    Exit Do
                                End If
                                
                            Loop
                        End If
                        
                    End With
                    
                    rstHoja.Close
                    Exit For
                    
                End If
                
            Next CiclaEmpresa
        
        End If
        
    Next Ciclo
        
        
        
    For CiclaEmpresa = 1 To Hasta

        WEmpresa = Empe(CiclaEmpresa, 1)
        txtOdbc = Empe(CiclaEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        For Ciclo = 1 To ZZLugarII
    
            XXArticulo = ZZHoja(Ciclo, 1)
            XXCantidad = ZZHoja(Ciclo, 2)
            XXLote = ZZHoja(Ciclo, 3)
            XXCantiLote = ZZHoja(Ciclo, 4)
            XXTerminado = ZZHoja(Ciclo, 5)
            XXPartida = ZZHoja(Ciclo, 6)
            XXTerminadoOriginal = ZZHoja(Ciclo, 7)
            XXClaveOriginal = ZZHoja(Ciclo, 8)
            XXLoteOriginal = ZZHoja(Ciclo, 9)
            XXProveedor = ZZHoja(Ciclo, 10)
            XXEmpresa = Val(ZZHoja(Ciclo, 11))
            XXOrden = Val(ZZHoja(Ciclo, 12))
            XXCarpeta = ZZHoja(Ciclo, 13)
            XXRemito = Val(ZZHoja(Ciclo, 14))
            XXFactura = Val(ZZHoja(Ciclo, 15))
            XXFechaCompo = ZZHoja(Ciclo, 16)
            XXCosto = Val(ZZHoja(Ciclo, 17))
            XXInforme = Val(ZZHoja(Ciclo, 18))
            XXDesProveedor = ZZHoja(Ciclo, 19)
            XXCantidadOriginal = ZZHoja(Ciclo, 20)
            XXDesArticulo = ZZHoja(Ciclo, 21)
            
            If Trim(XXProveedor) = "" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Laudo = " + "'" + XXLote + "'"
                ZSql = ZSql + " and Laudo.Articulo = " + "'" + XXArticulo + "'"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                
                    XXOrden = rstLaudo!Orden
                    XXInforme = rstLaudo!Informe
                    XXEmpresa = CiclaEmpresa
                    rstLaudo.Close
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Informe"
                    ZSql = ZSql + " Where Informe.Informe = " + "'" + Str$(XXInforme) + "'"
                    ZSql = ZSql + " and Informe.Articulo = " + "'" + XXArticulo + "'"
                    spInforme = ZSql
                    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                    If rstInforme.RecordCount > 0 Then
                        XXRemito = rstInforme!Remito
                        XXProveedor = rstInforme!Proveedor
                        rstInforme.Close
                    End If
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Orden = " + "'" + Str$(XXOrden) + "'"
                    ZSql = ZSql + " and Orden.Articulo = " + "'" + XXArticulo + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        XXCosto = rstOrden!Precio
                        XXCarpeta = rstOrden!Carpeta
                        XXFechaCompo = rstOrden!Fecha
                        rstOrden.Close
                    End If
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Articulo"
                    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + XXArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        XXDesArticulo = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                    
                    ZZHoja(Ciclo, 10) = XXProveedor
                    ZZHoja(Ciclo, 11) = Str$(XXEmpresa)
                    ZZHoja(Ciclo, 12) = Str$(XXOrden)
                    ZZHoja(Ciclo, 13) = XXCarpeta
                    ZZHoja(Ciclo, 14) = Str$(XXRemito)
                    ZZHoja(Ciclo, 16) = XXFechaCompo
                    ZZHoja(Ciclo, 17) = Str$(XXCosto)
                    ZZHoja(Ciclo, 18) = Str$(XXInforme)
                    ZZHoja(Ciclo, 21) = XXDesArticulo
                    
                End If
            
            End If
        
        Next Ciclo
        
    Next CiclaEmpresa
        
    Call Conecta_Empresa
        
    For Ciclo = 1 To ZZLugarII

        XXArticulo = ZZHoja(Ciclo, 1)
        XXCantidad = ZZHoja(Ciclo, 2)
        XXLote = ZZHoja(Ciclo, 3)
        XXCantiLote = ZZHoja(Ciclo, 4)
        XXTerminado = ZZHoja(Ciclo, 5)
        XXPartida = ZZHoja(Ciclo, 6)
        XXTerminadoOriginal = ZZHoja(Ciclo, 7)
        XXClaveOriginal = ZZHoja(Ciclo, 8)
        XXLoteOriginal = ZZHoja(Ciclo, 9)
        XXProveedor = ZZHoja(Ciclo, 10)
        XXEmpresa = Val(ZZHoja(Ciclo, 11))
        XXOrden = Val(ZZHoja(Ciclo, 12))
        XXCarpeta = ZZHoja(Ciclo, 13)
        XXRemito = Val(ZZHoja(Ciclo, 14))
        XXFactura = Val(ZZHoja(Ciclo, 15))
        XXFechaCompo = ZZHoja(Ciclo, 16)
        XXCosto = Val(ZZHoja(Ciclo, 17))
        XXInforme = Val(ZZHoja(Ciclo, 18))
        XXDesProveedor = ZZHoja(Ciclo, 19)
        XXCantidadOriginal = ZZHoja(Ciclo, 20)
        XXDesArticulo = ZZHoja(Ciclo, 21)
        XXCantidadTotal = ZZHoja(Ciclo, 22)
        
        If Trim(XXProveedor) <> "" Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM IvaComp"
            ZSql = ZSql + " Where IvaComp.Remito = " + "'" + Trim(Str$(XXRemito)) + "'"
            ZSql = ZSql + " and IvaComp.Proveedor = " + "'" + XXProveedor + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                ZZHoja(Ciclo, 15) = rstIvaComp!Numero
                ZZHoja(Ciclo, 16) = rstIvaComp!Fecha
                rstIvaComp.Close
                    Else
                Rem Stop
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + XXProveedor + "'"
            spProveedor = ZSql
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                ZZHoja(Ciclo, 19) = RstProveedor!Nombre
                RstProveedor.Close
            End If
        
        End If
        
    Next Ciclo
        
        
        
        
        
        
    
    With rstTrazabilidad
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
    
    For Ciclo = 1 To ZZLugarII
    
        XXArticulo = ZZHoja(Ciclo, 1)
        XXCantidad = ZZHoja(Ciclo, 2)
        XXLote = ZZHoja(Ciclo, 3)
        XXCantiLote = ZZHoja(Ciclo, 4)
        XXTerminado = ZZHoja(Ciclo, 5)
        XXPartida = ZZHoja(Ciclo, 6)
        XXTerminadoOriginal = ZZHoja(Ciclo, 7)
        XXClaveOriginal = ZZHoja(Ciclo, 8)
        XXLoteOriginal = ZZHoja(Ciclo, 9)
        XXProveedor = ZZHoja(Ciclo, 10)
        XXEmpresa = Val(ZZHoja(Ciclo, 11))
        XXOrden = Val(ZZHoja(Ciclo, 12))
        XXCarpeta = ZZHoja(Ciclo, 13)
        XXRemito = Val(ZZHoja(Ciclo, 14))
        XXFactura = Val(ZZHoja(Ciclo, 15))
        XXFechaCompo = ZZHoja(Ciclo, 16)
        XXCosto = Val(ZZHoja(Ciclo, 17))
        XXInforme = Val(ZZHoja(Ciclo, 18))
        XXDesProveedor = ZZHoja(Ciclo, 19)
        XXCantidadOriginal = ZZHoja(Ciclo, 20)
        XXDesArticulo = ZZHoja(Ciclo, 21)
        XXCantidadTotal = ZZHoja(Ciclo, 22)
        
        Auxi1 = XXLote
        Call Ceros(Auxi1, 8)
        
        Auxi2 = XXLoteOriginal
        Call Ceros(Auxi1, 8)
        
        XXClave = XXClaveOriginal + XXArticulo + Auxi2 + Auxi1
    
        Rem If XXTerminadoOriginal = "PT-16851-104" Then Stop
    
        With rstTrazabilidad
            .Index = "Clave"
            .Seek "=", XXClave
            If .NoMatch = True Then
                .AddNew
                !Clave = XXClave
                !Terminado = XXTerminado
                !terminadoii = XXTerminadoOriginal
                !Partida = XXPartida
                !Articulo = XXArticulo
                !Lote = XXLote
                !LoteOriginal = XXLoteOriginal
                !CantiLote = Val(XXCantiLote)
                !ClaveOriginal = XXClaveOriginal
                !Proveedor = XXProveedor
                !Empresa = Val(XXEmpresa)
                !Orden = Val(XXOrden)
                !Carpeta = Val(XXCarpeta)
                !Remito = Val(XXRemito)
                !FActura = Val(XXFactura)
                !FechaCompro = XXFechaCompo
                !Costo = XXCosto
                !Informe = Val(XXInforme)
                !CantiArticulo = Val(XXCantidadOriginal)
                !DesProveedor = XXDesProveedor
                !DesArticulo = XXDesArticulo
                !CantiTotal = Val(XXCantidadTotal)
                .Update
                    Else
                .Edit
                !CantiLote = !CantiLote + Val(XXCantiLote)
                .Update
            End If
        End With
        
    Next Ciclo
    
    With rstTrazabilidad
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                ZZTerminado = !terminadoii
                                
                ClavePrecios = Cliente.Text + ZZTerminado
                ZZDesTerminado = ""
        
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    ZZDesTerminado = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                
                !DesTerminado = ZZDesTerminado
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "Trazabilidad.rpt"
        Listado.Destination = 1
            Else
        If Tipo.ListIndex = 1 Then
            Listado.ReportFileName = "Trazabilidad.rpt"
            Listado.Destination = 0
                Else
            Listado.ReportFileName = "TrazabilidadExcel.rpt"
            Listado.Destination = 0
        End If
    End If

    Listado.WindowTitle = "Trazabilidad"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
     
    Listado.Action = 1

End Sub

Private Sub WVector1_DblClick()
        
    WVector1.Col = 4
    XXClave = WVector1.Text
    WVector1.Col = 1
    XXCodigo = WVector1.Text
    WVector1.Col = 3
    XXCantidadTotal = Val(WVector1.Text)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.Clave = " + "'" + XXClave + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        Erase ZZLote
    
        ZZLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
        ZZLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
        
        ZZLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
        ZZLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "0", rstEstadistica!lote2)
        
        ZZLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
        ZZLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "0", rstEstadistica!lote3)
        
        ZZLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
        ZZLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "0", rstEstadistica!lote4)
        
        ZZLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
        ZZLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "0", rstEstadistica!lote5)
                
        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
        
        If Len(Trim(WLoteAdicional)) = 98 Then
        
            ZZLote(6, 2) = Val(Mid$(WLoteAdicional, 1, 8))
            ZZLote(6, 1) = Val(Mid$(WLoteAdicional, 9, 6))
            
            ZZLote(7, 2) = Val(Mid$(WLoteAdicional, 15, 8))
            ZZLote(7, 1) = Val(Mid$(WLoteAdicional, 23, 6))
            
            ZZLote(8, 2) = Val(Mid$(WLoteAdicional, 29, 8))
            ZZLote(8, 1) = Val(Mid$(WLoteAdicional, 37, 6))
            
            ZZLote(9, 2) = Val(Mid$(WLoteAdicional, 43, 8))
            ZZLote(9, 1) = Val(Mid$(WLoteAdicional, 51, 6))
            
            ZZLote(10, 2) = Val(Mid$(WLoteAdicional, 57, 8))
            ZZLote(10, 1) = Val(Mid$(WLoteAdicional, 65, 6))
            
            ZZLote(11, 1) = Val(Mid$(WLoteAdicional, 71, 8))
            ZZLote(11, 2) = Val(Mid$(WLoteAdicional, 79, 6))
            
            ZZLote(12, 1) = Val(Mid$(WLoteAdicional, 85, 8))
            ZZLote(12, 2) = Val(Mid$(WLoteAdicional, 93, 6))
            
        End If
        
        rstEstadistica.Close
        
    End If
       
       
       
    XXLote = ZZLote(1, 1)
    XXCantidad = ZZLote(1, 2)
       
    If Val(WEmpresa) = 1 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        Hasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        Hasta = 4
    End If
    
    XEmpresa = WEmpresa
    
    For CiclaEmpresa = 1 To Hasta

        WEmpresa = Empe(CiclaEmpresa, 1)
        txtOdbc = Empe(CiclaEmpresa, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Hoja = " + "'" + XXLote + "'"
        ZSql = ZSql + " and Hoja.Producto = " + "'" + XXCodigo + "'"
        ZSql = ZSql + " Order by clave"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            
            Erase ZZHoja
            ZZLugar = 0
            
            With rstHoja
            
                .MoveFirst
                If .NoMatch = False Then
                    Do
                    
                        ZZLugar = ZZLugar + 1
                        
                        ZZHoja(ZZLugar, 1) = rstHoja!Tipo
                        ZZHoja(ZZLugar, 2) = rstHoja!Articulo
                        ZZHoja(ZZLugar, 3) = rstHoja!Terminado
                        ZZHoja(ZZLugar, 4) = Str$(rstHoja!Cantidad)
                        ZZHoja(ZZLugar, 5) = Str$(rstHoja!lote1)
                        ZZHoja(ZZLugar, 6) = Str$(rstHoja!Canti1)
                        ZZHoja(ZZLugar, 7) = Str$(rstHoja!lote2)
                        ZZHoja(ZZLugar, 8) = Str$(rstHoja!Canti2)
                        ZZHoja(ZZLugar, 9) = Str$(rstHoja!lote3)
                        ZZHoja(ZZLugar, 10) = Str$(rstHoja!Canti3)
                        
                        .MoveNext
                        
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                    Loop
                End If
                
            End With
                
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            EmpresaFabrica = CiclaEmpresa
            
            
            
            
            rstHoja.Close
            Exit For
        End If
        
    Next CiclaEmpresa
    
    Call Conecta_Empresa
    
  Rem  Stop
    
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Call Limpia_Vector
    
    Renglon = 0
    Numero.SetFocus

End Sub


Private Sub Form_Load()

    Call Limpia_Vector
    
    Tipo.Clear
    
    Tipo.AddItem "Normal"
    Tipo.AddItem "Normal Panta"
    Tipo.AddItem "Excel"
    
    Tipo.ListIndex = 0

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Rem Numero.SetFocus
     
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Proceso1_Click()

    WNeto = 0
    
    Call Limpia_Vector
    
    Renglon = 0
    Erase Auxiliar
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "01" + "'"
    ZSql = ZSql + " and Estadistica.NUmero = " + "'" + Numero.Text + "'"
    ZSql = ZSql + " Order by Estadistica.Clave"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Renglon = Renglon + 1
                    
                    WVector1.Row = Renglon
            
                    WVector1.Col = 1
                    WVector1.Text = rstEstadistica!Articulo
                    Auxi1 = rstEstadistica!Articulo
                
                    dada = Str$(rstEstadistica!Cantidad)
                    WVector1.Col = 3
                    WVector1.Text = Pusing("###,###.##", dada)
                
                    WVector1.Col = 4
                    WVector1.Text = rstEstadistica!Clave
                
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
    
    For Da = 1 To XRenglon
    
        Auxi1 = Auxiliar(Da, 1)
        
        If Left$(Auxi1, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Renglon = Renglon + 1
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
            Case Else
                ClavePrecios = Cliente.Text + Auxi1
        
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Renglon = Renglon + 1
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
        End Select
    Next Da
    
    WVector1.Col = 1
    WVector1.Row = 1
    WVector1.TopRow = 1

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "01" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
        
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!pago2
                WVendedor = rstCliente!vendedor
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
            Numero.SetFocus
                
        End If
    End If
End Sub



Private Sub Limpia_Vector()

    WVector1.Clear
    
    WVector1.ColWidth(0) = 150
    WVector1.ColWidth(1) = 1500
    WVector1.ColWidth(2) = 4000
    WVector1.ColWidth(3) = 1500
    WVector1.ColWidth(4) = 10
    WVector1.ColWidth(5) = 10
    WVector1.ColWidth(6) = 10
    
    WVector1.Row = 0
    
    WVector1.Col = 1
    WVector1.Text = "Producto"
    
    WVector1.Col = 2
    WVector1.Text = "Descripcion"
    
    WVector1.Col = 3
    WVector1.Text = "Cantidad"
    
    WVector1.Col = 4
    WVector1.Text = ""
    
    WVector1.Col = 5
    WVector1.Text = ""
    
    WVector1.Col = 6
    WVector1.Text = ""
    
End Sub
