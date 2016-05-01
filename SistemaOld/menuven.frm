VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Ventas"
   ClientHeight    =   7830
   ClientLeft      =   2550
   ClientTop       =   870
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   7350
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command11 
      Caption         =   "Command1"
      Height          =   1575
      Left            =   5880
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Menu Maestros 
      Caption         =   "Maestros"
      Begin VB.Menu Rubros 
         Caption         =   "Ingreso de Rubros"
      End
      Begin VB.Menu Vendedores 
         Caption         =   "Ingeso de Vendedores"
      End
      Begin VB.Menu Pago 
         Caption         =   "Ingresos de Condiciones de Pago"
      End
      Begin VB.Menu camb 
         Caption         =   "Ingreso de Cambios"
      End
      Begin VB.Menu Lineas 
         Caption         =   "Ingreso de Lineas de Venta"
      End
      Begin VB.Menu LineasMp 
         Caption         =   "Ingreso de Familia de Materias Primas"
      End
      Begin VB.Menu Envases 
         Caption         =   "Ingreso de Envases"
      End
      Begin VB.Menu Clientes 
         Caption         =   "Ingreso de Clientes"
      End
      Begin VB.Menu Precios 
         Caption         =   "Ingreso de Precios Por Cliente"
      End
      Begin VB.Menu Compo 
         Caption         =   "Ingreso de Composicion de Productos"
      End
      Begin VB.Menu Modif 
         Caption         =   "Modificacion de Precios"
      End
      Begin VB.Menu Gasimpo 
         Caption         =   "Ingreso de Conceptos de Gastos de Importacion"
      End
      Begin VB.Menu CompoVersion 
         Caption         =   "Consulta de Versiones de Composicion de Productos Terminados"
      End
      Begin VB.Menu ConsultaRevusines 
         Caption         =   "Consulta de Revisiones de Ensayos"
      End
      Begin VB.Menu CargaLista 
         Caption         =   "Ingreso de Listas de Precios"
      End
      Begin VB.Menu miraageda 
         Caption         =   "Agenda de Clientes"
      End
      Begin VB.Menu zfdasf 
         Caption         =   "salva de facturas"
         Visible         =   0   'False
      End
      Begin VB.Menu asdfasd 
         Caption         =   "salva de nota de crduito"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Nov 
      Caption         =   "Novedades"
      Begin VB.Menu Pedido 
         Caption         =   "Ingreso de Pedidos"
      End
      Begin VB.Menu Factura 
         Caption         =   "Emision de Factura/Remito (U$S)"
      End
      Begin VB.Menu FacturaII 
         Caption         =   "Emision de Factura/Remito ($)"
      End
      Begin VB.Menu FacturaCon 
         Caption         =   "Emision de Factura de Remitos ya emitidos"
      End
      Begin VB.Menu Devol 
         Caption         =   "Ingreso de Devolucion"
      End
      Begin VB.Menu Varios 
         Caption         =   "Ingreso de Comprobantes Varios"
      End
      Begin VB.Menu Consig 
         Caption         =   "Ingreso de Remitos a Facturar"
      End
      Begin VB.Menu Devcon 
         Caption         =   "Devolucion de Mercaderia de Exportacion"
      End
      Begin VB.Menu prgfactuprovisoria 
         Caption         =   "Ingreso de Factura de Exportacion Provisoria"
      End
      Begin VB.Menu FactuExpo 
         Caption         =   "Ingreso de Factura de Exportacion"
      End
      Begin VB.Menu VariosExpo 
         Caption         =   "Ingreso de Factura de Exportacion por Conceptos Varios"
      End
      Begin VB.Menu variosii 
         Caption         =   "Emision de Notas de Debito/Credito por Diferecnia de Cambio "
         Enabled         =   0   'False
      End
      Begin VB.Menu VariosIII 
         Caption         =   "Emision de Notas de Debito/Credito por Diferecnia de Cambio (Acreditacion)"
         Enabled         =   0   'False
      End
      Begin VB.Menu Autoriza 
         Caption         =   "Autorizacion de Pedidos"
      End
      Begin VB.Menu Modped 
         Caption         =   "Actualizacion de Pedidos"
      End
      Begin VB.Menu MovGas 
         Caption         =   "Ingreso de Gastos de Importacion"
      End
      Begin VB.Menu PedidoDevol 
         Caption         =   "Ingreso de Solicitud de Devolucion de Mercaderia"
      End
      Begin VB.Menu MovGasParcial 
         Caption         =   "Ingreso de Gastos de Importacion Parciales"
      End
      Begin VB.Menu OrdenImpo 
         Caption         =   "Ingreso de Ordenes de Compra de Importacion"
      End
      Begin VB.Menu PedidoOrdenTRabajo 
         Caption         =   "Ingreso de Pedidos de Desarrollo"
      End
      Begin VB.Menu ConsultaDesarrollo 
         Caption         =   "Consulta de Pedidos de Desarrollos"
      End
      Begin VB.Menu AutorizaDesarrollo 
         Caption         =   "Autorizacion de Desarrollos"
      End
      Begin VB.Menu VerificaCosto 
         Caption         =   "Verificacion de Cambio de Costo de Productos Terminados"
      End
      Begin VB.Menu ConsultaHojaRuta 
         Caption         =   "Consulta de Hoja de Ruta (Cot)"
      End
      Begin VB.Menu ConsultaHojaRutaII 
         Caption         =   "Consulta de Hoja de Ruta (Cliente)"
      End
      Begin VB.Menu varios105 
         Caption         =   "Ingreso de Comprobantes Varios ( Iva 10.5%)"
         Enabled         =   0   'False
      End
      Begin VB.Menu PedidoSol 
         Caption         =   "Ingreso de Solicitud de Pedidos de Venta"
      End
      Begin VB.Menu Variosletrab 
         Caption         =   "Factura  de Conceptos Varios ""B"""
      End
      Begin VB.Menu factub 
         Caption         =   "Factura Productos ""B"" Dolar"
      End
      Begin VB.Menu factubPesos 
         Caption         =   "Factura Productos ""B"" Pesos"
      End
      Begin VB.Menu FacturaConB 
         Caption         =   "Emision de Factura ""B"" de Remitos ya emitidos"
      End
      Begin VB.Menu EnvioEmailClie 
         Caption         =   "Envio de Email a Clientes"
      End
      Begin VB.Menu asdf 
         Caption         =   "Centro de Control de Importaciones"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu CtaCte1 
         Caption         =   "Consulta de Cuenta Corriente de Clientes por pantalla"
      End
      Begin VB.Menu CtaCteCli 
         Caption         =   "Cuenta Corriente de Clientes"
      End
      Begin VB.Menu SalCtaCteCli 
         Caption         =   "Saldos de Cuenta Corriente de Clientes"
      End
      Begin VB.Menu IvaVentas 
         Caption         =   "Subdiario de Iva Ventas"
      End
      Begin VB.Menu pedpen 
         Caption         =   "Listado de Pedido Pendientes"
      End
      Begin VB.Menu Cash 
         Caption         =   "Listado de Cash Flow"
      End
      Begin VB.Menu Ivaprv 
         Caption         =   "Listado de Ventas por Provincias"
      End
      Begin VB.Menu Movcon 
         Caption         =   "Listado de Mercaderia en remitos a facturar por cliente"
      End
      Begin VB.Menu Movcon1 
         Caption         =   "Listado de Mercaderia en remitos a facturar por Articulo"
      End
      Begin VB.Menu Estamal 
         Caption         =   "Listado de Ventas fuera de fecha"
      End
      Begin VB.Menu Ctacte9 
         Caption         =   "Listado de Cuenta Corriente a Fecha"
      End
      Begin VB.Menu Listgas 
         Caption         =   "Listado de Gastos de Importacion por Carpeta"
      End
      Begin VB.Menu CostoOrden 
         Caption         =   "Listado del Calculo de Costo de Importacion por Carpeta"
      End
      Begin VB.Menu Vendia 
         Caption         =   "Listado de Ventas Diarias Pendientes de Pago"
      End
      Begin VB.Menu NotaPend 
         Caption         =   "Listado de Notas de Debito por Diferencia de Cambio Pendientes"
      End
      Begin VB.Menu ListIbVen 
         Caption         =   "Listado de Percepciones de Ingresos Brutos"
      End
      Begin VB.Menu AnalisisPedido 
         Caption         =   "Analisis de Cumplimiento de Pedidos de Venta"
      End
      Begin VB.Menu LIstClieVend 
         Caption         =   "Listado de Clientes por Vendedor"
      End
      Begin VB.Menu ordpendyII 
         Caption         =   "Listado de Ordenes Compras Pendientes de Materia Prima"
      End
      Begin VB.Menu CostoOrdenParcial 
         Caption         =   "Listado del Calculo de Costo de Nacionalizacion de Mercaderia"
      End
      Begin VB.Menu ListaPrecios 
         Caption         =   "Listado de Precios (Grupo)"
      End
      Begin VB.Menu ListaPreciosCompa 
         Caption         =   "Listado de Precios Comparativo (Grupo)"
      End
      Begin VB.Menu ListaPreciosCompaCliente 
         Caption         =   "Listado de Precios Comparativo (Cliente)"
      End
      Begin VB.Menu ListadoGeneral 
         Caption         =   "Listado General de Productos"
      End
      Begin VB.Menu CtaCTeAnalitico 
         Caption         =   "Listado de Cuenta Corriente Analitico"
      End
      Begin VB.Menu ProyCtaCteAnalitico 
         Caption         =   "Listado de Proyeccion de Cuentas Corrientes de Clientes Analitico"
      End
      Begin VB.Menu ListaDevol 
         Caption         =   "Listado de Analisis de Devoluciones"
      End
      Begin VB.Menu pedpenpellital 
         Caption         =   "Listado de Pedidos Pendientes de Fazon de Pellital"
      End
      Begin VB.Menu ListaMinutas 
         Caption         =   "Listado de Minutas"
      End
      Begin VB.Menu VerificaVenta 
         Caption         =   "Verificacion de PT sin Ventas"
      End
      Begin VB.Menu venricaventasdy 
         Caption         =   "Verificacion de Dy sin Ventas"
      End
      Begin VB.Menu ctactefechaproye 
         Caption         =   "Proyecion de Ctacte A fecha"
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "Procesos"
      Begin VB.Menu Cierre 
         Caption         =   "Cierre del Mes"
      End
      Begin VB.Menu Veriform 
         Caption         =   "Verificacion de Formulas"
      End
      Begin VB.Menu Grabacd 
         Caption         =   "Grabacion de Datos Electronicamente"
      End
      Begin VB.Menu ActuaOrden 
         Caption         =   "Actualizacion de Costos de Importacion"
      End
      Begin VB.Menu Aaa 
         Caption         =   "Procesos Varios"
         Visible         =   0   'False
      End
      Begin VB.Menu PasaPrecio 
         Caption         =   "PasaPrecios"
      End
      Begin VB.Menu ListaVerificaCosto 
         Caption         =   "Veriricacion de Cambios de Costos"
      End
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstAtributo As Recordset
Dim spAtributo As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstComposicionVersion As Recordset
Dim spComposicionVersion As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim rstCargaIV As Recordset
Dim spCargaIV As String
Dim rstCargaIVVersion As Recordset
Dim spCargaIVVersion As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String

Dim Atri(10, 100) As Integer

Private Sub Arti_Click()
    PrgArti.Show
End Sub

Private Sub Aaa_Click()
    PrgFcsalva.Show
End Sub

Private Sub ActuaOrden_Click()
    PrgActuaOrden.Show
End Sub

Private Sub Aju_Click()
    ProcesoActivate = 0
    PrgAju.Show
End Sub

Private Sub AnalisisPedido_Click()
    PrgAnalisisPedido.Show
End Sub

Private Sub asdf_Click()
    PrgCentroImportacion.Show
End Sub

Private Sub asdfasd_Click()
    PrgDevolSalva.Show
End Sub

Private Sub Autoriza_Click()
    PrgAutoriza.Show
End Sub

Private Sub AutorizaDesarrollo_Click()
    PrgAutorizaDesarrollo.Show
End Sub

Private Sub Camb_Click()
    PrgCambios.Show
End Sub

Private Sub cambioasd_Click()
    PrgCentro.Show
End Sub

Private Sub Camiones_Click()
    PrgCamiones.Show
End Sub

Private Sub CargaLista_Click()
    PrgCargaLista.Show
End Sub

Private Sub Cash_Click()
    PrgCash.Show
End Sub

Private Sub Choferes_Click()
    PrgChoferes.Show
End Sub

Private Sub Clientes_Click()
    prgcliente.Show
End Sub

Private Sub Command1_Click()

    Dim ZZVector(5000, 10) As String
    Dim ZZLugar As Integer
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Order by Codigo"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstTerminado!Codigo, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstTerminado!Codigo
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZArticulo = "PT-12" + Mid$(ZZVector(Ciclo, 1), 6, 10)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Terminado SET "
        ZSql = ZSql + "Codigo = " + "'" + ZZArticulo + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZZVector(Ciclo, 1) + "'"
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Order by Codigo"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstTerminado!Codigo, 2)) = "NW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstTerminado!Codigo
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZArticulo = "NK-12" + Mid$(ZZVector(Ciclo, 1), 6, 10)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Terminado SET "
        ZSql = ZSql + "Codigo = " + "'" + ZZArticulo + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZZVector(Ciclo, 1) + "'"
        spTerminado = ZSql
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Composicion"
    ZSql = ZSql + " Order by Clave"
    spComposicion = ZSql
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicion.RecordCount > 0 Then
        With rstComposicion
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstComposicion!Clave, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstComposicion!Clave
                        ZZVector(ZZLugar, 2) = rstComposicion!Terminado
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicion.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClave = "PT-12" + Mid$(ZZVector(Ciclo, 1), 6, 100)
        ZZTerminado = "PT-12" + Mid$(ZZVector(Ciclo, 2), 6, 100)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Composicion SET "
        ZSql = ZSql + "Clave = " + "'" + ZZClave + "',"
        ZSql = ZSql + "Terminado = " + "'" + ZZTerminado + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZVector(Ciclo, 1) + "'"
        spComposicion = ZSql
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ComposicionVersion"
    ZSql = ZSql + " Order by Clave"
    spComposicionVersion = ZSql
    Set rstComposicionVersion = db.OpenRecordset(spComposicionVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicionVersion.RecordCount > 0 Then
        With rstComposicionVersion
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstComposicionVersion!Clave, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstComposicionVersion!Clave
                        ZZVector(ZZLugar, 2) = rstComposicionVersion!Terminado
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicionVersion.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClave = "PT-12" + Mid$(ZZVector(Ciclo, 1), 6, 100)
        ZZTerminado = "PT-12" + Mid$(ZZVector(Ciclo, 2), 6, 100)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE ComposicionVersion SET "
        ZSql = ZSql + "Clave = " + "'" + ZZClave + "',"
        ZSql = ZSql + "Terminado = " + "'" + ZZTerminado + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZVector(Ciclo, 1) + "'"
        spComposicionVersion = ZSql
        Set rstComposicionVersion = db.OpenRecordset(spComposicionVersion, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecifUnifica"
    ZSql = ZSql + " Order by Producto"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        With rstEspecifUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstEspecifUnifica!Producto, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstEspecifUnifica!Producto
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZProducto = "PT-12" + Mid$(ZZVector(Ciclo, 1), 6, 100)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE EspecifUnifica SET "
        ZSql = ZSql + "Producto = " + "'" + ZZProducto + "'"
        ZSql = ZSql + " Where Producto = " + "'" + ZZVector(Ciclo, 1) + "'"
        spEspecifUnifica = ZSql
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecifUnificaVersion"
    ZSql = ZSql + " Order by Producto"
    spEspecifUnificaVersion = ZSql
    Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnificaVersion.RecordCount > 0 Then
        With rstEspecifUnificaVersion
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstEspecifUnificaVersion!Producto, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstEspecifUnificaVersion!Clave
                        ZZVector(ZZLugar, 2) = rstEspecifUnificaVersion!Producto
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnificaVersion.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClave = Left$(ZZVector(Ciclo, 1), 4) + "PT-12" + Mid$(ZZVector(Ciclo, 1), 10, 100)
        ZZProducto = "PT-12" + Mid$(ZZVector(Ciclo, 2), 6, 100)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE EspecifUnificaVersion SET "
        ZSql = ZSql + "Producto = " + "'" + ZZProducto + "',"
        ZSql = ZSql + "Clave = " + "'" + ZZClave + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZVector(Ciclo, 1) + "'"
        spEspecifUnificaVersion = ZSql
        Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIV"
    ZSql = ZSql + " Order by Clave"
    spCargaIV = ZSql
    Set rstCargaIV = db.OpenRecordset(spCargaIV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIV.RecordCount > 0 Then
        With rstCargaIV
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZTerminado = IIf(IsNull(rstCargaIV!Terminado), "", rstCargaIV!Terminado)
                    If UCase(Left$(ZZTerminado, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstCargaIV!Clave
                        ZZVector(ZZLugar, 2) = rstCargaIV!Terminado
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIV.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClave = "PT-12" + Mid$(ZZVector(Ciclo, 1), 6, 100)
        ZZTerminado = "PT-12" + Mid$(ZZVector(Ciclo, 2), 6, 100)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaIV SET "
        ZSql = ZSql + "Terminado = " + "'" + ZZTerminado + "',"
        ZSql = ZSql + "Clave = " + "'" + ZZClave + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZVector(Ciclo, 1) + "'"
        spCargaIV = ZSql
        Set rstCargaIV = db.OpenRecordset(spCargaIV, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIVVersion"
    ZSql = ZSql + " Order by Clave"
    spCargaIVVersion = ZSql
    Set rstCargaIVVersion = db.OpenRecordset(spCargaIVVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIVVersion.RecordCount > 0 Then
        With rstCargaIVVersion
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstCargaIVVersion!Terminado, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstCargaIVVersion!Clave
                        ZZVector(ZZLugar, 2) = rstCargaIVVersion!Terminado
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIVVersion.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClave = "PT-12" + Mid$(ZZVector(Ciclo, 1), 6, 100)
        ZZTerminado = "PT-12" + Mid$(ZZVector(Ciclo, 2), 6, 100)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaIVVersion SET "
        ZSql = ZSql + "Terminado = " + "'" + ZZTerminado + "',"
        ZSql = ZSql + "Clave = " + "'" + ZZClave + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZVector(Ciclo, 1) + "'"
        spCargaIVVersion = ZSql
        Set rstCargaIVVersion = db.OpenRecordset(spCargaIVVersion, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PreciosMP"
    ZSql = ZSql + " Order by Clave"
    spPreciosMp = ZSql
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
        With rstPreciosMp
            .MoveFirst
            Do
                If .EOF = False Then
                    If UCase(Left$(rstPreciosMp!Articulo, 2)) = "DW" Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar, 1) = rstPreciosMp!Clave
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPreciosMp.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZClavePrecio = ZZVector(Ciclo, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM PreciosMp"
        ZSql = ZSql + " Where PreciosMp.Clave = " + "'" + ZZClavePrecio + "'"
        spPreciosMp = ZSql
        Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
        If rstPreciosMp.RecordCount > 0 Then
        
            ZZClave = rstPreciosMp!Clave
            ZZCliente = rstPreciosMp!Cliente
            ZZArticulo = rstPreciosMp!Articulo
            ZZPrecio = rstPreciosMp!Precio
            ZZFecha1 = rstPreciosMp!Fecha1
            ZZFactura1 = rstPreciosMp!Factura1
            ZZPrecio1 = rstPreciosMp!Precio1
            ZZCantidad1 = rstPreciosMp!Cantidad1
            ZZFecha2 = rstPreciosMp!Fecha2
            ZZFactura2 = rstPreciosMp!Factura2
            ZZPrecio2 = rstPreciosMp!Precio2
            ZZCantidad2 = rstPreciosMp!Cantidad2
            ZZFecha3 = rstPreciosMp!Fecha3
            ZZFactura3 = rstPreciosMp!Factura3
            ZZPrecio3 = rstPreciosMp!Precio3
            ZZCantidad3 = rstPreciosMp!Cantidad3
            ZZFecha4 = rstPreciosMp!Fecha4
            ZZFactura4 = rstPreciosMp!Factura4
            ZZPrecio4 = rstPreciosMp!Precio4
            ZZCantidad4 = rstPreciosMp!Cantidad4
            ZZFecha5 = rstPreciosMp!Fecha5
            ZZFactura5 = rstPreciosMp!Factura5
            ZZPrecio5 = rstPreciosMp!Precio5
            ZZCantidad5 = rstPreciosMp!Cantidad5
            ZZDate = rstPreciosMp!WDate
            ZZFecha = rstPreciosMp!Fecha
            ZZPago = rstPreciosMp!Pago
            
            rstPreciosMp.Close
            
            ZZClave = Left$(ZZClave, 6) + "PT-12" + Mid$(ZZClave, 10, 100)
            ZZTerminado = "PT-12" + Mid$(ZZArticulo, 4, 100)
            ZZDescripcion = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Terminado"
            ZSql = ZSql + " Where Terminado.Codigo = " + "'" + ZZTerminado + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZZDescripcion = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            
                    
            XParam = "'" + ZZClave + "','" + ZZCliente + "','" + ZZTerminado + "','" + Str$(ZZPrecio) + "','" _
                     + ZZDescripcion + "','" _
                     + ZZFecha1 + "','" + ZZFactura1 + "','" + Str$(ZZPrecio1) + "','" + Str$(ZZCantidad1) + "','" _
                     + ZZFecha2 + "','" + ZZFactura2 + "','" + Str$(ZZPrecio2) + "','" + Str$(ZZCantidad2) + "','" _
                     + ZZFecha3 + "','" + ZZFactura3 + "','" + Str$(ZZPrecio3) + "','" + Str$(ZZCantidad3) + "','" _
                     + ZZFecha4 + "','" + ZZFactura4 + "','" + Str$(ZZPrecio4) + "','" + Str$(ZZCantidad4) + "','" _
                     + ZZFecha5 + "','" + ZZFactura5 + "','" + Str$(ZZPrecio5) + "','" + Str$(ZZCantidad5) + "','" _
                     + ZZDate + "','" + ZZFecha + "','" + Str$(ZZPago) + "'"
            Set rstPrecios = db.OpenRecordset("AltaPrecios1 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Precios SET "
            ZSql = ZSql & "Estado = " + "'" + "0" + "'"
            ZSql = ZSql & " Where Clave = " + "'" + ZZClave + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Stop
    


End Sub

Private Sub Command11_Click()
    
    Dim ZZCliente(1000) As String
    ZZLugar = 0
    ZZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.OrdFecha >= " + "'" + "20130901" + "'"
    ZSql = ZSql + " and CtaCte.Tipo = " + "'" + "01" + "'"
    ZSql = ZSql + " Order by ctacte.Cliente"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
        With rstCtacte
        
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    If ZZPasa = 0 Then
                        ZZPasa = 1
                        ZZLugar = ZZLugar + 1
                        ZZCorte = rstCtacte!Cliente
                    End If
                    
                    If ZZCorte <> rstCtacte!Cliente Then
                        ZZLugar = ZZLugar + 1
                        ZZCorte = rstCtacte!Cliente
                    End If
                    
                    ZZCliente(ZZLugar) = ZZCorte
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            End If
            
        End With
        rstCtacte.Close
    
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Cliente SET "
    ZSql = ZSql + " MarcaI = " + "'" + ClaveRecibo + "',"
    ZSql = ZSql + " MarcaII = " + "'" + XRetIbCiudad + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    For Ciclo = 1 To ZZLugar
    
        ZZMarcaI = "X"
        ZZMarcaII = ""
    
        WCliente = ZZCliente(Ciclo)
        
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.OrdFecha >= " + "'" + "20130901" + "'"
        ZSql = ZSql + " and CtaCte.Tipo = " + "'" + "04" + "'"
        ZSql = ZSql + " and CtaCte.Impre = " + "'" + "ND" + "'"
        ZSql = ZSql + " and CtaCte.Cliente = " + "'" + WCliente + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            ZZMarcaII = "X"
            rstCtacte.Close
        End If
        
        
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Cliente SET "
        ZSql = ZSql + " MarcaI = " + "'" + ZZMarcaI + "',"
        ZSql = ZSql + " MarcaII = " + "'" + ZZMarcaII + "'"
        ZSql = ZSql + " Where Cliente = " + "'" + WCliente + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Next Ciclo
    
    
    Stop
    
    

End Sub

Private Sub Compo_Click()
    PrgCompo.Show
End Sub

Private Sub CompoVersion_Click()
    PrgCompoVersion.Show
End Sub


Private Sub ConsultaDesarrollo_Click()
    PrgConsultaDesarrollo.Show
End Sub

Private Sub ConsultaHojaRuta_Click()
    PrgConsultaHojaRuta.Show
End Sub

Private Sub ConsultaHojaRutaII_Click()
    PrgConsultaHojaRutaII.Show
End Sub

Private Sub ConsultaRevusines_Click()
    PrgConsultaEnsayo.Show
End Sub

Private Sub CostoOrden_Click()
    PrgCostoOrden.Show
End Sub

Private Sub CostoOrdenParcial_Click()
    PrgCostoOrdenParcial.Show
End Sub

Private Sub CtaCte1_Click()
    PrgCtaCte2.Show
End Sub

Private Sub Ctacte9_Click()
    PrgCtaCteFec.Show
End Sub

Private Sub CtaCTeAnalitico_Click()
    PrgCtaCteAnalitico.Show
End Sub

Private Sub CtaCteCli_Click()
    PrgCtaCte.Show
End Sub

Private Sub ctactefechaproye_Click()
    PrgCtaCtefechaProye.Show
End Sub

Private Sub Devcon_Click()
    PrgDevolExpo.Show
End Sub

Private Sub Devol_Click()
    PrgDevolPesos.Show
End Sub

Private Sub Devol21_Click()
    PrgDevol21.Show
End Sub

Private Sub Envases_Click()
    PrgEnv.Show
End Sub

Private Sub EnvioEmailClie_Click()
    PrgEnvioEmailClie.Show
End Sub

Private Sub Estamal_Click()
    PrgEstamal.Show
End Sub

Private Sub factub_Click()
    Rem WPasaMoneda = 0
    Rem PrgFactuLetraB.Show
End Sub

Private Sub factubPesos_Click()
    WPasaMoneda = 1
    PrgFactuLetraB.Show
End Sub

Private Sub FactuExpo_Click()
    WPasaMoneda = 0
    Rem PrgFactu.Show
    PrgFactuexpo.Show
End Sub

Private Sub Factura_Click()
    WPasaMoneda = 0
    PrgFactu.Show
End Sub

Private Sub factux_Click()
    PrgFactux.Show
End Sub

Private Sub Impreped_Click()
    PrgImpreped.Show
End Sub

Private Sub FacturaCon_Click()
    WPasaMoneda = 1
    PrgFactuRemitoActualiza.Show
End Sub

Private Sub FacturaConB_Click()
    PrgFactuRemitoActualizaB.Show
End Sub

Private Sub FacturaII_Click()
    WPasaMoneda = 1
    PrgFactu.Show
    Rem PrgFactup.Show
End Sub

Private Sub fsdf_Click()
    Form1.Show
End Sub

Private Sub consig_Click()
    PrgFactuRemito.Show
End Sub

Private Sub facturemito_Click()

End Sub

Private Sub Gasimpo_Click()
    PrgGasimpo.Show
End Sub

Private Sub Grabacd_Click()
    PrgGrabaCd.Show
End Sub

Private Sub HojaRuta_Click()
    PrgHojaRuta.Show
End Sub

Private Sub Ivaprv_Click()
    PrgIvavenPrv.Show
End Sub

Private Sub IvaVentas_Click()
    PrgIvaven.Show
End Sub

Private Sub Lineas_Click()
    PrgLinea.Show
End Sub

Private Sub LineasMp_Click()
    PrgLineaMp.Show
End Sub

Private Sub ListaDevol_Click()
    PrgListaDevol.Show
End Sub

Private Sub ListadoGeneral_Click()
    PrgListaGeneral.Show
End Sub

Private Sub ListaMinutas_Click()
    PrgListaMinutas.Show
End Sub

Private Sub ListaPrecios_Click()
    PrgListaPrecios.Show
End Sub

Private Sub ListaPreciosCompa_Click()
    PrgListaPreciosCompa.Show
End Sub

Private Sub ListaPreciosCompaCliente_Click()
    PrgListaPreciosCompaCliente.Show
End Sub

Private Sub ListaVerificaCosto_Click()
    PrgListaVerificaCosto.Show
End Sub

Private Sub LIstClieVend_Click()
    PrgListClieVend.Show
End Sub

Private Sub Listgas_Click()
    PrgListgas.Show
End Sub

Private Sub ListIbVen_Click()
    PrgListIbVen.Show
End Sub

Private Sub miraageda_Click()
    PrgMiraAgenda.Show
End Sub

Private Sub Modif_Click()
    PrgModif.Show
End Sub

Private Sub Modped_Click()
    PrgModpedNuevo.Show
End Sub

Private Sub ModpedExp_Click()
    PrgModpedExp.Show
End Sub

Private Sub Movcon_Click()
    PrgMovcon.Show
End Sub

Private Sub Movcon1_Click()
    PrgMovcon1.Show
End Sub

Private Sub Movgas_Click()
    PrgMovgas.Show
End Sub

Private Sub Muestra_Click()
    PrgMuestra.Show
End Sub

Private Sub movguiaauto_Click()
    PrgMovguiaAuto.Show
End Sub

Private Sub MovGasParcial_Click()
    PrgMovgasParcial.Show
End Sub

Private Sub NotaPend_Click()
    PrgNotaPend.Show
End Sub

Private Sub OrdenImpo_Click()
    Rem If Val(WEmpresa) = 1 Then
        PrgOrdenImpo.Show
    Rem End If
End Sub

Private Sub ordpendyII_Click()
    PrgOrdPenDyII.Show
End Sub

Private Sub PasaPrecio_Click()
    PrgPasaPrecio.Show
End Sub

Private Sub Pedido_Click()
    PrgPedido.Show
End Sub

Private Sub PedidoDevol_Click()
    PrgPedidodevol.Show
End Sub

Private Sub PedidoOrdenTRabajo_Click()
    PrgpedidoOrdenTrabajo.Show
End Sub

Private Sub PedidoSol_Click()
    PrgPedidoSol.Show
End Sub

Private Sub PedPen_Click()
    PrgPedPen.Show
End Sub

Private Sub pedpenpellital_Click()
    PrgPedPenPellital.Show
End Sub

Private Sub Precios_Click()
    PrgPrecio.Show
End Sub

Private Sub Prueba_Click()
    prgPrueba.Show
End Sub

Private Sub prgdevolx_Click()
    PrgDevx.Show
End Sub

Private Sub prgfactuprovisoria_Click()
    PrgFactuProvi.Show
End Sub

Private Sub pruebpeddo_Click()
    PrgPedPenII.Show
End Sub

Private Sub ProyCtaCteAnalitico_Click()
    PrgProyCtaCteAnalitico.Show
End Sub

Private Sub Rubros_Click()
    PrgRubro.Show
End Sub

Private Sub SalCtaCteCli_Click()
    PrgSaldoCta.Show
End Sub

Private Sub Terminado_Click()
    PrgTermi.Show
End Sub

Private Sub Ultima_Click()
    PrgUltima.Show
End Sub

Private Sub Salva_Click()
    PrgSalva.Show
End Sub

Private Sub SolGuia_Click()
    PrgSolGuia.Show
End Sub

Private Sub solguiatotal_Click()
    PrgMiraSolGuia.Show
End Sub

Private Sub solhoja_Click()
    PrgSolHojaII.Show
End Sub

Private Sub Varios_Click()
    PrgVarios.Show
End Sub

Private Sub varios105_Click()
    PrgVarios105.Show
End Sub

Private Sub VariosExpo_Click()
    PrgVariosEste.Show
End Sub

Private Sub variosii_Click()
    PrgVariosII.Show
End Sub

Private Sub VariosIII_Click()
    PrgVariosIII.Show
End Sub

Private Sub Variosletrab_Click()
    PrgVariosLetraB.Show
End Sub

Private Sub Vendedores_Click()
    PrgVendedor.Show
End Sub

Private Sub pago_Click()
    PrgCondPago.Show
End Sub

Private Sub Cambio_Click()
    frmLoginCotiza.Show
End Sub

Private Sub Fin_Click()
    Close
    End
End Sub

Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

    If Wempresa = "" Then
        Wempresa = "0001"
    End If

    If Wempresa = "" Then
        Empresa.Show
        Empresa.SetFocus
        Wempresa = 1
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(Wempresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de ventas : " + !Nombre
            End If
        End With
    End If
    
    XOperador = Str$(WOperador)
    XProceso = "2"
    WAtributo1 = "0000000000000000000000000000000"
    WAtributo2 = "0000000000000000000000000000000"
    WAtributo3 = "0000000000000000000000000000000"
    WAtributo4 = "0000000000000000000000000000000"
    WAtributo5 = "0000000000000000000000000000000"
    WAtributo6 = "0000000000000000000000000000000"
    WAtributo7 = "0000000000000000000000000000000"
    WAtributo8 = "0000000000000000000000000000000"
    WAtributo9 = "0000000000000000000000000000000"
    WAtributo10 = "000000000000000000000000000000"
    
    XParam = "'" + XOperador + "','" _
                 + XProceso + "'"
    spAtributo = "ConsultaAtributo " + XParam
    Set rstAtributo = db.OpenRecordset(spAtributo, dbOpenSnapshot, dbSQLPassThrough)
    If rstAtributo.RecordCount > 0 Then
        WAtributo1 = rstAtributo!Atributo1 + "00000000000000000000000000000"
        WAtributo2 = rstAtributo!Atributo2 + "000000000000000000000000000000"
        WAtributo3 = rstAtributo!Atributo3 + "000000000000000000000000000000000000"
        WAtributo4 = rstAtributo!Atributo4 + "00000000000000000000000000000000"
        WAtributo5 = rstAtributo!Atributo5 + "00000000000000000000000000000000"
        WAtributo6 = rstAtributo!Atributo6 + "00000000000000000000000000000000"
        WAtributo7 = rstAtributo!Atributo7 + "00000000000000000000000000000000"
        WAtributo8 = rstAtributo!Atributo8 + "00000000000000000000000000000000"
        WAtributo9 = rstAtributo!Atributo9 + "00000000000000000000000000000000"
        WAtributo10 = rstAtributo!Atributo10 + "00000000000000000000000000000"
        rstAtributo.Close
    End If
    
    For Ciclo = 1 To 10
        Select Case Ciclo
            Case 1
                Auxiliar = WAtributo1
            Case 2
                Auxiliar = WAtributo2
            Case 3
                Auxiliar = WAtributo3
            Case 4
                Auxiliar = WAtributo4
            Case 5
                Auxiliar = WAtributo5
            Case 6
                Auxiliar = WAtributo6
            Case 7
                Auxiliar = WAtributo7
            Case 8
                Auxiliar = WAtributo8
            Case 9
                Auxiliar = WAtributo9
            Case 10
                Auxiliar = WAtributo10
            Case Else
        End Select
        For Ciclo1 = 1 To 32
            aa = Ciclo
            AA1 = Ciclo1
            Atri(Ciclo, Ciclo1) = Val(Mid$(Auxiliar, Ciclo1, 1))
        Next Ciclo1
    Next Ciclo
            
    Menu.Rubros.Enabled = Atri(1, 1)
    Menu.Vendedores.Enabled = Atri(1, 2)
    Menu.Pago.Enabled = Atri(1, 3)
    Menu.camb.Enabled = Atri(1, 4)
    Menu.Lineas.Enabled = Atri(1, 5)
    Menu.LineasMp.Enabled = Atri(1, 6)
    Menu.Envases.Enabled = Atri(1, 7)
    Menu.Clientes.Enabled = Atri(1, 8)
    Menu.Precios.Enabled = Atri(1, 9)
    Menu.Compo.Enabled = Atri(1, 10)
    Menu.Modif.Enabled = Atri(1, 11)
    Menu.Gasimpo.Enabled = Atri(1, 12)
    Menu.CompoVersion.Enabled = Atri(1, 13)
    Menu.ConsultaRevusines.Enabled = Atri(1, 14)
    Menu.CargaLista.Enabled = Atri(1, 15)
    Rem by nan
    Menu.miraageda.Enabled = Atri(1, 16)
    Menu.zfdasf.Enabled = Atri(1, 17)
    Menu.asdfasd.Enabled = Atri(1, 18)
    Rem  end by nan
    
    
    
    
    Menu.Pedido.Enabled = Atri(2, 1)
    Menu.Factura.Enabled = Atri(2, 2)
    Menu.FacturaII.Enabled = Atri(2, 3)
    Menu.FacturaCon.Enabled = Atri(2, 4)
    Menu.Devol.Enabled = Atri(2, 5)
    Menu.Varios.Enabled = Atri(2, 6)
    Menu.Consig.Enabled = Atri(2, 7)
    Menu.Devcon.Enabled = Atri(2, 8)
    Menu.prgfactuprovisoria.Enabled = Atri(2, 9)
    Menu.FactuExpo.Enabled = Atri(2, 10)
    Menu.VariosExpo.Enabled = Atri(2, 11)
    Rem Menu.variosii.Enabled = Atri(2, 12)
    Rem Menu.VariosIII.Enabled = Atri(2, 13)
    Menu.variosii.Enabled = False
    Menu.VariosIII.Enabled = False
    Menu.Autoriza.Enabled = Atri(2, 14)
    Menu.Modped.Enabled = Atri(2, 15)
    Menu.MovGas.Enabled = Atri(2, 16)
    Menu.PedidoDevol.Enabled = Atri(2, 17)
    Menu.MovGasParcial.Enabled = Atri(2, 18)
    Menu.OrdenImpo.Enabled = Atri(2, 19)
    Menu.PedidoOrdenTRabajo.Enabled = Atri(2, 20)
    Menu.ConsultaDesarrollo.Enabled = Atri(2, 21)
    Menu.AutorizaDesarrollo.Enabled = Atri(2, 22)
    Menu.VerificaCosto.Enabled = Atri(2, 23)
    Menu.ConsultaHojaRuta.Enabled = Atri(2, 24)
    Menu.ConsultaHojaRutaII.Enabled = Atri(2, 25)
    Menu.PedidoSol.Enabled = Atri(2, 26)
    Menu.Variosletrab.Enabled = Atri(2, 27)
    Menu.factub.Enabled = Atri(2, 28)
    Menu.factubPesos.Enabled = Atri(2, 29)
    Menu.FacturaConB.Enabled = Atri(2, 30)
    Menu.EnvioEmailClie.Enabled = Atri(2, 31)
    
    
    
    
    
    
    
    Menu.CtaCte1.Enabled = Atri(3, 1)
    Menu.CtaCteCli.Enabled = Atri(3, 2)
    Menu.SalCtaCteCli.Enabled = Atri(3, 3)
    Menu.IvaVentas.Enabled = Atri(3, 4)
    Menu.pedpen.Enabled = Atri(3, 5)
    Menu.Cash.Enabled = Atri(3, 6)
    Menu.Ivaprv.Enabled = Atri(3, 7)
    Menu.Movcon.Enabled = Atri(3, 8)
    Menu.Movcon1.Enabled = Atri(3, 9)
    Menu.Estamal.Enabled = Atri(3, 10)
    Menu.Ctacte9.Enabled = Atri(3, 11)
    Menu.Listgas.Enabled = Atri(3, 12)
    Menu.CostoOrden.Enabled = Atri(3, 13)
    Menu.Vendia.Enabled = Atri(3, 14)
    Menu.NotaPend.Enabled = Atri(3, 15)
    Menu.ListIbVen.Enabled = Atri(3, 16)
    Menu.AnalisisPedido.Enabled = Atri(3, 17)
    Menu.LIstClieVend.Enabled = Atri(3, 18)
    Menu.ordpendyII.Enabled = Atri(3, 19)
    Menu.CostoOrdenParcial.Enabled = Atri(3, 20)
    Menu.ListaPrecios.Enabled = Atri(3, 21)
    Menu.ListaPreciosCompa.Enabled = Atri(3, 22)
    Menu.ListaPreciosCompaCliente.Enabled = Atri(3, 23)
    Menu.ListadoGeneral.Enabled = Atri(3, 24)
    Menu.CtaCTeAnalitico.Enabled = Atri(3, 25)
    Menu.ProyCtaCteAnalitico.Enabled = Atri(3, 26)
    Menu.ListaDevol.Enabled = Atri(3, 27)
    Menu.pedpenpellital.Enabled = Atri(3, 28)
    Menu.ListaMinutas.Enabled = Atri(3, 29)
    Menu.VerificaVenta.Enabled = Atri(3, 30)
    Menu.venricaventasdy.Enabled = Atri(3, 31)
        
    
    
    
    
    
    
    Menu.Cierre.Enabled = Atri(4, 1)
    Menu.Veriform.Enabled = Atri(4, 2)
    Menu.Grabacd.Enabled = Atri(4, 3)
    Menu.ActuaOrden.Enabled = Atri(4, 4)
    Menu.ActuaOrden.Enabled = Atri(4, 5)
    Menu.Aaa.Enabled = Atri(4, 6)
    Menu.PasaPrecio.Enabled = Atri(4, 7)
    Menu.ListaVerificaCosto = Atri(4, 8)
    
    Rem Menu.Fin.Enabled = Atri(4, 5)

End Sub

Private Sub Vendia_Click()
    PrgVendia.Show
End Sub

Private Sub venricaventasdy_Click()
    PrgVerificaVentaDy.Show
End Sub

Private Sub VerificaCosto_Click()
    PrgVerificaCosto.Show
End Sub

Private Sub VerificaVenta_Click()
    PrgVerificaVenta.Show
End Sub

Private Sub Veriform_Click()
    PrgVeriForm.Show
End Sub

Private Sub zfdasf_Click()
    WPasaMoneda = 0
    PrgFactuSalva.Show
End Sub
