VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM LiberaTerminado"
    spLiberaTerminado = Sql1 + Sql2
    Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstLiberaTerminado.RecordCount > 0 Then
        rstLiberaTerminado.MoveLast
        WCodigoMayor = IIf(IsNull(rstLiberaTerminado!CodigoMayor), "0", rstLiberaTerminado!CodigoMayor)
        Lote = Str$(WCodigoMayor)
        rstLiberaTerminado.Close
            Else
        Lote = "0"
    End If
        
    WCodigo = Str$(Val(Lote) + 1)
    WProducto = Producto.Text
    WFecha = fecha.Text
    WFechaord = Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
    WPartida = Partida.Text
    WPartiOri = ""
    WValor1 = Valor1.Text
    WValor2 = valor2.Text
    WValor3 = Valor3.Text
    WValor4 = valor4.Text
    WValor5 = valor5.Text
    WValor6 = valor6.Text
    WValor7 = valor7.Text
    WValor8 = valor8.Text
    WValor9 = valor9.Text
    WValor10 = valor10.Text
    WEnsayo = Ensayo.Text
    WAspecto = Aspecto.Text
    WObservaciones = Observaciones.Text
    WConfecciono = Confecciono.Text
    WMarca = "N"
    WCliente = Cliente.Text
    WObserva = Observa.Text
    WCantidad = Cantidad.Text
    WOrigen = "L"
    WTipo = "PT"
    WImpreProdI = "N"
    WImpreProdII = "N"
    WImpreProdIII = "N"
    WImpreVentas = "N"
    WTipoPro = ""
            
    XTipoPro = ""
    XCodigo = Val(Mid$(Producto.Text, 4, 5))
    If Left$(Producto.Text, 2) = "DY" Or Left$(Producto.Text, 2) = "DW" Then
        XTipoPro = "CO"
            Else
        If XCodigo >= 0 And XCodigo <= 999 Then
            XTipoPro = "CO"
                Else
            If XCodigo >= 11000 And XCodigo <= 11999 Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 25000 And XCodigo <= 25999 Then
                    XTipoPro = "FA"
                        Else
                    If XCodigo >= 2300 And XCodigo <= 2399 Then
                        XTipoPro = "BI"
                            Else
                        XTipoPro = "PT"
                    End If
                End If
            End If
        End If
    End If
                
    ZLinea = 0
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZLinea = rstTerminado!Linea
        rstTerminado.Close
    End If
                
    Select Case ZLinea
        Case 8
            XTipoPro = "PG"
        Case 10, 20
            XTipoPro = "FA"
        Case Else
    End Select
                
    WTipoPro = XTipoPro
            
    Select Case WTipoPro
        Case "CO", "PG"
            WImpreProdI = "S"
        Case "BI", "PT"
            WImpreProdII = "S"
        Case "FA"
            WImpreProdIII = "S"
        Case Else
    End Select
            
            
    ZSql = ""
    ZSql = ZSql & "INSERT INTO LiberaTerminado ("
    ZSql = ZSql & "Codigo, "
    ZSql = ZSql & "Producto, "
    ZSql = ZSql & "Fecha, "
    ZSql = ZSql & "OrdFecha, "
    ZSql = ZSql & "Partida, "
    ZSql = ZSql & "PartiOri, "
    ZSql = ZSql & "Valor1, "
    ZSql = ZSql & "Valor2, "
    ZSql = ZSql & "Valor3, "
    ZSql = ZSql & "Valor4, "
    ZSql = ZSql & "Valor5, "
    ZSql = ZSql & "Valor6, "
    ZSql = ZSql & "Valor7, "
    ZSql = ZSql & "Valor8, "
    ZSql = ZSql & "Valor9, "
    ZSql = ZSql & "Valor10, "
    ZSql = ZSql & "Ensayo, "
    ZSql = ZSql & "Aspecto, "
    ZSql = ZSql & "Observaciones, "
    ZSql = ZSql & "Confecciono, "
    ZSql = ZSql & "Marca, "
    ZSql = ZSql & "Cliente, "
    ZSql = ZSql & "Cantidad, "
    ZSql = ZSql & "Observa, "
    ZSql = ZSql & "Origen, "
    ZSql = ZSql & "Tipo, "
    ZSql = ZSql & "ImpreProdI, "
    ZSql = ZSql & "ImpreProdII, "
    ZSql = ZSql & "ImpreProdIII, "
    ZSql = ZSql & "ImpreVentas, "
    ZSql = ZSql & "TipoPro) "
    ZSql = ZSql & "Values ("
    ZSql = ZSql & "'" + WCodigo + "',"
    ZSql = ZSql & "'" + WProducto + "',"
    ZSql = ZSql & "'" + WFecha + "',"
    ZSql = ZSql & "'" + WOrdFecha + "',"
    ZSql = ZSql & "'" + WPartida + "',"
    ZSql = ZSql & "'" + WPartiOri + "',"
    ZSql = ZSql & "'" + WValor1 + "',"
    ZSql = ZSql & "'" + WValor2 + "',"
    ZSql = ZSql & "'" + WValor3 + "',"
    ZSql = ZSql & "'" + WValor4 + "',"
    ZSql = ZSql & "'" + WValor5 + "',"
    ZSql = ZSql & "'" + WValor6 + "',"
    ZSql = ZSql & "'" + WValor7 + "',"
    ZSql = ZSql & "'" + WValor8 + "',"
    ZSql = ZSql & "'" + WValor9 + "',"
    ZSql = ZSql & "'" + WValor10 + "',"
    ZSql = ZSql & "'" + WEnsayo + "',"
    ZSql = ZSql & "'" + WAspecto + "',"
    ZSql = ZSql & "'" + WObservaciones + "',"
    ZSql = ZSql & "'" + WConfecciono + "',"
    ZSql = ZSql & "'" + WMarca + "',"
    ZSql = ZSql & "'" + WCliente + "',"
    ZSql = ZSql & "'" + WCantidad + "',"
    ZSql = ZSql & "'" + WObserva + "',"
    ZSql = ZSql & "'" + WOrigen + "',"
    ZSql = ZSql & "'" + WTipo + "',"
    ZSql = ZSql & "'" + WImpreProdI + "',"
    ZSql = ZSql & "'" + WImpreProdII + "',"
    ZSql = ZSql & "'" + WImpreProdIII + "',"
    ZSql = ZSql & "'" + WImpreVentas + "',"
    ZSql = ZSql & "'" + WTipoPro + "')"
         
    spLiberaTerminado = ZSql
    Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
