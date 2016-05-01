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

                    
        Rem *******************certificados*****************
    
        Rem
        Rem certificado de analisis
        Rem
        Rem For ZZCiclo = 1 To 12
        Rem sacar dos sentenvcias de abajo
        Rem y restaurar el ciclo ok
    
        Rem ZZRequiereCertificado = 0
        
        If ZZRequiereCertificado = 1 Then
            ZZZZhasta = 12
                Else
            ZZZZhasta = 0
        End If
        
        For ZZCiclo = 1 To ZZZZhasta
            
            Select Case ZZCiclo
                Case 1
                    ZZLugar = 5
                Case 2
                    ZZLugar = 7
                Case 3
                    ZZLugar = 9
                Case 4
                    ZZLugar = 11
                Case 5
                    ZZLugar = 13
                Case 6
                    ZZLugar = 15
                Case 7
                    ZZLugar = 17
                Case 8
                    ZZLugar = 19
                Case 9
                    ZZLugar = 21
                Case 10
                    ZZLugar = 23
                Case 11
                    ZZLugar = 25
                Case Else
                    ZZLugar = 27
            End Select
            
            If Val(Auxiliar(DA, ZZLugar)) <> 0 Then
        
                ZZEntra = "N"
        
                If Left$(UCase(Articulo), 2) = "PT" Then
                
                    XCodigo = Val(Mid$(Articulo, 4, 5))
                    If XCodigo >= 0 And XCodigo <= 999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 11000 And XCodigo <= 12999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 25000 And XCodigo <= 25999 Then
                                XTipoPro = "FA"
                                    Else
                                If XCodigo >= 2300 And XCodigo <= 2399 Then
                                    XTipoPro = "BI"
                                        Else
                                    If XCodigo >= 40000 And XCodigo <= 41000 Then
                                        XTipoPro = "TA"
                                            Else
                                        XTipoPro = "PT"
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                    If Left$(Articulo, 2) = "YQ" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YH" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YP" Then
                        XTipoPro = "PT"
                    End If
                    If Left$(Articulo, 2) = "YF" Then
                        XTipoPro = "FA"
                    End If
            
                    ZLinea = 0
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
            
                    Select Case ZLinea
                        Case 8
                            XTipoPro = "PG"
                        Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                            XTipoPro = "FA"
                        Case Else
                    End Select
            
                    If XTipoPro <> "FA" And XTipoPro <> "TA" Then
                    Rem If XTipoPro = "CO" Then
                    
                        XEmpresa = Wempresa
                        
                        Select Case Val(Wempresa)
                            Case 1, 3, 5, 6, 7, 10, 11
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                Wempresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = Auxiliar(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = Auxiliar(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                        ZArticulo = Articulo
                        ZProducto = Articulo
                        ZLote = Auxiliar(DA, ZZLugar)
                        Rem ZCantidad = Cantidad
                        ZCantidad = Auxiliar(DA, ZZLugar + 1)
                        ZCliente = Cliente.Text
                            
                            
                        Erase ZOpcion
                        Erase ZValor
                        Erase ZEnsayo
                        Erase ZStd
                        Erase ZDescri
                        Erase ZDescriII
                            
                        WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        ZVersion = 0
                        
                        ZZEntra = "N"
                        
                        ZSql = ""
                        ZSql = ZSql & "Select *"
                        ZSql = ZSql & " FROM AltaCertificado"
                        ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                        ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZCliente + "'"
                        spAltaCertificado = ZSql
                        Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstAltaCertificado.RecordCount > 0 Then
                            ZOpcion(1) = rstAltaCertificado!Opcion1
                            ZOpcion(2) = rstAltaCertificado!Opcion2
                            ZOpcion(3) = rstAltaCertificado!Opcion3
                            ZOpcion(4) = rstAltaCertificado!Opcion4
                            ZOpcion(5) = rstAltaCertificado!Opcion5
                            ZOpcion(6) = rstAltaCertificado!Opcion6
                            ZOpcion(7) = rstAltaCertificado!Opcion7
                            ZOpcion(8) = rstAltaCertificado!Opcion8
                            ZOpcion(9) = rstAltaCertificado!Opcion9
                            ZOpcion(10) = rstAltaCertificado!Opcion10
                            rstAltaCertificado.Close
                            ZZEntra = "S"
                        End If
                        
                        If ZZEntra = "N" Then
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM AltaCertificado"
                            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "S00102" + "'"
                            spAltaCertificado = ZSql
                            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstAltaCertificado.RecordCount > 0 Then
                                ZOpcion(1) = rstAltaCertificado!Opcion1
                                ZOpcion(2) = rstAltaCertificado!Opcion2
                                ZOpcion(3) = rstAltaCertificado!Opcion3
                                ZOpcion(4) = rstAltaCertificado!Opcion4
                                ZOpcion(5) = rstAltaCertificado!Opcion5
                                ZOpcion(6) = rstAltaCertificado!Opcion6
                                ZOpcion(7) = rstAltaCertificado!Opcion7
                                ZOpcion(8) = rstAltaCertificado!Opcion8
                                ZOpcion(9) = rstAltaCertificado!Opcion9
                                ZOpcion(10) = rstAltaCertificado!Opcion10
                                rstAltaCertificado.Close
                                ZZEntra = "S"
                            End If
                        End If
                        
                        If ZZEntra = "S" Then
                            If ZOpcion(1) = 0 And ZOpcion(2) = 0 And ZOpcion(3) = 0 And ZOpcion(4) = 0 And ZOpcion(5) = 0 And ZOpcion(6) = 0 And ZOpcion(7) = 0 And ZOpcion(8) = 0 And ZOpcion(9) = 0 And ZOpcion(10) = 0 Then
                                ZZEntra = "N"
                            End If
                        End If
                        
                        If ZZEntra = "N" Then
                            m$ = "El Certificado de Analisis de " + Articulo + " no se ha encontrado"
                            a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                        End If
                  
                        If ZZEntra = "S" Then
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    CargaEmpresa(1, 1) = "0001"
                                    CargaEmpresa(1, 2) = "Empresa01"
                                    CargaEmpresa(2, 1) = "0003"
                                    CargaEmpresa(2, 2) = "Empresa03"
                                    CargaEmpresa(3, 1) = "0005"
                                    CargaEmpresa(3, 2) = "Empresa05"
                                    CargaEmpresa(4, 1) = "0006"
                                    CargaEmpresa(4, 2) = "Empresa06"
                                    CargaEmpresa(5, 1) = "0007"
                                    CargaEmpresa(5, 2) = "Empresa07"
                                    CargaEmpresa(6, 1) = "0010"
                                    CargaEmpresa(6, 2) = "Empresa10"
                                    CargaEmpresa(7, 1) = "0011"
                                    CargaEmpresa(7, 2) = "Empresa11"
                                    ZHasta1 = 7
                                Case Else
                                    CargaEmpresa(1, 1) = "0002"
                                    CargaEmpresa(1, 2) = "Empresa02"
                                    CargaEmpresa(2, 1) = "0004"
                                    CargaEmpresa(2, 2) = "Empresa04"
                                    CargaEmpresa(3, 1) = "0008"
                                    CargaEmpresa(3, 2) = "Empresa08"
                                    CargaEmpresa(4, 1) = "0009"
                                    CargaEmpresa(4, 2) = "Empresa09"
                                    ZHasta1 = 4
                            End Select
                                
                            For ZCiclo = 1 To ZHasta1
                                
                                Wempresa = CargaEmpresa(ZCiclo, 1)
                                txtOdbc = CargaEmpresa(ZCiclo, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        
                                If Val(ZLote) > 99999 Then
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 6)
                                        Else
                                    ZZLote = ZLote
                                    Call Ceros(ZZLote, 5)
                                End If
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Prueter"
                                ZSql = ZSql + " Where Prueter.Lote = " + "'" + ZLote + "'"
                                spPrueter = ZSql
                                Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
                                If rstPrueter.RecordCount > 0 Then
                                
                                    If Left$(rstPrueter!prueba, 1) = "2" Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                    If rstPrueter!Producto <> ZProducto Then
                                        rstPrueter.Close
                                        Exit Sub
                                    End If
                                        
                                        
                                    WFechaord = Right$(rstPrueter!Fecha, 4) + Mid$(rstPrueter!Fecha, 4, 2) + Left$(rstPrueter!Fecha, 2)
                                            
                                    ZValor(1) = rstPrueter!Valor1
                                    ZValor(2) = rstPrueter!valor2
                                    ZValor(3) = rstPrueter!Valor3
                                    ZValor(4) = rstPrueter!valor4
                                    ZValor(5) = rstPrueter!valor5
                                    ZValor(6) = rstPrueter!valor6
                                    ZValor(7) = rstPrueter!valor7
                                    ZValor(8) = rstPrueter!valor8
                                    ZValor(9) = rstPrueter!valor9
                                    ZValor(10) = rstPrueter!valor10
                                        
                                    rstPrueter.Close
                                    
                                    WFechaElaboracion = ""
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Hoja"
                                    ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZLote + "'"
                                    spHoja = ZSql
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        Rem WFechaElaboracion = Mid$(rstHoja!fechaIng, 4, 7)
                                        ZZHoja = rstHoja!Hoja
                                        ZZProducto = rstHoja!Producto
                                        ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                                        ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                                        ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                                        
                                        If ZZFechaRevalida <> "  /  /    " And ZZFechaRevalida <> "00/00/0000" Then
                                            WFecha = ZZFechaRevalida
                                            WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        End If
                                        
                                        ZZFecha = rstHoja!Fecha
                                        ZZMeses = ""
                                        rstHoja.Close
                                        
                                        spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                                            rstTerminado.Close
                                        End If
                                        
                                        If Val(ZZMeses) <> 0 Then
                                        
                                            If Val(ZZRevalida) <> 0 Then
                                                WVida = Val(ZZMesesRevalida)
                                                WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                                                WAno = Val(Right$(ZZFechaRevalida, 4))
                                                    Else
                                                WVida = Val(ZZMeses)
                                                WMes = Val(Mid$(ZZFecha, 4, 2))
                                                WAno = Val(Right$(ZZFecha, 4))
                                            End If
                                            
                                            For Ciclo = 1 To WVida
                                                WMes = WMes + 1
                                                If WMes > 12 Then
                                                    WAno = WAno + 1
                                                    WMes = 1
                                                End If
                                            Next Ciclo
                                            ZMes = Str$(WMes)
                                            ZAno = Str$(WAno)
                                            Call Ceros(ZMes, 2)
                                            Call Ceros(ZAno, 4)
                                            WFechaElaboracion = ZMes + "/" + ZAno
                                            
                                        End If
                                        
                                    End If
                                        
                                    If Left$(ZArticulo, 2) = "SE" Then
                                        WProducto = "SE" + Mid$(ZArticulo, 3, 10)
                                            Else
                                        WProducto = "PT" + Mid$(ZArticulo, 3, 10)
                                    End If
                                        
                                    Select Case Val(Wempresa)
                                        Case 1, 3, 5, 6, 7, 10, 11
                                            Wempresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            Wempresa = "0004"
                                            txtOdbc = "Empresa04"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                        
                                    LlamaImprime = "N"
                                    
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM EspecifUnificaVersion"
                                    ZSql = ZSql + " Where EspecifUnificaVersion.Producto = " + "'" + WProducto + "'"
                                    ZSql = ZSql + " Order by EspecifUnificaVersion.Producto, EspecifUnificaVersion.Version"
                                    spEspecifUnificaVersion = ZSql
                                    Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEspecifUnificaVersion.RecordCount > 0 Then
                                        With rstEspecifUnificaVersion
                                            .MoveFirst
                                            Do
                                                If .EOF = False Then
                                                    
                                                    WDesde = Right$(rstEspecifUnificaVersion!FechaInicio, 4) + Mid$(rstEspecifUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecifUnificaVersion!FechaInicio, 2)
                                                    WHasta = Right$(rstEspecifUnificaVersion!FechaFinal, 4) + Mid$(rstEspecifUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecifUnificaVersion!FechaFinal, 2)
                                                            
                                                    If WDesde <= WFechaord And WHasta >= WFechaord Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                                
                                                    End If
                                                        
                                                    If WDesde > WFechaord And LlamaImprime = "N" Then
                                                            
                                                        ZEnsayo(1) = rstEspecifUnificaVersion!Ensayo1
                                                        ZEnsayo(2) = rstEspecifUnificaVersion!Ensayo2
                                                        ZEnsayo(3) = rstEspecifUnificaVersion!Ensayo3
                                                        ZEnsayo(4) = rstEspecifUnificaVersion!Ensayo4
                                                        ZEnsayo(5) = rstEspecifUnificaVersion!Ensayo5
                                                        ZEnsayo(6) = rstEspecifUnificaVersion!Ensayo6
                                                        ZEnsayo(7) = rstEspecifUnificaVersion!Ensayo7
                                                        ZEnsayo(8) = rstEspecifUnificaVersion!Ensayo8
                                                        ZEnsayo(9) = rstEspecifUnificaVersion!Ensayo9
                                                        ZEnsayo(10) = rstEspecifUnificaVersion!Ensayo10
                                                                
                                                        ZStd(1, 1) = rstEspecifUnificaVersion!Valor1
                                                        ZStd(2, 1) = rstEspecifUnificaVersion!valor2
                                                        ZStd(3, 1) = rstEspecifUnificaVersion!Valor3
                                                        ZStd(4, 1) = rstEspecifUnificaVersion!valor4
                                                        ZStd(5, 1) = rstEspecifUnificaVersion!valor5
                                                        ZStd(6, 1) = rstEspecifUnificaVersion!valor6
                                                        ZStd(7, 1) = rstEspecifUnificaVersion!valor7
                                                        ZStd(8, 1) = rstEspecifUnificaVersion!valor8
                                                        ZStd(9, 1) = rstEspecifUnificaVersion!valor9
                                                        ZStd(10, 1) = rstEspecifUnificaVersion!valor10
                                                                
                                                        ZStd(1, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor11), "", rstEspecifUnificaVersion!Valor11)
                                                        ZStd(2, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor22), "", rstEspecifUnificaVersion!Valor22)
                                                        ZStd(3, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor33), "", rstEspecifUnificaVersion!Valor33)
                                                        ZStd(4, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor44), "", rstEspecifUnificaVersion!Valor44)
                                                        ZStd(5, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor55), "", rstEspecifUnificaVersion!Valor55)
                                                        ZStd(6, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor66), "", rstEspecifUnificaVersion!Valor66)
                                                        ZStd(7, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor77), "", rstEspecifUnificaVersion!Valor77)
                                                        ZStd(8, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor88), "", rstEspecifUnificaVersion!Valor88)
                                                        ZStd(9, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor99), "", rstEspecifUnificaVersion!Valor99)
                                                        ZStd(10, 2) = IIf(IsNull(rstEspecifUnificaVersion!Valor1010), "", rstEspecifUnificaVersion!Valor1010)
                                                                
                                                        ZStd(1, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde1), "", rstEspecifUnificaVersion!Desde1)
                                                        ZStd(2, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde2), "", rstEspecifUnificaVersion!Desde2)
                                                        ZStd(3, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde3), "", rstEspecifUnificaVersion!Desde3)
                                                        ZStd(4, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde4), "", rstEspecifUnificaVersion!Desde4)
                                                        ZStd(5, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde5), "", rstEspecifUnificaVersion!Desde5)
                                                        ZStd(6, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde6), "", rstEspecifUnificaVersion!Desde6)
                                                        ZStd(7, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde7), "", rstEspecifUnificaVersion!Desde7)
                                                        ZStd(8, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde8), "", rstEspecifUnificaVersion!Desde8)
                                                        ZStd(9, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde9), "", rstEspecifUnificaVersion!Desde9)
                                                        ZStd(10, 3) = IIf(IsNull(rstEspecifUnificaVersion!Desde10), "", rstEspecifUnificaVersion!Desde10)
                                                                
                                                        ZStd(1, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta1), "", rstEspecifUnificaVersion!Hasta1)
                                                        ZStd(2, 4) = IIf(IsNull(rstEspecifUnificaVersion!HAsta2), "", rstEspecifUnificaVersion!HAsta2)
                                                        ZStd(3, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta3), "", rstEspecifUnificaVersion!Hasta3)
                                                        ZStd(4, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta4), "", rstEspecifUnificaVersion!Hasta4)
                                                        ZStd(5, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta5), "", rstEspecifUnificaVersion!Hasta5)
                                                        ZStd(6, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta6), "", rstEspecifUnificaVersion!Hasta6)
                                                        ZStd(7, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta7), "", rstEspecifUnificaVersion!Hasta7)
                                                        ZStd(8, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta8), "", rstEspecifUnificaVersion!Hasta8)
                                                        ZStd(9, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta9), "", rstEspecifUnificaVersion!Hasta9)
                                                        ZStd(10, 4) = IIf(IsNull(rstEspecifUnificaVersion!Hasta10), "", rstEspecifUnificaVersion!Hasta10)
                                                                
                                                        ZVersion = rstEspecifUnificaVersion!Version
                                                        LlamaImprime = "S"
                                                    End If
                                                    
                                                    .MoveNext
                                                        Else
                                                    Exit Do
                                                End If
                                            Loop
                                        End With
                                        rstEspecifUnificaVersion.Close
                                    End If
                                    
                                    If LlamaImprime = "N" Then
                                    
                                        Sql1 = "Select Ensayo1,ensayo2,ensayo3,ensayo4,ensayo5,ensayo6,ensayo7,ensayo8,ensayo9,ensayo10,valor1,valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010"
                                        Sql2 = " FROM EspecifUnifica"
                                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEspecifUnifica.RecordCount > 0 Then
                                                
                                            ZEnsayo(1) = rstEspecifUnifica!Ensayo1
                                            ZEnsayo(2) = rstEspecifUnifica!Ensayo2
                                            ZEnsayo(3) = rstEspecifUnifica!Ensayo3
                                            ZEnsayo(4) = rstEspecifUnifica!Ensayo4
                                            ZEnsayo(5) = rstEspecifUnifica!Ensayo5
                                            ZEnsayo(6) = rstEspecifUnifica!Ensayo6
                                            ZEnsayo(7) = rstEspecifUnifica!Ensayo7
                                            ZEnsayo(8) = rstEspecifUnifica!Ensayo8
                                            ZEnsayo(9) = rstEspecifUnifica!Ensayo9
                                            ZEnsayo(10) = rstEspecifUnifica!Ensayo10
                                                                
                                            ZStd(1, 1) = rstEspecifUnifica!Valor1
                                            ZStd(2, 1) = rstEspecifUnifica!valor2
                                            ZStd(3, 1) = rstEspecifUnifica!Valor3
                                            ZStd(4, 1) = rstEspecifUnifica!valor4
                                            ZStd(5, 1) = rstEspecifUnifica!valor5
                                            ZStd(6, 1) = rstEspecifUnifica!valor6
                                            ZStd(7, 1) = rstEspecifUnifica!valor7
                                            ZStd(8, 1) = rstEspecifUnifica!valor8
                                            ZStd(9, 1) = rstEspecifUnifica!valor9
                                            ZStd(10, 1) = rstEspecifUnifica!valor10
                                                                
                                            ZStd(1, 2) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                                            ZStd(2, 2) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                                            ZStd(3, 2) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                                            ZStd(4, 2) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                                            ZStd(5, 2) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                                            ZStd(6, 2) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                                            ZStd(7, 2) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                                            ZStd(8, 2) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                                            ZStd(9, 2) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                                            ZStd(10, 2) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                                            
                                           Rem by nan
                                           rstEspecifUnifica.Close
                                          
                                          
                                        End If
                                          
                                        Rem by nan 21-5-2014
                                        Sql1 = "Select desde1,desde2,desde3,desde4,desde5,desde6,desde7,desde8,desde9,desde10,hasta1,hasta2,hasta3,hasta4,hasta5,hasta6,hasta7,hasta8,hasta9,hasta10,version"
                                        Sql2 = " FROM EspecifUnifica"
                                        Sql3 = " Where EspecifUnifica.Producto = " + "'" + ZProducto + "'"
                                        spEspecifUnifica = Sql1 + Sql2 + Sql3
                                        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEspecifUnifica.RecordCount > 0 Then
                                                  
                                                                      
                                            ZStd(1, 3) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                                            ZStd(2, 3) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                                            ZStd(3, 3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                                            ZStd(4, 3) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                                            ZStd(5, 3) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                                            ZStd(6, 3) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                                            ZStd(7, 3) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                                            ZStd(8, 3) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                                            ZStd(9, 3) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                                            ZStd(10, 3) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                                                                                 
                                            ZStd(1, 4) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                                            ZStd(2, 4) = IIf(IsNull(rstEspecifUnifica!HAsta2), "", rstEspecifUnifica!HAsta2)
                                            ZStd(3, 4) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                                            ZStd(4, 4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                                            ZStd(5, 4) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                                            ZStd(6, 4) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                                            ZStd(7, 4) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                                            ZStd(8, 4) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                                            ZStd(9, 4) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                                            ZStd(10, 4) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                                                                
                                            ZVersion = rstEspecifUnifica!Version
                                            rstEspecifUnifica.Close
                                            LlamaImprime = "S"
                                        End If
                                
                                    End If
                                    
                                    If LlamaImprime = "S" Then
                                        
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(1) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(1) = rstEnsayo!Descripcion
                                            ZDescriII(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(2) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(2) = rstEnsayo!Descripcion
                                            ZDescriII(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(3) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(3) = rstEnsayo!Descripcion
                                            ZDescriII(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(4) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(4) = rstEnsayo!Descripcion
                                            ZDescriII(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(5) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(5) = rstEnsayo!Descripcion
                                            ZDescriII(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(6) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(6) = rstEnsayo!Descripcion
                                            ZDescriII(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(7) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(7) = rstEnsayo!Descripcion
                                            ZDescriII(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(8) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(8) = rstEnsayo!Descripcion
                                            ZDescriII(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(9) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(9) = rstEnsayo!Descripcion
                                            ZDescriII(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                            
                                        spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(10) + "'"
                                        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnsayo.RecordCount > 0 Then
                                            ZDescri(10) = rstEnsayo!Descripcion
                                            ZDescriII(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
                                            rstEnsayo.Close
                                        End If
                                                
                                        Call Conecta_Empresa
                                        
                                        XEmpresa = Wempresa
                                        Select Case Val(XEmpresa)
                                            Case 1, 3, 5, 6, 7, 10, 11
                                                Wempresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case 2, 4, 8, 9
                                                Wempresa = "0008"
                                                txtOdbc = "Empresa08"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                        End Select
                                        
                                        ZImpreVto = 0
                                        ZRazon = ""
                                        spCliente = "ConsultaCliente " + "'" + ZCliente + "'"
                                        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstCliente.RecordCount > 0 Then
                                            ZRazon = Left$(rstCliente!Razon, 50)
                                            ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                                            rstCliente.Close
                                        End If
                                        
                                        ZZImpreVtoTermi = 0
                                        spTerminado = "ConsultaTerminado " + "'" + ZArticulo + "'"
                                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstTerminado.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstTerminado!Descripcion), "", rstTerminado!Descripcion)
                                            ZZImpreVtoTermi = IIf(IsNull(rstTerminado!ImpreVto), "0", rstTerminado!ImpreVto)
                                            rstTerminado.Close
                                        End If
                                        
                                        Rem If ZZImpreVtoTermi = 0 Then
                                        Rem     If ZImpreVto <> 1 Then
                                        Rem         Rem WFechaElaboracion = ""
                                        Rem     End If
                                        Rem End If
                        
                                        Rem
                                        Rem SI ES COLORANTE NO IMPRIME
                                        Rem LA FECHA DE VENCIMIENTO
                                        Rem
                                        XCodigo = Val(Mid$(ZProducto, 4, 5))
                                        XTipoPro = ""
                                        If Val(Wempresa) = 1 Then
                                            If XCodigo >= 0 And XCodigo <= 999 Then
                                                WFechaElaboracion = ""
                                                XTipoPro = "CO"
                                                    Else
                                                If XCodigo >= 11000 And XCodigo <= 12999 Then
                                                    WFechaElaboracion = ""
                                                    XTipoPro = "CO"
                                                        Else
                                                    XTipoPro = ""
                                                End If
                                            End If
                                        End If
                                        
                                        
                                        
                                        
                                            
                                        ZCliente = UCase(ZCliente)
                                        ZArticulo = UCase(ZArticulo)
                                        ZClave = ZCliente + ZArticulo
                        
                                        spPrecios = "ConsultaPrecios " + "'" + ZClave + "'"
                                        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstPrecios.RecordCount > 0 Then
                                            ZDesArticulo = IIf(IsNull(rstPrecios!Descripcion), "", rstPrecios!Descripcion)
                                            rstPrecios.Close
                                        End If
                                        
                                        Call Conecta_Empresa
                                                
                                        ZSql = "DELETE Certificado"
                                        spCertificado = ZSql
                                        Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                            
                                        LugarMetodo = 0
                                                
                                        For CiclaMetodo = 1 To 10
                                                
                                            If ZOpcion(CiclaMetodo) = 1 Then
                                                
                                                LugarMetodo = LugarMetodo + 1
                                                    
                                                ZOrden = ""
                                                ZClave1 = ZLote
                                                Call Ceros(ZClave1, 6)
                                                ZClave2 = Str$(LugarMetodo)
                                                Call Ceros(ZClave2, 2)
                                                ZClave = ZClave1 + ZClave2
                                                ZMetodo = ZEnsayo(CiclaMetodo)
                                                
                                                If Val(ZStd(CiclaMetodo, 3)) <> 0 And Val(ZStd(CiclaMetodo, 4)) <> 0 Then
                                                    ZValorNormalI = Trim(ZStd(CiclaMetodo, 3)) + " - " + Trim(ZStd(CiclaMetodo, 4)) + " " + Trim(ZDescriII(CiclaMetodo)) + " " + Left$(ZStd(CiclaMetodo, 1), 50)
                                                    ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                                        Else
                                                    ZValorNormalI = Left$(ZStd(CiclaMetodo, 1), 50)
                                                    ZValorNormalII = Left$(ZStd(CiclaMetodo, 2), 50)
                                                End If
                                                ZValorPartidaI = Left$(ZValor(CiclaMetodo), 80)
                                                
                                                ZValorNormalI = Trim(ZValorNormalI)
                                                ZCanti = 80 - Len(ZValorNormalI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalI = Space$(ZCanti) + ZValorNormalI
                                                
                                                ZValorNormalII = Trim(ZValorNormalII)
                                                ZCanti = 80 - Len(ZValorNormalII)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorNormalII = Space$(ZCanti) + ZValorNormalII
                                                
                                                ZValorPartidaI = Trim(ZValorPartidaI)
                                                ZCanti = 80 - Len(ZValorPartidaI)
                                                ZCanti = Int(ZCanti / 2)
                                                ZValorPartidaI = Space$(ZCanti) + ZValorPartidaI
                                                
                                                ZValorPartidaII = ""
                                                ZObservacionesI = ""
                                                ZObservacionesII = ""
                                                ZObservacionesIII = "Version " + ZVersion
                                                ZObservacionesIV = ""
                                                ZObservacionesV = ""
                                                ZObservacionesVI = ""
                                                If Val(Wempresa) = 1 Then
                                                    ZEmpresa = "Surfactan S.A."
                                                        Else
                                                    ZEmpresa = "Pellital S.A."
                                                End If
                                                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                                ZFechaII = WFechaElaboracion
                                                
                                                ZExamen = Trim(ZDescri(CiclaMetodo))
                                                ZExamenII = ""
                                                ZHasta = Len(Trim(ZExamen))
                                                If ZHasta > 25 Then
                                                    For Cicla = ZHasta To 1 Step -1
                                                        If Mid(ZExamen, Cicla, 1) = Space(1) Then
                                                            ZExamenII = Mid(ZExamen, Cicla + 1, 25)
                                                            ZExamen = Mid(ZExamen, 1, Cicla)
                                                            Exit For
                                                        End If
                                                    Next Cicla
                                                End If
                                                        
                                                        
                                                        
                                                aa = Auxiliar(1, 1)
                                                aa = Auxiliar(1, 2)
                                                aa = Auxiliar(1, 3)
                                                aa = Auxiliar(1, 4)
                                                aa = Auxiliar(1, 5)
                                                aa = Auxiliar(1, 6)
                                                aa = Auxiliar(1, 7)
                                                aa = Auxiliar(1, 8)
                                                aa = Auxiliar(1, 9)
                                                aa = Auxiliar(1, 10)
                                                aa = Auxiliar(1, 11)
                                                aa = Auxiliar(1, 12)
                                                aa = Auxiliar(1, 13)
                                                aa = Auxiliar(1, 14)
                                                aa = Auxiliar(1, 15)
                                                aa = Auxiliar(1, 16)
                                                aa = Auxiliar(1, 17)
                                                aa = Auxiliar(1, 18)
                                                aa = Auxiliar(1, 19)
                                                aa = Auxiliar(1, 20)
                                                        
                                                        
                                                ZSql = ""
                                                ZSql = ZSql + "INSERT INTO Certificado ("
                                                ZSql = ZSql + "Clave ,"
                                                ZSql = ZSql + "Partida ,"
                                                ZSql = ZSql + "Renglon ,"
                                                ZSql = ZSql + "Razon ,"
                                                ZSql = ZSql + "Orden ,"
                                                ZSql = ZSql + "Terminado ,"
                                                ZSql = ZSql + "Descripcion ,"
                                                ZSql = ZSql + "Fecha ,"
                                                ZSql = ZSql + "FechaII ,"
                                                ZSql = ZSql + "Cantidad ,"
                                                ZSql = ZSql + "Examen ,"
                                                ZSql = ZSql + "ExamenII ,"
                                                ZSql = ZSql + "ValorPartidaI ,"
                                                ZSql = ZSql + "ValorPartidaII ,"
                                                ZSql = ZSql + "ValorNormalI ,"
                                                ZSql = ZSql + "ValorNormalII ,"
                                                ZSql = ZSql + "Observaciones1 ,"
                                                ZSql = ZSql + "Observaciones2 ,"
                                                ZSql = ZSql + "Observaciones3 ,"
                                                ZSql = ZSql + "Observaciones4 ,"
                                                ZSql = ZSql + "Observaciones5 ,"
                                                ZSql = ZSql + "Observaciones6 ,"
                                                ZSql = ZSql + "Metodo ,"
                                                ZSql = ZSql + "Empresa )"
                                                ZSql = ZSql + "Values ("
                                                ZSql = ZSql + "'" + ZClave + "',"
                                                ZSql = ZSql + "'" + ZLote + "',"
                                                ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                                ZSql = ZSql + "'" + ZRazon + "',"
                                                ZSql = ZSql + "'" + ZOrden + "',"
                                                ZSql = ZSql + "'" + ZArticulo + "',"
                                                ZSql = ZSql + "'" + ZDesArticulo + "',"
                                                ZSql = ZSql + "'" + ZFecha + "',"
                                                ZSql = ZSql + "'" + ZFechaII + "',"
                                                ZSql = ZSql + "'" + ZCantidad + "',"
                                                ZSql = ZSql + "'" + ZExamen + "',"
                                                ZSql = ZSql + "'" + ZExamenII + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                                ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                                ZSql = ZSql + "'" + ZValorNormalI + "',"
                                                ZSql = ZSql + "'" + ZValorNormalII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesI + "',"
                                                ZSql = ZSql + "'" + ZObservacionesII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                                ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesV + "',"
                                                ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                                ZSql = ZSql + "'" + ZMetodo + "',"
                                                ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                                spCertificado = ZSql
                                                Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                        
                                            End If
                                                            
                                        Next CiclaMetodo
                                            
                                        Do
                                            
                                            If LugarMetodo = 10 Then
                                                Exit Do
                                            End If
                                                
                                            LugarMetodo = LugarMetodo + 1
                                                    
                                            ZOrden = ""
                                            ZClave1 = ZLote
                                            Call Ceros(ZClave1, 6)
                                            ZClave2 = Str$(LugarMetodo)
                                            Call Ceros(ZClave2, 2)
                                            ZClave = ZClave1 + ZClave2
                                            ZMetodo = ""
                                            ZExamen = ""
                                            ZValorNormalI = ""
                                            ZValorNormalII = ""
                                            ZValorPartidaI = ""
                                            ZValorPartidaII = ""
                                            ZObservacionesI = ""
                                            ZObservacionesII = ""
                                            ZObservacionesIII = "Version " + ZVersion
                                            ZObservacionesIV = ""
                                            ZObservacionesV = ""
                                            ZObservacionesVI = ""
                                            If Val(Wempresa) = 1 Then
                                                ZEmpresa = "Surfactan S.A."
                                                    Else
                                                ZEmpresa = "Pellital S.A."
                                            End If
                                            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                                            ZFechaII = WFechaElaboracion
                                            ZExamenII = ""
                                                        
                                            ZSql = ""
                                            ZSql = ZSql + "INSERT INTO Certificado ("
                                            ZSql = ZSql + "Clave ,"
                                            ZSql = ZSql + "Partida ,"
                                            ZSql = ZSql + "Renglon ,"
                                            ZSql = ZSql + "Razon ,"
                                            ZSql = ZSql + "Orden ,"
                                            ZSql = ZSql + "Terminado ,"
                                            ZSql = ZSql + "Descripcion ,"
                                            ZSql = ZSql + "Fecha ,"
                                            ZSql = ZSql + "Cantidad ,"
                                            ZSql = ZSql + "Examen ,"
                                            ZSql = ZSql + "ValorPartidaI ,"
                                            ZSql = ZSql + "ValorPartidaII ,"
                                            ZSql = ZSql + "ValorNormalI ,"
                                            ZSql = ZSql + "ValorNormalII ,"
                                            ZSql = ZSql + "Observaciones1 ,"
                                            ZSql = ZSql + "Observaciones2 ,"
                                            ZSql = ZSql + "Observaciones3 ,"
                                            ZSql = ZSql + "Observaciones4 ,"
                                            ZSql = ZSql + "Observaciones5 ,"
                                            ZSql = ZSql + "Observaciones6 ,"
                                            ZSql = ZSql + "Metodo ,"
                                            ZSql = ZSql + "Empresa )"
                                            ZSql = ZSql + "Values ("
                                            ZSql = ZSql + "'" + ZClave + "',"
                                            ZSql = ZSql + "'" + ZLote + "',"
                                            ZSql = ZSql + "'" + Str$(CiclaMetodo) + "',"
                                            ZSql = ZSql + "'" + ZRazon + "',"
                                            ZSql = ZSql + "'" + ZOrden + "',"
                                            ZSql = ZSql + "'" + ZArticulo + "',"
                                            ZSql = ZSql + "'" + ZDesArticulo + "',"
                                            ZSql = ZSql + "'" + ZFecha + "',"
                                            ZSql = ZSql + "'" + ZCantidad + "',"
                                            ZSql = ZSql + "'" + ZExamen + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaI + "',"
                                            ZSql = ZSql + "'" + ZValorPartidaII + "',"
                                            ZSql = ZSql + "'" + ZValorNormalI + "',"
                                            ZSql = ZSql + "'" + ZValorNormalII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesI + "',"
                                            ZSql = ZSql + "'" + ZObservacionesII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIII + "',"
                                            ZSql = ZSql + "'" + ZObservacionesIV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesV + "',"
                                            ZSql = ZSql + "'" + ZObservacionesVI + "',"
                                            ZSql = ZSql + "'" + ZMetodo + "',"
                                            ZSql = ZSql + "'" + ZEmpresa + "')"
                            
                                            spCertificado = ZSql
                                            Set rstCertificado = db.OpenRecordset(spCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                                
                                        Loop
                                                
                                        Listado.WindowTitle = "Certificado de Analisis"
                                        Listado.WindowTop = 0
                                        Listado.WindowLeft = 0
                                        Listado.WindowWidth = Screen.Width
                                        Listado.WindowHeight = Screen.Height
                        
                                        Listado.Destination = 1
                                    Rem     Listado.Destination = 0
                                                
                                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                                            Listado.ReportFileName = "CertificadoNuevo.rpt"
                                                Else
                                            Listado.ReportFileName = "CertificadoPelli.rpt"
                                        End If
                                                    
                                        DbConnect = db.Connect
                                        DSQ = getDatabase(DbConnect)
                        
                                        Listado.SQLQuery = "SELECT Certificado.Clave, Certificado.Partida, Certificado.Razon, Certificado.Orden, Certificado.Descripcion, Certificado.Fecha, Certificado.Cantidad, Certificado.Examen, Certificado.ValorPartidaI, Certificado.ValorPartidaII, Certificado.ValorNormalI, Certificado.ValorNormalII, Certificado.Observaciones3, Certificado.Metodo, Certificado.FechaII, Certificado.ExamenII " _
                                                        + "From " _
                                                        + DSQ + ".dbo.Certificado Certificado " _
                                                        + "Where " _
                                                        + "Certificado.Partida >= 0 AND " _
                                                         + "Certificado.Partida <= 999999"
                                                        
                           
                                                        
                                        Listado.Destination = 1
                                        Rem Listado.Destination = 0
                                        Listado.CopiesToPrinter = 1
                                        Rem BY NAN 29-4-2015
                                        ZZDescriArticuloPDF = Left(ZZDescriArticuloPDF, 12)
                                        
                                        If Trim(ZEmailFactura) <> "" Then
                                            Listado.ReportFileName = "Certificadopdf.rpt"
                                            Listado.Destination = crptToFile
                                            Listado.PrintFileType = crptWinWord
                                            Listado.PrintFileName = "c:\pdfprintii\" + ZZDescriArticuloPDF + ZLote + ".doc"
                                            ZZDesdedoc = "c:\pdfprintii\" + ZZDescriArticuloPDF + ZLote + ".doc"
                                            ZZDesdePdf = "c:\pdfprintii\" + ZZDescriArticuloPDF + ZLote + ".pdf"
                                        End If
                        
                                        Listado.Connect = Connect()
                                        Listado.Action = 1
                                   
                                        If Trim(ZEmailFactura) <> "" Then
                                            ZZLugarEnviaII = ZZLugarEnviaII + 1
                                            ZZEnviaPdfII(ZZLugarEnviaII, 1) = Articulo
                                            ZZEnviaPdfII(ZZLugarEnviaII, 2) = ZZDesdePdf
                                            ZZEnviaPdfII(ZZLugarEnviaII, 3) = ZZDesdedoc
                                            ZZEnviaPdfII(ZZLugarEnviaII, 4) = ZZDescriArticulo
                                            ZZEnviaPdfII(ZZLugarEnviaII, 5) = ZLote
                                        End If
                                                
                                    End If
                                          
                                End If
                                    
                            Next ZCiclo
                            
                            Select Case Val(XEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    Wempresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    Wempresa = "0004"
                                    txtOdbc = "Empresa04"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                            
                        End If
                        
                        Call Conecta_Empresa
                        
                    End If
                
                        Else
                        
                    If Left$(Articulo, 2) = "DY" Then
                        
                        Auxi = Mid$(Articulo, 1, 3) + Mid$(Articulo, 6, 7)
                        ZZCambia = "N"
                        
                        ZSql = ""
                        ZSql = ZSql & "Select *"
                        ZSql = ZSql & " FROM Laudo"
                        ZSql = ZSql & " Where Laudo.Laudo = " + "'" + Auxiliar(DA, ZZLugar) + "'"
                        ZSql = ZSql & " and Laudo.Articulo = " + "'" + Auxi + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            ZZPartiOri = Trim(rstLaudo!PartiOri)
                            rstLaudo.Close
                            ZZCambia = "S"
                            ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                            ZZEstado = Dir(ZZRuta)
                            ZZEstado = Trim(ZZEstado)
                            If ZZEstado <> "" Then
                            
                                ZZNombreArchi = 1
                                Do
                                    Auxi = Str$(ZZNombreArchi)
                                    Call Ceros(Auxi, 8)
                                    ZZNombreArchiII = "C:\pdfprint\" + Auxi + ".pdf"
                                    
                                    ZZRutaII = ZZNombreArchiII
                                    ZZEstadoII = Dir(ZZRutaII)
                                    ZZEstadoII = Trim(ZZEstadoII)
                                    If ZZEstadoII = "" Then
                                        Exit Do
                                    End If
                                    ZZNombreArchi = ZZNombreArchi + 1
                                Loop
                                
                                If Trim(ZEmailFactura) <> "" Then
                                    ZZRutaII = "C:\pdfprintII\CertificadodeSeguridad" + ZZDescriArticuloPDF + ZZPartiOri + ".pdf"
                                    ZZRutadoc = "C:\pdfprintII\CertificadodeSeguridad" + ZZDescriArticuloPDF + ZZPartiOri + ".doc"
                                End If
                                
                                FileCopy ZZRuta, ZZRutaII
                                ZZLugarEnvia = ZZLugarEnvia + 1
                                ZZEnviaPdf(ZZLugarEnvia, 1) = Articulo
                                ZZEnviaPdf(ZZLugarEnvia, 2) = ZZRutaII
                                ZZEnviaPdf(ZZLugarEnvia, 3) = ZZRutadoc
                                ZZEnviaPdf(ZZLugarEnvia, 4) = ZZDescriArticulo
                                ZZEnviaPdf(ZZLugarEnvia, 5) = ZZPartidaOri
                                Rem RetVal = Shell("C:\pdfprint\pdfprint " + ZZNombreArchiII, 6)
                                Rem RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docprf " + Chr$(34) + ZZNombreArchiII, 6)
                                ZZImprePdf = "S"
                                Rem TiempoPausa = 2 ' Asigna hora de inicio.
                                Rem Inicio = Timer  ' Establece la hora de inicio.
                                Rem Do While Timer < Inicio + TiempoPausa
                                Rem     DoEvents    ' Cambia a otros procesos.
                                Rem Loop
                            
                                Rem Select Case ZZVersion
                                Rem     Case 1
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t /o" + ZZRuta + " ", 6)
                                Rem     Case 2
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case 3
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case 4
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case 5
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem     Case Else
                                Rem         RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                Rem         Rem RetVal = Shell("C:\Impre\pdfprint " + ZZRuta + " ", 6)
                                Rem End Select
                                
                                    Else
                                 
                                m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                
                            End If
                        End If
                        
                        If ZZCambia = "N" Then
                                        
                            XEmpresa = Wempresa
                                    
                            Wempresa = "0006"
                            txtOdbc = "Empresa06"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM Laudo"
                            ZSql = ZSql & " Where Laudo.Laudo = " + "'" + Auxiliar(DA, ZZLugar) + "'"
                            ZSql = ZSql & " and Laudo.Articulo = " + "'" + Auxi + "'"
                            spLaudo = ZSql
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                ZZPartiOri = Trim(rstLaudo!PartiOri)
                                rstLaudo.Close
                                ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                                ZZEstado = Dir(ZZRuta)
                                ZZEstado = Trim(ZZEstado)
                                If ZZEstado <> "" Then
                                
                                    ZZNombreArchi = 1
                                    Do
                                        Auxi = Str$(ZZNombreArchi)
                                        Call Ceros(Auxi, 8)
                                        ZZNombreArchiII = "C:\pdfprint\" + Auxi + ".pdf"
                                        
                                        ZZRutaII = ZZNombreArchiII
                                        ZZEstadoII = Dir(ZZRutaII)
                                        ZZEstadoII = Trim(ZZEstadoII)
                                        If ZZEstadoII = "" Then
                                            Exit Do
                                        End If
                                        ZZNombreArchi = ZZNombreArchi + 1
                                    Loop
                                    
                                    If Trim(ZEmailFactura) <> "" Then
                                        ZZRutaII = "C:\pdfprintii\CertificadodeSeguridad" + ZZDescriArticulo + ZZPartiOri + ".pdf"
                                        ZZRutadoc = "C:\pdfprintII\CertificadodeSeguridad" + ZZDescriArticuloPDF + ZZPartiOri + ".doc"
                                    End If
                                    
                                    FileCopy ZZRuta, ZZRutaII
                                    ZZLugarEnvia = ZZLugarEnvia + 1
                                    ZZEnviaPdf(ZZLugarEnvia, 1) = Articulo
                                    ZZEnviaPdf(ZZLugarEnvia, 2) = ZZRutaII
                                    ZZEnviaPdf(ZZLugarEnvia, 3) = ZZRutadoc
                                    ZZEnviaPdf(ZZLugarEnvia, 4) = ZZDescriArticulo
                                    ZZEnviaPdf(ZZLugarEnvia, 5) = ZZPartidaOri
                                    Rem RetVal = Shell("C:\pdfprint\pdfprint " + ZZNombreArchiII, 6)
                                    Rem RetVal = Shell("C:\pdfprint\pdfprint -printer " + Chr$(34) + "docprf " + Chr$(34) + ZZNombreArchiII, 6)
                                    ZZImprePdf = "S"
                                    Rem TiempoPausa = 2 ' Asigna hora de inicio.
                                    Rem Inicio = Timer  ' Establece la hora de inicio.
                                    Rem Do While Timer < Inicio + TiempoPausa
                                    Rem     DoEvents    ' Cambia a otros procesos.
                                    Rem Loop
                                
                                    Rem Select Case ZZVersion
                                    Rem     Case 1
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 7.0\Reader\AcroRd32.exe /t /o" + ZZRuta + " ", 6)
                                    Rem     Case 2
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem     Case 3
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem      Case 4
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 8.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem     Case 5
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\Acrobat 9.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem     Case Else
                                    Rem         RetVal = Shell("C:\Archivos de programa\Adobe\reader 10.0\Reader\AcroRd32.exe /t " + ZZRuta + " ", 6)
                                    Rem         Rem RetVal = Shell("C:\Impre\pdfprint " + ZZRuta + " ", 6)
                                    Rem End Select
                                        Else
                                        
                                    m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                    a% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                
                                End If
                            End If
                            
                            Call Conecta_Empresa
                        
                        End If
                        
                    End If
                    
                End If
            End If
                
        Next ZZCiclo
        
    Next DA
    




















Private Sub Calcula_Cae()
    
    Dim WSAA As Object, WSFEX As Object
    Dim dst_cmp  As Integer
    
    
    
    On Error GoTo ManejoError
    
    
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
    
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEXv1
    tra = WSAA.CreateTRA("wsfex")
    Debug.Print tra
    
    
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Rem Path = CurDir() + "\"
    ZPath = "c:\salva\"
    
    Select Case Val(Wempresa)
        Case 1
            ZNombre = "surfa"
            ZCuit = "30549165083"
        Case Else
            ZNombre = "pellital"
            ZCuit = "30610524598"
    End Select
    
    

     ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    
    
    ' Llamar al web service para autenticar:
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms") ' Producción



    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 24 horas
    ' (este período se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electrónica de Exportación
    Set WSFEX = CreateObject("WSFEX")
   
    
    
    ' Setear tocken y sing de autorización (pasos previos)
    WSFEX.Token = WSAA.Token
    WSFEX.Sign = WSAA.Sign
    
    
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEX.Cuit = ZCuit
    
    
    
    ' Conectar al Servicio Web de Facturación
    ok = WSFEX.Conectar("https://servicios1.afip.gov.ar/WSFEX/service.asmx") ' homologación
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEX.Dummy
    Debug.Print "appserver status", WSFEX.AppServerStatus
    Debug.Print "dbserver status", WSFEX.DbServerStatus
    Debug.Print "authserver status", WSFEX.AuthServerStatus
       
    ' Establezco los valores de la factura a autorizar:
    tipo_cbte = 19 ' FC Expo (ver tabla de parámetros)
    Select Case Val(Wempresa)
        Case 1
            punto_vta = 6
        Case Else
            punto_vta = 3
    End Select
    
    
    ' Obtengo el último número de comprobante y le agrego 1
    
    Debug.Print WSFEX.XmlRequest
    Debug.Print WSFEX.XmlResponse
    
    
    Cbte_Nro = WSFEX.GetLastCMP(tipo_cbte, punto_vta) + 1 '16
    ZZComprobante = Cbte_Nro
    
    
    fecha_cbte = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    tipo_expo = 1 ' tipo de exportación (ver tabla de parámetros)
    permiso_existente = "N"
    dst_cmp = Val(ZZPais)
    XXCliente = WRazon
    cuit_pais_cliente = ZZCuit
    domicilio_cliente = WDireccion
    id_impositivo = ZZCuitII
    Rem ZZCuitII
    moneda_id = "DOL" ' para reales, "DOL" o "PES" (ver tabla de parámetros)
    Rem moneda_ctz = 0.5   PARIDAD
    moneda_ctz = Val(Paridad.Text)
    obs_comerciales = "..."
    obs = "..."
    forma_pago = Pago1.Text
    incoterms = CipLista.Text  ' (ver tabla de parámetros)
    incoterms_ds = ""
    idioma_cbte = Idioma.ListIndex  ' (ver tabla de parámetros)
    IMP_TOTAL = Total.Caption
   
    ' Creo una factura (internamente, no se llama al WebService):
    Rem ok = WSFEXv1.CrearFactura(tipo_cbte, punto_vta, Cbte_Nro, fecha_cbte, _
    REM         IMP_TOTAL, tipo_expo, permiso_existente, dst_cmp, _
    REM         XXCliente, cuit_pais_cliente, domicilio_cliente, _
    REM         id_impositivo, moneda_id, moneda_ctz, _
    REM         obs_comerciales, obs, forma_pago, incoterms, _
    REM         idioma_cbte, incoterms_ds)
    
    ' Creo una factura (internamente, no se llama al WebService):
    ok = WSFEX.CrearFactura(tipo_cbte, punto_vta, Cbte_Nro, fecha_cbte, _
            IMP_TOTAL, tipo_expo, permiso_existente, dst_cmp, _
            XXCliente, cuit_pais_cliente, domicilio_cliente, _
            id_impositivo, moneda_id, moneda_ctz, _
            obs_comerciales, obs, forma_pago, incoterms, _
            idioma_cbte)
    
    
    
    
    ' Agrego un item:
    
    For ZZCiclo = 1 To 80
    
        ZZArticulo = ZZVector(ZZCiclo, 1)
        ZZCantidad = ZZVector(ZZCiclo, 2)
        ZZPrecio = ZZVector(ZZCiclo, 3)
        
        If Trim(ZZArticulo) <> "" Then
    
            If Left$(ZZArticulo, 2) = "PT" Or Left$(ZZArticulo, 2) = "PE" Then
                ClavePrecios = Cliente.Text + ZZArticulo
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    ZZDescripcion = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                        Else
                WArti = Left$(ZZArticulo, 3) + Right$(ZZArticulo, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            End If
    
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            Rem If Val(IMP_TOTAL) <> 0 Then
                ' lo agrego a la factura (internamente, no se llama al WebService):
                Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            Rem End If
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
            
        End If
        
    Next ZZCiclo
    
    If Val(Wempresa) <> 1 Then
    
        If Val(Flete.Text) <> 0 Then
    
            ZZArticulo = "Flete"
            ZZCantidad = "1"
            ZZPrecio = Flete.Text
            ZZDescripcion = "Flete"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
        
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
        
        End If
        
        If Val(Seguro.Text) <> 0 Then
    
            ZZArticulo = "Seguro"
            ZZCantidad = "1"
            ZZPrecio = Seguro.Text
            ZZDescripcion = "Seguro"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
        End If
        
        If Val(Gastos.Text) <> 0 Then
    
            ZZArticulo = "Gastos"
            ZZCantidad = "1"
            ZZPrecio = Gastos.Text
            ZZDescripcion = "Gastos"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
        End If
        
        If Val(Descuento.Text) <> 0 Then
    
            ZZArticulo = "Dto"
            ZZCantidad = "1"
            ZZPrecio = Str$(Val(Gastos.Text) * -1)
            ZZDescripcion = "Descuento"
        
            XXCodigo = ZZArticulo
            XXDs = ZZDescripcion
            qty = ZZCantidad
            XXPrecio = ZZPrecio
            umed = 1 ' Ver tabla de parámetros (unidades de medida)
            IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
            Bonif = ""
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            Rem ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, imp_total, Bonif)
            
            ' lo agrego a la factura (internamente, no se llama al WebService):
            ok = WSFEX.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL)
            
        End If
        
        
    End If
    
    
    
    
    
    Rem OJO
    Rem If Val(WEmpresa) = 1 Then
    Rem
    Rem     If Val(Gastos.Text) <> 0 Then
    Rem
    Rem         ZZArticulo = "Gastos"
    Rem         ZZCantidad = "1"
    Rem         ZZPrecio = Gastos.Text
    Rem         ZZDescripcion = "Gastos"
    Rem
    Rem         XXCodigo = ZZArticulo
    Rem         XXDs = ZZDescripcion
    Rem         qty = ZZCantidad
    Rem         XXPrecio = ZZPrecio
    Rem         umed = 1 ' Ver tabla de parámetros (unidades de medida)
    Rem         IMP_TOTAL = Trim(Str$(Val(ZZPrecio) * Val(ZZCantidad))) ' importe total final del artículo
    Rem         Bonif = ""
    Rem
    Rem         ' lo agrego a la factura (internamente, no se llama al WebService):
    Rem         ok = WSFEXv1.AgregarItem(XXCodigo, Trim(XXDs), qty, umed, XXPrecio, IMP_TOTAL, Bonif)
    Rem     End If
    Rem
    Rem End If
    
    
    
    
    
    
    ' Agrego un permiso (ver manual para el desarrollador)
    Rem id = "99999AAXX999999A"
    Rem dst = Val(ZZPais)
    Rem ok = WSFEXv1.AgregarPermiso(id, dst)
        
        
        
        
    ' Agrego un comprobante asociado (ver manual para el desarrollador)
    Rem tipo_cbte_asoc = 19
    Rem punto_vta_asoc = 2
    Rem cbte_nro_asoc = 1
    Rem ok = WSFEXv1.AgregarCmpAsoc(tipo_cbte_asoc, punto_vta_asoc, cbte_nro_asoc)
        
        
        
    'id = "99000000000100" ' número propio de transacción
    ' obtengo el último ID y le adiciono 1 (advertencia: evitar overflow!)
    Rem id = CStr(CCur(WSFEXv1.GetLastID()) + 1)
    
    
    
    ' Llamo al WebService de Autorización para obtener el CAE
    Rem Cae = WSFEXv1.Authorize(id)
    Rem Debug.Print WSFEXv1.XmlRequest
    Rem Debug.Print WSFEXv1.XmlResponse
    Rem Cae.Text = Cae
        
        
        
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    Rem If Cae = "" Or WSFEXv1.Resultado <> "A" Then
    Rem     MsgBox "No se asignó CAE (Rechazado). Observación (motivos): " & WSFEXv1.obs, vbInformation + vbOKOnly
    Rem ElseIf WSFEXv1.obs <> "" And WSFEXv1.obs <> "00" Then
    Rem     MsgBox "Se asignó CAE pero con advertencias. Observación (motivos): " & WSFEXv1.obs, vbInformation + vbOKOnly
    Rem End If
    
    
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Rem Debug.Print WSFEXv1.XmlRequest
    Rem Debug.Print WSFEXv1.XmlResponse
    
    Rem MsgBox "Resultado:" & WSFEXv1.Resultado & " CAE: " & Cae & " Reproceso: " & WSFEXv1.Reproceso & " Obs: " & WSFEXv1.obs & " Nro: " & ZZComprobante, vbInformation + vbOKOnly
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    Rem For Each evento In WSFEXv1.Eventos
    Rem     If evento <> "0: " Then
    Rem         MsgBox "Evento: " & evento, vbInformation
    Rem     End If
    Rem Next
    
    ' Buscar la factura
    Rem cae2 = WSFEXv1.GetCMP(tipo_cbte, punto_vta, Cbte_Nro)
    
    Rem Debug.Print "Fecha Comprobante:", WSFEXv1.FechaCbte
    Rem Debug.Print "Importe Total:", WSFEXv1.ImpTotal
    
    Rem If Cae <> cae2 Then
    Rem     MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!"
    Rem         Else
    Rem     MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
    Rem     ZZGrabaFactura = "S"
    Rem End If
    
    
    
    
    
    
    
    
    'id = "99000000000100" ' número propio de transacción
    ' obtengo el último ID y le adiciono 1 (advertencia: evitar overflow!)
    id = CStr(CCur(WSFEX.GetLastID()) + 1)
    
    
    
    ' Llamo al WebService de Autorización para obtener el CAE
    Cae = WSFEX.Authorize(id)
    Debug.Print WSFEX.XmlRequest
    Debug.Print WSFEX.XmlResponse
    Cae.Text = Cae
        
        
        
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    If Cae = "" Or WSFEX.Resultado <> "A" Then
        MsgBox "No se asignó CAE (Rechazado). Observación (motivos): " & WSFEX.obs, vbInformation + vbOKOnly
    ElseIf WSFEX.obs <> "" And WSFEX.obs <> "00" Then
        MsgBox "Se asignó CAE pero con advertencias. Observación (motivos): " & WSFEX.obs, vbInformation + vbOKOnly
    End If
    
    
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Debug.Print WSFEX.XmlRequest
    Debug.Print WSFEX.XmlResponse
    
    MsgBox "Resultado:" & WSFEX.Resultado & " CAE: " & Cae & " Reproceso: " & WSFEX.Reproceso & " Obs: " & WSFEX.obs & " Nro: " & ZZComprobante, vbInformation + vbOKOnly
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEX.Eventos
        If evento <> "0: " Then
            MsgBox "Evento: " & evento, vbInformation
        End If
    Next
    
    ' Buscar la factura
    cae2 = WSFEX.GetCMP(tipo_cbte, punto_vta, Cbte_Nro)
    
    Debug.Print "Fecha Comprobante:", WSFEX.FechaCbte
    Debug.Print "Importe Total:", WSFEX.ImpTotal
    
    If Cae <> cae2 Then
        MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!"
            Else
        MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
        ZZGrabaFactura = "S"
    End If
    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
ManejoError:
    ' Si hubo error:
    Debug.Print WSFEXv1.XmlRequest
    Debug.Print WSFEXv1.XmlResponse
    
    
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    Debug.Print WSFEXv1.XmlRequest
    Debug.Assert False

End Sub




