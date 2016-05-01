VERSION 5.00
Begin VB.Form PrgProc1Auto 
   AutoRedraw      =   -1  'True
   Caption         =   "Reproceso de Stock de Materias Primas"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
End
Attribute VB_Name = "PrgProc1Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private Vector(20000, 3) As String
Dim Empe(12, 10) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim XParam As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Private xLote(100, 7) As String
Dim WFechaCierre As String
Dim WOrdFechaCierre As String

Private Sub Form_Load()
    Call Proceso_Click
End Sub


Private Sub Proceso_Click()

    Empe(1, 1) = "0001"
    Empe(1, 2) = "Empresa01"
    Empe(2, 1) = "0002"
    Empe(2, 2) = "Empresa02"
    Empe(3, 1) = "0003"
    Empe(3, 2) = "Empresa03"
    Empe(4, 1) = "0004"
    Empe(4, 2) = "Empresa04"
    Empe(5, 1) = "0005"
    Empe(5, 2) = "Empresa05"
    Empe(6, 1) = "0006"
    Empe(6, 2) = "Empresa06"
    Empe(7, 1) = "0007"
    Empe(7, 2) = "Empresa07"
    Empe(8, 1) = "0008"
    Empe(8, 2) = "Empresa08"
    Empe(9, 1) = "0009"
    Empe(9, 2) = "Empresa09"
    Empe(10, 1) = "0010"
    Empe(10, 2) = "Empresa10"
    Empe(11, 1) = "0011"
    Empe(11, 2) = "Empresa11"
    
    For A = WDesdeEmpresa To WHastaEmpresa
        
    WEmpresa = Empe(A, 1)
    txtOdbc = Empe(A, 2)
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Erase Vector
    Renglon = 0
    
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    With rstArticulo
            .MoveFirst
            Do
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstArticulo!Codigo = "DY-402-510" Then
                
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstArticulo!Codigo
                    Vector(Renglon, 2) = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
                    Vector(Renglon, 3) = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                
                Rem Rem End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    rstArticulo.Close
    
    
    
    For Da = 1 To Renglon
    
        WEntradas = 0
        WSalidas = 0
        WArticulo = Vector(Da, 1)
        XCodigo = Vector(Da, 1)
        WFechaCierre = Vector(Da, 2)
        WOrdFechaCierre = Vector(Da, 3)
        XDate = Date$
        
        Call calcula_datos
        
        XEntradas = Str$(WEntradas)
        XSalidas = Str$(WSalidas)
        
        XParam = "'" + XCodigo + "','" _
                + XEntradas + "','" _
                + XSalidas + "','" _
                + XDate + "'"
                                           
        spArticulo = "ModificaArticuloMovimientos " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Next Da
    
    Next A
    
    PrgProc1Auto.Hide
    Unload Me
    PrgProc2Auto.Show

End Sub

Private Sub calcula_datos()


    Rem PROCESA LOS LAUDOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoRepro" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                If rstLaudo!Marca = "X" And WSaldo = 0 Then
                        Else
                    If rstLaudo!articulo = WArticulo Then
                        WEntradas = WEntradas + rstLaudo!liberada
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstLaudo.Close
    End If
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    Rem spHoja = "ListaHojaRepro" + XParam
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Rem spHoja = "ListaHojaRepro " + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                        Else
                        
                    Rem If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                    If rstHoja!Tipo = "M" Then
                    
                        xLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        xLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        xLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        xLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        xLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        xLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                        
                        If Val(xLote(1, 1)) = 0 Then
                            xLote(1, 1) = rstHoja!Lote
                            xLote(1, 2) = rstHoja!Cantidad
                        End If
                        
                        For Da = 1 To 3
                        
                            If xLote(Da, 2) = "" Then
                                xLote(Da, 2) = "0"
                            End If
                        
                            WCanti = xLote(Da, 2)
                            If WCanti <> 0 Then
                
                                WArticulo = rstHoja!articulo
                                WCanti = xLote(Da, 2)
                                WFecha = rstHoja!Fecha
                                WHoja = rstHoja!Hoja
                                WLote = xLote(Da, 1)
                                
                                WSalidas = WSalidas + WCanti
                                
                            End If
                        Next Da
                        
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstHoja!Articulo > WArticulo Then
                Rem     Exit Do
                Rem End If
                
            Loop
            End If
        
        End With
        
        rstHoja.Close
        
    End If
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovvar = "ListaMovvarRepro1" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
        With rstMovvar
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovvar!Marca = "X" Then
                        Else
                    If rstMovvar!Tipo = "M" And rstMovvar!articulo = WArticulo Then
                        If rstMovvar!Movi = "E" Then
                            WEntradas = WEntradas + rstMovvar!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovvar!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
        rstMovvar.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaRepro1" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                        Else
                    If rstMovguia!Tipo = "M" And rstMovguia!articulo = WArticulo Then
                        If rstMovguia!Movi = "E" Then
                            WEntradas = WEntradas + rstMovguia!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovguia!Cantidad
                        End If
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
            
        End With
        rstMovguia.Close
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spMovlab = "ListaMovlabRepro1" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
        With rstMovlab
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstMovlab!Marca = "X" Then
                        Else
                    If rstMovlab!Tipo = "M" And rstMovlab!articulo = WArticulo Then
                        If rstMovlab!Movi = "E" Then
                            WEntradas = WEntradas + rstMovlab!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovlab!Cantidad
                        End If
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
    
    
    
    Rem PROCESA LAS VENTAS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spEstadistica = "ListaEstadisticaReproDy " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
            Do
                If .EOF = True Then
                    Exit Do
                End If
                If rstEstadistica!Marca = "X" Then
                        Else
                    WTipo = rstEstadistica!Tipo
                    xLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    xLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                    xLote(2, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote2)
                    xLote(2, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti2)
                    xLote(3, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote3)
                    xLote(3, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti3)
                    xLote(4, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote4)
                    xLote(4, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti4)
                    xLote(5, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote5)
                    xLote(5, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti5)
                        
                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                    
                    If Len(Trim(WLoteAdicional)) = 98 Then
                        xLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                        xLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                        xLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                        xLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                        xLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                        xLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                        xLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                        xLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                        xLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                        xLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                        xLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                        xLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                        xLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                        xLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                            Else
                        xLote(6, 1) = "0"
                        xLote(6, 2) = "0"
                        xLote(7, 1) = "0"
                        xLote(7, 2) = "0"
                        xLote(8, 1) = "0"
                        xLote(8, 2) = "0"
                        xLote(9, 1) = "0"
                        xLote(9, 2) = "0"
                        xLote(10, 1) = "0"
                        xLote(10, 2) = "0"
                        xLote(11, 1) = "0"
                        xLote(11, 2) = "0"
                        xLote(12, 1) = "0"
                        xLote(12, 2) = "0"
                    End If
                        
                    For Da = 1 To 12
                        WLote = xLote(Da, 1)
                        Auxi = xLote(Da, 2)
                        Auxi = Pusing("###,###.##", Auxi)
                        WCantidad = Val(Auxi)
                        
                        If WCantidad <> 0 Then
                            If WTipo = 2 Then
                                WEntradas = WEntradas + Abs(WCantidad)
                                    Else
                                WSalidas = WSalidas + WCantidad
                            End If
                        End If
                        
                    Next Da
                    
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
            End If
        End With
    End If

End Sub





