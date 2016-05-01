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
Private Vector(10000) As String
Dim Empe(10, 10) As String
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
    
    For A = 1 To 2
        
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
                
                Renglon = Renglon + 1
                
                Vector(Renglon) = rstArticulo!Codigo
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
        WArticulo = Vector(Da)
        XCodigo = Vector(Da)
        XDate = Date$
        
        Rem If WArticulo = "CO-886-100" Then Stop
        
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

    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
        rstArticulo.Close
    End If


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
                
                Rem If WArticulo = "AA-238-100" Then Stop
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)

                If rstLaudo!Marca = "X" And WSaldo = 0 Then
                
                        Else
                    
                    If rstLaudo!Articulo = WArticulo Then
                        WEntradas = WEntradas + rstLaudo!Liberada
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
                        
                    If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                        XX = rstHoja!Clave
                        WSalidas = WSalidas + rstHoja!Cantidad
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstHoja!Articulo > WArticulo Then
                    Exit Do
                End If
                
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
                        
                    If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
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
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
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
                
                    If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
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
    
    spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
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
                
                    If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = WArticulo Then
                    
                        WArticulo = rstEstadistica!ArticuloDy
                        WFecha = rstEstadistica!Fecha
                        WCodigo = rstEstadistica!Numero
                        WObservaciones = ""
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
                        
                        For Da = 1 To 5
                        
                            WLote = xLote(Da, 1)
                            WCantidad = xLote(Da, 2)
                    
                            If Val(WCantidad) <> 0 Then
                                If WTipo = 2 Then
                                    WEntradas = WEntradas + Abs(Val(WCantidad))
                                        Else
                                    WSalidas = WSalidas + WCantidad
                                End If
                            End If
                            
                        Next Da
                        
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
    
End Sub

