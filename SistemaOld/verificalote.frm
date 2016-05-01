VERSION 5.00
Begin VB.Form PrgVerificaLote 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Saldos de Lotes de Productos Terminados"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
End
Attribute VB_Name = "PrgVerificaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String

Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String

Dim XParam As String
Dim Vector(10000, 5) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase Vector
    Renglon = 0
    
    spHoja = "ListaHojaTotal"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    If rstHoja!Marca <> "X" And rstHoja!Renglon = 1 And rstHoja!Saldo <> 0 Then
                        
                        If rstHoja!producto >= "PT-11000-000" And rstHoja!producto <= "PT-11999-999" Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstHoja!Hoja
                            Vector(Renglon, 2) = rstHoja!producto
                        End If
                        
                        If rstHoja!producto >= "PT-00000-000" And rstHoja!producto <= "PT-00999-999" Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstHoja!Hoja
                            Vector(Renglon, 2) = rstHoja!producto
                        End If
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    XParam = "'" + "AA-00000-000" + "','" _
                 + "ZZ-99999-999" + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstMovguia!Marca <> "X" And rstMovguia!Saldo <> 0 Then
                    
                    If rstMovguia!Tipo = "T" Then
                    
                    If rstMovguia!terminado >= "PT-11000-000" And rstMovguia!terminado <= "PT-11999-999" Then
                    
                        WMovi = rstMovguia!Movi
                        If WMovi = "S" Then
                            Lote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                Else
                            Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        End If
                        If WMovi = "E" Then
                            Entra = "S"
                            For dada = 1 To Renglon
                                If Vector(dada, 1) = Lote Then
                                    Entra = "N"
                                    Exit For
                                End If
                            Next dada
                            If Entra = "S" Then
                                If Lote <> 0 Then
                                    Renglon = Renglon + 1
                                    Vector(Renglon, 1) = Lote
                                    Vector(Renglon, 2) = rstMovguia!terminado
                                End If
                            End If
                        End If
                        
                    End If
                    
                    If rstMovguia!terminado >= "PT-00000-000" And rstMovguia!terminado <= "PT-00999-999" Then
                    
                        WMovi = rstMovguia!Movi
                        If WMovi = "S" Then
                            Lote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                Else
                            Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        End If
                        If WMovi = "E" Then
                            Entra = "S"
                            For dada = 1 To Renglon
                                If Vector(dada, 1) = Lote Then
                                    Entra = "N"
                                    Exit For
                                End If
                            Next dada
                            If Entra = "S" Then
                                If Lote <> 0 Then
                                    Renglon = Renglon + 1
                                    Vector(Renglon, 1) = Lote
                                    Vector(Renglon, 2) = rstMovguia!terminado
                                End If
                            End If
                        End If
                        
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
    
    For dada = 1 To Renglon
    
        WLote = Vector(dada, 1)
        WTerminado = Vector(dada, 2)
        WOrdFecha = "19000101"
        
        XParam = "'" + WTerminado + "','" _
                     + WTerminado + "'"
        spEstadistica = "ListaEstadisticaDesdeHasta" + XParam
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
    
            With rstEstadistica
    
                .MoveFirst
                
                If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZOrdFecha = rstEstadistica!ordfecha
                    If WOrdFecha < ZOrdFecha Then
                        WOrdFecha = ZOrdFecha
                    End If
                
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
                End If
            End With
            rstEstadistica.Close
        End If
        
        Fecha1 = Right$(WOrdFecha, 2) + "/" + Mid$(WOrdFecha, 5, 2) + "/" + Left$(WOrdFecha, 4)
        Fecha2 = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
        ZMeses = DateDiff("m", Fecha1, Fecha2)
        
        If ZMeses >= 24 Then
                    
            If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DK" Or Left$(WTerminado, 2) = "DW" Or Left$(WTerminado, 2) = "NW" Then
        
                ZEntra = "N"
                
                If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DK" Then
                    ZArti = "DY-" + Right$(WTerminado, 7)
                        Else
                    ZArti = "DW-" + Right$(WTerminado, 7)
                End If
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZArti + "'"
                ZSql = ZSql + " and Laudo.Lote = " + "'" + WLote + "'"
                ZSql = ZSql + " Order by Laudo.Laudo"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                                
                    rstLaudo.Close
                                    
                    ZEntra = "S"
                    ZMarcaEstado = "N"
                    ZMarcaEstadoII = "V"
                                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Laudo SET "
                    ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "',"
                    ZSql = ZSql + "EstadoII  = " + "'" + ZMarcaEstadoII + "'"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZArti + "'"
                    ZSql = ZSql + " and Laudo.Lote = " + "'" + WLote + "'"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                    
                End If
                        
                If ZEntra = "N" Then
                                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZArti + "'"
                    ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                    ZSql = ZSql + " Order by Guia.Saldo desc"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        rstMovguia.Close
                                        
                        ZMarcaEstado = "N"
                        ZMarcaEstadoII = "V"
                                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Guia SET "
                        ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "',"
                        ZSql = ZSql + "EstadoII  = " + "'" + ZMarcaEstadoII + "'"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZArti + "'"
                        ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        
                    End If
                End If
            
                    Else
                
                ZEntra = "N"
            
                XParam = "'" + WLote + "','" _
                            + WTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                                
                    rstHoja.Close
                                    
                    ZEntra = "S"
                    ZMarcaEstado = "N"
                    ZMarcaEstadoII = "V"
                                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "',"
                    ZSql = ZSql + "EstadoII  = " + "'" + ZMarcaEstadoII + "'"
                    ZSql = ZSql + " Where Hoja.Producto = " + "'" + WTerminado + "'"
                    ZSql = ZSql + " and Hoja.Hoja = " + "'" + WLote + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    
                End If
                
                If ZEntra = "N" Then
                                
                    XParam = "'" + WTerminado + "','" _
                                + WLote + "'"
                                                
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                                    
                        rstMovguia.Close
                                        
                        ZMarcaEstado = "N"
                        ZMarcaEstadoII = "V"
                                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Guia SET "
                        ZSql = ZSql + "Estado  = " + "'" + ZMarcaEstado + "',"
                        ZSql = ZSql + "EstadoII  = " + "'" + ZMarcaEstadoII + "'"
                        ZSql = ZSql + " Where Guia.Terminado = " + "'" + WTerminado + "'"
                        ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        
                    End If
                End If
            
            End If
        End If
    
    Next dada
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Close
    PrgVerificaLote.Hide
    End
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub
