VERSION 5.00
Begin VB.Form Impresion_varios 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Impresion_varios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Proceso()

    Erase WVectorII
    LugarVectorII = 0

    
    WSalidaError = ""
    On Error GoTo Control_error
    

                
    Rem PROCESA Las compras
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Informe.Articulo = " + "'" + WArticulo + "'"
    ZSql = ZSql + " and Informe.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Informe.FechaOrd <= " + "'" + WHasta + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
    
        With rstInforme
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                LugarVectorII = LugarVectorII + 1
                
                WVectorII(LugarVectorII, 1) = "1"
                WVectorII(LugarVectorII, 2) = rstInforme!Fecha
                WVectorII(LugarVectorII, 3) = Str$(rstInforme!Cantidad)
                WVectorII(LugarVectorII, 4) = rstInforme!Remito
                WVectorII(LugarVectorII, 5) = rstInforme!informe
                WVectorII(LugarVectorII, 6) = rstInforme!Orden
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstInforme.Close
    End If
    
    For Ciclo = 1 To LugarVectorII
    
        WOrden = WVectorII(LugarVectorII, 6)
        
        WProveedor = ""
        spOrden = "ListaOrden" + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!proveedor
            WVectorII(Ciclo, 7) = rstOrden!proveedor
            rstOrden.Close
        End If
        
        WDEsProveedor = ""
                
        spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WVectorII(Ciclo, 8) = IIf(IsNull(RstProveedor!cufe), "", RstProveedor!cufe)
            WVectorII(Ciclo, 9) = IIf(IsNull(RstProveedor!cufeii), "", RstProveedor!cufeii)
            WVectorII(Ciclo, 10) = IIf(IsNull(RstProveedor!cufeiii), "", RstProveedor!cufeiii)
            RstProveedor.Close
        End If
        
    Next Ciclo
    
    
    
    
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
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
                
                If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "2"
                        WVectorII(LugarVectorII, 2) = rstHoja!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstHoja!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstHoja!hoja

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
    End If
    
    
    
    
    
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then

        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                
                XFec = Right$(rstMovvar!Fecha, 4) + Mid$(rstMovvar!Fecha, 4, 2) + Left$(rstMovvar!Fecha, 2)
                
                If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "3"
                        WVectorII(LugarVectorII, 2) = rstMovvar!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstMovvar!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstMovvar!codigo
                        WVectorII(LugarVectorII, 5) = rstMovvar!Movi

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
    
    
    
    
    
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                
                If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec And rstMovguia!codigo < 900000 Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "4"
                        WVectorII(LugarVectorII, 2) = rstMovguia!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstMovguia!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstMovguia!codigo
                        WVectorII(LugarVectorII, 5) = rstMovguia!Movi
                        WVectorII(LugarVectorII, 6) = rstMovguia!Destino
                        WVectorII(LugarVectorII, 7) = rstMovguia!Tipomov

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
    
    spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                
                XFec = Right$(rstMovlab!Fecha, 4) + Mid$(rstMovlab!Fecha, 4, 2) + Left$(rstMovlab!Fecha, 2)
                
                If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                
                    If WDesde <= XFec And WHasta >= XFec Then
                    
                        LugarVectorII = LugarVectorII + 1
                        
                        WVectorII(LugarVectorII, 1) = "4"
                        WVectorII(LugarVectorII, 2) = rstMovlab!Fecha
                        WVectorII(LugarVectorII, 3) = Str$(rstMovlab!Cantidad)
                        WVectorII(LugarVectorII, 4) = rstMovlab!codigo
                        WVectorII(LugarVectorII, 5) = rstMovlab!Movi

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
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    Resume Next
    
End Sub




