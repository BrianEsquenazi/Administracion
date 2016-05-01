VERSION 5.00
Begin VB.Form PrgMiraLote 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Marca de Lote/Control de Lote"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "PrgMiraLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstInventario As Recordset
Dim spInventario As String
Dim XParam As String
Dim Vector(5000, 10) As String

Private Sub Cancelar_Click()
    PrgMiraLote.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()
    
    Erase Vector
    Renglon = 0

    spInventario = "ListaInventarioTotal"
    Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
    If rstInventario.RecordCount > 0 Then
        
    With rstInventario
    
        .MoveFirst
        
        If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Renglon = Renglon + 1
                Vector(Renglon, 1) = rstInventario!Tipo
                Vector(Renglon, 2) = rstInventario!Articulo
                Vector(Renglon, 3) = rstInventario!Terminado
                Vector(Renglon, 4) = Str$(rstInventario!Cantidad)
                Vector(Renglon, 5) = rstInventario!Lote
                
                .MoveNext
                        
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
    End With
            
    rstInventario.Close
            
    End If
    
    
    For Ciclo = 1 To Renglon
    
            WTipo = Vector(Ciclo, 1)
            WArticulo = Vector(Ciclo, 2)
            WTerminado = Vector(Ciclo, 3)
            WCantidad = Vector(Ciclo, 4)
            WLote = Vector(Ciclo, 5)
                
            Rem If Left$(WArticulo, 2) <> "NI" And Left$(WTerminado, 2) <> "NI" Then
                
                If WTipo = "M" Then
                
                    WControla = 0
                    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                        rstArticulo.Close
                    End If
            
                    If WControla = 1 Then
                    
                        m$ = "La Materia Prima  " + WArticulo + " no posee marca de controla Lote"
                        ca% = MsgBox(m$, 0, "Control de Carga de Inventario")
                        
                            Else
                            
                        WBusqueda = "N"
                        
                        XParam = "'" + WLote + "','" _
                                    + WArticulo + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WBusqueda = "S"
                            rstLaudo.Close
                        End If
                
                        If WBusqueda = "N" Then
                        
                            XParam = "'" + WArticulo + "','" _
                                    + WLote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WBusqueda = "S"
                                rstMovguia.Close
                            End If
                            
                        End If
                        
                        If WBusqueda = "N" Then
                            m$ = "No existe el lote " + WLote + " de la materia prima  " + WArticulo
                            ca% = MsgBox(m$, 0, "Control de Carga de Inventario")
                        End If
                        
                    End If
                                
                                Else
                                
                    WControla = 0
                    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    If WControla = 1 Then
                    
                        m$ = "El Producto Terminado  " + WTerminado + " no posee marca de controla Lote"
                        ca% = MsgBox(m$, 0, "Control de Carga de Inventario")
                        
                            Else
                            
                        WBusqueda = "N"
                        
                        XParam = "'" + WLote + "','" _
                                + WTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WBusqueda = "S"
                            rstHoja.Close
                        End If
                        
                        If WBusqueda = "N" Then
                        
                            XParam = "'" + WTerminado + "','" _
                                    + WLote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WBusqueda = "S"
                                rstMovguia.Close
                            End If
                            
                        End If
                        
                        If WBusqueda = "N" Then
                            m$ = "No existe el lote " + WLote + " de la producto terminado  " + WTerminado
                            ca% = MsgBox(m$, 0, "Control de Carga de Inventario")
                        End If
                        
                    End If
                    
                End If
                
            Rem End If
                
    Next Ciclo
    
    Call Cancelar_Click

End Sub

