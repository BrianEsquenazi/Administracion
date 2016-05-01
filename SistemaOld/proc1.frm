VERSION 5.00
Begin VB.Form PrgProc1 
   AutoRedraw      =   -1  'True
   Caption         =   "Reproceso de Stock de Materias Primas"
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
Attribute VB_Name = "PrgProc1"
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
Private XLote(100, 7) As String
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
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim XParam As String

Private Sub Cancelar_Click()

    PrgProc1.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

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
          Rem BY NAN
           Rem    If rstArticulo!Codigo = "DS-001-100" Then
          Rem  Stop
              Renglon = Renglon + 1
                Vector(Renglon) = rstArticulo!Codigo
                
         Rem    End If
                
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
    
    Call Cancelar_Click

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
                
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)

                If rstLaudo!Marca = "X" And WSaldo = 0 Then
                
                        Else
                    
                    Rem If rstLaudo!Articulo = WArticulo Then
                        WEntradas = WEntradas + rstLaudo!Liberada
                    Rem End If
                
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
                    
                        XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                        
                        If Val(XLote(1, 1)) = 0 Then
                            XLote(1, 1) = rstHoja!Lote
                            XLote(1, 2) = rstHoja!Cantidad
                        End If
                        
                        For Da = 1 To 3
                        
                            If XLote(Da, 2) = "" Then
                                XLote(Da, 2) = "0"
                            End If
                        
                            WCanti = XLote(Da, 2)
                            If WCanti <> 0 Then
                
                                WArticulo = rstHoja!Articulo
                                WCanti = XLote(Da, 2)
                                WFecha = rstHoja!Fecha
                                WHoja = rstHoja!Hoja
                                WLote = XLote(Da, 1)
                                
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
                        
                    Rem If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                    If rstMovvar!Tipo = "M" Then
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
                        
                    Rem If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                    If rstMovguia!Tipo = "M" Then
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
                
                    Rem If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                    If rstMovlab!Tipo = "M" Then
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
        rstMovlab.Close
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
                    
                    If (rstEstadistica!Tipo = 2 And Left$(WArticulo, 2) = rstEstadistica!Tipopro) Or rstEstadistica!Tipo = 1 Then
                    
                        WArticulo = rstEstadistica!ArticuloDy
                        WFecha = rstEstadistica!Fecha
                        WCodigo = rstEstadistica!Numero
                        WObservaciones = ""
                        WTipo = rstEstadistica!Tipo
                        
                        XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                        XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                        XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                        XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                        XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                        XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                        XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                        XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                        
                        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                        
                        If Len(Trim(WLoteAdicional)) = 98 Then
                            XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                            XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                            XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                            XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                            XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                            XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                            XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                            XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                            XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                            XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                            XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                            XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                            XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                            XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                                Else
                            XLote(6, 1) = "0"
                            XLote(6, 2) = "0"
                            XLote(7, 1) = "0"
                            XLote(7, 2) = "0"
                            XLote(8, 1) = "0"
                            XLote(8, 2) = "0"
                            XLote(9, 1) = "0"
                            XLote(9, 2) = "0"
                            XLote(10, 1) = "0"
                            XLote(10, 2) = "0"
                            XLote(11, 1) = "0"
                            XLote(11, 2) = "0"
                            XLote(12, 1) = "0"
                            XLote(12, 2) = "0"
                        End If
                        
                        For Da = 1 To 12
                        
                            WLote = XLote(Da, 1)
                            WCantidad = XLote(Da, 2)
                    
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
    
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgProc1.Caption = "Reproceso de Stock de Materia Prima :  " + !Nombre
        End If
    End With

End Sub
