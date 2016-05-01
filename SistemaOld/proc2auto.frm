VERSION 5.00
Begin VB.Form PrgProc2Auto 
   AutoRedraw      =   -1  'True
   Caption         =   "Reprocesos de Productos Terminados"
   ClientHeight    =   7170
   ClientLeft      =   225
   ClientTop       =   975
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   11655
End
Attribute VB_Name = "PrgProc2Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WTerminado As String
Private WEntradas As Double
Private WSalidas As Double
Private WProceso As Double
Private Vector(20000) As String
Dim Empe(12, 10) As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
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
Dim rstConsig As Recordset
Dim spConsig As String
Dim XParam As String

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
            
        spTerminado = "ListaTerminado"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
        
            With rstTerminado
                .MoveFirst
                    
                    Do
                    
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        Renglon = Renglon + 1
                      Vector(Renglon) = rstTerminado!codigo
                        
                        .MoveNext
                        
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                    Loop
                    
            End With
            rstTerminado.Close
        
        End If
        
        
        Sql1 = "UPDATE Hoja SET "
        Sql2 = " Realant = 0"
        Sql3 = " Where Realant IS NULL"
        spHoja = Sql1 + Sql2 + Sql3
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        
        Rem Renglon = 1
        Rem Vector(Renglon) = "PT-30392-100"
        
        For Da = 1 To Renglon
        
            WEntradas = 0
            WSalidas = 0
            WTerminado = Vector(Da)
            XCodigo = Vector(Da)
            XDate = Date$
            
            Call calcula_datos
            
            XEntradas = Str$(WEntradas)
            XSalidas = Str$(WSalidas)
            
            XParam = "'" + XCodigo + "','" _
                    + XEntradas + "','" _
                    + XSalidas + "','" _
                    + XDate + "'"
                                               
            spTerminado = "ModificaTerminadoMovimientos " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            WProceso = 0
            
            ZSql = ""
            ZSql = ZSql + "Select Hoja.Marca, Hoja.Real, Hoja.Teorico, Hoja.Renglon, Hoja.Producto"
            ZSql = ZSql + " FROM Hoja"
            ZSql = ZSql + " Where Hoja.Producto = " + "'" + XCodigo + "'"
            ZSql = ZSql + " and Hoja.Marca <> 'X'"
            ZSql = ZSql + " and Hoja.Real = 0"
            ZSql = ZSql + " and Hoja.RealAnt = 0"
            ZSql = ZSql + " and Hoja.Teorico <> 0"
            ZSql = ZSql + " and Hoja.Renglon = 1"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
        
                With rstHoja
        
                    .MoveFirst
                
                    If .NoMatch = False Then
                
                        Do
                
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                            WProceso = WProceso + rstHoja!Teorico
                    
                            .MoveNext
                    
                            If .EOF = True Then
                                Exit Do
                            End If
                    
                        Loop
                
                    End If
                
                End With
                rstHoja.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Terminado SET "
            ZSql = ZSql + " Proceso = " + "'" + Str$(WProceso) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + XCodigo + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Da
    
    Next A
    
    PrgProc2Auto.Hide
    Unload Me
    PrgVeriSaldosOrden.Show

End Sub

Private Sub calcula_datos()

    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
        rstTerminado.Close
    End If


    Rem PROCESA LAS ESTADISTICAS
    
    
    ZTipoPro = Left$(UCase(WTerminado), 2)
    
    Select Case ZTipoPro
        Case "PT"
            XParam = "'" + WTerminado + "','" _
                         + WTerminado + "'"
            spEstadistica = "ListaEstadisticaRepro" + XParam
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
                        
                        Rem aa1 = rstEstadistica!cantidad
                        Rem aa2 = rstEstadistica!numero
                    
                        If Val(rstEstadistica!Tipo) = 1 Then
                            WSalidas = WSalidas + rstEstadistica!cantidad
                                Else
                            WEntradas = WEntradas + Abs(rstEstadistica!cantidad)
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
             
        Case "NK"
            ZTerminado = "PT" + Mid$(WTerminado, 3, 10)
            XParam = "'" + ZTerminado + "','" _
                         + ZTerminado + "'"
            spEstadistica = "ListaEstadisticaRepro" + XParam
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
                
                            If Val(rstEstadistica!Tipo) <> 1 Then
                                WSalidas = WSalidas + Abs(rstEstadistica!cantidad)
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
        
        Case Else
        
    End Select
    
    
    
    Rem PROCESA LAS HOJAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaRepro1" + XParam
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
                            
                If rstHoja!Tipo = "T" Then
                
                    WSalidas = WSalidas + rstHoja!cantidad
                
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
    
    Rem PROCESA LAS HOJAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaRepro2" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
            
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 And rstHoja!Real <> 0 Then
                
                    WEntradas = WEntradas + rstHoja!Real
                    
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
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovvar = "ListaMovvarRepro" + XParam
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
                
                If rstMovvar!Tipo = "T" Then
                
                    If rstMovvar!Movi = "E" Then
                        WEntradas = WEntradas + rstMovvar!cantidad
                            Else
                        WSalidas = WSalidas + rstMovvar!cantidad
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
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovguia = "ListaMovguiaRepro" + XParam
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
                
                If rstMovguia!Tipo = "T" Then
                
                    If rstMovguia!Movi = "E" Then
                        WEntradas = WEntradas + rstMovguia!cantidad
                            Else
                        WSalidas = WSalidas + rstMovguia!cantidad
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
    
    
    Rem PROCESA LOS MOVIMIENTOS DE LABORATORIO
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovlab = "ListaMovlabRepro" + XParam
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
                
                If rstMovlab!Tipo = "T" Then
                
                    If rstMovlab!Movi = "E" Then
                        WEntradas = WEntradas + rstMovlab!cantidad
                                Else
                        WSalidas = WSalidas + rstMovlab!cantidad
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
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spConsig = "ListaConsigRepro" + XParam
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount > 0 Then
    
        With rstConsig
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstConsig!Marca <> "X" Then
                    WCantidad = rstConsig!cantidad - rstConsig!Facturado
                    WSalidas = WSalidas + WCantidad
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstConsig.Close
    End If
    
    
        
    Rem PROCESA LOS las devoluciones de mercaderia
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntDev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntDev.RecordCount > 0 Then
    
        With rstEntDev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEntDev!Marca = "X" Then
                
                        Else
                
                WCantidad = rstEntDev!cantidad
                WEntradas = WEntradas + WCantidad
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstEntDev.Close
        
    End If

End Sub
