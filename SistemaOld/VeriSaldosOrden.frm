VERSION 5.00
Begin VB.Form PrgVeriSaldosOrden 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Ordenes de Compra Pendientes"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
End
Attribute VB_Name = "PrgVeriSaldosOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WOrden As String
Private WClave As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim Vector(20000, 2) As String
Dim Empe(12, 10) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double

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
    
    Rem spOrden = "ModificaOrdenReproceso0"
    Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql & "UPDATE Orden SET "
    ZSql = ZSql & "Recibida = Cantidad"
    ZSql = ZSql & " Where Orden < 900000"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql & "UPDATE Orden SET "
    ZSql = ZSql & "Saldo = Cantidad - Recibida"
    ZSql = ZSql & " Where Orden < 900000"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    
    spArticulo = "ModificaArticuloPedido0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem spOrden = "ListaOrdenTotal"
    
    ZSql = ""
    ZSql = ZSql + "Select Orden.Clave, Orden.FechaOrd, Orden.Orden"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.FechaOrd > " + "'" + "20130101" + "'"
    ZSql = ZSql + " and Orden.Orden < " + "'" + "900000" + "'"
    ZSql = ZSql + " Order by Orden.Orden"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    If !FechaOrd > "20130101" Then
                        If rstOrden!orden < 900000 Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstOrden!Clave
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WClave = Vector(Ciclo, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select orden.Clave, Orden.Orden, Orden.Articulo, Orden.Cantidad"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Clave = " + "'" + WClave + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WOrden = rstOrden!orden
            WArticulo = rstOrden!articulo
            WCantidad = Str$(rstOrden!cantidad)
            rstOrden.Close
        End If
        
        WResta = 0
        
        If Val(WOrden) < 900000 Then
        
            Rem XParam = "'" + WOrden + "','" _
            REM              + WArticulo + "'"
            Rem spInforme = "ListaInformeOrdenArticulo" + XParam
            
            ZSql = ""
            ZSql = ZSql + "Select Informe.Clave, Informe.Orden, Informe.Articulo, Informe,Resta "
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Orden = " + "'" + WOrden + "'"
            ZSql = ZSql + " and Informe.Articulo = " + "'" + WArticulo + "'"
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
                        
                            WResta = WResta + rstInforme!Resta
                
                            .MoveNext
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                        Loop
                    End If
                End With
                rstInforme.Close
            End If
        
            Rem XParam = "'" + WClave + "','" _
            rem              + Str$(WResta) + "'"
            Rem spOrden = "ModificaOrdenReproceso " + XParam
            Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
        
            Rem If WArticulo = "DF-010-100" Then Stop
            
            If WResta > Val(WCantidad) Then
                WResta = WCantidad
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Orden SET "
            ZSql = ZSql & "Recibida = " + Str$(WResta) + ", "
            ZSql = ZSql & "Saldo = Cantidad - " + Str$(WResta) + " "
            ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
            WDife = Val(WCantidad) - WResta
            XParam = "'" + WArticulo + "','" _
                         + Str$(WDife) + "'"
            spArticulo = "ModificaArticuloPedido " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next Ciclo
    
    
    Erase Vector
    Renglon = 0
    
    
    
    Rem spOrden = "ListaOrdenTotal"
    
    ZSql = ""
    ZSql = ZSql & "UPDATE Orden SET "
    ZSql = ZSql & "Recibida = Cantidad"
    ZSql = ZSql & " Where Recibida > Cantidad and Orden > 900000"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql & "UPDATE Orden SET "
    ZSql = ZSql & "Saldo = Cantidad - Recibida"
    ZSql = ZSql & " Where Orden > 900000"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "Select Orden.Clave, Orden.FechaOrd, Orden.Orden, Orden.Saldo, Orden.Articulo"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.FechaOrd > " + "'" + "20130101" + "'"
    ZSql = ZSql + " and Orden.Orden >= " + "'" + "900000" + "'"
    ZSql = ZSql + " and Orden.Saldo <> 0"
    ZSql = ZSql + " Order by Orden.Orden"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    If !FechaOrd > "20020101" Then
                        If rstOrden!orden >= 900000 And rstOrden!Saldo <> 0 Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstOrden!articulo
                            Vector(Renglon, 2) = Str$(rstOrden!Saldo)
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WArticulo = Vector(Ciclo, 1)
        WCantidad = Val(Vector(Ciclo, 2))
            
        XParam = "'" + WArticulo + "','" _
                     + Str$(WCantidad) + "'"
        spArticulo = "ModificaArticuloPedido " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Next A
    
    PrgVeriSaldosOrden.Hide
    Unload Me
    PrgVeriSaldosInforme.Show
    
End Sub

