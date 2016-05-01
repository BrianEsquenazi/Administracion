VERSION 5.00
Begin VB.Form PrgProcVto 
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
Attribute VB_Name = "PrgProcVto"
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
Private VectorVto(5000, 7) As String
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

Dim ZMes As String
Dim ZAno As String
Dim XEmpresa As String

Private Sub Form_Load()
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    ZZFechaActual = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)

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
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " MarcaVencida = " + "'" + "" + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Guia SET "
        ZSql = ZSql + " MarcaVencida = " + "'" + "" + "'"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
        Erase VectorVto
        Renglon = 0
        
        
        ZSql = ""
        ZSql = ZSql + "Select Hoja.Hoja, Hoja.Marca, Hoja.Real, Hoja.Teorico, Hoja.Renglon, Hoja.Producto, Hoja.Revalida, Hoja.MesesRevalida, Hoja.Fecha, Hoja.FechaRevalida"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Marca <> 'X'"
        ZSql = ZSql + " and Hoja.Saldo <> 0"
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
            
                    Renglon = Renglon + 1
                    VectorVto(Renglon, 1) = rstHoja!Hoja
                    VectorVto(Renglon, 2) = rstHoja!producto
                    VectorVto(Renglon, 3) = "H"
                    VectorVto(Renglon, 4) = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                    VectorVto(Renglon, 5) = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                    VectorVto(Renglon, 6) = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                    VectorVto(Renglon, 7) = rstHoja!Fecha
                    
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
        ZSql = ZSql + "Select Guia.Clave, Guia.Codigo, Guia.Marca, Guia.Saldo, Guia.Lote, Guia.Terminado"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Guia.Saldo <> 0"
        ZSql = ZSql + " and Guia.Tipo = 'T'"
        spGuia = ZSql
        Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
        If rstGuia.RecordCount > 0 Then

            With rstGuia
        
                .MoveFirst
                
                If .NoMatch = False Then
                
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    Renglon = Renglon + 1
                    VectorVto(Renglon, 1) = rstGuia!Clave
                    VectorVto(Renglon, 2) = rstGuia!Terminado
                    VectorVto(Renglon, 3) = "G"
                    VectorVto(Renglon, 4) = rstGuia!Lote
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                
                End If
                
            End With
            
            rstGuia.Close
    
        End If
        
        For Da = 1 To Renglon
        
            ZZTipoMov = VectorVto(Da, 3)
            
            Select Case ZZTipoMov
                Case "H"
                    ZZHoja = VectorVto(Da, 1)
                    ZZProducto = VectorVto(Da, 2)
                    ZZTipoMov = VectorVto(Da, 3)
                    ZZRevalida = VectorVto(Da, 4)
                    ZZMesesRevalida = VectorVto(Da, 5)
                    ZZFechaRevalida = VectorVto(Da, 6)
                    ZZFecha = VectorVto(Da, 7)
                    ZZMeses = ""
                    
                    spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                        rstTerminado.Close
                    End If
                    
                    If Val(ZZMeses) <> 0 Then
                    
                        Rem VERIFICA EL 75 %
                    
                        If Val(ZZRevalida) <> 0 Then
                        
                            WVida = Int(Val(ZZMesesRevalida) * 0.75)
                            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                            WAno = Val(Right$(ZZFechaRevalida, 4))
                            
                                Else
                                
                            WVida = Int(Val(ZZMeses) * 0.75)
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
                        ZZOrdVto = ZAno + ZMes + "01"
                        
                        If ZZOrdVto < ZZFechaActual Then
                        
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Hoja SET "
                            ZSql = ZSql + " MarcaVencida = " + "'" + "S" + "'"
                            ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZZHoja + "'"
                            spHoja = ZSql
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                        
                        
                        Rem VERIFICA EL 100 %
                        
                        If Val(ZZRevalida) <> 0 Then
                        
                            WVida = Int(Val(ZZMesesRevalida))
                            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                            WAno = Val(Right$(ZZFechaRevalida, 4))
                            
                                Else
                                
                            WVida = Int(Val(ZZMeses))
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
                        ZZOrdVto = ZAno + ZMes + "01"
                        
                        If ZZOrdVto < ZZFechaActual Then
                        
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Hoja SET "
                            ZSql = ZSql + " MarcaVencida = " + "'" + "V" + "'"
                            ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZZHoja + "'"
                            spHoja = ZSql
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                
                    End If
                    
                Case "G"
                    ZZClave = VectorVto(Da, 1)
                    ZZProducto = VectorVto(Da, 2)
                    ZZTipoMov = VectorVto(Da, 3)
                    ZZLote = VectorVto(Da, 4)
                    ZZRevalida = ""
                    ZZMesesRevalida = ""
                    ZZFechaRevalida = "  /  /    "
                    ZZFecha = "  /  /    "
                    ZZMeses = ""
                    
                    spTerminado = "ConsultaTerminado " + "'" + ZZProducto + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZZMeses = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
                        rstTerminado.Close
                    End If
                    
                    If Val(ZZMeses) <> 0 Then
                
                        XEmpresa = WEmpresa
                        
                        For CiclaEmpresa = 1 To 9
            
                            WEmpresa = Empe(CiclaEmpresa, 1)
                            txtOdbc = Empe(CiclaEmpresa, 2)
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Hoja"
                            ZSql = ZSql + " Where Hoja.Hoja = " + "'" + ZZLote + "'"
                            ZSql = ZSql + " and Hoja.Producto = " + "'" + ZZProducto + "'"
                            spHoja = ZSql
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
                                ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
                                ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "  /  /    ", rstHoja!FechaRevalida)
                                ZZFecha = rstHoja!Fecha
                                rstHoja.Close
                                Exit For
                            End If
                            
                        Next CiclaEmpresa
                        
                        Call Conecta_Empresa
                        
                        Rem VERIFICA EL 75%
                        
                        If Val(ZZRevalida) <> 0 Then
                        
                            WVida = Int(Val(ZZMesesRevalida) * 0.75)
                            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                            WAno = Val(Right$(ZZFechaRevalida, 4))
                            
                                Else
                                
                            WVida = Int(Val(ZZMeses) * 0.75)
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
                        ZZOrdVto = ZAno + ZMes + "01"
                        
                        If ZZOrdVto < ZZFechaActual Then
                        
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Guia SET "
                            ZSql = ZSql + " MarcaVencida = " + "'" + "S" + "'"
                            ZSql = ZSql + " Where Guia.Clave = " + "'" + ZZClave + "'"
                            spGuia = ZSql
                            Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                        
                        Rem VERIFICA EL 100%
                        
                        If Val(ZZRevalida) <> 0 Then
                        
                            WVida = Int(Val(ZZMesesRevalida))
                            WMes = Val(Mid$(ZZFechaRevalida, 4, 2))
                            WAno = Val(Right$(ZZFechaRevalida, 4))
                            
                                Else
                                
                            WVida = Int(Val(ZZMeses))
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
                        ZZOrdVto = ZAno + ZMes + "01"
                        
                        If ZZOrdVto < ZZFechaActual Then
                                    
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Guia SET "
                            ZSql = ZSql + " MarcaVencida = " + "'" + "V" + "'"
                            ZSql = ZSql + " Where Guia.Clave = " + "'" + ZZClave + "'"
                            spGuia = ZSql
                            Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                
                    End If
                
                Case Else
                
            End Select
        Next Da
    
    Next A
    
    
    
    PrgProcVto.Hide
    Unload Me
    Close
    End

End Sub

Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub


