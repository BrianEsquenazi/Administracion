VERSION 5.00
Begin VB.Form PrgVeriSaldosInforme 
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
Attribute VB_Name = "PrgVeriSaldosInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WClave As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim Vector(20000) As String
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
    
    spInforme = "ModificaInformeProcesoSaldo"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    spArticulo = "ModificaArticuloLaboratorio0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    XParam = "'" + "20020101" + "'"
    spInforme = "ModificaInformeProceso0 " + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    spInforme = "ListaInformeTotal"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                    If !FechaOrd > "20130101" Then
                        Renglon = Renglon + 1
                        Vector(Renglon) = rstInforme!Clave
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If

    For Ciclo = 1 To Renglon
        
        WClave = Vector(Ciclo)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Informe"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
        spInforme = ZSql
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            WInforme = Str$(rstInforme!Informe)
            WArticulo = rstInforme!articulo
            WCantidad = Str$(rstInforme!Cantidad)
            rstInforme.Close
        End If
        
        WResta = 0
        
        XParam = "'" + WInforme + "','" _
                 + WArticulo + "'"
        spLaudo = "ListaLaudoInforme " + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
    
            With rstLaudo
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        WLiberada = rstLaudo!liberada
                        WDevuelta = rstLaudo!devuelta
                        WSuma = WLiberada + WDevuelta
                        
                        WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        WDevueltaAnt = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                        WSumaAnt = WLiberadaAnt + WDevueltaAnt
                        
                        If WSumaAnt <> 0 Then
                            WResta = WResta + WSumaAnt
                                Else
                            WResta = WResta + WSuma
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
        
        XParam = "'" + WClave + "','" _
                 + Str$(WResta) + "'"
        spInforme = "ModificaInformeProceso " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
        WDife = Val(WCantidad) - WResta
        
        If WDife > 0 Then
            XParam = "'" + WArticulo + "','" _
                 + Str$(WDife) + "'"
            spArticulo = "ModificaArticuloLaboratorio " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Ciclo
    
    spInforme = "ModificaInformeProcesoDife"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    Next A

    PrgVeriSaldosInforme.Hide
    Unload Me
    PrgProcPedPenAuto.Show
    
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


