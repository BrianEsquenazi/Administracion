VERSION 5.00
Begin VB.Form PrgCierreStkAnt 
   AutoRedraw      =   -1  'True
   Caption         =   "Cierre de Stock"
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
Attribute VB_Name = "PrgCierreStkAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEntdev As Recordset
Dim spEntdev As String
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
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Private Uno As String
Private Dos As String
Private Tres As String
Private Auxi As String
Private Auxi1 As String
Private Auxi2 As String
Dim WVector(50000, 10) As String

Dim WDesde1 As String
Dim WHasta1 As String
Dim WDesde2 As String
Dim WHasta2 As String
Dim WDesde3 As String
Dim WHasta3 As String
Dim WDesde4 As String
Dim WHasta4 As String
Dim WDesde5 As String
Dim WHasta5 As String
Dim WDesde6 As String
Dim WHasta6 As String
Dim WDesde7 As String
Dim WHasta7 As String
Dim WDesde8 As String
Dim WHasta8 As String

Private Sub Cancelar_Click()

    PrgCierreStkAnt.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

     Rem HERNAN ACTUALIZACION
     Rem PARAMETROS DE MATERIA PRIMA

 Rem   WDesde1 = "AA-000-000"
  Rem  WHasta1 = "XX-030-999"
     
  Rem   WDesde2 = "DS-049-100"
  Rem   WHasta2 = "DS-049-100"
    
 Rem    WDesde5 = "EC-000-100"
 Rem    WHasta5 = "SR-999-100"
    
 Rem    WDesde6 = "WA-000-000"
 Rem    WHasta6 = "XV-999-999"
    
 Rem    WDesde7 = "DA-005-100"
 Rem    WHasta7 = "DQ-410-100"
        
   Rem WDesde2 = "DA-005-100"
   Rem WHasta2 = "DQ-410-100"
        
   Rem WDesde5 = "CD-020-100"
   Rem WHasta5 = "CM-000-100"
        
   Rem WDesde6 = "DS-049-100"
   Rem WHasta6 = "DS-049-100"
        
  Rem  WDesde7 = "DA-005-100"
  Rem  WHasta7 = "DA-005-100"
        
 Rem   WDesde8 = "CD-020-100"
 Rem   WHasta8 = "CM-000-100"
    
    
    
    Rem PARAMETROS DE PRODUCTO TERMINADO
    
    WDesde3 = "NK-25024-000"
  WHasta3 = "NK-25024-999"
   
  Rem  WDesde4 = "RE-05106-000"
  Rem  WHasta4 = "RE-25301-999"
    
    T$ = "Cierre de Inventario"
    m$ = "!!! ATENCION !!!   Se actualizara los datos del sistema para mantener la ficha historica de movimientos de stock, Desea realizar el proceso      "
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% <> 6 Then
        Exit Sub
    End If
    
    
    
    
 Rem   Stop

    Rem Procesa los Laudos
    
    Rem WEmpresa = "0015"
   Rem txtOdbc = "Empresa15"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Erase WVector
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Order by Laudo"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WArticulo = IIf(IsNull(rstLaudo!Articulo), "", rstLaudo!Articulo)
                
                Pasa = "N"
                
                If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                    Pasa = "S"
                End If
                If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                    Pasa = "S"
                End If
                
                If Pasa = "S" Then
                
                    WLiberadaAnt = IIf(IsNull(rstLaudo!LiberadaAnt), "0", rstLaudo!LiberadaAnt)
                    WDevueltaAnt = IIf(IsNull(rstLaudo!DevueltaAnt), "0", rstLaudo!DevueltaAnt)
                    WMarca = IIf(IsNull(rstLaudo!Marca), "", rstLaudo!Marca)
                    WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                    WDevuelta = IIf(IsNull(rstLaudo!Devuelta), "0", rstLaudo!Devuelta)
                    WSaldo = IIf(IsNull(rstLaudo!Devuelta), "0", rstLaudo!Saldo)
                
                    If WLiberadaAnt = 0 And WDevueltaAnt = 0 Then
                        If WLiberada <> 0 Or WDevuelta <> 0 Then
                            Lugar = Lugar + 1
                            WVector(Lugar, 1) = rstLaudo!Clave
                            WVector(Lugar, 2) = WMarca
                            WVector(Lugar, 3) = Str$(WLiberada)
                            WVector(Lugar, 4) = Str$(WDevuelta)
                            WVector(Lugar, 5) = Str$(WSaldo)
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
        rstLaudo.Close
    End If
    
  Rem  WEmpresa = "0007"
  Rem  txtOdbc = "Empresa07"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Laudo SET "
        ZSql = ZSql & "MarcaAnt = " + "'" + WVector(Ciclo, 2) + "',"
        ZSql = ZSql & "SaldoAnt = " + "'" + WVector(Ciclo, 5) + "',"
        ZSql = ZSql & "LiberadaAnt = " + "'" + WVector(Ciclo, 3) + "',"
        ZSql = ZSql & "DevueltaAnt = " + "'" + WVector(Ciclo, 4) + "'"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo, 1) + "'"
                
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    
    
    
    
    
    Rem Procesa las Hojas
    
    
    
   Rem WEmpresa = "0015"
   Rem txtOdbc = "Empresa15"
   Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
   Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase WVector
    Lugar = 0


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Order by Hoja"
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
                
                Rem If rstHoja!hoja = 302121 Then Stop
                
                XTipoPro = ""
                
                WTerminado = IIf(IsNull(rstHoja!Producto), "", rstHoja!Producto)
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                
                Pasa = "N"
                
                If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                    Pasa = "S"
                End If
                If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                    Pasa = "S"
                End If
                
                If Pasa = "S" Then
                    WRealAnt = IIf(IsNull(rstHoja!RealAnt), "0", rstHoja!RealAnt)
                    WReal = IIf(IsNull(rstHoja!Real), "0", rstHoja!Real)
                    WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WMarca = IIf(IsNull(rstHoja!Marca), "", rstHoja!Marca)
                
                    If WRealAnt = 0 And WReal <> 0 Then
                        Lugar = Lugar + 1
                        WVector(Lugar, 1) = rstHoja!Clave
                        WVector(Lugar, 2) = Str$(WReal)
                        WVector(Lugar, 3) = Str$(WSaldo)
                        WVector(Lugar, 4) = WMarca
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
    
Rem    WEmpresa = "0007"
Rem    txtOdbc = "Empresa07"
Rem    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
Rem    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Hoja SET "
        ZSql = ZSql & "MarcaAnt = " + "'" + WVector(Ciclo, 4) + "',"
        ZSql = ZSql & "SaldoAnt = " + "'" + WVector(Ciclo, 3) + "',"
        ZSql = ZSql & "RealAnt = " + "'" + WVector(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo, 1) + "'"
                
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    
    
    Rem Procesa las guias
    
Rem    WEmpresa = "0015"
Rem    txtOdbc = "Empresa15"
Rem    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
Rem    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Order by Codigo"
    spMovguia = ZSql
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTipo = rstMovguia!Tipo
                WArticulo = rstMovguia!Articulo
                WTerminado = rstMovguia!Terminado
                
                If WTipo = "M" Then
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                End If
                        
                If Pasa = "S" Then
                
                    WCantidadAnt = IIf(IsNull(rstMovguia!CantidadAnt), "0", rstMovguia!CantidadAnt)
                    WCantidad = IIf(IsNull(rstMovguia!Cantidad), "0", rstMovguia!Cantidad)
                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    WMarca = IIf(IsNull(rstMovguia!Marca), "", rstMovguia!Marca)
                
                    If WCantidadAnt = 0 And WCantidad <> 0 Then
                        Lugar = Lugar + 1
                        WVector(Lugar, 1) = rstMovguia!Clave
                        WVector(Lugar, 2) = Str$(WCantidad)
                        WVector(Lugar, 3) = Str$(WSaldo)
                        WVector(Lugar, 4) = WMarca
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
    
   Rem WEmpresa = "0007"
  Rem  txtOdbc = "Empresa07"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Guia SET "
        ZSql = ZSql & "MarcaAnt = " + "'" + WVector(Ciclo, 4) + "',"
        ZSql = ZSql & "SaldoAnt = " + "'" + WVector(Ciclo, 3) + "',"
        ZSql = ZSql & "CantidadAnt = " + "'" + WVector(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo, 1) + "'"
                
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las Movimientos Varios
    
  Rem  WEmpresa = "0015"
  Rem  txtOdbc = "Empresa15"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Movvar"
    ZSql = ZSql + " Order by Codigo"
    spMovvar = ZSql
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTipo = rstMovvar!Tipo
                WArticulo = rstMovvar!Articulo
                WTerminado = rstMovvar!Terminado
                
                If WTipo = "M" Then
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                End If
                        
                If Pasa = "S" Then
                
                    WMarcaAnt = IIf(IsNull(rstMovvar!Marcaant), "", rstMovvar!Marcaant)
                    WMarca = IIf(IsNull(rstMovvar!Marca), "", rstMovvar!Marca)
                
                    If Trim(WMarcaAnt) <> Trim(WMarca) Then
                        Lugar = Lugar + 1
                        WVector(Lugar, 1) = rstMovvar!Clave
                        WVector(Lugar, 2) = WMarca
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
    
  Rem  WEmpresa = "0007"
  Rem  txtOdbc = "Empresa07"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Movvar SET "
        ZSql = ZSql & "MarcaAnt = " + "'" + WVector(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo, 1) + "'"
                
        spMovvar = ZSql
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las Movimientos de Laboratorio
    
  Rem  WEmpresa = "0015"
  Rem  txtOdbc = "Empresa15"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Movlab"
    ZSql = ZSql + " Order by Codigo"
    spMovlab = ZSql
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTipo = rstMovlab!Tipo
                WArticulo = rstMovlab!Articulo
                WTerminado = rstMovlab!Terminado
                
                If WTipo = "M" Then
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                End If
                        
                If Pasa = "S" Then
                
                    WMarcaAnt = IIf(IsNull(rstMovlab!Marcaant), "", rstMovlab!Marcaant)
                    WMarca = IIf(IsNull(rstMovlab!Marca), "", rstMovlab!Marca)
                
                    If Trim(WMarcaAnt) <> Trim(WMarca) Then
                        Lugar = Lugar + 1
                        WVector(Lugar, 1) = rstMovlab!Clave
                        WVector(Lugar, 2) = WMarca
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
    
  Rem  WEmpresa = "0007"
  Rem  txtOdbc = "Empresa07"
  Rem  strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
  Rem  Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Movlab SET "
        ZSql = ZSql & "MarcaAnt = " + "'" + WVector(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo, 1) + "'"
                
        spMovlab = ZSql
        Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las estadisticas
    
   Rem WEmpresa = "0015"
   Rem txtOdbc = "Empresa15"
   Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
   Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Order by Clave"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                
                WTerminado = IIf(IsNull(rstEstadistica!Articulo), "", rstEstadistica!Articulo)
                WArticulo = Left$(WTerminado, 3) + Right$(WTerminado, 7)
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                
                If Left$(WTerminado, 2) = "PT" Or Left$(WTerminado, 2) = "YQ" Or Left$(WTerminado, 2) = "YF" Then
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                
                End If
                        
                If Pasa = "S" Then
                
                    WMarcaAnt = IIf(IsNull(rstEstadistica!Marcaant), "", rstEstadistica!Marcaant)
                    WMarca = IIf(IsNull(rstEstadistica!Marca), "", rstEstadistica!Marca)
                    Rem If WTerminado = "PT-00741-100" Then Stop
                    
                    WMarcaAnt = Trim(WMarcaAnt)
                    WMarca = Trim(WMarca)
                
                    If WMarcaAnt <> WMarca Then
                        Lugar = Lugar + 1
                        WVector(Lugar, 1) = rstEstadistica!Clave
                        WVector(Lugar, 2) = WMarca
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
    
    Rem WEmpresa = "0007"
    Rem txtOdbc = "Empresa07"
    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Estadistica SET "
        ZSql = ZSql & "MarcaAnt = " + "'" + WVector(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo, 1) + "'"
                
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las entradas de devoluciones
    
   Rem WEmpresa = "0015"
   Rem txtOdbc = "Empresa15"
   Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
   Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Erase WVector
    Lugar = 0

    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EntDev"
    ZSql = ZSql + " Order by Codigo"
    spEntdev = ZSql
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XTipoPro = ""
                
                WTerminado = rstEntdev!Terminado
                WArticulo = Left$(rstEntdev!Terminado, 3) + Right$(rstEntdev!Terminado, 7)
                XCodigo = Val(Mid$(WTerminado, 4, 5))
                Rem BY NAN AGREGO NK 23-11
                If Left$(WTerminado, 2) = "NK" Or Left$(WTerminado, 2) = "PT" Or Left$(WTerminado, 2) = "YQ" Or Left$(WTerminado, 2) = "YF" Then
                
                    Pasa = "N"
                    
                    If WTerminado >= WDesde3 And WTerminado <= WHasta3 Then
                        Pasa = "S"
                    End If
                    If WTerminado >= WDesde4 And WTerminado <= WHasta4 Then
                        Pasa = "S"
                    End If
                    
                        Else
                
                    Pasa = "N"
                    
                    If WArticulo >= WDesde1 And WArticulo <= WHasta1 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde2 And WArticulo <= WHasta2 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde5 And WArticulo <= WHasta5 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde6 And WArticulo <= WHasta6 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde7 And WArticulo <= WHasta7 Then
                        Pasa = "S"
                    End If
                    If WArticulo >= WDesde8 And WArticulo <= WHasta8 Then
                        Pasa = "S"
                    End If
                
                End If
                        
                If Pasa = "S" Then
                
                    WMarcaAnt = IIf(IsNull(rstEntdev!Marcaant), "", rstEntdev!Marcaant)
                    WMarca = IIf(IsNull(rstEntdev!Marca), "", rstEntdev!Marca)
                
                    If Trim(WMarcaAnt) <> Trim(WMarca) Then
                        Lugar = Lugar + 1
                        WVector(Lugar, 1) = rstEntdev!Clave
                        WVector(Lugar, 2) = WMarca
                    End If
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstEntdev.Close
    End If
    
   Rem WEmpresa = "0007"
   Rem txtOdbc = "Empresa07"
   Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
   Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql & "UPDATE EntDev SET "
        ZSql = ZSql & "MarcaAnt = " + "'" + WVector(Ciclo, 2) + "'"
        ZSql = ZSql & " Where Clave = " + "'" + WVector(Ciclo, 1) + "'"
                
        spEntdev = ZSql
        Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa los movimientos varios
    Rem
    Rem spMovvar = "ModificaMovvarMarcaAnt"
    Rem Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem Procesa las movimientos varios de labaratorio
    Rem
    Rem spMovlab = "ModificaMovlabMarcaAnt"
    Rem Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem Procesa las estadisticas
    Rem
    Rem spEstadistica = "ModificaEstadisticaMarcaAnt"
    Rem Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem Procesa las devoluciones
    Rem
    Rem spEntdev = "ModificaEntdevMarcaAnt"
    Rem Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            
    Call Cancelar_Click

End Sub


Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgCierreStkAnt.Caption = "Cierre de Stock :  " + !Nombre
        End If
    End With

End Sub
