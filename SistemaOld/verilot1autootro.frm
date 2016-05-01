VERSION 5.00
Begin VB.Form PrgVerilot1AUTOOtro 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Saldos de Lotes de Materias Primas"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.CommandButton Acepta 
      Caption         =   "Acepta"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Campo3 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Campo2 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Campo1 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PROCESANDO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
End
Attribute VB_Name = "PrgVerilot1AUTOOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
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
Dim Vector(10000, 6) As String
Private xLote(100, 7) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double
Private WPartiOri As String
Dim Empe(10, 10) As String

Private Sub Acepta_Click()

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
    
    For A = WDesdeEmpresa To WHastaEmpresa
        
    WEmpresa = Empe(A, 1)
    txtOdbc = Empe(A, 2)
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Erase Vector
    Renglon = 0
    
    Pasa = 0
    Corte = ""
    
    Sql1 = "Select Laudo.Marca,Laudo.Laudo,Laudo.Saldo,Laudo.Articulo,Laudo.PartiOri,Laudo.Lote"
    Sql2 = " FROM Laudo"
    Sql3 = " Where Laudo.Marca <> " + "'" + "X" + "'"
    Sql4 = " Order by PartiOri,Clave"
    spLaudo = Sql1 + Sql2 + Sql3 + Sql4
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If .EOF = False Then
                 
                    If rstLaudo!Articulo = "DY-306-701" Then
                
                    WCompara = Trim(rstLaudo!PartiOri)
                    If Left$(rstLaudo!Articulo, 2) <> "DY" And Left$(rstLaudo!Articulo, 2) <> "DW" And Left$(rstLaudo!Articulo, 2) <> "DS" Then
                        WCompara = ""
                    End If
                    
                    If WCompara = "" Then
                        WCompara = rstLaudo!Laudo
                    End If
                    
                    If Pasa = 0 Then
                        Pasa = 1
                        WCorte = WCompara
                        Wlaudo = rstLaudo!Laudo
                        WArticulo = rstLaudo!Articulo
                        If WCorte = "" Then
                            WCorte = Wlaudo
                        End If
                        Saldo = 0
                        Dada = 0
                        DadaII = 0
                    End If
                    
                    If WCorte <> WCompara Or WArticulo <> rstLaudo!Articulo Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = Wlaudo
                        Vector(Renglon, 2) = Str$(Saldo)
                        Vector(Renglon, 3) = WArticulo
                        Vector(Renglon, 4) = WCorte
                        Vector(Renglon, 5) = "1"
                        WCorte = WCompara
                        Wlaudo = rstLaudo!Laudo
                        WArticulo = rstLaudo!Articulo
                        If WCorte = "" Then
                            WCorte = Wlaudo
                        End If
                        Saldo = 0
                        Dada = 0
                        DadaII = 0
                    End If
                    
                    Saldo = Saldo + rstLaudo!Saldo
                    Dada = Dada + 1
                    If rstLaudo!Saldo <> 0 Then
                        DadaII = DadaII + 1
                    End If
                    
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstLaudo.Close
    End If
    
    If Pasa <> 0 Then
        Renglon = Renglon + 1
        Vector(Renglon, 1) = Wlaudo
        Vector(Renglon, 2) = Str$(Saldo)
        Vector(Renglon, 3) = WArticulo
        Vector(Renglon, 4) = WCorte
        Vector(Renglon, 5) = "1"
        Vector(Renglon, 6) = Str$(DadaII)
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM Guia"
    Sql3 = " Where (Guia.Marca <> " + "'" + "X" + "'"
    Sql4 = " or Guia.Saldo <> 0)"
    Sql5 = " and Guia.Tipo = " + "'" + "M" + "'"
    Sql6 = " and Guia.Movi = " + "'" + "E" + "'"
    Sql7 = " Order by Clave"
    spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                                
                    If rstMovguia!Articulo = "DY-306-701" Then
                    
                    WCantidad = rstMovguia!Cantidad
                    WMovi = rstMovguia!Movi
                    Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                    PartiOri = Trim(IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri))
                    
                    Entra = "S"
                    For Dada = 1 To Renglon
                        If Left$(rstMovguia!Articulo, 2) <> "DY" And Left$(rstMovguia!Articulo, 2) <> "DW" And Left$(rstMovguia!Articulo, 2) <> "DS" Then
                            If Vector(Dada, 1) = Lote And Vector(Dada, 3) = rstMovguia!Articulo Then
                                Vector(Dada, 2) = Str$(Val(Vector(Dada, 2)) + rstMovguia!Saldo)
                                If rstMovguia!Saldo <> 0 Then
                                    Vector(Renglon, 6) = Str$(Val(Vector(Renglon, 6)) + 1)
                                End If
                                Entra = "N"
                                Exit For
                            End If
                                Else
                            If Vector(Dada, 4) = PartiOri And Vector(Dada, 3) = rstMovguia!Articulo Then
                                Vector(Dada, 2) = Str$(Val(Vector(Dada, 2)) + rstMovguia!Saldo)
                                If rstMovguia!Saldo <> 0 Then
                                    Vector(Renglon, 6) = Str$(Val(Vector(Renglon, 6)) + 1)
                                End If
                                Entra = "N"
                                Exit For
                            End If
                        End If
                    Next Dada
                    
                    If Entra = "S" Then
                        If Lote <> "" Then
                            Renglon = Renglon + 1
                            Q = rstMovguia!codigo
                            Vector(Renglon, 1) = Lote
                            Vector(Renglon, 2) = Str$(rstMovguia!Saldo)
                            Vector(Renglon, 3) = rstMovguia!Articulo
                            Vector(Renglon, 4) = ""
                            Vector(Renglon, 5) = "2"
                            Vector(Renglon, 6) = "1"
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
    
    For Dada = 1 To Renglon
    
        WLote = Vector(Dada, 1)
        WSaldo = Val(Vector(Dada, 2))
        WArticulo = Vector(Dada, 3)
        WPartiOri = RTrim(Vector(Dada, 4))
        WOrigen = RTrim(Vector(Dada, 5))
        WPuntas = Val(Vector(Dada, 6))
        XSaldo = 0
        
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
            WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
            rstArticulo.Close
        End If
        
        ZZArticulo = WArticulo
        If Val(WEmpresa) = 2 Or Val(WEmpresa) = 4 Or Val(WEmpresa) = 8 Then
            ZZArticulo = "ZZ"
        End If
        
        
        If Left$(ZZArticulo, 2) = "DY" Or Left$(ZZArticulo, 2) = "DW" Or Left$(ZZArticulo, 2) = "DS" Then
        
            If WPartiOri <> "" Then
        
                Sql1 = "Select *"
                Sql2 = " FROM Laudo"
                Sql3 = " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
                Sql4 = " and Laudo.Articulo = " + "'" + WArticulo + "'"
                Sql5 = " Order by Clave"
                spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    With rstLaudo
                    .MoveFirst
                        Do
                            If .EOF = False Then
                                WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                                If WLiberada <> 0 Then
                                    XSaldo = XSaldo + rstLaudo!Liberada
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstLaudo.Close
                End If
        
                    Else
                    
                If WOrigen = 1 Then
                    XParam = "'" + WLote + "','" _
                                 + WArticulo + "'"
                    spLaudo = "ListaLaudoArticulo" + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        If WLiberada <> 0 Then
                            XSaldo = XSaldo + rstLaudo!Liberada
                        End If
                        rstLaudo.Close
                    End If
                End If
            
            End If
            
                    Else
                    
            XParam = "'" + WLote + "','" _
                                 + WArticulo + "'"
            spLaudo = "ListaLaudoArticulo" + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                If WLiberada <> 0 Then
                    XSaldo = XSaldo + rstLaudo!Liberada
                End If
                rstLaudo.Close
            End If
        
        End If
        
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
                        
                        XFecff = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                        If XFecff >= WOrdFechaCierre Then
                        Xmarca = IIf(IsNull(rstHoja!Marca), "", rstHoja!Marca)
                        If !Tipo = "M" And Xmarca <> "X" Then
                        
                            sdf = rstHoja!Clave
                
                            xLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                            xLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                            xLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                            xLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                            xLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                            xLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                    
                            If Val(xLote(1, 1)) = 0 And rstHoja!Lote <> 0 Then
                                xLote(1, 1) = rstHoja!Lote
                                xLote(1, 2) = rstHoja!Cantidad
                            End If
                    
                            For Da = 1 To 3
                                If Val(xLote(Da, 1)) = Val(WLote) Then
                                    XSaldo = XSaldo - xLote(Da, 2)
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
            rstHoja.Close
        End If
    
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
                
                        If !Tipo = "M" And !Marca <> "X" Then
                            ZLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                            If Val(WLote) = Val(ZLote) Then
                                If rstMovvar!Movi = "E" Then
                                    XSaldo = XSaldo + rstMovvar!Cantidad
                                        Else
                                    XSaldo = XSaldo - rstMovvar!Cantidad
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
                        
                        Da = rstMovguia!Clave
                        WMarca = IIf(IsNull(rstMovguia!Marca), "", rstMovguia!Marca)
                
                        If rstMovguia!Tipo = "M" And WMarca <> "X" Then
                        
                            ZLote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            ZPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            If ZLote = 0 Then
                                ZLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            End If
                            
                            If (Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Or Left$(WArticulo, 2) = "DS") And Trim(WPartiOri) <> "" Then
                    
                                If Trim(ZPartiOri) = Trim(WPartiOri) Then
                                    If rstMovguia!Movi = "E" Then
                                        XSaldo = XSaldo + rstMovguia!Cantidad
                                            Else
                                        XSaldo = XSaldo - rstMovguia!Cantidad
                                    End If
                                End If
                                
                                        Else
                                        
                                If Val(WLote) = Val(ZLote) Then
                                    If rstMovguia!Movi = "E" Then
                                        XSaldo = XSaldo + rstMovguia!Cantidad
                                            Else
                                        XSaldo = XSaldo - rstMovguia!Cantidad
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
                
                        If !Tipo = "M" And !Marca <> "X" Then
                
                            WCantidad = rstMovlab!Cantidad
                            WMovi = rstMovlab!Movi
                            ZLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                    
                            If Val(WLote) = Val(ZLote) Then
                        
                                If WMovi = "E" Then
                                    XSaldo = XSaldo + WCantidad
                                        Else
                                    XSaldo = XSaldo - WCantidad
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
                                ZLote = xLote(Da, 1)
                                WCantidad = xLote(Da, 2)
                                If Val(WLote) = Val(ZLote) Then
                                    If Val(WCantidad) <> 0 Then
                                        If rstEstadistica!Tipo = 2 Then
                                            XSaldo = XSaldo + Abs(Val(WCantidad))
                                                Else
                                            XSaldo = XSaldo - WCantidad
                                        End If
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
        
        
        If WPuntas > 1 And WPatrtiOri <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Laudo SET "
            ZSql = ZSql + " Saldo = 0"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo + "'"
            ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WPartiOri + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        dd = WLote
        aad = WArticulo
        asdd = WPartiOri
        
        Call Redondeo(XSaldo)
        Call Redondeo(WSaldo)
        
        If XSaldo <> WSaldo Then
            
            If WOrigen = 1 Then
            
                XParam = "'" + WLote + "','" _
                             + WArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WClave = rstLaudo!Clave
                    ZSaldo = Str$(XSaldo)
                    WDate = Date$
                    rstLaudo.Close
                            
                    XParam = "'" + WClave + "','" _
                             + WDate + "','" _
                             + ZSaldo + "'"
                    spLaudo = "ModificaLaudoSaldo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                        Else
                                
                XParam = "'" + WArticulo + "','" _
                             + WLote + "'"
                spMovguia = "ListaMovguiaLote " + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WClave = rstMovguia!Clave
                    ZSaldo = Str$(XSaldo)
                    WDate = Date$
                    rstMovguia.Close
                        
                    XParam = "'" + WClave + "','" _
                                 + WDate + "','" _
                                 + ZSaldo + "'"
                    spMovguia = "ModificaMovguiaSaldo " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                End If
                            
            End If
            
        End If
        
        
        
    
    Next Dada
    
    Next A
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgVerilot1AUTO.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub
