VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVerilot1 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Saldos de Lotes de Materias Primas"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlotemat.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "verilot1.frx":0000
      Left            =   120
      List            =   "verilot1.frx":0007
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgVerilot1"
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
Dim Vector(10000, 4) As String
Private XLote(100, 7) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double
Private WPartiOri As String

Private Sub Acepta_Click()

    Rem Open "lpt1" For Output As #1
    Open WEmpresa + "MP.TXT" For Output As #1
    
    Erase Vector
    Renglon = 0
    
    Pasa = 0
    Corte = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Laudo.Marca,Laudo.Laudo,Laudo.Saldo,Laudo.Articulo,Laudo.PartiOri,Laudo.Lote"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.Marca <> " + "'" + "X" + "'"
    ZSql = ZSql + " Order by PartiOri,Clave"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WCompara = Trim(rstLaudo!PartiOri)
                    If Left$(rstLaudo!Articulo, 2) <> "DY" And Left$(rstLaudo!Articulo, 2) <> "DS" And Left$(rstLaudo!Articulo, 2) <> "DQ" Then
                        WCompara = ""
                    End If
                    
                    If WCompara = "" Then
                        WCompara = rstLaudo!Laudo
                    End If
                    
                    If Pasa = 0 Then
                        Pasa = 1
                        WCorte = WCompara
                        WLaudo = rstLaudo!Laudo
                        WArticulo = rstLaudo!Articulo
                        If WCorte = "" Then
                            WCorte = WLaudo
                        End If
                        Saldo = 0
                        dada = 0
                    End If
                    
                    If WCorte <> WCompara Or WArticulo <> rstLaudo!Articulo Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = WLaudo
                        Vector(Renglon, 2) = Str$(Saldo)
                        Vector(Renglon, 3) = WArticulo
                        Vector(Renglon, 4) = WCorte
                        WCorte = WCompara
                        WLaudo = rstLaudo!Laudo
                        WArticulo = rstLaudo!Articulo
                        If WCorte = "" Then
                            WCorte = WLaudo
                        End If
                        Saldo = 0
                        dada = 0
                    End If
                    
                    Saldo = Saldo + rstLaudo!Saldo
                    dada = dada + 1
                    
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
        Vector(Renglon, 1) = WLaudo
        Vector(Renglon, 2) = Str$(Saldo)
        Vector(Renglon, 3) = WArticulo
        Vector(Renglon, 4) = WCorte
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where (Guia.Marca <> " + "'" + "X" + "'"
    ZSql = ZSql + " or Guia.Saldo <> 0)"
    ZSql = ZSql + " and Guia.Tipo = " + "'" + "M" + "'"
    ZSql = ZSql + " and Guia.Movi = " + "'" + "E" + "'"
    ZSql = ZSql + " Order by Clave"
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
                    
                    WCantidad = rstMovguia!Cantidad
                    WMovi = rstMovguia!Movi
                    If WMovi = "S" Then
                        Lote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            Else
                        Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                    End If
                    
                    Entra = "S"
                    For dada = 1 To Renglon
                        If Vector(dada, 1) = Lote And Vector(dada, 3) = rstMovguia!Articulo Then
                            Vector(dada, 2) = Str$(Val(Vector(dada, 2)) + rstMovguia!Saldo)
                            Entra = "N"
                            Exit For
                        End If
                    Next dada
                    If Entra = "S" Then
                        If Lote <> "" Then
                            Renglon = Renglon + 1
                            Q = rstMovguia!Codigo
                            Vector(Renglon, 1) = Lote
                            Vector(Renglon, 2) = Str$(rstMovguia!Saldo)
                            Vector(Renglon, 3) = rstMovguia!Articulo
                            Vector(Renglon, 4) = ""
                        End If
                    End If
                    
                    Rem End If
                
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
        WSaldo = Val(Vector(dada, 2))
        WArticulo = Vector(dada, 3)
        WPartiOri = RTrim(Vector(dada, 4))
        XSaldo = 0
        
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
            WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
            rstArticulo.Close
        End If
        
        If (Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DS" Or Left$(WArticulo, 2) = "DQ") And WPartiOri <> "" Then
        
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
            ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArticulo + "'"
            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
            spLaudo = ZSql
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
        
            XParam = "'" + WLote + "','" _
                        + WArticulo + "'"
    
            spLaudo = "ListaLaudoArticulo" + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
        
                Rem WArticulo = rstLaudo!Articulo
                WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                If WLiberada <> 0 Then
                    XSaldo = XSaldo + rstLaudo!Liberada
                End If
                rstLaudo.Close
           
                    Else
            
                XParam = "'" + WLote + "'"
                spMovguia = "ListaMovguiaLoteSolo" + XParam
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    Rem WArticulo = rstMovguia!Articulo
                    rstMovguia.Close
                End If
                
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
                
                            XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                            XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                            XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                            XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                            XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                            XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                    
                            If Val(XLote(1, 1)) = 0 And rstHoja!Lote <> 0 Then
                                XLote(1, 1) = rstHoja!Lote
                                XLote(1, 2) = rstHoja!Cantidad
                            End If
                    
                            For Da = 1 To 3
                                If Val(XLote(Da, 1)) = Val(WLote) Then
                                    XSaldo = XSaldo - XLote(Da, 2)
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
                            If ZLote = 0 Then
                                ZLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            End If

                            If Val(WLote) = Val(ZLote) Then
                                If rstMovguia!Movi = "E" Then
                                    XSaldo = XSaldo + rstMovguia!Cantidad
                                        Else
                                    XSaldo = XSaldo - rstMovguia!Cantidad
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
                        
                        If (rstEstadistica!Tipo = 2 And Left$(WArticulo, 2) = rstEstadistica!Tipopro) Or rstEstadistica!Tipo = 1 Then
                    
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
                                ZLote = XLote(Da, 1)
                                WCantidad = XLote(Da, 2)
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
        
        dd = WLote
        aad = WArticulo
        
        Call Redondeo(XSaldo)
        Call Redondeo(WSaldo)
        
        If XSaldo <> WSaldo Or XSaldo < 0 Then
            Print #1, WLote, WArticulo, WSaldo, XSaldo
        End If
    
    Next dada
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Close
    PrgVerilot1.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaMat
End Sub
