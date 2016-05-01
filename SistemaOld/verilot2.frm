VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVerilot2 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Saldos de Lotes de Productos Terminados"
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
      ItemData        =   "verilot2.frx":0000
      Left            =   120
      List            =   "verilot2.frx":0007
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
Attribute VB_Name = "PrgVerilot2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim XParam As String
Dim Vector(10000, 4) As String
Private XLote(100, 7) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double

Private Sub Acepta_Click()

    Rem Open "lpt1" For Output As #1
    Open WEmpresa + "pt.TXT" For Output As #1

    Erase Vector
    Renglon = 0
    
    spHoja = "ListaHojaTotal"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Rem If rstHoja!Hoja = 45805 Then
                    
                        If rstHoja!Marca <> "X" And rstHoja!Renglon = 1 And Left$(rstHoja!Producto, 2) = "PT" Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = rstHoja!Hoja
                            Vector(Renglon, 2) = rstHoja!Saldo
                            Vector(Renglon, 3) = rstHoja!Producto
                            Vector(Renglon, 4) = rstHoja!Real
                        End If
                        
                        Rem End If
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
    End If
    
    XParam = "'" + "AA-00000-000" + "','" _
                 + "ZZ-99999-999" + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstMovguia!Marca <> "X" And rstMovguia!Saldo <> 0 Then
                    
                    If rstMovguia!Tipo = "T" Then
                    
                        WCantidad = rstMovguia!Cantidad
                        WMovi = rstMovguia!Movi
                        If WMovi = "S" Then
                            Lote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                Else
                            Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        End If
                        If WMovi = "E" Then
                            Entra = "S"
                            For dada = 1 To Renglon
                                If Vector(dada, 1) = Lote Then
                                    Entra = "N"
                                    Exit For
                                End If
                            Next dada
                            If Entra = "S" Then
                                If Lote <> 0 Then
                                    Renglon = Renglon + 1
                                    Vector(Renglon, 1) = Lote
                                    Vector(Renglon, 2) = rstMovguia!Saldo
                                    Vector(Renglon, 3) = rstMovguia!Terminado
                                    Vector(Renglon, 4) = 0
                                End If
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
    
    

    For dada = 1 To Renglon
    
        WLote = Vector(dada, 1)
        WSaldo = Vector(dada, 2)
        WTerminado = Vector(dada, 3)
        XSaldo = Vector(dada, 4)
        
        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
            WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
            rstTerminado.Close
        End If
        
        
        XParam = "'" + WTerminado + "','" _
                     + WTerminado + "'"
        spEstadistica = "ListaEstadisticaDesdeHasta" + XParam
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
    
            With rstEstadistica
    
                .MoveFirst
                
                If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    Xmarca = IIf(IsNull(rstEstadistica!Marca), "", rstEstadistica!Marca)
                    
                    If Xmarca <> "X" Then
                    
                        aa = rstEstadistica!Numero
                    
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
                        
                        If XLote(1, 2) = 0 Then
                            XLote(1, 2) = rstEstadistica!Cantidad
                        End If
                        
                        For Da = 1 To 12
                        
                            ZLote = XLote(Da, 1)
                            WCantidad = XLote(Da, 2)
                
                            If Val(ZLote) = Val(WLote) Then
                                If rstEstadistica!Tipo = 1 Then
                                    XSaldo = XSaldo - WCanti1
                                        Else
                                    XSaldo = XSaldo + WCanti1
                                End If
                            End If
                            
                        Next Da
                    
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
    
        XParam = "'" + WTerminado + "','" _
                     + WTerminado + "'"
        spHoja = "ListaHojaTerminadoDesdeHasta" + XParam
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
                
                                If rstHoja!lote1 = Val(WLote) Then
                                    XSaldo = XSaldo - rstHoja!Canti1
                                End If
                                If rstHoja!lote2 = Val(WLote) Then
                                    XSaldo = XSaldo - rstHoja!Canti2
                                End If
                                If rstHoja!lote3 = Val(WLote) Then
                                    XSaldo = XSaldo - rstHoja!Canti3
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
            rstHoja.Close
        End If
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spMovvar = "ListaMovvarTerminadoDesdeHasta" + XParam
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovvar.RecordCount > 0 Then
    
            With rstMovvar
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If rstMovvar!Marca <> "X" Then
                
                        If rstMovvar!Tipo = "T" Then
                            WCantidad = rstMovvar!Cantidad
                            WMovi = rstMovvar!Movi
                            Lote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                            If Val(Lote) = Val(WLote) Then
                                If WMovi = "E" Then
                                    XSaldo = XSaldo + WCantidad
                                        Else
                                    XSaldo = XSaldo - WCantidad
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
            rstMovvar.Close
        End If
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
    
            With rstMovguia
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstMovguia!Marca <> "X" Then
                
                        If rstMovguia!Tipo = "T" Then
                            WCantidad = rstMovguia!Cantidad
                            WMovi = rstMovguia!Movi
                            If WMovi = "S" Then
                                Lote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                    Else
                                Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            End If
                            If Val(Lote) = Val(WLote) Then
                                If WMovi = "E" Then
                                    XSaldo = XSaldo + WCantidad
                                        Else
                                    XSaldo = XSaldo - WCantidad
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
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spConsig = "ListaConsigTerminado" + XParam
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
                            WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                            Lote = IIf(IsNull(rstConsig!Lote), "0", rstConsig!Lote)
                            If WCantidad <> 0 Then
                                If Val(Lote) = Val(WLote) Then
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
            rstConsig.Close
        End If
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spMovlab = "ListaMovlabTerminadoDesdeHasta" + XParam
        Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovlab.RecordCount > 0 Then
    
            With rstMovlab
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If rstMovlab!Marca <> "X" Then
                
                        If rstMovlab!Tipo = "T" Then
                            WCantidad = rstMovlab!Cantidad
                            WMovi = rstMovlab!Movi
                            Lote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                            If Val(Lote) = Val(WLote) Then
                                If WMovi = "E" Then
                                    XSaldo = XSaldo + WCantidad
                                        Else
                                    XSaldo = XSaldo - WCantidad
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
            rstMovlab.Close
        End If
        
        dd = WLote
        aad = WTerminado
        
        Call Redondeo(XSaldo)
        Call Redondeo(WSaldo)
        
        If XSaldo <> WSaldo Or XSaldo < 0 Then
            Print #1, WLote, WTerminado, WSaldo, XSaldo
        End If
    
    Next dada
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Close
    PrgVerilot2.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub
